#!/usr/bin/env python3
"""
cull_retirees.py — MLBC retiree stat extractor

Replaces the "_RetireesCalc.xlsm" macro workbook. Given a season's
retirements file, filters the four raw stat exports down to just the
retired players and writes the results out — either as CSVs, into a
SQLite database (with full history across seasons), or both.

Fixes vs. the old workbook:
  - Year is a parameter, not hardcoded text baked into 8 places
  - No row-count ceiling (old macro silently dropped anything past
    row 300 in retirements, or row 1285/6000 in the stat files)
  - Never overwrites: CSV outputs are written per-year; DB rows are
    keyed by (season_year, id) so re-running a year replaces just
    that year's slice instead of clobbering everything
  - Reports any retiree IDs that don't match any stat file at all,
    and any duplicate IDs in the retirements list, so a bad export
    is caught instead of silently producing an incomplete result
  - One consolidated run instead of 4 near-identical copy/paste subs

Usage:
    python cull_retirees.py --year 2081 --input-dir /path/to/season/files

    # also load results into a SQLite analytics DB (accumulates history)
    python cull_retirees.py --year 2081 --input-dir . --db mlbc_players.db

    # also append this season's retirees onto the site's cumulative files
    # (CareerBatStat_retired.csv / CareerPitStat_retired.csv) so they're
    # ready to drop straight into the PHP import — skips anyone already
    # present, and matches whatever header convention the file already uses
    python cull_retirees.py --year 2081 --input-dir . \
        --append-batting /path/to/CareerBatStat_retired.csv \
        --append-pitching /path/to/CareerPitStat_retired.csv

    # also append this season's retirees' season-by-season lines onto
    # retire_batting.csv / retire_pitching.csv, remapped to that file's
    # narrower column set (drops the _rank/rookie/league columns and
    # reorders the rest to match)
    python cull_retirees.py --year 2081 --input-dir . \
        --append-retire-batting /path/to/retire_batting.csv \
        --append-retire-pitching /path/to/retire_pitching.csv

Expected input filenames in --input-dir (same convention as before):
    retirements_{year}.csv           (fn, ln, team, oldid)
    CareerBatStat{year}.csv          (ID, ...)
    CareerPitStat{year}.csv          (ID, ...)
    leagueleaders_batting_{year}.csv (id, year, ...)
    leagueleaders_pitching_{year}.csv(id, year, ...)
"""

import argparse
import sqlite3
import sys
from pathlib import Path

import pandas as pd

# (label, input filename template, id column name, output filename template)
STAT_FILES = [
    ("career_batting", "CareerBatStat{year}.csv", "ID", "_RET_CareerBatStat{year}.csv"),
    ("career_pitching", "CareerPitStat{year}.csv", "ID", "_RET_CareerPitStat{year}.csv"),
    ("league_leaders_batting", "leagueleaders_batting_{year}.csv", "id",
     "_RET_leagueleaders_batting{year}.csv"),
    ("league_leaders_pitching", "leagueleaders_pitching_{year}.csv", "id",
     "_RET_leagueleaders_pitching{year}.csv"),
]

# Column order for the site's cumulative retire_batting.csv / retire_pitching.csv,
# confirmed against real samples of those files. These are a subset of the raw
# leagueleaders_batting/pitching columns (all the _rank / rookie / league-rank /
# rook,league,name,fn,ln,num,lvl,team,eligible,pfa columns are dropped) in a
# different order. Any column here not found in the source export is left blank
# in-output. Currently that's "gs_lead" for pitching, which doesn't exist
# in the leagueleaders_pitching export — Travis (league admin) confirmed
# he's been filling that column with 0 by hand, so we match that.
RETIRE_BATTING_COLUMNS = [
    "id", "year", "level", "g", "ab", "abbb", "avg", "obp", "slg", "ops", "fpoints",
    "h", "b2", "b3", "hr", "bb", "so", "rbi", "sb", "cs", "r", "u1", "hbp", "gidp",
    "ibb", "sf", "sh", "pab", "phit",
    "avg_lead", "obp_lead", "slg_lead", "ab_lead", "h_lead", "b2_lead", "b3_lead",
    "hr_lead", "r_lead", "rbi_lead", "bb_lead", "so_lead", "sb_lead", "cs_lead",
]
RETIRE_PITCHING_COLUMNS = [
    "id", "year", "level", "outs", "ip", "whip", "era", "pops", "ptb", "fpoints",
    "h", "b2", "b3", "hr", "bb", "k", "er", "r", "sb", "cs", "w", "l", "sv", "bs",
    "cg", "sho", "gs", "u1", "g", "u2", "bk", "hb", "wp", "u3", "hld", "qs", "pitch_num",
    "era_lead", "w_lead", "l_lead", "sv_lead", "bs_lead", "g_lead", "gs_lead",
    "cg_lead", "sho_lead", "ip_lead", "k_lead", "pops_lead", "whip_lead",
]


def load_retirements(input_dir: Path, year: int) -> pd.DataFrame:
    path = input_dir / f"retirements_{year}.csv"
    if not path.exists():
        sys.exit(f"ERROR: retirements file not found: {path}")

    ret = pd.read_csv(path)
    required = {"fn", "ln", "team", "oldid"}
    missing = required - set(ret.columns)
    if missing:
        sys.exit(f"ERROR: {path.name} is missing expected column(s): {sorted(missing)}")

    dupes = ret[ret.duplicated("oldid", keep=False)]
    if not dupes.empty:
        print(f"WARNING: {len(dupes)} duplicate oldid rows in {path.name}:")
        print(dupes.to_string(index=False))

    print(f"Loaded {len(ret)} retirement rows ({ret['oldid'].nunique()} unique IDs) from {path.name}")
    return ret


def process_stat_file(
    input_dir: Path,
    output_dir: Path,
    year: int,
    label: str,
    in_template: str,
    id_col: str,
    out_template: str,
    retired_ids: set,
) -> pd.DataFrame | None:
    in_path = input_dir / in_template.format(year=year)
    if not in_path.exists():
        print(f"SKIP: {label} — file not found: {in_path.name}")
        return None

    try:
        df = pd.read_csv(in_path)
    except (UnicodeDecodeError, pd.errors.ParserError) as e:
        print(f"SKIP: {label} — {in_path.name} could not be read as a CSV "
              f"({type(e).__name__}: it may be corrupted or in the wrong format). "
              f"Get a fresh copy of this file and re-run for this year.")
        return None
    if id_col not in df.columns:
        # Some seasons' exports capitalize the ID column differently
        # (e.g. "ID" instead of "id") — match case-insensitively before
        # giving up, since this is a real inconsistency we've seen in
        # practice, not a sign of a genuinely different file format.
        case_match = next((c for c in df.columns if c.lower() == id_col.lower()), None)
        if case_match:
            df = df.rename(columns={case_match: id_col})
        else:
            print(f"SKIP: {label} — {in_path.name} has no '{id_col}' column — "
                  f"check the export format. Other files for this year were still processed.")
            return None

    filtered = df[df[id_col].isin(retired_ids)].copy()

    matched_ids = set(filtered[id_col].unique())
    print(f"{label}: {len(filtered)} rows for {len(matched_ids)} retired IDs "
          f"(of {len(retired_ids)} total retirees) from {in_path.name}")

    out_path = output_dir / out_template.format(year=year)
    filtered.to_csv(out_path, index=False)
    print(f"  -> wrote {out_path}")

    filtered.attrs["matched_ids"] = matched_ids
    return filtered


def remap_columns(df: pd.DataFrame, dest_columns: list, label: str) -> pd.DataFrame:
    """Reduce/reorder a raw leagueleaders_* export down to the narrower
    column set the site's retire_batting/retire_pitching files use.
    Matches column names case-insensitively (some seasons' exports use
    "Year"/"ID" instead of "year"/"id") before concluding a column is
    truly absent. Any dest column genuinely not present in the source
    is filled with 0, matching the site's existing convention for
    columns it doesn't track upstream."""
    col_lookup = {c.lower(): c for c in df.columns}
    missing = [c for c in dest_columns if c.lower() not in col_lookup]
    if missing:
        print(f"  NOTE ({label}): source export has no column(s) {missing} — "
              f"filled with 0 in the output (matches site convention).")
    out = pd.DataFrame(index=df.index)
    for col in dest_columns:
        src_col = col_lookup.get(col.lower())
        out[col] = df[src_col] if src_col is not None else 0
    return out


def append_to_cumulative(filtered: pd.DataFrame, target_path: Path, label: str, key_cols: list = None):
    """Append this season's rows onto an existing cumulative file
    (e.g. CareerBatStat_retired.csv or retire_batting.csv), skipping rows
    already present (matched on key_cols — defaults to just the first
    column, e.g. player ID; pass ["id", "year"] for season-level files
    where the same player has one row per season).
    Auto-detects whether the existing file has a header row by sniffing
    whether its first cell parses as a number, and writes back in the
    same convention so it stays a drop-in replacement for the PHP import.
    """
    key_cols = key_cols or [filtered.columns[0]]

    if not target_path.exists() or target_path.stat().st_size == 0:
        filtered.to_csv(target_path, index=False, header=False)
        print(f"  Append ({label}): {target_path.name} didn't exist (or was empty) — created it with "
              f"{len(filtered)} row(s), no header (matches raw export convention)")
        return

    with open(target_path) as f:
        first_line = f.readline().strip()
    first_cell = first_line.split(",")[0].strip('"')
    has_header = not first_cell.lstrip("-").isdigit()

    existing = pd.read_csv(target_path, header=0 if has_header else None)
    if len(existing.columns) != len(filtered.columns):
        print(f"  WARNING ({label}): {target_path.name} has {len(existing.columns)} columns, "
              f"but the source export has {len(filtered.columns)}. Skipping append — "
              f"check the file format before running again.")
        return
    existing.columns = filtered.columns

    existing_keys = set(existing[key_cols].astype(str).agg("|".join, axis=1))
    filtered_keys = filtered[key_cols].astype(str).agg("|".join, axis=1)
    new_rows = filtered[~filtered_keys.isin(existing_keys)]

    if new_rows.empty:
        print(f"  Append ({label}): no new rows to add to {target_path.name} "
              f"(all {len(filtered)} already present, matched on {key_cols})")
        return

    combined = pd.concat([existing, new_rows], ignore_index=True)
    combined.to_csv(target_path, index=False, header=has_header)
    print(f"  Append ({label}): added {len(new_rows)} new row(s) to {target_path.name} "
          f"(header={'yes' if has_header else 'no'}, now {len(combined)} total rows)")


def load_into_db(db_path: Path, year: int, results: dict):
    conn = sqlite3.connect(db_path)
    for label, df in results.items():
        if df is None or df.empty:
            continue
        table = f"retirees_{label}"
        df = df.copy()
        df.insert(0, "season_year", year)

        # Replace just this year's slice instead of duplicating on re-run
        exists = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
        ).fetchone()
        if exists:
            conn.execute(f"DELETE FROM {table} WHERE season_year = ?", (year,))
        df.to_sql(table, conn, if_exists="append", index=False)
        conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_year ON {table}(season_year)")
        print(f"  DB: {table} now has {conn.execute(f'SELECT COUNT(*) FROM {table}').fetchone()[0]} "
              f"total rows across all seasons loaded")
    conn.commit()
    conn.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--year", type=int, required=True, help="Season year, e.g. 2081")
    parser.add_argument("--input-dir", type=Path, default=Path("."), help="Directory with the raw season export CSVs")
    parser.add_argument("--output-dir", type=Path, default=None,
                         help="Where to write _RET_ CSVs (default: same as --input-dir)")
    parser.add_argument("--db", type=Path, default=None,
                         help="Optional path to a SQLite DB to load results into (accumulates across seasons)")
    parser.add_argument("--append-batting", type=Path, default=None,
                         help="Path to the site's cumulative CareerBatStat_retired.csv to append this season's retirees onto")
    parser.add_argument("--append-pitching", type=Path, default=None,
                         help="Path to the site's cumulative CareerPitStat_retired.csv to append this season's retirees onto")
    parser.add_argument("--append-retire-batting", type=Path, default=None,
                         help="Path to the site's retire_batting.csv (season-level, narrower column set) to append this season's retirees onto")
    parser.add_argument("--append-retire-pitching", type=Path, default=None,
                         help="Path to the site's retire_pitching.csv (season-level, narrower column set) to append this season's retirees onto")
    args = parser.parse_args()

    input_dir = args.input_dir
    output_dir = args.output_dir or input_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    ret = load_retirements(input_dir, args.year)
    retired_ids = set(ret["oldid"])

    results = {}
    all_matched_ids = set()
    for label, in_template, id_col, out_template in STAT_FILES:
        df = process_stat_file(input_dir, output_dir, args.year, label, in_template, id_col, out_template, retired_ids)
        results[label] = df
        if df is not None:
            all_matched_ids |= df.attrs.get("matched_ids", set())

    unmatched = retired_ids - all_matched_ids
    if unmatched:
        unmatched_rows = ret[ret["oldid"].isin(unmatched)]
        print(f"\nWARNING: {len(unmatched)} retired ID(s) had no rows in ANY stat file "
              f"(likely never played a game, or an ID mismatch — worth a manual check):")
        print(unmatched_rows.to_string(index=False))

    if args.append_batting or args.append_pitching:
        print()
    if args.append_batting:
        if results.get("career_batting") is not None:
            append_to_cumulative(results["career_batting"], args.append_batting, "career_batting")
        else:
            print(f"  Append (career_batting): skipped — CareerBatStat{args.year}.csv was not found")
    if args.append_pitching:
        if results.get("career_pitching") is not None:
            append_to_cumulative(results["career_pitching"], args.append_pitching, "career_pitching")
        else:
            print(f"  Append (career_pitching): skipped — CareerPitStat{args.year}.csv was not found")

    if args.append_retire_batting:
        if results.get("league_leaders_batting") is not None:
            remapped = remap_columns(results["league_leaders_batting"], RETIRE_BATTING_COLUMNS, "retire_batting")
            append_to_cumulative(remapped, args.append_retire_batting, "retire_batting", key_cols=["id", "year"])
        else:
            print(f"  Append (retire_batting): skipped — leagueleaders_batting_{args.year}.csv was not found")
    if args.append_retire_pitching:
        if results.get("league_leaders_pitching") is not None:
            remapped = remap_columns(results["league_leaders_pitching"], RETIRE_PITCHING_COLUMNS, "retire_pitching")
            append_to_cumulative(remapped, args.append_retire_pitching, "retire_pitching", key_cols=["id", "year"])
        else:
            print(f"  Append (retire_pitching): skipped — leagueleaders_pitching_{args.year}.csv was not found")

    if args.db:
        print(f"\nLoading results into {args.db} ...")
        load_into_db(args.db, args.year, results)

    print("\nDone.")


if __name__ == "__main__":
    main()
