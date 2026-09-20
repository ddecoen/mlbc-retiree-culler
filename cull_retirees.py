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

    df = pd.read_csv(in_path)
    if id_col not in df.columns:
        sys.exit(f"ERROR: {in_path.name} has no '{id_col}' column — check the export format")

    filtered = df[df[id_col].isin(retired_ids)].copy()

    matched_ids = set(filtered[id_col].unique())
    print(f"{label}: {len(filtered)} rows for {len(matched_ids)} retired IDs "
          f"(of {len(retired_ids)} total retirees) from {in_path.name}")

    out_path = output_dir / out_template.format(year=year)
    filtered.to_csv(out_path, index=False)
    print(f"  -> wrote {out_path}")

    filtered.attrs["matched_ids"] = matched_ids
    return filtered


def append_to_cumulative(filtered: pd.DataFrame, target_path: Path, label: str):
    """Append this season's retiree rows onto an existing cumulative file
    (e.g. CareerBatStat_retired.csv), skipping IDs already present.
    Auto-detects whether the existing file has a header row by sniffing
    whether its first cell parses as a number, and writes back in the
    same convention so it stays a drop-in replacement for the PHP import.
    """
    id_col = filtered.columns[0]  # "ID" — first column in the raw export

    if not target_path.exists():
        filtered.to_csv(target_path, index=False, header=False)
        print(f"  Append ({label}): {target_path.name} didn't exist — created it with "
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

    existing_ids = set(existing[id_col].astype(str))
    new_rows = filtered[~filtered[id_col].astype(str).isin(existing_ids)]

    if new_rows.empty:
        print(f"  Append ({label}): no new retirees to add to {target_path.name} "
              f"(all {len(filtered)} already present)")
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

    if args.db:
        print(f"\nLoading results into {args.db} ...")
        load_into_db(args.db, args.year, results)

    print("\nDone.")


if __name__ == "__main__":
    main()
