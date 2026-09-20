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

    if args.db:
        print(f"\nLoading results into {args.db} ...")
        load_into_db(args.db, args.year, results)

    print("\nDone.")


if __name__ == "__main__":
    main()
