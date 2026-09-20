# MLBC Retiree Culler

Extracts retired players' career and league-leader stat lines from a
season's raw exports. Replaces the old `_RetireesCalc.xlsm` Excel/VBA
workbook with a single Python script.

## What it does

Each season, once a `retirements_{year}.csv` file exists (a list of
players marked retired that season), this script filters the four raw
stat exports down to just those players and writes the results out —
either as CSVs, into a SQLite database that accumulates history across
every season, or both.

**This tool does not decide who retires.** That list still has to come
from wherever it currently comes from (the sim engine export, or a
manual process). The script only extracts and archives the stat lines
for players already on that list.

## Requirements

- Python 3.9+
- `pandas` (`pip install pandas`)

## Input files

Put these in one folder (same naming convention as the old workbook):

| File | Columns expected |
|---|---|
| `retirements_{year}.csv` | `fn, ln, team, oldid` |
| `CareerBatStat{year}.csv` | `ID, ...` |
| `CareerPitStat{year}.csv` | `ID, ...` |
| `leagueleaders_batting_{year}.csv` | `id, year, ...` |
| `leagueleaders_pitching_{year}.csv` | `id, year, ...` |

Any file that's missing is skipped with a warning — the script won't
silently produce an incomplete result without telling you.

## Usage

Basic run — reproduces the old workbook's four `_RET_*.csv` outputs:

```bash
python cull_retirees.py --year 2081 --input-dir /path/to/season/files
```

Also load the results into a SQLite analytics DB (creates
`retirees_career_batting`, `retirees_career_pitching`,
`retirees_league_leaders_batting`, `retirees_league_leaders_pitching`
tables tagged by `season_year`, accumulating across seasons):

```bash
python cull_retirees.py --year 2081 --input-dir . --db mlbc_players.db
```

Re-running the same `--year` replaces just that year's rows in the DB
(no duplicates) and overwrites just that year's CSVs — it won't touch
other seasons' data.

### Options

| Flag | Required | Description |
|---|---|---|
| `--year` | yes | Season year, e.g. `2081` |
| `--input-dir` | no (default: current dir) | Folder containing the raw exports |
| `--output-dir` | no (default: same as `--input-dir`) | Where to write `_RET_*.csv` files |
| `--db` | no | Path to a SQLite DB to load results into |

## What changed vs. the old workbook

- Year is a command-line argument, not hardcoded text baked into 8
  places in the macro
- No row-count ceiling — the old macro silently dropped anything past
  row 300 in the retirements file, or row 1285/6000 in the stat files
- Never silently clobbers history — CSVs are per-year, and DB loads
  are keyed by `(season_year, id)`
- Flags duplicate IDs in the retirements list and any retired ID that
  doesn't match a row in *any* stat file, so a bad export gets caught
  instead of producing a quietly incomplete result
- One script instead of four copy-pasted macro subs

## Troubleshooting

- **"retirements file not found"** — check the year in `--year` matches
  the year in the filenames in `--input-dir`.
- **"missing expected column(s)"** — the retirements export changed
  format; check the header row against the table above.
- **Warning about unmatched retired IDs** — usually means a player
  retired without ever appearing in a stat file (e.g. never played),
  or the ID in the retirements list doesn't match the ID space used in
  the other exports. Worth a manual check rather than ignoring.
