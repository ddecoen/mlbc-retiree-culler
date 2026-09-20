# MLBC Retiree Culler

Extracts retired players' career and league-leader stat lines from a
season's raw exports. Replaces the old `_RetireesCalc.xlsm` Excel/VBA
workbook with a single Python script.

## What it does

Each season, once a `retirements_{year}.csv` file exists (a list of
players marked retired that season), this script filters the four raw
stat exports down to just those players and writes the results out —
either as dated CSVs, into a SQLite database that accumulates history
across every season, appended directly onto the site's cumulative
`CareerBatStat_retired.csv` / `CareerPitStat_retired.csv` files, or
any combination of the three.

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

Also append this season's retirees directly onto the site's cumulative
career files — the ones that feed the `careerleaders_batting.php` /
pitching PHP import:

```bash
python cull_retirees.py --year 2081 --input-dir . \
    --append-batting /path/to/CareerBatStat_retired.csv \
    --append-pitching /path/to/CareerPitStat_retired.csv
```

This is the piece that plugs into the site's existing update process:
it reads whichever of `CareerBatStat_retired.csv` /
`CareerPitStat_retired.csv` already exists, figures out on its own
whether that file has a header row or not (matches its own convention
either way), skips any retiree already present so re-running is safe,
and appends only the new ones. If the target file doesn't exist yet,
it's created fresh. Column count is checked before writing — if it
doesn't match the source export, nothing is appended and you get a
warning instead of a corrupted file.

The league-leader files get the same treatment, but need remapping
first — `retire_batting.csv`/`retire_pitching.csv` use a narrower,
reordered column set than the raw `leagueleaders_batting`/`pitching`
exports (all the `_rank`/rookie/league-rank columns are dropped):

```bash
python cull_retirees.py --year 2081 --input-dir . \
    --append-retire-batting /path/to/retire_batting.csv \
    --append-retire-pitching /path/to/retire_pitching.csv
```

Since these are season-by-season rows (one row per player per season,
not one row per player), duplicates are checked on `(id, year)` rather
than `id` alone, so appending a new season never collides with a
player's earlier season rows already in the file.

**Known gap:** `retire_pitching.csv` expects a `gs_lead` column that
doesn't exist anywhere in the `leagueleaders_pitching` export (it has
`g_lead` but not `gs_lead`). Travis (league admin) confirmed he's been
filling that column with `0` by hand, so the script does the same —
it's not computed from real data, just matched to existing convention.

### Options

| Flag | Required | Description |
|---|---|---|
| `--year` | yes | Season year, e.g. `2081` |
| `--input-dir` | no (default: current dir) | Folder containing the raw exports |
| `--output-dir` | no (default: same as `--input-dir`) | Where to write `_RET_*.csv` files |
| `--db` | no | Path to a SQLite DB to load results into |
| `--append-batting` | no | Path to the cumulative `CareerBatStat_retired.csv` to append this season's retirees onto |
| `--append-pitching` | no | Path to the cumulative `CareerPitStat_retired.csv` to append this season's retirees onto |
| `--append-retire-batting` | no | Path to `retire_batting.csv` (season-level, narrower columns) to append this season's retirees onto |
| `--append-retire-pitching` | no | Path to `retire_pitching.csv` (season-level, narrower columns) to append this season's retirees onto |

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
- Can append straight onto the site's live cumulative career files,
  instead of that merge being a manual step every season

## Troubleshooting

- **"retirements file not found"** — check the year in `--year` matches
  the year in the filenames in `--input-dir`.
- **"missing expected column(s)"** — the retirements export changed
  format; check the header row against the table above.
- **Warning about unmatched retired IDs** — usually means a player
  retired without ever appearing in a stat file (e.g. never played),
  or the ID in the retirements list doesn't match the ID space used in
  the other exports. Worth a manual check rather than ignoring.
- **"has N columns, but the source export has M" (append mode)** — the
  cumulative file's format doesn't match the raw export shape anymore
  (someone added/reordered columns by hand). Nothing gets appended in
  this case; fix the column mismatch before re-running.
