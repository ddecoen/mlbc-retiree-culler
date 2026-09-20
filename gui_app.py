#!/usr/bin/env python3
"""
MLBC Retiree Updater — point-and-click GUI wrapper around cull_retirees.py

Built for the league admin (no command line, no flags to remember).
Reuses the exact same, already-tested logic from cull_retirees.py —
this file only adds a window around it.

Workflow:
  1. Browse to the folder with this season's raw export files
     (retirements_{year}.csv, CareerBatStat{year}.csv, etc.) — the
     year is detected automatically from the filenames in that folder.
  2. Browse to the four site files that need updating:
       CareerBatStat_retired.csv, CareerPitStat_retired.csv,
       retire_batting.csv, retire_pitching.csv
  3. Click "Run Update".

The four site file paths are remembered between runs (saved next to
the program), so next season only step 1 needs to change.
"""

import json
import re
import sys
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import cull_retirees as core

CONFIG_PATH = Path(sys.argv[0]).resolve().parent / "mlbc_updater_config.json"

FIELDS = [
    ("career_batting_retired", "CareerBatStat_retired.csv"),
    ("career_pitching_retired", "CareerPitStat_retired.csv"),
    ("retire_batting", "retire_batting.csv"),
    ("retire_pitching", "retire_pitching.csv"),
]


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text())
        except Exception:
            return {}
    return {}


def save_config(cfg: dict):
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2))
    except Exception:
        pass  # non-fatal — just means paths won't be remembered next time


class TextRedirector:
    """Sends print() output into the log widget instead of a (possibly
    nonexistent, in a --noconsole build) console."""
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.configure(state="normal")
        self.widget.insert("end", text)
        self.widget.see("end")
        self.widget.configure(state="disabled")
        self.widget.update_idletasks()

    def flush(self):
        pass


def detect_year(folder: Path):
    """Look for retirements_{year}.csv in the folder and return the year(s) found."""
    years = []
    for f in folder.glob("retirements_*.csv"):
        m = re.match(r"retirements_(\d{4})\.csv$", f.name)
        if m:
            years.append(int(m.group(1)))
    return sorted(years)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MLBC Retiree Updater")
        self.geometry("720x560")
        self.resizable(True, True)

        self.cfg = load_config()
        self.vars = {}

        pad = {"padx": 8, "pady": 6}

        # --- Season folder ---
        frame_season = ttk.LabelFrame(self, text="Step 1 — This season's export files")
        frame_season.pack(fill="x", **pad)

        self.vars["season_folder"] = tk.StringVar(value=self.cfg.get("season_folder", ""))
        self.year_label_var = tk.StringVar(value="Year: (pick a folder)")

        row = ttk.Frame(frame_season)
        row.pack(fill="x", padx=8, pady=6)
        ttk.Entry(row, textvariable=self.vars["season_folder"]).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Browse...", command=self.browse_season_folder).pack(side="left", padx=(6, 0))

        ttk.Label(frame_season, textvariable=self.year_label_var).pack(anchor="w", padx=8, pady=(0, 6))

        # --- Site files ---
        frame_site = ttk.LabelFrame(self, text="Step 2 — Your site's cumulative files (remembered after first run)")
        frame_site.pack(fill="x", **pad)

        for key, default_name in FIELDS:
            self.vars[key] = tk.StringVar(value=self.cfg.get(key, ""))
            row = ttk.Frame(frame_site)
            row.pack(fill="x", padx=8, pady=4)
            ttk.Label(row, text=default_name, width=26).pack(side="left")
            ttk.Entry(row, textvariable=self.vars[key]).pack(side="left", fill="x", expand=True)
            ttk.Button(row, text="Browse...", command=lambda k=key: self.browse_file(k)).pack(side="left", padx=(6, 0))

        # --- Run button ---
        ttk.Button(self, text="Run Update", command=self.run_update).pack(pady=10)

        # --- Log ---
        frame_log = ttk.LabelFrame(self, text="Log")
        frame_log.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(frame_log, state="disabled", height=15)
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

        if self.vars["season_folder"].get():
            self.refresh_year_label()

    def browse_season_folder(self):
        path = filedialog.askdirectory(title="Select the folder with this season's export files")
        if path:
            self.vars["season_folder"].set(path)
            self.refresh_year_label()

    def browse_file(self, key):
        current = self.vars[key].get()
        initial = str(Path(current).parent) if current else None
        path = filedialog.askopenfilename(
            title=f"Select {dict(FIELDS)[key]}",
            initialdir=initial,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.vars[key].set(path)

    def refresh_year_label(self):
        folder = Path(self.vars["season_folder"].get())
        if not folder.exists():
            self.year_label_var.set("Year: folder not found")
            return
        years = detect_year(folder)
        if not years:
            self.year_label_var.set("Year: no retirements_XXXX.csv found in that folder")
        elif len(years) == 1:
            self.year_label_var.set(f"Year: {years[0]} (detected automatically)")
        else:
            self.year_label_var.set(f"Year: multiple found {years} — put only one season's files in this folder")

    def run_update(self):
        folder = Path(self.vars["season_folder"].get())
        if not folder.exists():
            messagebox.showerror("Missing folder", "Pick a valid season files folder first.")
            return
        years = detect_year(folder)
        if len(years) != 1:
            messagebox.showerror(
                "Can't determine season year",
                "That folder needs exactly one retirements_XXXX.csv file so the year can be detected.",
            )
            return
        year = years[0]

        missing_targets = [dict(FIELDS)[k] for k, _ in FIELDS if not self.vars[k].get()]
        if missing_targets:
            messagebox.showerror("Missing file(s)", "Please select a location for:\n" + "\n".join(missing_targets))
            return

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.log)
        try:
            print(f"=== Running update for season {year} ===\n")
            ret = core.load_retirements(folder, year)
            retired_ids = set(ret["oldid"])

            results = {}
            all_matched_ids = set()
            for label, in_template, id_col, out_template in core.STAT_FILES:
                df = core.process_stat_file(folder, folder, year, label, in_template, id_col, out_template, retired_ids)
                results[label] = df
                if df is not None:
                    all_matched_ids |= df.attrs.get("matched_ids", set())

            unmatched = retired_ids - all_matched_ids
            if unmatched:
                unmatched_rows = ret[ret["oldid"].isin(unmatched)]
                print(f"\nWARNING: {len(unmatched)} retired ID(s) had no rows in ANY stat file:")
                print(unmatched_rows.to_string(index=False))

            print()
            if results.get("career_batting") is not None:
                core.append_to_cumulative(results["career_batting"], Path(self.vars["career_batting_retired"].get()), "career_batting")
            if results.get("career_pitching") is not None:
                core.append_to_cumulative(results["career_pitching"], Path(self.vars["career_pitching_retired"].get()), "career_pitching")
            if results.get("league_leaders_batting") is not None:
                remapped = core.remap_columns(results["league_leaders_batting"], core.RETIRE_BATTING_COLUMNS, "retire_batting")
                core.append_to_cumulative(remapped, Path(self.vars["retire_batting"].get()), "retire_batting", key_cols=["id", "year"])
            if results.get("league_leaders_pitching") is not None:
                remapped = core.remap_columns(results["league_leaders_pitching"], core.RETIRE_PITCHING_COLUMNS, "retire_pitching")
                core.append_to_cumulative(remapped, Path(self.vars["retire_pitching"].get()), "retire_pitching", key_cols=["id", "year"])

            print("\n=== Done. Upload the four files above to the site as usual. ===")
            messagebox.showinfo("Done", "Update finished — see the log for details.\nUpload the four files to the site as usual.")

        except SystemExit as e:
            print(f"\nSTOPPED: {e}")
            messagebox.showerror("Error", f"Something went wrong:\n{e}\n\nSee the log for details.")
        except Exception:
            tb = traceback.format_exc()
            print(f"\nUNEXPECTED ERROR:\n{tb}")
            messagebox.showerror("Error", "An unexpected error occurred.\nSee the log for details.")
        finally:
            sys.stdout = old_stdout

        # Remember the site file locations (not the season folder — that changes every year)
        for key, _ in FIELDS:
            self.cfg[key] = self.vars[key].get()
        self.cfg["season_folder"] = str(folder.parent)  # remember the parent, to speed up next Browse
        save_config(self.cfg)


if __name__ == "__main__":
    app = App()
    app.mainloop()
