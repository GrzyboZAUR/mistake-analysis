import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
from pathlib import Path
import threading
import traceback
from datetime import datetime
import sys
from report_engine_s2 import generate_report

# ── PATH DETECTION (.exe vs .py) ────────────────────────────────────────
if getattr(sys, "frozen", False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).parent

# ── RELATIVE PATHS (all inside project folder) ──────────────────────────
DATA_FILE = _BASE_DIR / "output" / "errors_combined.xlsx"
STAN_FILE = _BASE_DIR / "output" / "headcount.xlsx"

LOG_DIR = _BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "log_reports.txt"
LOG_DIR.mkdir(parents=True, exist_ok=True)


# ── LOGGING ─────────────────────────────────────────────────────────────
def log(msg=""):
    """Print to console AND append to log file."""
    print(msg)
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            for line in str(msg).split("\n"):
                f.write(f"[{timestamp}] {line}\n")
    except Exception:
        pass


def _global_exception_handler(exc_type, exc_value, exc_tb):
    """Catch every unhandled exception and write it to the log."""
    error = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    log("\n" + "=" * 60)
    log("❌ UNEXPECTED ERROR:")
    log(error)
    log("=" * 60)


sys.excepthook = _global_exception_handler

# ── Startup log ─────────────────────────────────────────────────────────
log("\n" + "#" * 60)
log(f"# START: gui  ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
log(f"# DATA_FILE: {DATA_FILE}")
log(f"# STAN_FILE: {STAN_FILE}")
log("#" * 60)

MONTHS = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}
MONTHS_REVERSE = {v: k for k, v in MONTHS.items()}


class ReportGenerator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Report Generator")
        self.resizable(False, False)

        # Check if input files exist - friendly error instead of stacktrace
        if not DATA_FILE.exists() or not STAN_FILE.exists():
            messagebox.showerror(
                "Error",
                f"Input files not found.\n\n"
                f"Expected:\n  {DATA_FILE}\n  {STAN_FILE}\n\n"
                f"Run process_errors.py first."
            )
            self.destroy()
            return

        try:
            self.fin = pd.read_excel(DATA_FILE)
            self.fin["date"] = pd.to_datetime(self.fin["date"])
            self.stan = pd.read_excel(STAN_FILE)
            self.stan["date"] = pd.to_datetime(self.stan["date"])
        except FileNotFoundError as e:
            messagebox.showerror(
                "Error",
                f"File not found:\n{e}\n\n"
                f"Run process_errors.py first."
            )
            self.destroy()
            return
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load files:\n{e}")
            self.destroy()
            return

        self.available_years = sorted(self.fin["date"].dt.year.unique(), reverse=True)
        self.available_months = sorted(self.fin["date"].dt.month.unique())
        self.available_stages = sorted(self.fin["greenhouse"].unique())

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # --- REPORT TYPE ---
        frame_type = ttk.LabelFrame(self, text="Report type")
        frame_type.grid(row=0, column=0, sticky="ew", **pad)

        self.var_report_type = tk.StringVar(value="miesiac")
        ttk.Radiobutton(frame_type, text="Monthly",
                        variable=self.var_report_type, value="miesiac",
                        command=self._toggle_type).grid(row=0, column=0, padx=10, pady=4)
        ttk.Radiobutton(frame_type, text="Per stage",
                        variable=self.var_report_type, value="szklarnia",
                        command=self._toggle_type).grid(row=0, column=1, padx=10, pady=4)

        # --- MONTHLY OPTIONS (multi-select) ---
        self.frame_monthly = ttk.LabelFrame(self, text="Period (select one or more months)")
        self.frame_monthly.grid(row=1, column=0, sticky="ew", **pad)

        # Year
        ttk.Label(self.frame_monthly, text="Year:").grid(
            row=0, column=0, padx=5, pady=4, sticky="w")
        self.var_year = tk.StringVar(value=str(self.available_years[0]))
        ttk.Combobox(self.frame_monthly, textvariable=self.var_year,
                     values=[str(y) for y in self.available_years],
                     width=8, state="readonly").grid(
            row=0, column=1, padx=5, pady=4, sticky="w")

        # Month list (multi-select)
        ttk.Label(self.frame_monthly, text="Months (Ctrl+click - multiple):").grid(
            row=1, column=0, columnspan=2, padx=5, pady=(4, 2), sticky="w")

        # Frame for listbox + scrollbar
        list_frame = ttk.Frame(self.frame_monthly)
        list_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=(0, 6), sticky="ew")

        self.list_months = tk.Listbox(
            list_frame, selectmode="extended",
            height=6, width=20, exportselection=False,
            font=("Segoe UI", 9),
        )
        self.list_months.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                  command=self.list_months.yview)
        scrollbar.pack(side="right", fill="y")
        self.list_months.config(yscrollcommand=scrollbar.set)

        # Fill list with months (sorted)
        for m in sorted(self.available_months):
            self.list_months.insert("end", MONTHS[m])

        # Default: select last available month
        if self.available_months:
            last_idx = len(self.available_months) - 1
            self.list_months.selection_set(last_idx)
            self.list_months.see(last_idx)

        # --- PER STAGE OPTIONS ---
        self.frame_stage = ttk.LabelFrame(self, text="Stage")
        self.frame_stage.grid(row=1, column=0, sticky="ew", **pad)
        self.frame_stage.grid_remove()

        ttk.Label(self.frame_stage, text="Stage:").grid(
            row=0, column=0, padx=5, pady=4, sticky="w")
        self.var_stage = tk.StringVar(
            value=self.available_stages[0] if self.available_stages else ""
        )
        ttk.Combobox(self.frame_stage, textvariable=self.var_stage,
                     values=self.available_stages, width=20,
                     state="readonly").grid(row=0, column=1, padx=5, pady=4, sticky="w")

        # Checkbox "All months"
        self.var_all_months = tk.BooleanVar(value=True)
        self.chk_all_months = ttk.Checkbutton(
            self.frame_stage,
            text="All months (no restriction)",
            variable=self.var_all_months,
            command=self._toggle_stage_months,
        )
        self.chk_all_months.grid(row=1, column=0, columnspan=2,
                                 padx=5, pady=(6, 2), sticky="w")

        ttk.Label(self.frame_stage, text="Selected months (Ctrl+click):").grid(
            row=2, column=0, columnspan=2, padx=5, pady=(4, 2), sticky="w")

        # Month list for stage
        stage_list_frame = ttk.Frame(self.frame_stage)
        stage_list_frame.grid(row=3, column=0, columnspan=2,
                              padx=5, pady=(0, 6), sticky="ew")

        self.list_stage_months = tk.Listbox(
            stage_list_frame, selectmode="extended",
            height=6, width=20, exportselection=False,
            font=("Segoe UI", 9),
        )
        self.list_stage_months.pack(side="left", fill="both", expand=True)

        scrollbar_stage = ttk.Scrollbar(stage_list_frame, orient="vertical",
                                        command=self.list_stage_months.yview)
        scrollbar_stage.pack(side="right", fill="y")
        self.list_stage_months.config(yscrollcommand=scrollbar_stage.set)

        # Fill list with months
        for m in sorted(self.available_months):
            self.list_stage_months.insert("end", MONTHS[m])

        # Default: list disabled ("All months" checked)
        self.list_stage_months.config(state="disabled")

        # --- COMMENT ---
        frame_comment = ttk.LabelFrame(self, text="Comment (optional)")
        frame_comment.grid(row=2, column=0, sticky="ew", **pad)

        self.text_comment = scrolledtext.ScrolledText(
            frame_comment, height=5, width=55, font=("Arial", 9))
        self.text_comment.pack(padx=5, pady=5)

        # ── "MORE OPTIONS" BUTTON ────────────────────────────────────
        self.var_more = tk.BooleanVar(value=False)
        self.btn_more = ttk.Button(
            self, text="▼  More options",
            command=self._toggle_more
        )
        self.btn_more.grid(row=3, column=0, pady=(4, 0))

        # ── ADVANCED FRAME (hidden by default) ───────────────────────
        self.frame_more = ttk.Frame(self)
        self.frame_more.grid(row=4, column=0, sticky="ew", padx=10, pady=4)
        self.frame_more.grid_remove()  # hidden initially

        # --- Title page ---
        frame_title = ttk.LabelFrame(self.frame_more, text="Title page")
        frame_title.pack(fill="x", pady=(0, 6))

        self.var_title_page = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_title, text="Add title page",
                        variable=self.var_title_page).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        ttk.Label(frame_title, text="Author:").grid(row=1, column=0, sticky="e", padx=5)
        self.var_author = tk.StringVar(value="Bartosz Grzybowski")
        ttk.Entry(frame_title, textvariable=self.var_author,
                  width=35).grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(frame_title, text="Position:").grid(row=2, column=0, sticky="e", padx=5)
        self.var_position = tk.StringVar(value="Data Administration Specialist")
        ttk.Entry(frame_title, textvariable=self.var_position,
                  width=35).grid(row=2, column=1, sticky="w", padx=5)

        ttk.Label(frame_title, text="E-mail:").grid(row=3, column=0, sticky="e", padx=5)
        self.var_email = tk.StringVar(value="bargrzybowski@gmail.com")
        ttk.Entry(frame_title, textvariable=self.var_email,
                  width=35).grid(row=3, column=1, sticky="w", padx=5)

        ttk.Label(frame_title, text="Title:").grid(row=4, column=0, sticky="e", padx=5)
        self.var_title = tk.StringVar(value="Employee Error Report")
        ttk.Entry(frame_title, textvariable=self.var_title,
                  width=35).grid(row=4, column=1, sticky="w", padx=5)

        ttk.Label(frame_title, text="Subtitle (optional):").grid(
            row=5, column=0, sticky="e", padx=5)
        self.var_subtitle = tk.StringVar(value="")
        ttk.Entry(frame_title, textvariable=self.var_subtitle,
                  width=35).grid(row=5, column=1, sticky="w", padx=5)

        # --- Sections ---
        frame_sections = ttk.LabelFrame(self.frame_more, text="Report sections")
        frame_sections.pack(fill="x", pady=(0, 6))

        # NOTE: keys stay Polish - report_engine.py reads them by these names
        self.var_sections = {
            "statystyki": tk.BooleanVar(value=True),
            "typy": tk.BooleanVar(value=True),
            "podzial": tk.BooleanVar(value=True),
            "obecnosc": tk.BooleanVar(value=True),
            "korelacja": tk.BooleanVar(value=False),
            "trend": tk.BooleanVar(value=False),
        }

        section_labels = {
            "statystyki": "General statistics (table)",
            "typy": "Error type distribution (table)",
            "podzial": "Error breakdown per location",
            "obecnosc": "Error attendance per person",
            "korelacja": "Correlation: headcount vs. errors",
            "trend": "Monthly trend",
        }

        self.section_checkboxes = {}
        for i, (key, label) in enumerate(section_labels.items()):
            cb = ttk.Checkbutton(frame_sections, text=label,
                                 variable=self.var_sections[key])
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=5, pady=1)
            self.section_checkboxes[key] = cb

        # --- GENERATE BUTTON ---
        self.btn = ttk.Button(self, text="🖨  Generate PDF", command=self._generate)
        self.btn.grid(row=5, column=0, pady=8)

        self.label_status = ttk.Label(self, text="", foreground="grey")
        self.label_status.grid(row=6, column=0, pady=(0, 6))

    def _toggle_type(self):
        if self.var_report_type.get() == "miesiac":
            self.frame_stage.grid_remove()
            self.frame_monthly.grid()
            # Enable correlation checkbox
            self.section_checkboxes["korelacja"].config(state="normal")
        else:
            self.frame_monthly.grid_remove()
            self.frame_stage.grid()
            # Disable and uncheck correlation (only 1 point per stage)
            self.section_checkboxes["korelacja"].config(state="disabled")
            self.var_sections["korelacja"].set(False)

    def _generate(self):
        self.btn.config(state="disabled")
        self.label_status.config(text="⏳ Generating report...", foreground="grey")

        # NOTE: keys in params must stay Polish - report_engine.py reads them
        params = {
            "typ": self.var_report_type.get(),
            "komentarz": self.text_comment.get("1.0", "end").strip(),
            "tytul_strona": self.var_title_page.get(),
            "tytul_dane": {
                "autor": self.var_author.get(),
                "stanowisko": self.var_position.get(),
                "email": self.var_email.get(),
                "tytul": self.var_title.get(),
                "podtytul": self.var_subtitle.get(),
            },
            "sekcje": {k: v.get() for k, v in self.var_sections.items()},
        }

        if params["typ"] == "miesiac":
            selected_idx = self.list_months.curselection()
            if not selected_idx:
                messagebox.showwarning("No selection",
                                       "Select at least one month.")
                self.btn.config(state="normal")
                self.label_status.config(text="", foreground="grey")
                return

            selected_months = [self.available_months[i] for i in selected_idx]
            params["miesiace"] = sorted(selected_months)
            params["rok"] = int(self.var_year.get())

        else:  # per stage
            params["etap"] = self.var_stage.get()

            # Months for the stage
            if self.var_all_months.get():
                params["miesiace_szklarnia"] = None  # no filter
            else:
                selected = self.list_stage_months.curselection()
                if not selected:
                    messagebox.showwarning(
                        "No selection",
                        "You unchecked 'All months' - select at least one "
                        "month or switch back to 'All months'."
                    )
                    self.btn.config(state="normal")
                    self.label_status.config(text="", foreground="grey")
                    return
                params["miesiace_szklarnia"] = sorted(
                    self.available_months[i] for i in selected
                )

        threading.Thread(target=self._generate_in_background,
                         args=(params,)).start()

    def _generate_in_background(self, params):
        try:
            files = generate_report(self.fin, self.stan, params)
            self.after(0, self._success, files)
        except Exception:
            self.after(0, self._error, traceback.format_exc())

    def _success(self, files):
        self.btn.config(state="normal")
        self.label_status.config(
            text=f"✅ Saved: {', '.join(files)}", foreground="green")
        messagebox.showinfo("Done", "Report generated:\n" + "\n".join(files))

    def _error(self, msg):
        log("\n" + "=" * 60)
        log("❌ GUI ERROR:")
        log(msg)
        log("=" * 60)
        self.btn.config(state="normal")
        self.label_status.config(text="❌ Error - check log", foreground="red")
        messagebox.showerror(
            "Error",
            f"An error occurred. Details saved to:\n{LOG_FILE}"
        )

    def _toggle_more(self):
        if self.var_more.get():
            self.frame_more.grid()
            self.btn_more.config(text="▲  Hide options")
        else:
            self.frame_more.grid_remove()
            self.btn_more.config(text="▼  More options")
        self.var_more.set(not self.var_more.get())

    def _toggle_stage_months(self):
        """Enable/disable the month list for per-stage reports."""
        if self.var_all_months.get():
            self.list_stage_months.config(state="disabled")
        else:
            self.list_stage_months.config(state="normal")

if __name__ == "__main__":
    app = ReportGenerator()
    app.mainloop()
