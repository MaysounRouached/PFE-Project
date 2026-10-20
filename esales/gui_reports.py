"""
Unified launcher for the reporting scripts — no code editing required.

Drop this file into ANY of your report folders alongside its matching
clean_*.py script (and its inputs). On startup it automatically detects
which report script(s) actually live in this same folder and only offers
those — so a copy sitting in your eSales folder only ever shows the eSales
report, a copy in your Channel Revenue folder only shows that one, and so
on. No risk of picking a report whose script isn't actually there.

If you keep two or more scripts in the same folder, all of those show up
too — nothing about this requires a single shared folder.
"""
import importlib
import io
import os
import queue
import sys
import threading
import tkinter as tk
from contextlib import redirect_stdout, redirect_stderr
from tkinter import filedialog, messagebox, ttk

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

# ---------------------------------------------------------------------------
# Full catalog of every report this launcher knows how to run. At startup,
# only the ones whose .py file actually exists next to this GUI are shown —
# see _detect_available_reports().
# ---------------------------------------------------------------------------
REPORT_CATALOG = {
    "Weekly Partners Performance Report": {
        "module": "clean_weekly_report",
        "fields": [
            ("RAW_FILE_PATH", "Raw Weekly Partner Data (.xlsx)", False),
            ("TARGET_FILE_PATH", "Partner Targets & Revenue (.csv)", False),
            ("HISTORICAL_FILE_PATH", "Historical Revenue (.csv)", False),
            ("RAW_DR_EXPORT_FILE", "Raw DR Export (.xlsx)", False),
            ("HISTORICAL_PL_FILE", "Historical PL (.xlsx)", False),
            ("CHANNEL_REVENUE_REPORT_FILE", "Channel Revenue Report output (.xlsx) \u2014 optional, for Channel/AM split", True),
        ],
    },
    "Channel Revenue Report": {
        "module": "clean_channel_revenue_report",
        "fields": [
            ("RAW_CONTRACT_FILE", "Raw Channel Revenue (.xlsx)", False),
            ("OPP_DR_FILE", "Opportunity / DR Mapping (.xlsx)", False),
            ("HISTORY_FILES", "History Data (.xlsx) \u2014 optional", True),
            ("REFERENCE_FILE", "Reference Data (.xlsx) \u2014 optional", True),
            ("TIER1_PO_REFERENCE_FILE", "Tier 1 PO Reference (.xlsx) \u2014 optional", True),
        ],
    },
    "eSales Report": {
        "module": "clean_esales_report",
        "fields": [
            ("CURRENT_FILE", "Current Week eSales Export (.xlsx)", False),
            ("DR_FILE", "DR Export (.xlsx)", False),
        ],
    },
    "Certifications Report": {
        "module": "clean_certifications",
        "fields": [
            ("RAW_FILE_PATH", "Raw Certifications Data (.xlsx)", False),
            ("UNIVERSITY_LOOKUP_PATH", "University Lookup (.csv) \u2014 optional, uses existing file if not reselected", True),
            ("LEVEL_RULES_PATH", "Partner Level Rules (.csv) \u2014 optional", True),
            ("PARTNER_TARGETS_PATH", "Partner Targets (.csv) \u2014 optional", True),
        ],
    },
}

RED = "#C8102E"
DARK_RED = "#7A1414"
GRAPHITE = "#2B2B2B"
LIGHT_GREY = "#F4F4F4"
MID_GREY = "#666666"
GREEN = "#2E7D32"
FAIL_RED = "#C62828"
LOG_BG = "#1B1B1B"
LOG_FG = "#E6E6E6"


def _detect_available_reports():
    available = {}
    for name, spec in REPORT_CATALOG.items():
        script_path = os.path.join(SCRIPT_DIR, spec["module"] + ".py")
        if os.path.exists(script_path):
            available[name] = spec
    return available


class ReportGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reporting Suite")
        self.geometry("820x680")
        self.minsize(700, 560)
        self.configure(bg="white")

        self.available_reports = _detect_available_reports()
        self.log_queue = queue.Queue()
        self.field_vars = {}

        self._setup_style()

        if not self.available_reports:
            self._build_no_scripts_found_layout()
            return

        self._build_layout()
        self._on_report_change()
        self.after(100, self._poll_log_queue)

    # -----------------------------------------------------------------
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.font_title = ("Segoe UI", 16, "bold")
        self.font_subtitle = ("Segoe UI", 9)
        self.font_section = ("Segoe UI", 10, "bold")
        self.font_label = ("Segoe UI", 10)
        self.font_mono = ("Consolas", 10)

        style.configure("TFrame", background="white")
        style.configure("Header.TFrame", background=RED)
        style.configure("Card.TFrame", background=LIGHT_GREY)
        style.configure("TLabel", background="white", font=self.font_label)
        style.configure("Header.TLabel", background=RED, foreground="white", font=self.font_title)
        style.configure("SubHeader.TLabel", background=RED, foreground="#FFE3E3", font=self.font_subtitle)
        style.configure("Section.TLabel", background="white", foreground=DARK_RED, font=self.font_section)
        style.configure("Card.TLabel", background=LIGHT_GREY, font=self.font_label)
        style.configure("Status.TLabel", background="white", font=("Segoe UI", 10, "bold"))

        style.configure("TCombobox", padding=4)
        style.configure(
            "Accent.TButton", font=("Segoe UI", 11, "bold"),
            background=RED, foreground="white", padding=(18, 10), borderwidth=0,
        )
        style.map("Accent.TButton",
                  background=[("active", DARK_RED), ("disabled", "#D9A0A6")],
                  foreground=[("disabled", "#FFFFFF")])
        style.configure(
            "Browse.TButton", font=("Segoe UI", 9), padding=(10, 5),
            background=GRAPHITE, foreground="white", borderwidth=0,
        )
        style.map("Browse.TButton", background=[("active", "#444444")])

    # -----------------------------------------------------------------
    def _build_no_scripts_found_layout(self):
        frame = ttk.Frame(self, padding=40)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="No report script found in this folder.",
            font=("Segoe UI", 13, "bold"), foreground=FAIL_RED,
        ).pack(pady=(0, 10))
        known = ", ".join(spec["module"] + ".py" for spec in REPORT_CATALOG.values())
        ttk.Label(
            frame,
            text=f"This launcher looks for one of:\n{known}\n\nin its own folder, and only "
                 f"offers reports whose script is actually present.\nPlace this file next to "
                 f"the matching clean_*.py script and its input files, then re-open it.",
            font=self.font_label, justify="center",
        ).pack()

    # -----------------------------------------------------------------
    def _build_layout(self):
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")
        inner = ttk.Frame(header, style="Header.TFrame", padding=(20, 14, 20, 14))
        inner.pack(fill="x")
        ttk.Label(inner, text="Reporting Suite", style="Header.TLabel").pack(anchor="w")
        n = len(self.available_reports)
        subtitle = (f"{n} report{'s' if n != 1 else ''} available in this folder"
                    if n > 1 else "1 report available in this folder")
        ttk.Label(inner, text=subtitle, style="SubHeader.TLabel").pack(anchor="w", pady=(2, 0))

        body = ttk.Frame(self, padding=20)
        body.pack(fill="both", expand=True)

        picker_row = ttk.Frame(body)
        picker_row.pack(fill="x", pady=(0, 14))
        ttk.Label(picker_row, text="Report:", style="Section.TLabel").pack(side="left")
        self.report_choice = tk.StringVar(value=list(self.available_reports.keys())[0])
        self.report_dropdown = ttk.Combobox(
            picker_row, textvariable=self.report_choice,
            values=list(self.available_reports.keys()),
            state="readonly", width=48, font=self.font_label,
        )
        self.report_dropdown.pack(side="left", padx=(12, 0))
        self.report_dropdown.bind("<<ComboboxSelected>>", lambda e: self._on_report_change())

        card = ttk.Frame(body, style="Card.TFrame", padding=16)
        card.pack(fill="x", pady=(0, 16))
        ttk.Label(card, text="INPUT FILES", style="Card.TLabel", font=("Segoe UI", 9, "bold"),
                  foreground=MID_GREY).pack(anchor="w", pady=(0, 8))
        self.fields_frame = ttk.Frame(card, style="Card.TFrame")
        self.fields_frame.pack(fill="x")

        action_row = ttk.Frame(body)
        action_row.pack(fill="x", pady=(0, 14))
        self.build_btn = ttk.Button(action_row, text="Build Report", style="Accent.TButton",
                                     command=self._on_build_clicked)
        self.build_btn.pack(side="left")
        self.status_label = ttk.Label(action_row, text="", style="Status.TLabel")
        self.status_label.pack(side="left", padx=16)

        ttk.Label(body, text="LOG", style="Card.TLabel", font=("Segoe UI", 9, "bold"),
                  foreground=MID_GREY, background="white").pack(anchor="w", pady=(0, 6))
        log_container = tk.Frame(body, bg=LOG_BG, bd=0, highlightthickness=1,
                                  highlightbackground="#3A3A3A")
        log_container.pack(fill="both", expand=True)
        self.log_text = tk.Text(
            log_container, wrap="word", state="disabled", bg=LOG_BG, fg=LOG_FG,
            font=self.font_mono, bd=0, padx=12, pady=10, insertbackground=LOG_FG,
        )
        self.log_text.pack(fill="both", expand=True)

    # -----------------------------------------------------------------
    def _on_report_change(self):
        for child in self.fields_frame.winfo_children():
            child.destroy()
        self.field_vars = {}

        report = self.available_reports[self.report_choice.get()]
        for config_name, label, optional in report["fields"]:
            row = ttk.Frame(self.fields_frame, style="Card.TFrame")
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=label, width=58, anchor="w", style="Card.TLabel").pack(side="left")
            var = tk.StringVar()
            self.field_vars[config_name] = var
            entry = tk.Entry(row, textvariable=var, font=self.font_label, bd=1, relief="solid")
            entry.pack(side="left", fill="x", expand=True, padx=8, ipady=3)
            ttk.Button(row, text="Browse...", style="Browse.TButton",
                       command=lambda v=var: self._browse(v)).pack(side="left")

    def _browse(self, var):
        path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            var.set(path)

    # -----------------------------------------------------------------
    def _on_build_clicked(self):
        report_name = self.report_choice.get()
        report = self.available_reports[report_name]

        missing = []
        for config_name, label, optional in report["fields"]:
            val = self.field_vars[config_name].get().strip()
            if not val and not optional:
                missing.append(label)
        if missing:
            messagebox.showerror(
                "Missing files",
                "Please browse to a file for:\n\n" + "\n".join(f"\u2022 {m}" for m in missing),
            )
            return

        self.build_btn.config(state="disabled")
        self.status_label.config(text=f"Building {report_name}...", foreground=MID_GREY)
        self._clear_log()

        thread = threading.Thread(target=self._run_report, args=(report_name, report), daemon=True)
        thread.start()

    def _run_report(self, report_name, report):
        module_name = report["module"]
        try:
            if module_name in sys.modules:
                del sys.modules[module_name]
            mod = importlib.import_module(module_name)

            for config_name, label, optional in report["fields"]:
                val = self.field_vars[config_name].get().strip()
                if val:
                    setattr(mod, config_name, val)

            class _Tee(io.StringIO):
                def write(_self, s):
                    self.log_queue.put(s)
                    return len(s)

            tee = _Tee()
            with redirect_stdout(tee), redirect_stderr(tee):
                mod.main()

            output_path = getattr(mod, "OUTPUT_PATH", None)
            self.log_queue.put(("__DONE__", output_path))

        except Exception as e:
            import traceback
            self.log_queue.put(traceback.format_exc())
            self.log_queue.put(("__ERROR__", str(e)))

    # -----------------------------------------------------------------
    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    self._on_build_finished(success=True, output_path=item[1])
                elif isinstance(item, tuple) and item[0] == "__ERROR__":
                    self._on_build_finished(success=False, error=item[1])
                else:
                    self._append_log(item)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _append_log(self, text):
        self.log_text.config(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _on_build_finished(self, success, output_path=None, error=None):
        self.build_btn.config(state="normal")
        if success:
            self.status_label.config(text="\u2713 Done", foreground=GREEN)
            full_path = os.path.abspath(output_path) if output_path else None
            answer = messagebox.askyesno(
                "Report built",
                f"Report generated successfully.\n\n{full_path or ''}\n\nOpen the output folder now?",
            )
            if answer and full_path:
                folder = os.path.dirname(full_path)
                try:
                    if sys.platform == "win32":
                        os.startfile(folder)
                    elif sys.platform == "darwin":
                        os.system(f'open "{folder}"')
                    else:
                        os.system(f'xdg-open "{folder}"')
                except Exception:
                    pass
        else:
            self.status_label.config(text="\u2717 Failed \u2014 see log", foreground=FAIL_RED)
            messagebox.showerror("Build failed", f"The report failed to build:\n\n{error}\n\nFull details are in the log.")


if __name__ == "__main__":
    app = ReportGUI()
    app.mainloop()