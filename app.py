"""
app.py – Attendance Report GUI (customtkinter version)
"""

import os
import sys
import threading
import calendar
from datetime import datetime
from tkinter import filedialog, messagebox
import tkinter as tk
import tkinter.ttk as ttk

try:
    import customtkinter as ctk
    HAS_CTK = True
except ImportError:
    HAS_CTK = False

from attendance_report import (
    read_workbook, load_punches, load_schedule,
    build_report, to_excel,
)


# ── appearance ───────────────────────────────────────────────────────────

if HAS_CTK:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    _BaseWindow = ctk.CTk
    _BaseFrame = ctk.CTkFrame
    _Label = ctk.CTkLabel
    _Button = ctk.CTkButton
    _OptionMenu = ctk.CTkOptionMenu
    _Entry = ctk.CTkEntry
    _Tabview = ctk.CTkTabview
else:
    _BaseWindow = tk.Tk
    _BaseFrame = tk.Frame
    _Label = tk.Label
    _Button = tk.Button
    _OptionMenu = tk.OptionMenu
    _Entry = tk.Entry
    _Tabview = None  # fallback uses ttk.Notebook


def _style_treeview(tree: ttk.Treeview, parent: tk.Misc) -> ttk.Style:
    """Apply a dark-mode-friendly style to a ttk.Treeview."""
    style = ttk.Style(parent)

    # Detect current appearance
    is_dark = False
    if HAS_CTK:
        is_dark = ctk.get_appearance_mode().lower() == "dark"

    if is_dark:
        bg = "#2b2b2b"
        fg = "#e0e0e0"
        sel_bg = "#1f6aa5"
        sel_fg = "#ffffff"
        heading_bg = "#333333"
        heading_fg = "#e0e0e0"
        field_bg = "#2b2b2b"
    else:
        bg = "#ffffff"
        fg = "#1a1a1a"
        sel_bg = "#0078d4"
        sel_fg = "#ffffff"
        heading_bg = "#e0e0e0"
        heading_fg = "#1a1a1a"
        field_bg = "#ffffff"

    style_name = "Custom.Treeview"
    style.configure(style_name,
                    background=bg,
                    foreground=fg,
                    fieldbackground=field_bg,
                    rowheight=28,
                    font=("Segoe UI", 10))
    style.configure(f"{style_name}.Heading",
                    background=heading_bg,
                    foreground=heading_fg,
                    font=("Segoe UI", 10, "bold"))
    style.map(style_name,
              background=[("selected", sel_bg)],
              foreground=[("selected", sel_fg)])

    tree.configure(style=style_name)
    return style


# ── main app ─────────────────────────────────────────────────────────────

class AttendanceApp:
    """Main application window."""

    def __init__(self):
        self.root = _BaseWindow() if HAS_CTK else tk.Tk()
        self.root.title("Attendance Report Generator")
        self.root.geometry("1300x820")
        self.root.minsize(900, 600)

        self.schedule_path: str = ""
        self.device_path: str = ""

        self._build_ui()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self):
        root = self.root

        # top frame – file pickers
        top = _BaseFrame(root)
        top.pack(fill="x", padx=16, pady=(12, 4))

        _Label(top, text="Schedule File:").grid(row=0, column=0,
                                                 sticky="w", padx=4, pady=4)
        self.sched_entry = _Entry(top, width=500)
        self.sched_entry.grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        _Button(top, text="Browse…", width=90,
                command=self._pick_schedule).grid(row=0, column=2,
                                                  padx=4, pady=4)

        _Label(top, text="Device Data:").grid(row=1, column=0,
                                               sticky="w", padx=4, pady=4)
        self.dev_entry = _Entry(top, width=500)
        self.dev_entry.grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        _Button(top, text="Browse…", width=90,
                command=self._pick_device).grid(row=1, column=2,
                                                padx=4, pady=4)

        top.grid_columnconfigure(1, weight=1)

        # middle frame – month picker + buttons
        mid = _BaseFrame(root)
        mid.pack(fill="x", padx=16, pady=4)

        _Label(mid, text="Month:").pack(side="left", padx=(4, 2))

        months = [f"{calendar.month_name[m]} {y}"
                  for y in range(2024, 2028)
                  for m in range(1, 13)]
        now = datetime.now()
        default_month = f"{calendar.month_name[now.month]} {now.year}"

        if HAS_CTK:
            self.month_var = ctk.StringVar(value=default_month)
            _OptionMenu(mid, variable=self.month_var,
                        values=months, width=170).pack(side="left", padx=4)
        else:
            self.month_var = tk.StringVar(value=default_month)
            tk.OptionMenu(mid, self.month_var, *months).pack(side="left",
                                                              padx=4)

        _Button(mid, text="Generate Report",
                command=self._on_generate).pack(side="left", padx=12)
        _Button(mid, text="Export to Excel",
                command=self._on_export).pack(side="left", padx=4)

        # status label
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = _Label(mid, textvariable=self.status_var)
        self.status_label.pack(side="right", padx=8)

        # tabview
        if HAS_CTK and _Tabview is not None:
            self.tabview = ctk.CTkTabview(root)
            self.tabview.pack(fill="both", expand=True, padx=16, pady=(4, 12))
            tab_summary = self.tabview.add("Employee Summary")
            tab_detail = self.tabview.add("Daily Detail")
        else:
            nb = ttk.Notebook(root)
            nb.pack(fill="both", expand=True, padx=16, pady=(4, 12))
            tab_summary = ttk.Frame(nb)
            tab_detail = ttk.Frame(nb)
            nb.add(tab_summary, text="Employee Summary")
            nb.add(tab_detail, text="Daily Detail")

        self.summary_tree = self._make_tree(tab_summary, [
            "Employee", "Present", "On Time", "Late",
            "Absent", "WFH", "Leave", "Leave Breakdown",
        ])
        self.detail_tree = self._make_tree(tab_detail, [
            "Employee", "Date", "Day", "Scheduled",
            "Status", "Clock In", "Clock Out", "Punch Count", "Remarks",
        ])

        # keep references for export
        self.daily_df = None
        self.summary_df = None

    # ── helpers ──────────────────────────────────────────────────────────

    def _make_tree(self, parent, columns) -> ttk.Treeview:
        """Create a scrollable, styled Treeview."""
        frame = _BaseFrame(parent) if HAS_CTK else tk.Frame(parent)
        frame.pack(fill="both", expand=True, padx=2, pady=2)

        tree = ttk.Treeview(frame, columns=columns, show="headings",
                            selectmode="browse")

        # column widths
        for col in columns:
            width = 140 if col in ("Employee", "Leave Breakdown", "Remarks") else 95
            tree.heading(col, text=col, anchor="w")
            tree.column(col, width=width, minwidth=60, anchor="w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        _style_treeview(tree, self.root)
        return tree

    def _set_entry(self, entry, text: str):
        if HAS_CTK:
            entry.delete(0, "end")
            entry.insert(0, text)
        else:
            entry.delete(0, tk.END)
            entry.insert(0, text)

    def _pick_schedule(self):
        path = filedialog.askopenfilename(
            title="Select Schedule File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.schedule_path = path
            self._set_entry(self.sched_entry, path)

    def _pick_device(self):
        path = filedialog.askopenfilename(
            title="Select Device Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.device_path = path
            self._set_entry(self.dev_entry, path)

    # ── generation ───────────────────────────────────────────────────────

    def _on_generate(self):
        if not self.schedule_path or not self.device_path:
            messagebox.showwarning("Missing files",
                                   "Please select both the schedule and device data files.")
            return

        self.status_var.set("Loading data…")
        self.root.update_idletasks()

        # run in thread to avoid freezing
        threading.Thread(target=self._generate_worker, daemon=True).start()

    def _generate_worker(self):
        try:
            self._update_status("Reading device data…")
            punch_xf = read_workbook(self.device_path)
            punches = load_punches(punch_xf)

            self._update_status("Reading schedule…")
            sched_xf = read_workbook(self.schedule_path)

            # parse month
            month_str = self.month_var.get()
            dt = datetime.strptime(month_str, "%B %Y")
            year, month = dt.year, dt.month

            schedule = load_schedule(sched_xf, year, month)

            self._update_status("Building report…")
            daily_df, summary_df = build_report(punches, schedule, year, month)

            self.daily_df = daily_df
            self.summary_df = summary_df

            # populate trees on main thread
            self.root.after(0, lambda: self._populate(daily_df, summary_df))
            self._update_status(
                f"Done – {len(summary_df)} employees, {len(daily_df)} rows")
        except Exception as exc:
            self._update_status(f"Error: {exc}")
            self.root.after(0, lambda e=exc: messagebox.showerror(
                "Error", f"Report generation failed:\n{e}"))

    def _update_status(self, text: str):
        self.root.after(0, lambda: self.status_var.set(text))

    def _populate(self, daily_df, summary_df):
        """Fill both Treeview widgets."""
        # summary
        self.summary_tree.delete(*self.summary_tree.get_children())
        cols = list(summary_df.columns)
        self.summary_tree.configure(columns=cols)
        for c in cols:
            self.summary_tree.heading(c, text=c, anchor="w")
            w = 160 if c in ("Employee", "Leave Breakdown") else 80
            self.summary_tree.column(c, width=w, minwidth=50, anchor="w")
        for _, row in summary_df.iterrows():
            vals = [str(row[c]) if not (isinstance(row[c], float) and row[c] != row[c]) else ""
                    for c in cols]
            self.summary_tree.insert("", "end", values=vals)

        # detail
        self.detail_tree.delete(*self.detail_tree.get_children())
        cols = list(daily_df.columns)
        self.detail_tree.configure(columns=cols)
        for c in cols:
            self.detail_tree.heading(c, text=c, anchor="w")
            w = 140 if c in ("Employee", "Remarks") else 90
            self.detail_tree.column(c, width=w, minwidth=50, anchor="w")
        for _, row in daily_df.iterrows():
            vals = [str(row[c]) if not (isinstance(row[c], float) and row[c] != row[c]) else ""
                    for c in cols]
            self.detail_tree.insert("", "end", values=vals)

    # ── export ───────────────────────────────────────────────────────────

    def _on_export(self):
        if self.daily_df is None or self.summary_df is None:
            messagebox.showwarning("No data",
                                   "Generate a report first before exporting.")
            return

        path = filedialog.asksaveasfilename(
            title="Save Report",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        try:
            self.status_var.set("Exporting…")
            self.root.update_idletasks()
            to_excel(self.daily_df, self.summary_df, path)
            self.status_var.set(f"Exported to {os.path.basename(path)}")
            messagebox.showinfo("Success", f"Report saved to:\n{path}")
        except Exception as exc:
            self.status_var.set(f"Export error: {exc}")
            messagebox.showerror("Export Error", str(exc))

    # ── run ──────────────────────────────────────────────────────────────

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = AttendanceApp()
    app.run()
