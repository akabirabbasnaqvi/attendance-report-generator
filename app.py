from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from attendance_report import build_report, load_punches, load_schedule, to_excel

try:
    import customtkinter as ctk
except ImportError:
    ctk = None


class AttendanceApp:
    def __init__(self) -> None:
        if ctk:
            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")
            self.root = ctk.CTk()
        else:
            self.root = tk.Tk()
        self.root.title("Attendance Report Generator")
        self.root.geometry("1250x760")
        self.attendance_path: Path | None = None
        self.schedule_path: Path | None = None
        self.schedule = None
        self.detail = None
        self.summary = None
        self.periods: list[tuple[int, int]] = []
        self._build_ui()

    def _build_ui(self) -> None:
        frame_type = ctk.CTkFrame if ctk else tk.Frame
        label_type = ctk.CTkLabel if ctk else tk.Label
        button_type = ctk.CTkButton if ctk else tk.Button
        combo_type = ctk.CTkComboBox if ctk else ttk.Combobox
        outer = frame_type(self.root)
        outer.pack(fill="both", expand=True, padx=18, pady=18)
        label_type(outer, text="Attendance Report Generator", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(0, 4))
        label_type(outer, text="Upload both workbooks, choose a month, and export the complete employee report.").pack(anchor="w", pady=(0, 14))

        controls = frame_type(outer)
        controls.pack(fill="x", pady=(0, 12))
        button_type(controls, text="Choose attendance file", command=self.choose_attendance).grid(row=0, column=0, padx=(0, 8), pady=4)
        self.attendance_label = label_type(controls, text="No attendance file selected", anchor="w")
        self.attendance_label.grid(row=0, column=1, sticky="w", padx=8)
        button_type(controls, text="Choose schedule file", command=self.choose_schedule).grid(row=1, column=0, padx=(0, 8), pady=4)
        self.schedule_label = label_type(controls, text="No schedule file selected", anchor="w")
        self.schedule_label.grid(row=1, column=1, sticky="w", padx=8)
        label_type(controls, text="Report month:").grid(row=2, column=0, sticky="w", pady=(10, 4))
        self.month_box = combo_type(controls, values=[], state="readonly", width=28)
        self.month_box.grid(row=2, column=1, sticky="w", padx=8, pady=(10, 4))
        button_type(controls, text="Generate report", command=self.generate, width=150).grid(row=2, column=2, padx=8, pady=(10, 4))
        self.export_button = button_type(controls, text="Export Excel", command=self.export, width=120, state="disabled")
        self.export_button.grid(row=2, column=3, padx=8, pady=(10, 4))
        self.status_label = label_type(outer, text="Select both files to begin.", anchor="w")
        self.status_label.pack(fill="x", pady=(0, 8))

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)
        self.summary_tree = self._make_tree(notebook, "summary")
        self.detail_tree = self._make_tree(notebook, "detail")

    def _make_tree(self, parent, kind):
        frame = ttk.Frame(parent)
        parent.add(frame, text="Employee Summary" if kind == "summary" else "Daily Detail")
        tree = ttk.Treeview(frame, show="headings")
        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def choose_attendance(self):
        path = filedialog.askopenfilename(title="Choose attendance file", filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")])
        if path:
            self.attendance_path = Path(path)
            self.attendance_label.configure(text=self.attendance_path.name)
            self._load_if_ready()

    def choose_schedule(self):
        path = filedialog.askopenfilename(title="Choose schedule file", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.schedule_path = Path(path)
            self.schedule_label.configure(text=self.schedule_path.name)
            self._load_if_ready()

    def _load_if_ready(self):
        if not self.attendance_path or not self.schedule_path:
            return
        try:
            load_punches(self.attendance_path)
            self.schedule = load_schedule(self.schedule_path)
            self.periods = sorted({(item.year, item.month) for item in self.schedule["Work Date"]})
            labels = [pd.Timestamp(year=y, month=m, day=1).strftime("%B %Y") for y, m in self.periods]
            self.month_box.configure(values=labels)
            if labels:
                self.month_box.set(labels[-1])
            self.status_label.configure(text=f"Loaded {len(self.periods)} months. All employees will be retained in the report.")
        except Exception as error:
            messagebox.showerror("Could not read files", str(error))

    def generate(self):
        if not self.attendance_path or self.schedule is None or not self.month_box.get():
            messagebox.showwarning("Missing input", "Choose both files and a report month first.")
            return
        try:
            index = list(self.month_box.cget("values")).index(self.month_box.get())
            year, month = self.periods[index]
            punches = load_punches(self.attendance_path)
            self.detail, self.summary = build_report(punches, self.schedule, year, month)
            self._fill_tree(self.summary_tree, self.summary)
            self._fill_tree(self.detail_tree, self.detail)
            self.export_button.configure(state="normal")
            self.status_label.configure(text=f"Generated {len(self.summary)} employees and {len(self.detail)} daily rows.")
        except Exception as error:
            messagebox.showerror("Report error", str(error))

    def _fill_tree(self, tree, data):
        tree.delete(*tree.get_children())
        tree.configure(columns=list(data.columns))
        for column in data.columns:
            tree.heading(column, text=column)
            tree.column(column, width=max(110, min(220, len(column) * 12)), anchor="w")
        for row in data.itertuples(index=False, name=None):
            tree.insert("", "end", values=["" if pd.isna(value) else value for value in row])

    def export(self):
        if self.detail is None or self.summary is None:
            return
        path = filedialog.asksaveasfilename(title="Save attendance report", defaultextension=".xlsx", initialfile="attendance_report.xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            Path(path).write_bytes(to_excel(self.detail, self.summary))
            self.status_label.configure(text=f"Saved report: {Path(path).name}")
            messagebox.showinfo("Report saved", f"Report saved to:\n{path}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    AttendanceApp().run()
