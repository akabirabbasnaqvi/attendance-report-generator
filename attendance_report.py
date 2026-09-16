from __future__ import annotations

import io
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd


GRACE_MINUTES = 15
NON_WORKING_STATES = {"OFF", "WFH", "LV", "LEAVE", "RESIGNED"}


def read_workbook(uploaded_file_or_path, **kwargs) -> pd.DataFrame:
    """Read xlsx and legacy xls files, including the duplicated .xls extension."""
    name = getattr(uploaded_file_or_path, "name", str(uploaded_file_or_path)).lower()
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"
    return pd.read_excel(uploaded_file_or_path, engine=engine, **kwargs)


def normalize_name(value: object) -> str:
    value = re.sub(r"\([^)]*\)", "", str(value or ""))
    return re.sub(r"[^a-z0-9]", "", value.lower())


def load_punches(source) -> pd.DataFrame:
    punches = read_workbook(source)
    punches = punches.rename(columns={"No.": "Name No.", "Employee ID": "Name No.", "Employee Name": "Name"})
    required = {"Name No.", "Name", "Date/Time"}
    missing = required - set(punches.columns)
    if missing:
        raise ValueError(f"Attendance file is missing columns: {', '.join(sorted(missing))}")
    punches = punches[["Name No.", "Name", "Date/Time"]].copy()
    punches["Date/Time"] = pd.to_datetime(punches["Date/Time"], dayfirst=True, errors="coerce")
    punches = punches.dropna(subset=["Date/Time"])
    punches["Work Date"] = punches["Date/Time"].dt.date
    punches["Name Key"] = punches["Name"].map(normalize_name)
    return punches.sort_values("Date/Time")


def load_schedule(source) -> pd.DataFrame:
    workbook = pd.ExcelFile(source, engine="openpyxl")
    records: list[dict] = []
    for sheet in workbook.sheet_names:
        raw = pd.read_excel(source, sheet_name=sheet, header=None, engine="openpyxl")
        date_row = next((i for i in range(min(5, len(raw))) if raw.iloc[i].map(pd.to_datetime, errors="coerce").notna().sum() >= 2), None)
        if date_row is None:
            continue
        dates = pd.to_datetime(raw.iloc[date_row], errors="coerce")
        name_row = next((i for i in range(date_row + 1, min(date_row + 4, len(raw))) if "staff name" in str(raw.iloc[i, 0]).lower()), None)
        if name_row is None:
            name_row = date_row + 1
        for row_index in range(name_row + 1, len(raw)):
            staff_name = raw.iat[row_index, 0]
            if pd.isna(staff_name) or str(staff_name).strip().lower() in {"punctuality", "late comming", "shift changes"}:
                continue
            for column_index, schedule_date in dates.items():
                if pd.isna(schedule_date):
                    continue
                value = str(raw.iat[row_index, column_index]).strip() if column_index < raw.shape[1] and not pd.isna(raw.iat[row_index, column_index]) else ""
                if value and normalize_name(value) != normalize_name(staff_name):
                    records.append({"Work Date": schedule_date.date(), "Schedule Name": str(staff_name).strip(), "Name Key": normalize_name(staff_name), "Schedule": value})
    if not records:
        raise ValueError("No dated schedule rows were found in the schedule workbook.")
    return pd.DataFrame(records).drop_duplicates(["Work Date", "Name Key"])


def parse_start_time(schedule_value: str) -> time | None:
    match = re.match(r"^\s*(\d{1,2})(?::(\d{2}))?(?::\d{2})?(?:\s*-|$)", schedule_value)
    if not match:
        return None
    return time(int(match.group(1)) % 24, int(match.group(2) or 0))


def expected_datetime(work_date: date, scheduled_start: time, punches: pd.Series) -> datetime:
    candidates = [datetime.combine(work_date, scheduled_start), datetime.combine(work_date, scheduled_start) + timedelta(hours=12)]
    if punches.empty:
        return candidates[1]
    actual = punches.iloc[0].to_pydatetime()
    return min(candidates, key=lambda candidate: abs((candidate - actual).total_seconds()))


def build_report(punches: pd.DataFrame, schedule: pd.DataFrame, year: int, month: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    selected_dates = pd.date_range(date(year, month, 1), date(year, month, 1) + pd.offsets.MonthEnd(1)).date
    schedule_month = schedule[schedule["Work Date"].isin(selected_dates)].copy()
    employees = punches[["Name No.", "Name", "Name Key"]].drop_duplicates("Name No.")
    rows: list[dict] = []
    for _, employee in employees.iterrows():
        employee_schedule = schedule_month[schedule_month["Name Key"] == employee["Name Key"]]
        if employee_schedule.empty:
            employee_schedule = schedule_month[schedule_month["Schedule Name"].map(normalize_name) == employee["Name Key"]]
        for _, planned in employee_schedule.iterrows():
            day_punches = punches[(punches["Name No."] == employee["Name No."]) & (punches["Work Date"] == planned["Work Date"])]
            schedule_value = planned["Schedule"]
            state = schedule_value.upper()
            first_punch = day_punches.iloc[0]["Date/Time"] if not day_punches.empty else pd.NaT
            start = parse_start_time(schedule_value)
            if start is not None and "-" not in schedule_value:
                schedule_value = f"{start.strftime('%I:%M %p')} shift"
            if state in NON_WORKING_STATES or start is None:
                if state == "OFF":
                    status = "Off / Not scheduled"
                elif state in NON_WORKING_STATES:
                    status = state.title()
                elif start is None:
                    status = "Unrecognized schedule"
                scheduled_display = schedule_value
            elif pd.isna(first_punch):
                status, scheduled_display = "Absent", schedule_value
            else:
                expected = expected_datetime(planned["Work Date"], start, day_punches["Date/Time"])
                status = "On time" if first_punch <= expected + timedelta(minutes=GRACE_MINUTES) else "Late"
                scheduled_display = expected.strftime("%I:%M %p")
            rows.append({"Employee ID": employee["Name No."], "Employee Name": employee["Name"], "Date": planned["Work Date"], "Scheduled": scheduled_display, "First Punch": "" if pd.isna(first_punch) else first_punch.strftime("%I:%M:%S %p"), "Status": status})
    detail = pd.DataFrame(rows).sort_values(["Employee Name", "Date"])
    summary = detail.groupby(["Employee ID", "Employee Name"], as_index=False).agg(
        **{"Working Days": ("Status", lambda values: int(values.isin(["On time", "Late", "Absent"]).sum())), "On Time": ("Status", lambda values: int((values == "On time").sum())), "Late": ("Status", lambda values: int((values == "Late").sum())), "Absent": ("Status", lambda values: int((values == "Absent").sum())), "Off / Not scheduled": ("Status", lambda values: int((values == "Off / Not scheduled").sum()))}
    )
    return detail, summary


def to_excel(detail: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Employee Summary", index=False)
        detail.to_excel(writer, sheet_name="Daily Detail", index=False)
    return output.getvalue()