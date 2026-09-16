"""
attendance_report.py – Attendance report engine (fixed version)
Reads a schedule workbook and a device-export workbook, matches employees
by fuzzy name logic, and produces per-day attendance rows plus a summary.
"""

import re
import math
import datetime as _dt
from datetime import datetime, timedelta, time as _time
from collections import defaultdict
from typing import Dict, List, Optional, Tuple
from openpyxl.styles import Font, PatternFill, Alignment

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

# ── constants ────────────────────────────────────────────────────────────
GRACE_MINUTES = 15

# canonical leave / non-working codes
_LEAVE_ALIASES: Dict[str, str] = {
    "off":                "OFF",
    "wfh":                "WFH",
    "lv":                 "LEAVE",
    "leave":              "LEAVE",
    "resigned":           "RESIGNED",
    "sl":                 "SL",
    "s/l":                "SL",
    "sick leave":         "SL",
    "cl":                 "CL",
    "contingency l":      "CL",
    "al":                 "AL",
    "annual leave":       "AL",
    "annual leaves":      "AL",
    "emergency leave":    "EL",
    "m. leave":           "ML",
    "maternity leave":    "ML",
    "unpaid":             "UNPAID",
    "unpaid leave":       "UNPAID",
    "absent":             "ABSENT",
    "join":               "JOIN",
    "compansated leaves": "COMP",
    "compensated leave":  "COMP",
    "compensated leaves": "COMP",
    "compansated leave":  "COMP",
    "comp off":           "COMP",
    "half day":           "HALF",
    "h/d":                "HALF",
}

_LEAVE_LABELS: Dict[str, str] = {
    "OFF":      "Off",
    "WFH":      "Work From Home",
    "LEAVE":    "Leave",
    "RESIGNED": "Resigned",
    "SL":       "Sick Leave",
    "CL":       "Contingency Leave",
    "AL":       "Annual Leave",
    "EL":       "Emergency Leave",
    "ML":       "Maternity Leave",
    "UNPAID":   "Unpaid Leave",
    "ABSENT":   "Absent",
    "JOIN":     "Joining",
    "COMP":     "Compensated Leave",
    "HALF":     "Half Day",
}

# all codes that mean "not working that day"
NON_WORKING_CODES = set(_LEAVE_LABELS.keys())

# codes that count toward leave totals in summary
_LEAVE_COUNT_CODES = {"SL", "CL", "AL", "EL", "ML", "UNPAID", "ABSENT",
                       "COMP", "LEAVE", "HALF"}

# ── junk-row filtering ──────────────────────────────────────────────────
_JUNK_ROW_NAMES = {
    "finance", "inventory", "surveillance department", "surveillance",
    "automation & it department", "automation", "it department",
    "reporting", "operations department", "operations",
    "leave policy", "working hours", "annual leaves", "annual leave",
    "punctuality", "late comming", "late coming", "shift changes",
    "department", "departments", "staff", "employee", "employees",
    "name", "names", "sr", "sr.", "sr.no", "s.no", "no.",
    "monday", "tuesday", "wednesday", "thursday", "friday",
    "saturday", "sunday",
    "note", "notes", "remarks", "total", "grand total",
}

_DAY_NAMES = {"monday", "tuesday", "wednesday", "thursday", "friday",
              "saturday", "sunday"}


def _is_junk_row(name: str) -> bool:
    """Return True if *name* is a header / policy row, not an employee."""
    low = name.strip().lower()
    if low in _JUNK_ROW_NAMES:
        return True
    if low in _DAY_NAMES:
        return True
    # very long strings are policy text, not names
    if len(low) > 60:
        return True
    # pure numbers
    if low.replace(".", "").replace("-", "").isdigit():
        return True
    return False


# ── time regexes ─────────────────────────────────────────────────────────
# requires colon – e.g. "9:00", "21:30", "10:00:00"
_TIME_RE = re.compile(
    r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$"
)
# bare hour only if >= 6 (avoids stray "1","2","3" artefacts)
_BARE_HOUR_RE = re.compile(r"^\s*(\d{1,2})\s*$")

# range like "9:00 - 5:00" or "9:00-17:00"
_RANGE_RE = re.compile(
    r"^\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*$"
)

# date-like string to skip (openpyxl sometimes returns datetime as str)
_DATE_STR_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")


# ── workbook I/O ─────────────────────────────────────────────────────────

def read_workbook(path: str) -> pd.ExcelFile:
    """Open .xls or .xlsx transparently."""
    if str(path).lower().endswith(".xls"):
        return pd.ExcelFile(path, engine="xlrd")
    return pd.ExcelFile(path, engine="openpyxl")


def normalize_name(raw: str) -> str:
    """Lower-case, collapse whitespace, strip punctuation."""
    s = str(raw).lower().strip()
    s = re.sub(r"[._]+", " ", s)       # dots and underscores → space
    s = re.sub(r"\s+", " ", s)         # collapse whitespace
    s = re.sub(r"[^a-z0-9 ]", "", s)   # drop remaining punctuation
    return s.strip()


def _name_tokens(norm: str) -> set:
    """Return the set of word-tokens from a normalized name."""
    return set(norm.split())


# ── fuzzy name matching ──────────────────────────────────────────────────

def _edit_distance(a: str, b: str) -> int:
    """Simple Levenshtein distance (enough for short tokens)."""
    la, lb = len(a), len(b)
    if la == 0:
        return lb
    if lb == 0:
        return la
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        curr = [i] + [0] * lb
        for j in range(1, lb + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            curr[j] = min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost)
        prev = curr
    return prev[lb]


def _close_substring(needle: str, haystack: str, max_dist: int = 1) -> bool:
    """True if *needle* appears inside *haystack* with ≤ *max_dist* edits."""
    ln, lh = len(needle), len(haystack)
    if ln > lh:
        return False
    for start in range(lh - ln + 1):
        if _edit_distance(needle, haystack[start:start + ln]) <= max_dist:
            return True
    return False


def _build_name_map(sched_names: List[str],
                    punch_names: List[str]) -> Dict[str, str]:
    """
    Map each *sched_name* → best matching *punch_name*.

    Strategy (in priority order):
      1. Exact normalised match
      2. Schedule name is a substring of an attendance name
      3. All word-tokens of the shorter name appear in the longer
      4. Reverse substring (attendance key inside schedule key)
      5. Fuzzy substring with edit-distance ≤ 1
    Returns a dict {normalised_sched_name: normalised_punch_name}.
    """
    norm_sched = {normalize_name(n): n for n in sched_names}
    norm_punch = {normalize_name(n): n for n in punch_names}
    mapping: Dict[str, str] = {}

    unmatched_sched = set(norm_sched.keys())
    used_punch = set()

    # pass 1 – exact
    for ns in list(unmatched_sched):
        if ns in norm_punch and ns not in used_punch:
            mapping[ns] = ns
            unmatched_sched.discard(ns)
            used_punch.add(ns)

    # pass 2 – schedule key is substring of punch key
    for ns in list(unmatched_sched):
        if len(ns) < 3:
            continue
        for np in sorted(norm_punch.keys()):
            if np in used_punch:
                continue
            if ns in np:
                mapping[ns] = np
                unmatched_sched.discard(ns)
                used_punch.add(np)
                break

    # pass 3 – word-token containment (all tokens of shorter in longer)
    for ns in list(unmatched_sched):
        stok = _name_tokens(ns)
        if len(stok) < 1:
            continue
        best = None
        best_extra = 999
        for np in sorted(norm_punch.keys()):
            if np in used_punch:
                continue
            ptok = _name_tokens(np)
            if stok <= ptok:                 # sched tokens ⊆ punch tokens
                extra = len(ptok) - len(stok)
                if extra < best_extra:
                    best = np
                    best_extra = extra
            elif ptok <= stok:               # reverse
                extra = len(stok) - len(ptok)
                if extra < best_extra:
                    best = np
                    best_extra = extra
        if best is not None:
            mapping[ns] = best
            unmatched_sched.discard(ns)
            used_punch.add(best)

    # pass 4 – reverse substring
    for ns in list(unmatched_sched):
        for np in sorted(norm_punch.keys()):
            if np in used_punch:
                continue
            if len(np) >= 3 and np in ns:
                mapping[ns] = np
                unmatched_sched.discard(ns)
                used_punch.add(np)
                break

    # pass 5 – fuzzy substring (edit dist ≤ 1)
    for ns in list(unmatched_sched):
        stok = ns.split()
        for np in sorted(norm_punch.keys()):
            if np in used_punch:
                continue
            ptok = np.split()
            # try matching each schedule token fuzzily against punch tokens
            matched_tokens = 0
            for st in stok:
                if len(st) < 3:
                    continue
                for pt in ptok:
                    if _close_substring(st, pt, max_dist=1) or \
                       _close_substring(pt, st, max_dist=1):
                        matched_tokens += 1
                        break
            meaningful_sched = [t for t in stok if len(t) >= 3]
            if meaningful_sched and matched_tokens >= len(meaningful_sched):
                mapping[ns] = np
                unmatched_sched.discard(ns)
                used_punch.add(np)
                break

    return mapping


# ── schedule cell helpers ────────────────────────────────────────────────

def _cell_to_schedule_str(val) -> Optional[str]:
    """
    Convert an openpyxl cell value to a schedule-string we can parse.
    Returns None if the cell should be skipped entirely.
    """
    if val is None:
        return None
    # datetime.time objects from openpyxl (time-formatted cells)
    if isinstance(val, _time):
        return val.strftime("%H:%M")
    # datetime / Timestamp objects – likely dates bleeding from header row
    if isinstance(val, (datetime, _dt.date, pd.Timestamp)):
        return None
    s = str(val).strip()
    if not s:
        return None
    # skip date-like strings
    if _DATE_STR_RE.match(s):
        return None
    return s


def _classify_schedule(raw: str) -> Tuple[str, Optional[str]]:
    """
    Given a schedule-cell string, return (kind, detail).
      kind = "time"  → detail is "HH:MM" 24-h start time
      kind = "range" → detail is "HH:MM" (the start portion)
      kind = "leave" → detail is canonical code (OFF, SL, etc.)
      kind = "unknown" → detail is None
    """
    low = raw.strip().lower()
    # check leave codes first
    if low in _LEAVE_ALIASES:
        return ("leave", _LEAVE_ALIASES[low])

    # range "9:00 - 17:00"
    m = _RANGE_RE.match(raw.strip())
    if m:
        return ("range", m.group(1))

    # explicit time with colon
    m = _TIME_RE.match(raw.strip())
    if m:
        h, mn = int(m.group(1)), int(m.group(2))
        return ("time", f"{h}:{mn:02d}")

    # bare hour >= 6
    m = _BARE_HOUR_RE.match(raw.strip())
    if m:
        h = int(m.group(1))
        if 6 <= h <= 23:
            return ("time", f"{h}:00")
        return ("unknown", None)

    return ("unknown", None)


# ── punch loading ────────────────────────────────────────────────────────

def load_punches(xf: pd.ExcelFile) -> Dict[str, Dict[str, List[datetime]]]:
    """
    Return {normalised_name: {date_str: [punch_datetimes]}}.
    """
    df = xf.parse(xf.sheet_names[0])
    df.columns = [str(c).strip() for c in df.columns]

    # find the datetime column
    dt_col = None
    for c in df.columns:
        if "date" in c.lower() and "time" in c.lower():
            dt_col = c
            break
    if dt_col is None:
        for c in df.columns:
            if "date" in c.lower():
                dt_col = c
                break
    if dt_col is None:
        raise ValueError("Cannot find a Date/Time column in device data.")

    name_col = None
    for c in df.columns:
        if c.lower().strip() == "name":
            name_col = c
            break
    if name_col is None:
        raise ValueError("Cannot find a 'Name' column in device data.")

    punches: Dict[str, Dict[str, List[datetime]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for _, row in df.iterrows():
        raw_name = row.get(name_col)
        raw_dt = row.get(dt_col)
        if pd.isna(raw_name) or pd.isna(raw_dt):
            continue
        name = normalize_name(str(raw_name))
        if not name:
            continue
        try:
            dt = pd.to_datetime(raw_dt, dayfirst=True)
        except Exception:
            continue
        date_str = dt.strftime("%Y-%m-%d")
        punches[name][date_str].append(dt.to_pydatetime())

    # sort each day's punches
    for name in punches:
        for ds in punches[name]:
            punches[name][ds].sort()

    return dict(punches)


# ── schedule loading ─────────────────────────────────────────────────────

def load_schedule(xf: pd.ExcelFile,
                  year: int,
                  month: int) -> Dict[str, Dict[str, str]]:
    """
    Return {normalised_name: {date_str: schedule_value_str}}.
    """
    target = datetime(year, month, 1)
    best_sheet = None
    best_diff = None

    for sn in xf.sheet_names:
        try:
            sheet_date = pd.to_datetime(sn, format="%b %Y")
        except Exception:
            try:
                sheet_date = pd.to_datetime(sn)
            except Exception:
                continue
        diff = abs((sheet_date.year - target.year) * 12 +
                   sheet_date.month - target.month)
        if best_diff is None or diff < best_diff:
            best_diff = diff
            best_sheet = sn

    if best_sheet is None:
        raise ValueError(f"No sheet found for {target:%B %Y}")

    df = xf.parse(best_sheet, header=None)

    # row 0 = dates, row 1 = day-of-week / department header
    date_row = df.iloc[0]

    # detect first_date_col: find where actual dates start
    # Column 0 is the name column; sometimes col 0 also has a date (duplicate)
    first_date_col = 1
    for ci in range(len(date_row)):
        val = date_row.iloc[ci]
        try:
            pd.to_datetime(val)
            first_date_col = ci
            break
        except Exception:
            continue
    # If col 0 has the same date as col 1, skip it (it's the name col)
    if first_date_col == 0 and len(date_row) > 1:
        try:
            d0 = pd.to_datetime(date_row.iloc[0])
            d1 = pd.to_datetime(date_row.iloc[1])
            if d0 == d1:
                first_date_col = 1
        except Exception:
            first_date_col = 1

    # map column index → date string
    col_dates: Dict[int, str] = {}
    for ci in range(first_date_col, len(date_row)):
        val = date_row.iloc[ci]
        try:
            dt = pd.to_datetime(val)
            if dt.month == month and dt.year == year:
                col_dates[ci] = dt.strftime("%Y-%m-%d")
        except Exception:
            continue

    if not col_dates:
        raise ValueError(f"No date columns found for {target:%B %Y}")

    schedule: Dict[str, Dict[str, str]] = defaultdict(dict)

    for ri in range(2, len(df)):
        raw_name = df.iloc[ri, 0]
        if pd.isna(raw_name):
            continue
        name_str = str(raw_name).strip()
        if not name_str:
            continue
        if _is_junk_row(name_str):
            continue

        norm = normalize_name(name_str)
        if not norm:
            continue

        for ci, date_str in col_dates.items():
            cell = df.iloc[ri, ci]
            sval = _cell_to_schedule_str(cell)
            if sval is None:
                continue
            kind, detail = _classify_schedule(sval)
            if kind == "unknown":
                continue
            # store raw schedule string
            if kind == "leave":
                schedule[norm][date_str] = detail   # canonical code
            else:
                schedule[norm][date_str] = sval      # time string

    return dict(schedule)


# ── time helpers ─────────────────────────────────────────────────────────

def parse_start_time(sched_val: str) -> Optional[_time]:
    """Parse a schedule value into a time-of-day for the shift start."""
    if not sched_val:
        return None

    # range "9:00 - 17:00"
    m = _RANGE_RE.match(sched_val.strip())
    if m:
        sched_val = m.group(1)

    # HH:MM
    m = _TIME_RE.match(sched_val.strip())
    if m:
        h, mn = int(m.group(1)), int(m.group(2))
        if 0 <= h <= 23 and 0 <= mn <= 59:
            return _time(h, mn)
        return None

    # bare hour
    m = _BARE_HOUR_RE.match(sched_val.strip())
    if m:
        h = int(m.group(1))
        if 6 <= h <= 23:
            return _time(h, 0)

    return None


def expected_datetime(shift_time: _time,
                      date_str: str,
                      punch: Optional[datetime] = None) -> datetime:
    """
    Return the expected clock-in datetime for a shift, resolving AM/PM
    ambiguity using the actual punch when available.
    """
    base_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    candidate = datetime.combine(base_date, shift_time)

    if shift_time.hour >= 12:
        # unambiguous PM/night
        return candidate

    # AM shift by default; but if hour < 12, also consider +12h (PM)
    candidate_pm = candidate + timedelta(hours=12)

    if punch is None:
        # no punch → assume the AM reading
        return candidate

    diff_am = abs((punch - candidate).total_seconds())
    diff_pm = abs((punch - candidate_pm).total_seconds())

    return candidate if diff_am <= diff_pm else candidate_pm


# ── build the report ─────────────────────────────────────────────────────

def _is_badge_id(name: str) -> bool:
    """Return True if *name* is purely numeric (a device badge ID, not a real name)."""
    return name.replace(" ", "").isdigit()


def build_report(punch_data: Dict[str, Dict[str, List[datetime]]],
                 sched_data: Dict[str, Dict[str, str]],
                 year: int,
                 month: int) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    ##Return (daily_df, summary_df, badge_df).
    Return (daily_df, summary_df, badge_df, comparison_df).

    badge_df contains punch records for employees whose device only
    recorded a numeric badge ID instead of a name.  They get their own
    sheet so they can be identified and cross-referenced.
    """
    # Separate badge-ID-only punches from named punches
    named_punch_data: Dict[str, Dict[str, List[datetime]]] = {}
    badge_punch_data: Dict[str, Dict[str, List[datetime]]] = {}
    for pname, pdays in punch_data.items():
        if _is_badge_id(pname):
            badge_punch_data[pname] = pdays
        else:
            named_punch_data[pname] = pdays

    # Build the name mapping (only for named employees)
    sched_names = list(sched_data.keys())
    punch_names = list(named_punch_data.keys())
    name_map = _build_name_map(sched_names, punch_names)


    # Build comparison of names between device data and schedule data.
    # name_map maps schedule names to matching device names.
    matched_punch_names = set(name_map.values())
    matched_schedule_names = set(name_map.keys())

    comparison_rows = []

    # Names or badge IDs found in device data but not matched to schedule.
    for device_name in sorted(punch_data):
        if device_name not in matched_punch_names:
            comparison_rows.append({
                "Category": "Device only",
                "Name or ID": device_name.replace("_", " ").title(),
                "Normalized Key": device_name,
                "Details": "Found in device attendance but not matched in schedule",
            })

    # Names found in schedule data but not matched to device data.
    for schedule_name in sorted(sched_data):
        if schedule_name not in matched_schedule_names:
            comparison_rows.append({
                "Category": "Schedule only",
                "Name or ID": schedule_name.replace("_", " ").title(),
                "Normalized Key": schedule_name,
                "Details": "Found in schedule but not matched in device attendance",
            })

    comparison_df = pd.DataFrame(
        comparison_rows,
        columns=[
            "Category",
            "Name or ID",
            "Normalized Key",
            "Details",
        ],
    )

    if not comparison_df.empty:
        comparison_df = comparison_df.sort_values(
            ["Category", "Name or ID"]
        ).reset_index(drop=True)



    # All days in the month
    first_day = _dt.date(year, month, 1)
    if month == 12:
        last_day = _dt.date(year + 1, 1, 1) - _dt.timedelta(days=1)
    else:
        last_day = _dt.date(year, month + 1, 1) - _dt.timedelta(days=1)
    num_days = last_day.day
    all_dates = [
        (first_day + _dt.timedelta(days=d)).strftime("%Y-%m-%d")
        for d in range(num_days)
    ]

    # Gather all NAMED employees (union of schedule + named punches)
    all_employees = set(sched_data.keys())
    # Also add punch-only named employees not matched
    matched_punch_names = set(name_map.values())
    for pn in named_punch_data:
        if pn not in matched_punch_names:
            all_employees.add(pn)

    daily_rows = []
    summary_rows = []

    for emp in sorted(all_employees):
        # resolve which punch key to use
        punch_key = name_map.get(emp, emp)
        emp_punches = named_punch_data.get(punch_key, {})
        emp_schedule = sched_data.get(emp, {})

        # If employee has neither schedule nor punches, skip
        if not emp_schedule and not emp_punches:
            continue

        # determine display name: prefer the original-cased version
        display_name = emp.replace("_", " ").title()

        # summary accumulators
        present_days = 0
        late_days = 0
        on_time_days = 0
        absent_days = 0
        wfh_days = 0
        leave_days = 0
        leave_detail: Dict[str, int] = defaultdict(int)  # code → count

        for date_str in all_dates:
            sched_val = emp_schedule.get(date_str)
            day_punches = emp_punches.get(date_str, [])

            # ── no schedule for this employee at all ──
            if not emp_schedule:
                if not day_punches:
                    continue  # skip days with no data at all
                # has punches but no schedule
                first_in = day_punches[0].strftime("%H:%M")
                last_out = day_punches[-1].strftime("%H:%M") if len(day_punches) > 1 else ""
                daily_rows.append({
                    "Employee": display_name,
                    "Date": date_str,
                    "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                    "Scheduled": "No schedule",
                    "Status": "Present (no schedule)",
                    "Clock In": first_in,
                    "Clock Out": last_out,
                    "Punch Count": len(day_punches),
                    "Remarks": "No schedule found",
                })
                present_days += 1
                continue

            # ── has schedule ──
            if sched_val is None:
                # no entry for this date in schedule → skip (not scheduled)
                if day_punches:
                    first_in = day_punches[0].strftime("%H:%M")
                    last_out = day_punches[-1].strftime("%H:%M") if len(day_punches) > 1 else ""
                    daily_rows.append({
                        "Employee": display_name,
                        "Date": date_str,
                        "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                        "Scheduled": "—",
                        "Status": "Present (unscheduled day)",
                        "Clock In": first_in,
                        "Clock Out": last_out,
                        "Punch Count": len(day_punches),
                        "Remarks": "",
                    })
                    present_days += 1
                continue

            # check if it's a leave / non-working code
            if sched_val in NON_WORKING_CODES:
                code = sched_val
                label = _LEAVE_LABELS.get(code, code)
                if code == "WFH":
                    wfh_days += 1
                elif code == "OFF":
                    pass  # don't count OFF as leave
                elif code == "RESIGNED":
                    pass
                else:
                    leave_days += 1
                    leave_detail[code] += 1

                clock_in = ""
                clock_out = ""
                if day_punches:
                    clock_in = day_punches[0].strftime("%H:%M")
                    if len(day_punches) > 1:
                        clock_out = day_punches[-1].strftime("%H:%M")

                daily_rows.append({
                    "Employee": display_name,
                    "Date": date_str,
                    "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                    "Scheduled": label,
                    "Status": label,
                    "Clock In": clock_in,
                    "Clock Out": clock_out,
                    "Punch Count": len(day_punches),
                    "Remarks": f"Punched despite {label}" if day_punches else "",
                })
                continue

            # ── it's a working shift ──
            shift_start = parse_start_time(sched_val)
            if shift_start is None:
                # can't parse → just record raw
                daily_rows.append({
                    "Employee": display_name,
                    "Date": date_str,
                    "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                    "Scheduled": sched_val,
                    "Status": "Schedule parse error",
                    "Clock In": "",
                    "Clock Out": "",
                    "Punch Count": len(day_punches),
                    "Remarks": f"Could not parse: {sched_val}",
                })
                continue

            sched_display = shift_start.strftime("%H:%M")

            if not day_punches:
                # absent
                absent_days += 1
                daily_rows.append({
                    "Employee": display_name,
                    "Date": date_str,
                    "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                    "Scheduled": sched_display,
                    "Status": "Absent",
                    "Clock In": "",
                    "Clock Out": "",
                    "Punch Count": 0,
                    "Remarks": "",
                })
                continue

            # has punches
            first_punch = day_punches[0]
            last_punch = day_punches[-1] if len(day_punches) > 1 else None

            exp_dt = expected_datetime(shift_start, date_str, first_punch)
            # diff_minutes = (first_punch - exp_dt).total_seconds() / 60.0

            # grace = timedelta(minutes=GRACE_MINUTES)
            # if first_punch <= exp_dt + grace:
            #     status = "On Time"
            #     on_time_days += 1
            # else:
            #     late_min = math.ceil(diff_minutes)
            #     status = f"Late ({late_min} min)"
            #     late_days += 1
            scheduled_minute = exp_dt.replace(second=0, microsecond=0)
            arrival_minute = first_punch.replace(second=0, microsecond=0)

            diff_minutes = int((arrival_minute - scheduled_minute).total_seconds() / 60)

            if diff_minutes <= GRACE_MINUTES:
                status = "On Time"
                on_time_days += 1
            else:
                late_min = diff_minutes
                status = f"Late ({late_min} min)"
                late_days += 1



            present_days += 1

            daily_rows.append({
                "Employee": display_name,
                "Date": date_str,
                "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                "Scheduled": sched_display,
                "Status": status,
                "Clock In": first_punch.strftime("%H:%M"),
                "Clock Out": last_punch.strftime("%H:%M") if last_punch else "",
                "Punch Count": len(day_punches),
                "Remarks": "",
            })

        # ── summary row ──
        leave_breakdown = ", ".join(
            f"{_LEAVE_LABELS.get(c, c)}: {n}" for c, n in sorted(leave_detail.items())
        )
        summary_rows.append({
            "Employee": display_name,
            "Present": present_days,
            "On Time": on_time_days,
            "Late": late_days,
            "Absent": absent_days,
            "WFH": wfh_days,
            "Leave": leave_days,
            "Leave Breakdown": leave_breakdown,
        })

    daily_df = pd.DataFrame(daily_rows)
    summary_df = pd.DataFrame(summary_rows)

    # sort
    if not daily_df.empty:
        daily_df = daily_df.sort_values(["Employee", "Date"]).reset_index(drop=True)
    if not summary_df.empty:
        summary_df = summary_df.sort_values("Employee").reset_index(drop=True)

    # ── Badge-ID punch sheet ──
    badge_rows = []
    for badge_id in sorted(badge_punch_data.keys(), key=lambda x: int(x) if x.isdigit() else 0):
        days = badge_punch_data[badge_id]
        for date_str in sorted(days.keys()):
            # only include dates within the target month
            try:
                d = datetime.strptime(date_str, "%Y-%m-%d").date()
                if d.month != month or d.year != year:
                    continue
            except Exception:
                continue

            day_punches = days[date_str]
            first_in = day_punches[0].strftime("%H:%M") if day_punches else ""
            last_out = day_punches[-1].strftime("%H:%M") if len(day_punches) > 1 else ""
            all_times = ", ".join(p.strftime("%H:%M:%S") for p in day_punches)

            badge_rows.append({
                "Badge ID": badge_id,
                "Date": date_str,
                "Day": datetime.strptime(date_str, "%Y-%m-%d").strftime("%A"),
                "Clock In": first_in,
                "Clock Out": last_out,
                "Punch Count": len(day_punches),
                "All Punch Times": all_times,
                "Remarks": "Name not registered on device",
            })

    badge_df = pd.DataFrame(badge_rows)
    if not badge_df.empty:
        badge_df = badge_df.sort_values(["Badge ID", "Date"]).reset_index(drop=True)

    #return daily_df, summary_df, badge_df
    return daily_df, summary_df, badge_df, comparison_df


# ── Excel export ─────────────────────────────────────────────────────────

# def to_excel(daily_df: pd.DataFrame,
#              summary_df: pd.DataFrame,
#              path: str,
#              badge_df: Optional[pd.DataFrame] = None) -> None:
#     """Write DataFrames to a single .xlsx with separate sheets."""
#     with pd.ExcelWriter(path, engine="openpyxl") as writer:
#         summary_df.to_excel(writer, sheet_name="Employee Summary", index=False)
#         daily_df.to_excel(writer, sheet_name="Daily Detail", index=False)
#         if badge_df is not None and not badge_df.empty:
#             badge_df.to_excel(writer, sheet_name="Badge ID Punches", index=False)

def to_excel(daily_df: pd.DataFrame,
             summary_df: pd.DataFrame,
             path: str,
             badge_df: Optional[pd.DataFrame] = None,
             comparison_df: Optional[pd.DataFrame] = None) -> None:
    """Write all report DataFrames to separate Excel worksheets."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary_df.to_excel(
            writer,
            sheet_name="Employee Summary",
            index=False,
        )

        daily_df.to_excel(
            writer,
            sheet_name="Daily Detail",
            index=False,
        )

        if badge_df is not None and not badge_df.empty:
            badge_df.to_excel(
                writer,
                sheet_name="Badge ID Punches",
                index=False,
            )

        if comparison_df is not None:
            comparison_df.to_excel(
                writer,
                sheet_name="Name Comparison",
                index=False,
            )

        header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        tab_colors = {
            "Employee Summary": "70AD47",
            "Daily Detail": "5B9BD5",
            "Badge ID Punches": "ED7D31",
            "Name Comparison": "A5A5A5",
        }

        for worksheet in writer.book.worksheets:
            worksheet.sheet_properties.tabColor = tab_colors.get(
                worksheet.title,
                "5B9BD5",
            )
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
            worksheet.row_dimensions[1].height = 28
            worksheet.freeze_panes = "A2"
            if worksheet.max_column > 0 and worksheet.max_row > 1:
                worksheet.auto_filter.ref = worksheet.dimensions
