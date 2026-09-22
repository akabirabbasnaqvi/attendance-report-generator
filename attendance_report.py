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

# ── employee-ID matching ────────────────────────────────────────────────
# The device export has a "No." column (each employee's device enrollment
# number) and the schedule workbook can have an "Employees ID" column with
# the same number. When both are present, matching by this ID is exact and
# always wins over fuzzy name matching, which is what previously produced
# errors whenever the schedule used a short/nickname and the device used
# the full legal name (or vice versa).
#
# load_punches()/load_schedule() populate the module-level maps below as a
# side effect of parsing each workbook; build_report() reads them to build
# ID-based overrides before falling back to fuzzy name matching. This keeps
# the public function signatures unchanged, so nothing else needs to change
# to call them.
_LAST_PUNCH_ID_TO_KEY: Dict[str, str] = {}   # normalised device ID -> punch dict key
_LAST_SCHEDULE_ID_MAP: Dict[str, str] = {}   # normalised schedule name -> normalised ID

_DEVICE_ID_COL_ALIASES = {
    "no", "employee no", "emp no", "badge no", "badge id",
    "employee id", "emp id", "id no", "id number", "staff no", "staff id",
}

_SCHEDULE_ID_HEADER_ALIASES = {
    "employees id", "employee id", "emp id", "emp id no",
    "staff id", "employee no", "emp no", "badge id", "badge no",
    "id no", "id number", "employee id no", "id",
}


def _normalize_id(val) -> Optional[str]:
    """Normalize a raw employee-ID cell value (int/float/str) to a clean,
    comparable string key. Returns None if the cell holds no usable ID."""
    if val is None:
        return None
    if isinstance(val, float):
        if val != val:  # NaN
            return None
        val = int(val) if val == int(val) else val
    if isinstance(val, int):
        return str(val)
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    # collapse a float-like string such as "45.0" -> "45"
    m = re.match(r"^(\d+)\.0+$", s)
    if m:
        s = m.group(1)
    # strip leading zeros from pure numeric IDs ("003" -> "3") for comparison
    if s.isdigit():
        s = str(int(s))
    return s.lower()


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
        xf = pd.ExcelFile(path, engine="xlrd")
    else:
        xf = pd.ExcelFile(path, engine="openpyxl")
    # Stash the source path on the object so downstream readers (e.g.
    # load_schedule's merged-cell handling) can re-open the workbook in a
    # non-read-only mode without changing this function's signature or
    # touching any caller (pd.ExcelFile allows arbitrary extra attributes).
    try:
        xf._source_path = str(path)
    except Exception:
        pass
    return xf


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
                    punch_names: List[str],
                    id_overrides: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    """
    Map each *sched_name* → best matching *punch_name*.

    Strategy (in priority order):
      1. Exact normalised match
      2. Exact employee-ID match (see _LAST_SCHEDULE_ID_MAP /
         _LAST_PUNCH_ID_TO_KEY), passed in via *id_overrides* — this runs
         AFTER the exact name match on purpose: a literal name match is
         unambiguous, while an ID column can have a data-entry typo (e.g.
         two different employee IDs transposed), so a disagreeing ID match
         should never override an exact name match. ID matching still
         resolves the common case an exact match can't: a schedule
         nickname/short-name with no literal match in the device export.
      3. Schedule name is a substring of an attendance name
      4. All word-tokens of the shorter name appear in the longer
      5. Reverse substring (attendance key inside schedule key)
      6. Fuzzy substring with edit-distance ≤ 1
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

    # pass 2 – exact employee-ID match (see docstring: runs after exact
    # name matching so a typo'd ID can never override an unambiguous name)
    if id_overrides:
        for ns, punch_key in id_overrides.items():
            if ns in unmatched_sched and punch_key in norm_punch and punch_key not in used_punch:
                mapping[ns] = punch_key
                unmatched_sched.discard(ns)
                used_punch.add(punch_key)

    # pass 3 – schedule key is substring of punch key
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

    # pass 4 – word-token containment (all tokens of shorter in longer)
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

    # pass 5 – reverse substring
    for ns in list(unmatched_sched):
        for np in sorted(norm_punch.keys()):
            if np in used_punch:
                continue
            if len(np) >= 3 and np in ns:
                mapping[ns] = np
                unmatched_sched.discard(ns)
                used_punch.add(np)
                break

    # pass 6 – fuzzy substring (edit dist ≤ 1)
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

    As a side effect, also populates the module-level _LAST_PUNCH_ID_TO_KEY
    map (normalised device employee-ID -> normalised name key) whenever the
    device workbook has an employee-ID column (e.g. "No."), so build_report()
    can match schedule rows to device punches by ID instead of by name.
    """
    global _LAST_PUNCH_ID_TO_KEY
    _LAST_PUNCH_ID_TO_KEY = {}

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

    # find the device employee-ID column (e.g. "No.")
    id_col = None
    for c in df.columns:
        key = c.strip().lower().rstrip(".")
        if key in _DEVICE_ID_COL_ALIASES:
            id_col = c
            break

    punches: Dict[str, Dict[str, List[datetime]]] = defaultdict(
        lambda: defaultdict(list)
    )
    name_to_id: Dict[str, str] = {}

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

        if id_col is not None and name not in name_to_id:
            id_val = _normalize_id(row.get(id_col))
            if id_val:
                name_to_id[name] = id_val

    # sort each day's punches
    for name in punches:
        for ds in punches[name]:
            punches[name][ds].sort()

    # build the ID -> punch-key alias map used by build_report()
    for name, id_val in name_to_id.items():
        _LAST_PUNCH_ID_TO_KEY.setdefault(id_val, name)

    return dict(punches)


# ── schedule loading ─────────────────────────────────────────────────────

def _fill_merged_cells(df: "pd.DataFrame", xf: pd.ExcelFile, sheet_name: str) -> None:
    """
    Forward-fill merged-cell values across their full span, in place, on the
    freshly-parsed (header=None) DataFrame `df`.

    Excel only stores a merged range's value in its top-left cell; every
    other cell in the range reads back as blank/None through pandas and
    openpyxl alike. A multi-day leave block (e.g. one "Annual Leaves" cell
    merged across 8 date columns) therefore only ever shows up on the first
    day unless we copy that top-left value into the rest of the range
    ourselves. This must run before any date/name/id parsing below.

    Reading `xf.book` directly does not work here because pandas opens the
    openpyxl workbook in read-only mode for `ExcelFile.parse()`, and
    read-only worksheets don't expose `.merged_cells`. So we re-open the
    same workbook file from disk (not read-only) just to read the merge
    geometry, using the path `read_workbook()` stashed on `xf`.
    """
    source_path = getattr(xf, "_source_path", None)
    if not source_path:
        return

    try:
        if str(source_path).lower().endswith(".xls"):
            # Legacy .xls via xlrd: merged_cells is a list of
            # (row_lo, row_hi, col_lo, col_hi) 0-indexed, half-open ranges —
            # already aligned with the 0-indexed, header=None DataFrame.
            book = xf.book  # xlrd.Book (not read-only-restricted)
            sheet = book.sheet_by_name(sheet_name)
            ranges = [
                (r0, r1 - 1, c0, c1 - 1)
                for (r0, r1, c0, c1) in getattr(sheet, "merged_cells", [])
            ]
        else:
            import openpyxl
            wb2 = openpyxl.load_workbook(source_path, data_only=True, read_only=False)
            if sheet_name not in wb2.sheetnames:
                return
            ws2 = wb2[sheet_name]
            # openpyxl ranges are 1-indexed and inclusive; convert to the
            # 0-indexed positions used by the header=None DataFrame.
            ranges = [
                (mr.min_row - 1, mr.max_row - 1, mr.min_col - 1, mr.max_col - 1)
                for mr in ws2.merged_cells.ranges
            ]

        n_rows, n_cols = df.shape
        for r0, r1, c0, c1 in ranges:
            if r0 < 0 or c0 < 0 or r0 >= n_rows or c0 >= n_cols:
                continue
            top_val = df.iat[r0, c0]
            if top_val is None or (isinstance(top_val, float) and top_val != top_val):
                continue
            r_end = min(r1, n_rows - 1)
            c_end = min(c1, n_cols - 1)
            for ri in range(r0, r_end + 1):
                for ci in range(c0, c_end + 1):
                    if ri == r0 and ci == c0:
                        continue
                    df.iat[ri, ci] = top_val
    except Exception:
        # Merge-fill is a best-effort enhancement; never let it break
        # report generation if a workbook can't be re-opened this way.
        return


def load_schedule(xf: pd.ExcelFile,
                  year: int,
                  month: int) -> Dict[str, Dict[str, str]]:
    """
    Return {normalised_name: {date_str: schedule_value_str}}.

    As a side effect, also populates the module-level _LAST_SCHEDULE_ID_MAP
    map (normalised schedule name -> normalised employee ID) whenever the
    schedule sheet has an employee-ID column (e.g. "Employees ID"), so
    build_report() can match schedule rows to device punches by ID instead
    of by name.
    """
    global _LAST_SCHEDULE_ID_MAP
    _LAST_SCHEDULE_ID_MAP = {}

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
    _fill_merged_cells(df, xf, best_sheet)

    # row 0 = dates, row 1 = day-of-week / department header
    date_row = df.iloc[0]

    # detect first_date_col: find where actual dates start
    first_date_col = 1
    for ci in range(len(date_row)):
        val = date_row.iloc[ci]
        try:
            pd.to_datetime(val)
            first_date_col = ci
            break
        except Exception:
            continue

    # Some schedule templates have a stray duplicate date in the label
    # column immediately before the real first date column (a merged-cell
    # rendering artifact). This label column holds the employee name (or,
    # when an "Employees ID" column has been inserted before it, the name
    # sits one column further right) — either way it is not real shift
    # data, so if the detected first date column repeats in the very next
    # column, skip the first one.
    if first_date_col + 1 < len(date_row):
        try:
            d0 = pd.to_datetime(date_row.iloc[first_date_col])
            d1 = pd.to_datetime(date_row.iloc[first_date_col + 1])
            if d0 == d1:
                first_date_col += 1
        except Exception:
            pass

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

    # detect an employee-ID column (e.g. "Employees ID"), checked against the
    # two header rows across any column that isn't itself a date column
    id_col = None
    header_rows = [df.iloc[0]]
    if len(df) > 1:
        header_rows.append(df.iloc[1])
    for hdr_row in header_rows:
        for ci in range(len(hdr_row)):
            if ci in col_dates:
                continue
            val = hdr_row.iloc[ci]
            if val is None:
                continue
            if isinstance(val, float) and val != val:  # NaN
                continue
            text = str(val).strip().lower().rstrip(".")
            text = re.sub(r"\s+", " ", text)
            if text in _SCHEDULE_ID_HEADER_ALIASES:
                id_col = ci
                break
        if id_col is not None:
            break

    # The employee name is normally in column 0. If an "Employees ID"
    # column was inserted right at column 0 (pushing the name one column
    # over, as some updated templates do), read the name from column 1
    # instead.
    name_col = 1 if id_col == 0 else 0

    schedule: Dict[str, Dict[str, str]] = defaultdict(dict)

    for ri in range(2, len(df)):
        raw_name = df.iloc[ri, name_col]
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

        if id_col is not None:
            id_val = _normalize_id(df.iloc[ri, id_col])
            if id_val:
                _LAST_SCHEDULE_ID_MAP[norm] = id_val

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


def _to_date(val) -> _dt.date:
    """Coerce a date, datetime, or 'YYYY-MM-DD' string into a date object."""
    if isinstance(val, _dt.datetime):
        return val.date()
    if isinstance(val, _dt.date):
        return val
    return datetime.strptime(str(val), "%Y-%m-%d").date()


def load_schedule_range(xf: pd.ExcelFile,
                        start_date,
                        end_date) -> Dict[str, Dict[str, str]]:
    """
    Like load_schedule(), but for an arbitrary [start_date, end_date] span
    instead of a single calendar month.

    Schedule workbooks store one sheet per month, so this calls
    load_schedule() once for every distinct (year, month) the range
    touches and merges the results, trimming each employee's day-map down
    to just the requested dates. A month with no matching sheet in the
    workbook is skipped rather than failing the whole range, unless the
    range matches no sheet at all.

    Also merges _LAST_SCHEDULE_ID_MAP across every sheet touched, so the
    ID-based matching in build_report() still sees the full picture even
    when the report spans more than one schedule sheet.
    """
    global _LAST_SCHEDULE_ID_MAP

    start_d = _to_date(start_date)
    end_d = _to_date(end_date)
    if end_d < start_d:
        start_d, end_d = end_d, start_d
    start_str = start_d.strftime("%Y-%m-%d")
    end_str = end_d.strftime("%Y-%m-%d")

    months = []
    cursor = _dt.date(start_d.year, start_d.month, 1)
    while cursor <= end_d:
        months.append((cursor.year, cursor.month))
        if cursor.month == 12:
            cursor = _dt.date(cursor.year + 1, 1, 1)
        else:
            cursor = _dt.date(cursor.year, cursor.month + 1, 1)

    merged: Dict[str, Dict[str, str]] = defaultdict(dict)
    merged_id_map: Dict[str, str] = {}
    any_sheet_found = False

    for (y, m) in months:
        try:
            month_sched = load_schedule(xf, y, m)
        except ValueError:
            # No sheet in this workbook for that month — skip it rather
            # than failing the whole date range.
            continue
        any_sheet_found = True
        merged_id_map.update(_LAST_SCHEDULE_ID_MAP)
        for name, days in month_sched.items():
            for date_str, val in days.items():
                if start_str <= date_str <= end_str:
                    merged[name][date_str] = val

    if not any_sheet_found:
        raise ValueError(
            f"No schedule sheet found covering {start_d:%d %b %Y} - {end_d:%d %b %Y}"
        )

    _LAST_SCHEDULE_ID_MAP = merged_id_map
    return dict(merged)


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


def _format_date_ranges(date_strs: List[str]) -> str:
    """
    Collapse a list of 'YYYY-MM-DD' strings into human-readable ranges,
    e.g. ['2026-09-15', ..., '2026-09-22'] -> 'Sep 15-22'.
    Non-consecutive dates become separate comma-joined parts, e.g.
    'Sep 5, Sep 15-22'.
    """
    if not date_strs:
        return ""
    dates = sorted({datetime.strptime(d, "%Y-%m-%d").date() for d in date_strs})

    runs: List[Tuple[_dt.date, _dt.date]] = []
    run_start = dates[0]
    run_end = dates[0]
    for d in dates[1:]:
        if (d - run_end).days == 1:
            run_end = d
        else:
            runs.append((run_start, run_end))
            run_start = d
            run_end = d
    runs.append((run_start, run_end))

    parts = []
    for start, end in runs:
        if start == end:
            parts.append(start.strftime("%b %d"))
        elif start.month == end.month and start.year == end.year:
            parts.append(f"{start.strftime('%b %d')}-{end.strftime('%d')}")
        else:
            parts.append(f"{start.strftime('%b %d')} - {end.strftime('%b %d')}")
    return ", ".join(parts)


_NIGHT_SHIFT_HOUR = 18   # a scheduled start at/after 6 PM counts as a night shift
_NIGHT_SHIFT_CUTOFF = _time(12, 0)  # early punches before noon can belong to the prior night


def _reassign_night_shift_punches(emp_schedule: Dict[str, str],
                                  emp_punches: Dict[str, List[datetime]]
                                  ) -> Dict[str, List[datetime]]:
    """
    Re-attribute punches that were logged just after midnight to the night
    shift they actually belong to, instead of the punch's own raw
    calendar date.

    The device logs every punch under its own calendar date. For a shift
    that starts late in the evening (e.g. 11:58 PM), the employee's
    clock-in often lands a few minutes into the *next* calendar date
    (e.g. 12:01 AM) — well within the grace period of the previous
    night's shift. Left as-is, that makes the report show:
      - the correct shift day as "Absent" (no punch was ever recorded
        under that date), and
      - the following day as if the employee clocked in around
        midnight, hours before that day's own (also late-night) shift.

    This walks the employee's schedule in date order; for any day whose
    shift starts at or after 6 PM, it pulls in punches from the *next*
    calendar day that occur before noon and re-attributes them to this
    shift day, removing them from the next day's own punch list so they
    are never counted twice. Day shifts (start before 6 PM) are left
    untouched, since their punches already land on the correct date.
    """
    if not emp_schedule or not emp_punches:
        return emp_punches

    output: Dict[str, List[datetime]] = {
        ds: list(plist) for ds, plist in emp_punches.items()
    }

    for date_str in sorted(emp_schedule.keys()):
        sched_val = emp_schedule[date_str]
        if not sched_val:
            continue
        kind, _detail = _classify_schedule(sched_val)
        if kind not in ("time", "range"):
            continue  # OFF / leave / unrecognised — no shift to roll punches into
        shift_start = parse_start_time(sched_val)
        if shift_start is None or shift_start.hour < _NIGHT_SHIFT_HOUR:
            continue  # a day shift; punches already land on the right date

        try:
            d = datetime.strptime(date_str, "%Y-%m-%d").date()
        except Exception:
            continue
        next_date_str = (d + timedelta(days=1)).strftime("%Y-%m-%d")

        # Don't steal punches from a next day that has its own legitimate
        # early/day shift (start before 6 PM) — only roll punches forward
        # when the next day is unscheduled, off/leave, or itself another
        # night shift.
        next_sched_val = emp_schedule.get(next_date_str)
        if next_sched_val:
            next_kind, _nd = _classify_schedule(next_sched_val)
            if next_kind in ("time", "range"):
                next_shift_start = parse_start_time(next_sched_val)
                if next_shift_start is not None and next_shift_start.hour < _NIGHT_SHIFT_HOUR:
                    continue

        next_day_punches = output.get(next_date_str)
        if not next_day_punches:
            continue

        claimed = [p for p in next_day_punches if p.time() < _NIGHT_SHIFT_CUTOFF]
        if not claimed:
            continue

        output[date_str] = sorted(output.get(date_str, []) + claimed)
        output[next_date_str] = [p for p in next_day_punches if p not in claimed]

    return output


def build_report(punch_data: Dict[str, Dict[str, List[datetime]]],
                 sched_data: Dict[str, Dict[str, str]],
                 start_date,
                 end_date) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    start_date / end_date may be datetime.date, datetime.datetime, or
    'YYYY-MM-DD' strings, and mark the inclusive range the report covers
    (previously this took a single calendar month via year/month).

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

    # ── exact employee-ID matching (authoritative, always wins) ──
    # Uses the ID maps that load_schedule()/load_punches() populated while
    # parsing each workbook ("Employees ID" in the schedule, "No." in the
    # device export). This resolves the cases fuzzy name matching gets
    # wrong — e.g. two different real people who happen to share a display
    # name, or a schedule nickname that doesn't fuzzy-match the device's
    # full legal name at all.
    id_overrides: Dict[str, str] = {}
    claimed_badge_ids: set = set()
    for sched_key, sched_id in _LAST_SCHEDULE_ID_MAP.items():
        if sched_key not in sched_data:
            continue
        punch_key = _LAST_PUNCH_ID_TO_KEY.get(sched_id)
        if not punch_key:
            continue
        if punch_key in named_punch_data:
            id_overrides[sched_key] = punch_key
        elif punch_key in badge_punch_data and sched_key not in named_punch_data:
            # device only ever recorded a badge number for this person; the
            # schedule's ID tells us who they really are, so claim their
            # punches out of the anonymous Badge ID sheet.
            named_punch_data[sched_key] = badge_punch_data[punch_key]
            id_overrides[sched_key] = sched_key
            claimed_badge_ids.add(punch_key)

    for bid in claimed_badge_ids:
        badge_punch_data.pop(bid, None)

    # Build the name mapping (only for named employees); ID matches above
    # win first, fuzzy name matching fills in the rest.
    sched_names = list(sched_data.keys())
    punch_names = list(named_punch_data.keys())
    name_map = _build_name_map(sched_names, punch_names, id_overrides=id_overrides)


    # Build comparison of names between device data and schedule data.
    # name_map maps schedule names to matching device names.
    matched_punch_names = set(name_map.values()) | claimed_badge_ids
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



    # All days in the requested [start_date, end_date] range (inclusive)
    start_d = _to_date(start_date)
    end_d = _to_date(end_date)
    if end_d < start_d:
        start_d, end_d = end_d, start_d
    num_days = (end_d - start_d).days + 1
    all_dates = [
        (start_d + _dt.timedelta(days=d)).strftime("%Y-%m-%d")
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

        # Night shifts (e.g. starting 11:58 PM) often log their clock-in a
        # few minutes into the next calendar date; re-attribute those
        # early punches back to the correct shift day before evaluating
        # attendance below.
        emp_punches = _reassign_night_shift_punches(emp_schedule, emp_punches)

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
        leave_detail: Dict[str, List[str]] = defaultdict(list)  # code → dates on leave
        off_days = 0
        off_dates: List[str] = []

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
                    off_days += 1
                    off_dates.append(date_str)
                elif code == "RESIGNED":
                    pass
                else:
                    leave_days += 1
                    leave_detail[code].append(date_str)

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
        leave_breakdown = "; ".join(
            f"{_LEAVE_LABELS.get(c, c)}: {_format_date_ranges(dates)}"
            for c, dates in sorted(leave_detail.items())
        )
        off_breakdown = _format_date_ranges(off_dates)
        summary_rows.append({
            "Employee": display_name,
            "Present": present_days,
            "On Time": on_time_days,
            "Late": late_days,
            "Absent": absent_days,
            "WFH": wfh_days,
            "Off": off_days,
            "Off Dates": off_breakdown,
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
            # only include dates within the requested range
            try:
                d = datetime.strptime(date_str, "%Y-%m-%d").date()
                if d < start_d or d > end_d:
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
