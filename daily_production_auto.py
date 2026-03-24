#!/usr/bin/env python3
"""
Daily Production Automation
=============================
Reads all crew daily report files from the 'Daily's' subfolder,
extracts today's production values, updates the master template,
and writes a crew/subcontractor contribution report.

Each run:
  - Writes today's job-level footage into the correct rows
  - Writes contractor-level DAILY sub-total formulas (auto-sum job rows)
  - Writes top-level DAILY summary formulas
  - Accumulates WEEKLY totals (resets each Monday)
  - Accumulates MONTHLY totals (resets on 1st of month)
  - Updates the monthly table for the current month (O14:W25)
  - Writes QUARTERLY formula rows (auto-sum monthly rows)
  - Writes a YEARLY formula row
  - Writes all dollar-value formula cells (footage x rate)
  - Saves a dated copy AND updates the master template file

Usage:
    python daily_production_auto.py                   # uses today's date
    python daily_production_auto.py --date 03-24-2026 # override date (MM-DD-YYYY)
    python daily_production_auto.py --dry-run          # preview only, no files written
"""

import sys
import re
import argparse
from datetime import date, datetime
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl is required.  Run:  pip install openpyxl")
    sys.exit(1)

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(r"c:\Users\JoshuaCollver\OneDrive\Working Folder for Claude.AI\Daily Production")
DAILY_DIR     = BASE_DIR / "Daily's"
TEMPLATE_FILE = BASE_DIR / "DAILY PRODUCTION TEMPLATE (7).xlsx"
TEMPLATE_SHEET = "DAILY PRODUCTION TEMPLATE"

# ── Contractor sections ────────────────────────────────────────────────────────
# (name, first_job_row, last_job_row, daily_row, weekly_row, monthly_row)
# first/last_job_row cover the FULL range including gap rows — SUM treats blanks as 0
CONTRACTOR_SECTIONS = [
    ("AT&T LG - 01",      11,  36,  37,  38,  39),
    ("NATCO - 08",        41,  42,  43,  44,  45),
    ("AT&T SVA - 01",     47,  63,  64,  65,  66),
    ("WINDSTREAM",        68,  81,  82,  83,  84),
    ("YELCOTT - 15",      86,  94,  95,  96,  97),
    ("ARTEL - 12",        99, 101, 102, 103, 104),
    ("FECC - 39",        106, 115, 116, 117, 118),
    ("Southern Pipeline", None, None, 122, 123, 124),
    ("AECC - 32",        131, 131, 132, 133, 134),
]

# Rows that hold contractor-level DAILY subtotals (used in top summary formulas)
ALL_DAILY_ROWS = [37, 43, 64, 82, 95, 102, 116, 122, 132]

# ── Job-area column letters (columns C–L) ─────────────────────────────────────
JOB_COL = {
    "Bore":          "C",
    "Rock Bore":     "D",
    "Bore Inc":      "E",
    "Bore Rock Inc": "F",
    "Trench":        "G",
    "Trench Rock":   "H",
    "Plow":          "I",
    "Aerial":        "J",
    "Cable":         "K",
    "Drops":         "L",
}

# ── Template daily/weekly/monthly update mapping (columns C–L → summary cells) ─
# Template column letters for writing TODAY's job values
TEMPLATE_COL = {
    "Bore":        "C",
    "Rock Bore":   "D",
    "Trench":      "G",
    "Trench Rock": "H",
    "Plow":        "I",
    "Aerial":      "J",
    "Cable":       "K",
    "Drops":       "L",
}
TEMPLATE_COMMENT_COL = "N"

# ── Summary block cell addresses ───────────────────────────────────────────────
# DAILY LF — these become formula cells (=SUM of contractor DAILY rows)
DAILY_LF_CELL = {
    "Bore":          "B3",
    "Rock Bore":     "B4",
    "Bore Inc":      "B5",
    "Bore Rock Inc": "B6",
    "Trench":        "I3",
    "Trench Rock":   "I4",
    "Plow":          "I5",
    "Aerial":        "I6",
    "Cable":         "I7",
    "Drops":         "I8",
}

# WEEKLY LF — written as numbers by the script (accumulated)
WEEKLY_CELL = {
    "Bore":        "D3",
    "Rock Bore":   "D4",
    "Trench":      "K3",
    "Trench Rock": "K4",
    "Plow":        "K5",
    "Aerial":      "K6",
    "Cable":       "K7",
    "Drops":       "K8",
}

# MONTHLY LF — written as numbers by the script (accumulated)
MONTHLY_CELL = {
    "Bore":        "F3",
    "Rock Bore":   "F4",
    "Trench":      "M3",
    "Trench Rock": "M4",
    "Plow":        "M5",
    "Aerial":      "M6",
    "Cable":       "M7",
    "Drops":       "M8",
}

# ── Day-of-week breakdown table (rows 2-6, cols O-W) ──────────────────────────
# Monday=row2, Tuesday=row3, Wednesday=row4, Thursday=row5, Friday=row6
# Column O holds the day label; P-W hold the daily LF values.
DAY_TABLE_FIRST_ROW = 2   # Monday
DAY_TABLE_COL = {
    "Drops":       "P",
    "Bore":        "Q",
    "Rock Bore":   "R",
    "Trench":      "S",
    "Trench Rock": "T",
    "Plow":        "U",
    "Aerial":      "V",
    "Cable":       "W",
}

# Dollar value formulas: daily $, weekly $, monthly $ (column C-G and J-N)
# Pattern: daily_lf_cell * rate,  weekly_lf_cell * rate,  monthly_lf_cell * rate
DOLLAR_FORMULA_MAP = [
    # (daily_$ cell, weekly_$ cell, monthly_$ cell, daily_lf_cell, weekly_lf_cell, monthly_lf_cell, rate_cell)
    ("C3", "E3", "G3", "B3", "D3", "F3", "Y$15"),   # Bore
    ("C4", "E4", "G4", "B4", "D4", "F4", "Z$15"),   # Rock Bore
    ("J3", "L3", "N3", "I3", "K3", "M3", "AA$15"),  # Trench
    ("J4", "L4", "N4", "I4", "K4", "M4", "AB$15"),  # Trench Rock
    ("J5", "L5", "N5", "I5", "K5", "M5", "AC$15"),  # Plow
    ("J6", "L6", "N6", "I6", "K6", "M6", "AD$15"),  # Aerial
    ("J7", "L7", "N7", "I7", "K7", "M7", "AE$15"),  # Cable
    ("J8", "L8", "N8", "I8", "K8", "M8", "AF$15"),  # Drops
]

# AD:AL value block — dollar values of daily/weekly/monthly by work type
# AE–AL = Bore, Rock Bore, Trench, Trench Rock, Plow, Aerial, Cable, Drops
AD_AL_MAP = [
    # (daily_$ col, weekly_$ col, monthly_$ col, daily_lf_cell, weekly_lf_cell, monthly_lf_cell, rate)
    ("AE", "B3", "D3", "F3", "Y$15"),   # Bore
    ("AF", "B4", "D4", "F4", "Z$15"),   # Rock Bore
    ("AG", "I3", "K3", "M3", "AA$15"),  # Trench
    ("AH", "I4", "K4", "M4", "AB$15"),  # Trench Rock
    ("AI", "I5", "K5", "M5", "AC$15"),  # Plow
    ("AJ", "I6", "K6", "M6", "AD$15"),  # Aerial
    ("AK", "I7", "K7", "M7", "AE$15"),  # Cable
    ("AL", "I8", "K8", "M8", "AF$15"),  # Drops
]

# ── Monthly table (O14:W25) ────────────────────────────────────────────────────
MONTH_TABLE_START_ROW = 14   # row 14 = JAN, row 25 = DEC
MONTH_TABLE_COL = {
    "Bore":        "Q",
    "Rock Bore":   "R",
    "Trench":      "S",
    "Trench Rock": "T",
    "Plow":        "U",
    "Aerial":      "V",
    "Cable":       "W",
    "Drops":       "P",
}
MONTHLY_TABLE_COLS = ["P", "Q", "R", "S", "T", "U", "V", "W"]
QUARTERLY_START_ROW = 27     # Q1=27, Q2=28, Q3=29, Q4=30
YEARLY_ROW = 31

WORK_TYPES = ["Bore", "Rock Bore", "Trench", "Trench Rock", "Plow", "Aerial", "Cable", "Drops"]

# ── Header patterns for daily report layout detection ─────────────────────────
HEADER_PATTERNS = [
    ("rock bore",    "Rock Bore"),
    ("rock trench",  "Rock Trench"),
    ("daily ttl",    "Daily Total"),
    ("daily total",  "Daily Total"),
    ("bore",         "Bore"),
    ("plow",         "Plow"),
    ("trench",       "Trench"),
    ("aerial",       "Aerial"),
    ("cable",        "Cable"),
    ("drops",        "Drops"),
    ("remarks",      "Remarks"),
    ("date",         "Date"),
]

JOB_RE = re.compile(r"\d{2}[-/]\d{4}[-/]\d{2}")
JOB_NUMBER_RE = re.compile(r"^\d{2}-\d{4}-\d{2}$")


# ═════════════════════════════════════════════════════════════════════════════
# Helpers
# ═════════════════════════════════════════════════════════════════════════════

def _to_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    return 0.0


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except (ValueError, AttributeError):
                pass
    return None


def _norm_job(v):
    if not v:
        return None
    m = JOB_RE.search(str(v).strip())
    if m:
        return re.sub(r"[/]", "-", m.group())
    return None


def _norm_crew(name):
    return str(name or "UNKNOWN").strip().upper()


# ═════════════════════════════════════════════════════════════════════════════
# Daily report reading
# ═════════════════════════════════════════════════════════════════════════════

def detect_layout(ws):
    def cv(addr):
        v = ws[addr].value
        return str(v).strip() if v is not None else None

    job_a = _norm_job(cv("K5"))
    job_b = _norm_job(cv("M5"))

    if job_a:
        job_number = job_a
        wo_number  = cv("K4")
        location   = cv("K7")
        crew       = cv("K8")
        supervisor = cv("B11")
    elif job_b:
        job_number = job_b
        wo_number  = cv("M4")
        location   = cv("M6")
        crew       = cv("M7")
        supervisor = cv("B11")
    else:
        job_number = None
        wo_number = location = crew = supervisor = None
        for r in range(4, 13):
            for c in range(1, 20):
                jn = _norm_job(ws.cell(row=r, column=c).value)
                if jn:
                    job_number = jn
                    break
            if job_number:
                break
        supervisor = cv("B11")

    header_row = None
    col_map    = {}
    for row in ws.iter_rows(min_row=10, max_row=20):
        row_headers = {}
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                low = cell.value.strip().lower()
                for key, wt in HEADER_PATTERNS:
                    if key in low:
                        row_headers[wt] = cell.column - 1
                        break
        if "Date" in row_headers and ("Bore" in row_headers or "Plow" in row_headers):
            header_row = row[0].row
            col_map    = row_headers
            break

    if header_row is None:
        header_row = 14
        col_map = {
            "Date": 0, "Plow": 1, "Bore": 2, "Rock Bore": 3,
            "Trench": 4, "Rock Trench": 5, "Aerial": 6, "Cable": 7,
            "Drops": 8, "Daily Total": 9, "Remarks": 10,
        }

    return {
        "job_number": job_number,
        "wo_number":  wo_number,
        "location":   location,
        "crew":       crew,
        "supervisor": supervisor,
        "header_row": header_row,
        "col_map":    col_map,
        "data_start": header_row + 1,
    }


def read_daily_report(filepath: Path, target_date: date) -> dict:
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    layout = detect_layout(ws)
    col    = layout["col_map"]

    report = {
        "filename":        filepath.name,
        "job_number":      layout["job_number"],
        "wo_number":       layout["wo_number"],
        "location":        layout["location"],
        "crew":            layout["crew"],
        "supervisor":      layout["supervisor"],
        "today_values":    {wt: 0.0 for wt in WORK_TYPES},
        "today_remarks":   [],
        "today_entries":   [],
        "all_entries":     [],
        "used_last_entry": False,
    }

    for row_tuple in ws.iter_rows(min_row=layout["data_start"], values_only=True):
        date_val = row_tuple[col.get("Date", 0)] if col.get("Date", 0) < len(row_tuple) else None
        if date_val is None:
            break
        row_date = _to_date(date_val)
        if row_date is None:
            continue

        def gc(key):
            idx = col.get(key)
            return _to_num(row_tuple[idx]) if idx is not None and idx < len(row_tuple) else 0.0

        # "Rock Trench" in daily reports maps to "Trench Rock" in our WORK_TYPES
        entry = {
            "date":        row_date,
            "Plow":        gc("Plow"),
            "Bore":        gc("Bore"),
            "Rock Bore":   gc("Rock Bore"),
            "Trench":      gc("Trench"),
            "Trench Rock": gc("Rock Trench") or gc("Trench Rock"),
            "Aerial":      gc("Aerial"),
            "Cable":       gc("Cable"),
            "Drops":       gc("Drops"),
            "total":       gc("Daily Total"),
            "remarks":     str(row_tuple[col["Remarks"]] if col.get("Remarks") and col["Remarks"] < len(row_tuple) and row_tuple[col["Remarks"]] else "").strip(),
        }
        report["all_entries"].append(entry)

    target_entries = [e for e in report["all_entries"] if e["date"] == target_date]
    if target_entries:
        report["today_entries"] = target_entries
    elif report["all_entries"]:
        report["today_entries"]   = [report["all_entries"][-1]]
        report["used_last_entry"] = True

    for wt in WORK_TYPES:
        report["today_values"][wt] = sum(e[wt] for e in report["today_entries"])

    report["today_remarks"] = [e["remarks"] for e in report["today_entries"] if e["remarks"]]
    return report


# ═════════════════════════════════════════════════════════════════════════════
# Deduplication
# ═════════════════════════════════════════════════════════════════════════════

def deduplicate_reports(reports: list) -> tuple:
    best       = {}
    all_by_job = defaultdict(list)

    for r in reports:
        jn = r["job_number"]
        if not jn:
            best[id(r)] = r
            continue
        all_by_job[jn].append(r)

    duplicates = []
    for jn, group in all_by_job.items():
        if len(group) == 1:
            best[jn] = group[0]
        else:
            def last_date(r):
                return r["all_entries"][-1]["date"] if r["all_entries"] else date.min
            group.sort(key=last_date, reverse=True)
            best[jn] = group[0]
            duplicates.extend(group[1:])

    return list(best.values()), duplicates


# ═════════════════════════════════════════════════════════════════════════════
# Template job-row map
# ═════════════════════════════════════════════════════════════════════════════

def build_job_row_map(ws) -> dict:
    mapping = {}
    for row in ws.iter_rows(min_col=2, max_col=2):
        cell = row[0]
        if cell.value and isinstance(cell.value, str):
            val = cell.value.strip()
            if JOB_NUMBER_RE.match(val):
                mapping[val] = cell.row
    return mapping


# ═════════════════════════════════════════════════════════════════════════════
# Formula writers
# ═════════════════════════════════════════════════════════════════════════════

def _sum_ref(col, rows):
    """Build =SUM(C37,C43,...) style formula from a column letter and list of rows."""
    refs = ",".join(f"{col}{r}" for r in rows)
    return f"=SUM({refs})"


def write_formulas(ws):
    """
    Write all structural Excel formulas into the worksheet.
    These formulas auto-recalculate when Excel opens the file.

    Formulas written:
      - Contractor DAILY rows: =SUM(job-row range)
      - Top DAILY summary (B3/B4/I3-I8): =SUM(contractor DAILY cells)
      - Dollar value cells (C3,E3,G3,J3-N8,AE3-AL5): = LF * rate
      - Dollar subtotals (Y17:Y19): =SUM(AE:AL row)
      - Quarterly rows (O27:W30): =SUM(3 monthly rows)
      - Yearly row (O31:W31): =SUM(all 12 monthly rows)
    """

    # ── 1. Contractor DAILY rows: =SUM(C{first}:C{last}) ──────────────────
    for _, first_row, last_row, daily_row, _, _ in CONTRACTOR_SECTIONS:
        if first_row is None:
            # No job rows (e.g., Southern Pipeline) — set to 0
            for col in JOB_COL.values():
                ws[f"{col}{daily_row}"] = 0
            continue
        for col in JOB_COL.values():
            ws[f"{col}{daily_row}"] = f"=SUM({col}{first_row}:{col}{last_row})"

    # ── 2. Top DAILY summary formulas ─────────────────────────────────────
    # Map each work type to its job-area column, then sum across all contractor DAILY rows
    daily_col_map = {
        "Bore":          "C",
        "Rock Bore":     "D",
        "Bore Inc":      "E",
        "Bore Rock Inc": "F",
        "Trench":        "G",
        "Trench Rock":   "H",
        "Plow":          "I",
        "Aerial":        "J",
        "Cable":         "K",
        "Drops":         "L",
    }
    for wt, summary_cell in DAILY_LF_CELL.items():
        col = daily_col_map[wt]
        ws[summary_cell] = _sum_ref(col, ALL_DAILY_ROWS)

    # ── 3. Dollar value cells in summary block (rows 3-8) ─────────────────
    for daily_cell, weekly_cell, monthly_cell, d_lf, w_lf, m_lf, rate in DOLLAR_FORMULA_MAP:
        ws[daily_cell]   = f"={d_lf}*{rate}"
        ws[weekly_cell]  = f"={w_lf}*{rate}"
        ws[monthly_cell] = f"={m_lf}*{rate}"

    # ── 4. AD:AL block (rows 3-5): daily/weekly/monthly $ per work type ───
    for col_letter, d_lf, w_lf, m_lf, rate in AD_AL_MAP:
        ws[f"{col_letter}3"] = f"={d_lf}*{rate}"   # DAILY $
        ws[f"{col_letter}4"] = f"={w_lf}*{rate}"   # WEEKLY $
        ws[f"{col_letter}5"] = f"={m_lf}*{rate}"   # MONTHLY $

    # ── 5. Dollar subtotals (Y17:Y19) ─────────────────────────────────────
    ws["Y17"] = "=SUM(AE3:AL3)"   # Daily total $
    ws["Y18"] = "=SUM(AE4:AL4)"   # Weekly total $
    ws["Y19"] = "=SUM(AE5:AL5)"   # Monthly total $

    # ── 6. Quarterly formulas (O27:W30) ───────────────────────────────────
    # Q1=JAN/FEB/MAR (rows 14-16), Q2=APR-JUN (17-19), Q3=JUL-SEP (20-22), Q4=OCT-DEC (23-25)
    quarter_ranges = [
        (QUARTERLY_START_ROW + 0, 14, 16),  # Q1
        (QUARTERLY_START_ROW + 1, 17, 19),  # Q2
        (QUARTERLY_START_ROW + 2, 20, 22),  # Q3
        (QUARTERLY_START_ROW + 3, 23, 25),  # Q4
    ]
    ws[f"O{QUARTERLY_START_ROW + 0}"] = "Q1"
    ws[f"O{QUARTERLY_START_ROW + 1}"] = "Q2"
    ws[f"O{QUARTERLY_START_ROW + 2}"] = "Q3"
    ws[f"O{QUARTERLY_START_ROW + 3}"] = "Q4"
    for q_row, m_start, m_end in quarter_ranges:
        for col in MONTHLY_TABLE_COLS:
            ws[f"{col}{q_row}"] = f"=SUM({col}{m_start}:{col}{m_end})"

    # ── 7. Yearly formula row ──────────────────────────────────────────────
    ws[f"O{YEARLY_ROW}"] = "YEARLY"
    for col in MONTHLY_TABLE_COLS:
        ws[f"{col}{YEARLY_ROW}"] = f"=SUM({col}{MONTH_TABLE_START_ROW}:{col}{MONTH_TABLE_START_ROW + 11})"


# ═════════════════════════════════════════════════════════════════════════════
# Weekly / Monthly accumulation
# ═════════════════════════════════════════════════════════════════════════════

def accumulate_periods(ws, daily_totals: dict, target_date: date):
    """
    Update weekly and monthly accumulated LF values in the summary block.
    Uses A2 to detect the last processed date and determine resets.
    Also updates the monthly table row for the current month.
    """
    # Read last processed date from A2
    last_run = None
    raw = ws["A2"].value
    if raw:
        try:
            last_run = datetime.strptime(str(raw).strip(), "%m/%d/%Y").date()
        except (ValueError, AttributeError):
            pass

    if last_run == target_date:
        print("  (accumulation skipped — already processed today)")
        return

    # Determine period boundaries
    new_week  = True
    new_month = True
    if last_run:
        if last_run.isocalendar()[:2] == target_date.isocalendar()[:2]:
            new_week = False
        if last_run.month == target_date.month and last_run.year == target_date.year:
            new_month = False

    period_label = []
    if new_week:
        period_label.append("new week")
    if new_month:
        period_label.append("new month")
    if period_label:
        print(f"  (period reset: {', '.join(period_label)})")

    # Update weekly and monthly LF cells
    for wt in WORK_TYPES:
        today_val    = daily_totals.get(wt, 0.0)
        weekly_cell  = WEEKLY_CELL.get(wt)
        monthly_cell = MONTHLY_CELL.get(wt)

        if weekly_cell:
            if new_week:
                ws[weekly_cell] = round(today_val) if today_val else 0
            else:
                existing = _to_num(ws[weekly_cell].value)
                ws[weekly_cell] = round(existing + today_val) if (existing + today_val) else 0

        if monthly_cell:
            if new_month:
                ws[monthly_cell] = round(today_val) if today_val else 0
            else:
                existing = _to_num(ws[monthly_cell].value)
                ws[monthly_cell] = round(existing + today_val) if (existing + today_val) else 0

    # Sync monthly table row for current month with the accumulated monthly values
    month_row = MONTH_TABLE_START_ROW + target_date.month - 1   # JAN=14 ... DEC=25
    for wt, col in MONTH_TABLE_COL.items():
        monthly_cell = MONTHLY_CELL.get(wt)
        if monthly_cell:
            ws[f"{col}{month_row}"] = ws[monthly_cell].value

    # ── Update day-of-week breakdown table (O2:W6) ────────────────────────
    weekday = target_date.weekday()   # 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri
    if weekday <= 4:                  # skip Saturday / Sunday
        # On Monday (start of new week) clear the whole table first so
        # stale data from the prior week doesn't linger in Tue-Fri rows.
        if new_week or weekday == 0:
            for r in range(DAY_TABLE_FIRST_ROW, DAY_TABLE_FIRST_ROW + 5):
                for col in DAY_TABLE_COL.values():
                    ws[f"{col}{r}"] = None

        day_row = DAY_TABLE_FIRST_ROW + weekday   # Mon=2, Tue=3, ...
        for wt, col in DAY_TABLE_COL.items():
            v = daily_totals.get(wt, 0.0)
            ws[f"{col}{day_row}"] = round(v) if v else None


# ═════════════════════════════════════════════════════════════════════════════
# Main template update
# ═════════════════════════════════════════════════════════════════════════════

def update_template(reports: list, target_date: date,
                    output_path: Path, dry_run: bool) -> tuple:

    wb  = openpyxl.load_workbook(TEMPLATE_FILE)
    ws  = wb[TEMPLATE_SHEET]
    job_row_map = build_job_row_map(ws)

    matched   = []
    unmatched = []

    # ── Write today's job-level footage ───────────────────────────────────
    for report in reports:
        job_num = report["job_number"]
        if not job_num:
            unmatched.append(report)
            continue
        row_num = job_row_map.get(job_num)
        if row_num is None:
            unmatched.append(report)
            continue

        vals = report["today_values"]
        for wt, col in TEMPLATE_COL.items():
            v = vals.get(wt, 0.0)
            ws[f"{col}{row_num}"] = round(v) if v else None

        remarks_text = " | ".join(report["today_remarks"])
        if remarks_text:
            ws[f"{TEMPLATE_COMMENT_COL}{row_num}"] = remarks_text

        matched.append(report)

    # ── Compute overall daily totals (across all reports) ─────────────────
    daily_totals = {wt: 0.0 for wt in WORK_TYPES}
    for report in reports:
        for wt in WORK_TYPES:
            daily_totals[wt] += report["today_values"].get(wt, 0.0)

    # ── Accumulate weekly / monthly ───────────────────────────────────────
    accumulate_periods(ws, daily_totals, target_date)

    # ── Write all structural formulas ─────────────────────────────────────
    write_formulas(ws)

    # ── Stamp processing date ─────────────────────────────────────────────
    ws["A2"] = target_date.strftime("%m/%d/%Y")

    # ── Save ──────────────────────────────────────────────────────────────
    if not dry_run:
        wb.save(output_path)
        # Also update the master template so accumulated values persist for tomorrow
        try:
            wb.save(TEMPLATE_FILE)
            print(f"  Master template updated: {TEMPLATE_FILE.name}")
        except PermissionError:
            print(f"  WARNING: Could not update master template (file may be open in Excel).")
            print(f"  Close the template and re-run, or copy {output_path.name} over it manually.")

    return matched, unmatched


# ═════════════════════════════════════════════════════════════════════════════
# Crew contribution report
# ═════════════════════════════════════════════════════════════════════════════

def split_crews(crew_field: str) -> list:
    if not crew_field:
        return ["UNKNOWN"]
    parts  = re.split(r"[,/]+", crew_field)
    result = [p.strip().upper() for p in parts if p.strip()]
    return result if result else ["UNKNOWN"]


def generate_crew_report(reports: list, target_date: date,
                          output_path: Path, dry_run: bool):
    crew_totals  = defaultdict(lambda: defaultdict(float))
    crew_details = defaultdict(list)
    all_remarks  = []

    for report in reports:
        raw_crew   = _norm_crew(report["crew"])
        crews      = split_crews(raw_crew)
        supervisor = str(report["supervisor"] or "").strip()
        job_num    = report["job_number"] or "N/A"
        location   = str(report["location"] or "").strip()
        vals       = report["today_values"]
        total_lf   = sum(vals.get(wt, 0.0) for wt in WORK_TYPES)
        flag       = " *" if report["used_last_entry"] else ""

        for wt in WORK_TYPES:
            crew_totals[raw_crew][wt] += vals.get(wt, 0.0)

        crew_details[raw_crew].append({
            "job":        job_num,
            "location":   location,
            "supervisor": supervisor,
            "values":     vals,
            "total":      total_lf,
            "remarks":    " | ".join(report["today_remarks"]),
            "flag":       flag,
            "filename":   report["filename"],
            "sub_crews":  crews if len(crews) > 1 else [],
        })

        for remark in report["today_remarks"]:
            if remark:
                all_remarks.append({
                    "crew":       raw_crew,
                    "supervisor": supervisor,
                    "job":        job_num,
                    "location":   location,
                    "remarks":    remark,
                })

    cwb      = openpyxl.Workbook()
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F497D")
    alt_fill = PatternFill("solid", fgColor="DCE6F1")
    c_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    l_align  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin     = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )

    def style_hdr(ws, row, n):
        for c in range(1, n + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = hdr_font; cell.fill = hdr_fill
            cell.alignment = c_align; cell.border = thin

    def sc(cell, alt=False, align=None):
        cell.fill      = alt_fill if alt else PatternFill()
        cell.alignment = align or c_align
        cell.border    = thin

    date_label = target_date.strftime("%B %d, %Y")

    # ── Crew Summary sheet ────────────────────────────────────────────────
    s1 = cwb.active
    s1.title = "Crew Summary"
    s1.append([f"Daily Production  --  {date_label}"])
    s1.cell(row=1, column=1).font = Font(bold=True, size=13)
    s1.row_dimensions[1].height = 22

    hdrs1 = ["Crew / Company"] + WORK_TYPES + ["Total LF"]
    for c, h in enumerate(hdrs1, 1):
        s1.cell(row=2, column=c, value=h)
    style_hdr(s1, 2, len(hdrs1))

    grand = {wt: 0.0 for wt in WORK_TYPES}
    for r, (crew, totals) in enumerate(sorted(crew_totals.items()), 3):
        alt = r % 2 == 0
        sc(s1.cell(row=r, column=1, value=crew), alt, l_align)
        row_total = 0.0
        for c, wt in enumerate(WORK_TYPES, 2):
            v = totals.get(wt, 0.0)
            grand[wt] += v
            sc(s1.cell(row=r, column=c, value=round(v) if v else None), alt)
            row_total += v
        tot_cell = s1.cell(row=r, column=len(hdrs1), value=round(row_total) if row_total else None)
        sc(tot_cell, alt); tot_cell.font = Font(bold=True)

    gt_row = len(crew_totals) + 3
    sc(s1.cell(row=gt_row, column=1, value="GRAND TOTAL"), align=l_align)
    for c, wt in enumerate(WORK_TYPES, 2):
        sc(s1.cell(row=gt_row, column=c, value=round(grand[wt]) if grand[wt] else None))
    sc(s1.cell(row=gt_row, column=len(hdrs1), value=round(sum(grand.values()))))
    style_hdr(s1, gt_row, len(hdrs1))

    s1.column_dimensions["A"].width = 30
    for c in range(2, len(hdrs1) + 1):
        s1.column_dimensions[get_column_letter(c)].width = 11

    # ── Job Detail sheet ──────────────────────────────────────────────────
    s2 = cwb.create_sheet("Job Detail")
    s2.append([f"Daily Production  --  {date_label}"])
    s2.cell(row=1, column=1).font = Font(bold=True, size=13)
    s2.row_dimensions[1].height = 22

    hdrs2 = ["Crew / Company", "Sub-Crews", "Supervisor", "Job Number",
              "Location"] + WORK_TYPES + ["Total LF", "Notes / Remarks", "Source File"]
    for c, h in enumerate(hdrs2, 1):
        s2.cell(row=2, column=c, value=h)
    style_hdr(s2, 2, len(hdrs2))

    r = 3
    for crew in sorted(crew_details.keys()):
        for info in crew_details[crew]:
            alt  = r % 2 == 0
            vals = info["values"]
            sub  = ", ".join(info["sub_crews"]) if info["sub_crews"] else ""
            sc(s2.cell(row=r, column=1, value=crew + info["flag"]), alt, l_align)
            sc(s2.cell(row=r, column=2, value=sub),                 alt, l_align)
            sc(s2.cell(row=r, column=3, value=info["supervisor"]),  alt, l_align)
            sc(s2.cell(row=r, column=4, value=info["job"]),         alt)
            sc(s2.cell(row=r, column=5, value=info["location"]),    alt, l_align)
            for c, wt in enumerate(WORK_TYPES, 6):
                v = vals.get(wt, 0.0)
                sc(s2.cell(row=r, column=c, value=round(v) if v else None), alt)
            sc(s2.cell(row=r, column=6 + len(WORK_TYPES),
                       value=round(info["total"]) if info["total"] else None), alt)
            sc(s2.cell(row=r, column=7 + len(WORK_TYPES), value=info["remarks"]),  alt, l_align)
            sc(s2.cell(row=r, column=8 + len(WORK_TYPES), value=info["filename"]), alt, l_align)
            r += 1

    s2.column_dimensions["A"].width = 30
    s2.column_dimensions["B"].width = 28
    s2.column_dimensions["C"].width = 20
    s2.column_dimensions["D"].width = 14
    s2.column_dimensions["E"].width = 24
    for c in range(6, 6 + len(WORK_TYPES) + 1):
        s2.column_dimensions[get_column_letter(c)].width = 11
    s2.column_dimensions[get_column_letter(7 + len(WORK_TYPES))].width = 42
    s2.column_dimensions[get_column_letter(8 + len(WORK_TYPES))].width = 42

    # ── Notes Review sheet ────────────────────────────────────────────────
    s3 = cwb.create_sheet("Notes Review")
    s3.append([f"Field Notes / Remarks -- {date_label}  (review for crew attribution)"])
    s3.cell(row=1, column=1).font = Font(bold=True, size=13)
    s3.row_dimensions[1].height = 22

    hdrs3 = ["Crew / Company", "Supervisor", "Job Number", "Location", "Remarks / Notes"]
    for c, h in enumerate(hdrs3, 1):
        s3.cell(row=2, column=c, value=h)
    style_hdr(s3, 2, len(hdrs3))

    if all_remarks:
        for r2, item in enumerate(all_remarks, 3):
            alt = r2 % 2 == 0
            sc(s3.cell(row=r2, column=1, value=item["crew"]),       alt, l_align)
            sc(s3.cell(row=r2, column=2, value=item["supervisor"]), alt, l_align)
            sc(s3.cell(row=r2, column=3, value=item["job"]),        alt)
            sc(s3.cell(row=r2, column=4, value=item["location"]),   alt, l_align)
            sc(s3.cell(row=r2, column=5, value=item["remarks"]),    alt, l_align)
    else:
        s3.cell(row=3, column=1, value="No remarks recorded for this date.")

    s3.column_dimensions["A"].width = 30
    s3.column_dimensions["B"].width = 20
    s3.column_dimensions["C"].width = 14
    s3.column_dimensions["D"].width = 24
    s3.column_dimensions["E"].width = 65

    if not dry_run:
        cwb.save(output_path)

    return crew_totals


# ═════════════════════════════════════════════════════════════════════════════
# Main
# ═════════════════════════════════════════════════════════════════════════════

def parse_args():
    p = argparse.ArgumentParser(description="Daily Production Automation")
    p.add_argument("--date",    help="Override date MM-DD-YYYY (default: today)")
    p.add_argument("--dry-run", action="store_true", help="Preview only, no files written")
    return p.parse_args()


def main():
    args    = parse_args()
    dry_run = args.dry_run

    if args.date:
        try:
            target_date = datetime.strptime(args.date, "%m-%d-%Y").date()
        except ValueError:
            print(f"ERROR: Invalid date format '{args.date}'. Use MM-DD-YYYY.")
            sys.exit(1)
    else:
        target_date = date.today()

    date_str        = target_date.strftime("%m-%d-%Y")
    output_template = BASE_DIR / f"Daily Production {date_str}.xlsx"
    output_crew     = BASE_DIR / f"Crew Report {date_str}.xlsx"

    sep = "-" * 72
    print(f"\n{sep}")
    print(f"  Daily Production Automation  --  {target_date.strftime('%A, %B %d, %Y')}")
    if dry_run:
        print("  [DRY RUN -- no files will be saved]")
    print(sep)
    print(f"\nReading reports from: {DAILY_DIR}\n")

    # ── Read all daily reports ─────────────────────────────────────────────
    raw_reports = []
    errors      = []

    for xlsx_file in sorted(DAILY_DIR.glob("*.xlsx")):
        try:
            raw_reports.append(read_daily_report(xlsx_file, target_date))
        except Exception as exc:
            errors.append((xlsx_file.name, str(exc)))
            print(f"  ERROR reading {xlsx_file.name}: {exc}")

    reports, duplicates = deduplicate_reports(raw_reports)

    print(f"  {'File':<50}  {'Job #':<15}  {'Crew':<22}  {'LF':>6}  {'Note'}")
    print(f"  {'-'*50}  {'-'*15}  {'-'*22}  {'-'*6}  {'-'*18}")
    for r in reports:
        job   = r["job_number"] or "NO JOB#"
        crew  = (r["crew"] or "N/A")[:22]
        total = sum(r["today_values"].values())
        note  = "[last entry used]" if r["used_last_entry"] else ""
        print(f"  {r['filename'][:50]:<50}  {job:<15}  {crew:<22}  {total:>6.0f}  {note}")

    if duplicates:
        print(f"\n  Duplicate files skipped:")
        for d in duplicates:
            print(f"    - {d['filename']}")

    print(f"\n  {len(raw_reports)} files read  |  {len(duplicates)} duplicates skipped  |  {len(errors)} errors")

    # ── Update master template ─────────────────────────────────────────────
    print(f"\n{sep}")
    print("  Updating master template and writing formulas...")
    matched, unmatched = update_template(reports, target_date, output_template, dry_run)
    print(f"  Matched {len(matched)} jobs  |  {len(unmatched)} unmatched")

    if unmatched:
        print("\n  Jobs NOT found in template (add manually):")
        for r in unmatched:
            print(f"    X  {(r['job_number'] or 'NO JOB#'):<22}  {r['filename']}")

    if not dry_run:
        print(f"  Saved -> {output_template.name}")

    # ── Crew report ────────────────────────────────────────────────────────
    print(f"\n{sep}")
    print("  Generating crew contribution report...")
    crew_totals = generate_crew_report(reports, target_date, output_crew, dry_run)

    print(f"\n  {'Crew':<30}  {'Bore':>6}  {'R.Bore':>6}  {'Trench':>7}  "
          f"{'Plow':>5}  {'Aerial':>6}  {'Cable':>5}  {'Drops':>5}  {'TOTAL':>7}")
    print(f"  {'-'*30}  {'-'*6}  {'-'*6}  {'-'*7}  {'-'*5}  {'-'*6}  {'-'*5}  {'-'*5}  {'-'*7}")
    for crew in sorted(crew_totals.keys()):
        t     = crew_totals[crew]
        total = sum(t.values())
        if total:
            print(
                f"  {crew[:30]:<30}  {t['Bore']:>6.0f}  {t['Rock Bore']:>6.0f}  "
                f"{t['Trench']:>7.0f}  {t['Plow']:>5.0f}  {t['Aerial']:>6.0f}  "
                f"{t['Cable']:>5.0f}  {t['Drops']:>5.0f}  {total:>7.0f}"
            )

    if not dry_run:
        print(f"  Saved -> {output_crew.name}")

    print(f"\n{sep}")
    print("  Done." if not dry_run else "  Dry run complete. No files written.")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
