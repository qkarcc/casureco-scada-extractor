"""
automate_report.py
==================
CASURECO IV — Distribution Lines & Power Quality Report Automation

Reads all available feeder SCADA files and writes extracted values
to a clean Excel file: "output_MONTH_YEAR.xlsx"

The output Excel has:
  1. Main data table  — all extracted values per feeder
  2. Source reference table (below) — exact cell address in TrendPage01_stat
     for every value, so the user can open the input file and verify.

Rules:
  - Inst. Active Power  "L" -> MAX  (TrendPage01_stat row 5, col L)
  - Total Reactive Power "M" -> MAX  (TrendPage01_stat row 5, col M)
  - Inst. Reactive Power "N" -> MAX  (TrendPage01_stat row 5, col N)
  - Off-Peak Voltage Va/Vb/Vc -> MIN (TrendPage01_stat row 4, col E/F/G)
  - Peak Voltage     Va/Vb/Vc -> MAX (TrendPage01_stat row 5, col E/F/G)
  - Off-Peak Current Ia/Ib/Ic/In -> MIN (row 4, col H/I/J/K)
  - Peak Current     Ia/Ib/Ic/In -> MAX (row 5, col H/I/J/K)
"""

import os
import math
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ===========================================================================
# CONFIG
# ===========================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_FILE = os.path.join(
    BASE_DIR,
    "DISTRIBUTION LINES AND POWER QUALITY (INPUT NUMBERS).xlsx"
)

FEEDERS = {
    "LF1": "Lagonoy F1.xlsx",
    "LF2": "Lagonoy F2.xlsx",
    "LF3": "Lagonoy F3.xlsx",
    # "LF4": "Lagonoy F4.xlsx",
    "PF1": "Presentacion F1.xlsx",
    "PF2": "Presentacion F2.xlsx",
    "PF3": "Presentacion F3.xlsx",
    "TF1": "Tigaon F1.xlsx",
    "TF2": "Tigaon F2.xlsx",
    "TF3": "Tigaon F3.xlsx",
}

OUTPUT_COL_ORDER = ["LF1", "LF2", "LF3", "LF4", "PF1", "PF2", "TF1", "TF2", "TF3"]

# ===========================================================================
# FIELD MAP  (TrendPage01 raw data sheet)
#   Header row is at Excel row 8 (pandas header=7)
#   Data starts at Excel row 9  (pandas row index 0 = Excel row 9)
#   EXCEL_ROW = pandas_idx + 9
#
#   Column letters in TrendPage01:
#     E=Voltage A,  F=Voltage B,  G=Voltage C
#     H=Current A,  I=Current B,  J=Current C,  K=Current N
#     L=Instantaneous Active Power
#     M=Total Reactive Power
#     N=Instantaneous Reactive Power
# ===========================================================================

TREND_HEADER_ROW = 8    # Excel row where headers live in TrendPage01
TREND_DATA_START = 9    # first Excel row of actual data

# field_key -> (col_letter, col_name, rule)
FIELD_META = {
    "peak_demand":    ("L", "Instantaneous Active Power",   "MAX"),
    "total_reactive": ("M", "Total Reactive Power",          "MAX"),
    "inst_reactive":  ("N", "Instantaneous Reactive Power",  "MAX"),
    "offpeak_va":     ("E", "Voltage A",                     "MIN"),
    "offpeak_vb":     ("F", "Voltage B",                     "MIN"),
    "offpeak_vc":     ("G", "Voltage C",                     "MIN"),
    "peak_va":        ("E", "Voltage A",                     "MAX"),
    "peak_vb":        ("F", "Voltage B",                     "MAX"),
    "peak_vc":        ("G", "Voltage C",                     "MAX"),
    "offpeak_ia":     ("H", "Current A",                     "MIN"),
    "offpeak_ib":     ("I", "Current B",                     "MIN"),
    "offpeak_ic":     ("J", "Current C",                     "MIN"),
    "offpeak_in":     ("K", "Current N",                     "MIN"),
    "peak_ia":        ("H", "Current A",                     "MAX"),
    "peak_ib":        ("I", "Current B",                     "MAX"),
    "peak_ic":        ("J", "Current C",                     "MAX"),
    "peak_in":        ("K", "Current N",                     "MAX"),
}

# ===========================================================================
# ROW DEFINITIONS  (label, field_key)
# ===========================================================================

ROWS = [
    ('Inst. Active Power "L"  [PEAK DEMAND KW]',  "peak_demand"),
    ('Total Reactive Power "M" [KWhr]',            "total_reactive"),
    ('Inst. Reactive Power "N" [KVARh]',           "inst_reactive"),
    ("Off-Peak Voltage  Va",                        "offpeak_va"),
    ("Off-Peak Voltage  Vb",                        "offpeak_vb"),
    ("Off-Peak Voltage  Vc",                        "offpeak_vc"),
    ("Peak Voltage      Va",                        "peak_va"),
    ("Peak Voltage      Vb",                        "peak_vb"),
    ("Peak Voltage      Vc",                        "peak_vc"),
    ("Off-Peak Current  Ia",                        "offpeak_ia"),
    ("Off-Peak Current  Ib",                        "offpeak_ib"),
    ("Off-Peak Current  Ic",                        "offpeak_ic"),
    ("Off-Peak Current  In",                        "offpeak_in"),
    ("Peak Current      Ia",                        "peak_ia"),
    ("Peak Current      Ib",                        "peak_ib"),
    ("Peak Current      Ic",                        "peak_ic"),
    ("Peak Current      In",                        "peak_in"),
]

# ===========================================================================
# EXTRACTION
# ===========================================================================

def safe_val(v, decimals=2):
    try:
        f = float(v)
        if math.isnan(f) or math.isinf(f):
            return "NS"
        return round(f, decimals)
    except (TypeError, ValueError):
        return "NS"


def extract(filepath):
    """
    Read TrendPage01 raw data directly.
    Returns dict: field_key -> {"value": float|"NS", "cell": "L9847"|"NS"}
    Header is at Excel row 8 (pandas header=7).
    Data starts at Excel row 9 -> pandas_idx 0.
    Excel row = pandas_idx + TREND_DATA_START
    """
    print(f"  Reading {os.path.basename(filepath)} ...", end="", flush=True)
    xl = pd.ExcelFile(filepath)
    df = pd.read_excel(xl, sheet_name="TrendPage01", header=TREND_HEADER_ROW - 1)
    # Drop completely empty rows
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)

    result = {}
    for field_key, (col_ltr, col_name, rule) in FIELD_META.items():
        if col_name not in df.columns:
            result[field_key] = {"value": "NS", "cell": "NS"}
            continue
        series = pd.to_numeric(df[col_name], errors="coerce").dropna()
        if series.empty:
            result[field_key] = {"value": "NS", "cell": "NS"}
            continue

        if rule == "MAX":
            pandas_idx = series.idxmax()
            raw_val    = series[pandas_idx]
        else:
            # Off-peak MIN: exclude zero/negative values (outages, no-load)
            if "Current" in col_name:
                # Off-peak current: also exclude values < 10 (noise/near-zero)
                live = series[series >= 10]
            else:
                # Off-peak voltage: exclude zero/negative only
                live = series[series > 0]

            if live.empty:
                result[field_key] = {"value": "NS", "cell": "NS"}
                continue
            pandas_idx = live.idxmin()
            raw_val    = live[pandas_idx]

        # Convert pandas row index -> Excel row number
        excel_row  = int(pandas_idx) + TREND_DATA_START
        cell_addr  = f"{col_ltr}{excel_row}"
        result[field_key] = {"value": safe_val(raw_val), "cell": cell_addr}

    print(" OK")
    return result


def all_ns():
    return {k: {"value": "NS", "cell": "NS"} for k in FIELD_META}


# ===========================================================================
# EXCEL BUILDER
# ===========================================================================

def build_excel(month_label, table):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = month_label[:31]

    # ---- Styles ------------------------------------------------------------
    title_font  = Font(bold=True, size=13)
    hdr_font    = Font(bold=True, size=11, color="FFFFFF")
    bold_font   = Font(bold=True, size=11)
    src_font    = Font(size=10,   italic=True, color="444444")
    ref_font    = Font(bold=True, size=11, color="C00000")   # red cell refs
    hdr_fill    = PatternFill("solid", fgColor="1F4E79")     # dark blue
    src_hdr_fill= PatternFill("solid", fgColor="375623")     # dark green
    alt_fill    = PatternFill("solid", fgColor="D9E1F2")     # light blue
    alt2_fill   = PatternFill("solid", fgColor="E2EFDA")     # light green (source table)
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left        = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin        = Side(style="thin")
    med         = Side(style="medium")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    N_FEED      = len(OUTPUT_COL_ORDER)
    # Col layout: 1=FIELD, 2..N+1=feeders
    FEED_START  = 2
    TOTAL_COLS  = 1 + N_FEED

    def set_cell(row, col, value, font=None, fill=None, align=None, bdr=None):
        c = ws.cell(row=row, column=col, value=value)
        if font:  c.font      = font
        if fill:  c.fill      = fill
        if align: c.alignment = align
        if bdr:   c.border    = bdr
        return c

    def hdr_cell(row, col, val, fill_style=None):
        f = fill_style or hdr_fill
        c = set_cell(row, col, val, font=hdr_font, fill=f, align=center, bdr=border)
        return c

    # ========================================================================
    # TABLE — Main data table
    # ========================================================================

    HDR_ROW = 2

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TOTAL_COLS)
    set_cell(1, 1,
             f"DISTRIBUTION LINES & POWER QUALITY  —  {month_label}  "
             f"(Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')})",
             font=title_font, align=center)

    # Header row: FIELD | LF1 … TF3
    hdr_cell(HDR_ROW, 1, "FIELD").alignment = left
    for i, code in enumerate(OUTPUT_COL_ORDER):
        hdr_cell(HDR_ROW, FEED_START + i, code)

    ws.freeze_panes = "A3"

    # Data rows
    for r_idx, (label, key) in enumerate(ROWS):
        xrow = HDR_ROW + 1 + r_idx
        fill = alt_fill if r_idx % 2 == 0 else None

        # Col 1 — field label
        set_cell(xrow, 1, label, font=bold_font, fill=fill, align=left, bdr=border)

        # Cols 2+ — feeder values
        for i, code in enumerate(OUTPUT_COL_ORDER):
            entry = table.get(code, {}).get(key, {"value": "NS", "cell": "NS"})
            val   = entry["value"] if isinstance(entry, dict) else entry
            set_cell(xrow, FEED_START + i, val, fill=fill, align=center, bdr=border)

    # ========================================================================
    # COLUMN WIDTHS & ROW HEIGHTS
    # ========================================================================

    ws.column_dimensions["A"].width = 42
    for i in range(N_FEED):
        ws.column_dimensions[get_column_letter(FEED_START + i)].width = 11

    ws.row_dimensions[1].height       = 24
    ws.row_dimensions[HDR_ROW].height = 22

    return wb


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    print("=" * 65)
    print("  CASURECO IV -- Distribution Lines & Power Quality Extractor")
    print("=" * 65)

    month_label = input(
        "\nEnter month label (e.g. APRIL 2026): "
    ).strip().upper()
    if not month_label:
        month_label = datetime.today().strftime("%B %Y").upper()
        print(f"  Using: {month_label}")

    # Extract from each feeder
    print("\nReading feeder files...")
    table = {}
    for code in OUTPUT_COL_ORDER:
        fname = FEEDERS.get(code)
        if fname is None:
            print(f"  {code}: not configured -> NS")
            table[code] = all_ns()
            continue
        fpath = os.path.join(BASE_DIR, fname)
        if not os.path.exists(fpath):
            print(f"  {code}: {fname} not found -> NS")
            table[code] = all_ns()
            continue
        table[code] = extract(fpath)

    # Print quick preview
    print(f"\nExtracted values preview:")
    for label, key in ROWS:
        col_ltr = FIELD_META[key][0]
        rule    = FIELD_META[key][2]
        parts   = []
        for code in OUTPUT_COL_ORDER:
            entry = table.get(code, {}).get(key, {"value": "NS", "cell": "NS"})
            if isinstance(entry, dict) and entry.get("value", "NS") != "NS":
                parts.append(f"{code}={entry['value']} (cell {entry['cell']})") 
        preview = "  ".join(parts) or "(all NS)"
        print(f"  [{rule} col {col_ltr}] {label[:36]:<36}  {preview}")

    # Build and save Excel
    print("\nBuilding Excel output...")
    wb = build_excel(month_label, table)

    safe_month   = month_label.replace(" ", "_")
    out_filename = f"output_{safe_month}.xlsx"
    out_path     = os.path.join(BASE_DIR, out_filename)
    wb.save(out_path)

    print(f"\nSaved: {out_path}")
    print("\n" + "=" * 55)
    input("  Done! Press ENTER to exit...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERROR: {e}")
        input("\nPress ENTER to exit...")
