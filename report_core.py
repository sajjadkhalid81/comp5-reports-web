"""
report_core.py  —  COMP5 QatarEnergy LNG / Saipem JV
Report generation logic for:
  1. TQ & SDR Weekly Report  (sheets: TQY + SDR from COMP5_-_TQ-SDR-CR-ARN-DCR_LOG.xlsx)
  2. Weekly COMP5 Issued Documents Report  (sheet: Issued Documents from COMP5_REGISTER.xlsx)
"""

from __future__ import annotations
from io import BytesIO
from datetime import datetime, date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ── Palette ────────────────────────────────────────────────────────────────
WHITE      = "FFFFFF"
BLACK      = "000000"
DARK_BLUE  = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "BDD7EE"
HEADER_GREY = "D9D9D9"

THIN = Side(border_style="thin", color="BFBFBF")
MED  = Side(border_style="medium", color="595959")

def _border(all_thin=True):
    s = THIN if all_thin else MED
    return Border(left=s, right=s, top=s, bottom=s)

def _cell(ws, row, col, value="", bold=False, color=BLACK, bg=None,
          align="left", wrap=False, size=9, border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", bold=bold, color=color, size=size)
    c.alignment = Alignment(horizontal=align, vertical="center",
                             wrap_text=wrap)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if border:
        c.border = _border()
    return c

def _hdr(ws, row, col, label, bg=DARK_BLUE, fg=WHITE, size=9):
    return _cell(ws, row, col, label, bold=True, color=fg, bg=bg,
                 align="center", wrap=True, size=size)

def _fmt_date(v):
    if v is None or (isinstance(v, float) and str(v) == "nan"):
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d-%b-%Y")
    return str(v).strip()

def _safe(v):
    if v is None:
        return ""
    if isinstance(v, float) and str(v) == "nan":
        return ""
    if isinstance(v, (datetime, date)):
        return _fmt_date(v)
    return str(v).strip()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — TQ & SDR REPORT
# ══════════════════════════════════════════════════════════════════════════════

TODAY = datetime.today().date()
DATE_STR = datetime.today().strftime("%d%b%Y").upper()


def _read_tqy(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(raw), sheet_name="TQY", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    # Normalise date columns
    for col in df.columns:
        if "date" in col.lower() or "due" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _read_sdr(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(raw), sheet_name="SDR", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        if "date" in col.lower() or "due" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _find_col(df: pd.DataFrame, *candidates: str) -> str | None:
    """Return first matching column name (case-insensitive partial match)."""
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        for k, v in cols_lower.items():
            if cand.lower() in k:
                return v
    return None


def _split_tqy(df: pd.DataFrame):
    """Split TQY dataframe into 5 buckets."""
    # Detect key columns flexibly
    reply_col  = _find_col(df, "date replied", "replied")
    status_col = _find_col(df, "cpy response", "status", "response")
    due_col    = _find_col(df, "due date", "respond due")
    issued_col = _find_col(df, "date issued", "date issue")

    mask_replied = df[reply_col].notna() if reply_col else pd.Series(False, index=df.index)
    mask_closed  = pd.Series(False, index=df.index)
    if status_col:
        mask_closed = mask_replied & df[status_col].astype(str).str.upper().str.contains("CLOS|ACCEPT|APPROV", na=False)

    ds_issued  = df.copy()   # ALL issued
    ds_closed  = df[mask_closed].copy() if status_col else df.iloc[0:0].copy()
    ds_repopen = df[mask_replied & ~mask_closed].copy()

    mask_open = ~mask_replied
    if due_col:
        due_dates = pd.to_datetime(df[due_col], errors="coerce")
        ds_valid   = df[mask_open & (due_dates.dt.date >= TODAY)].copy()
        ds_expired = df[mask_open & (due_dates.dt.date < TODAY)].copy()
        # Add days remaining
        ds_valid["Days Remaining"]   = (due_dates[ds_valid.index].dt.date.apply(
            lambda d: (d - TODAY).days if pd.notna(d) else ""))
        ds_expired["Days Remaining"] = (due_dates[ds_expired.index].dt.date.apply(
            lambda d: (TODAY - d).days if pd.notna(d) else ""))
    else:
        ds_valid   = df[mask_open].copy()
        ds_expired = df.iloc[0:0].copy()

    return ds_issued, ds_valid, ds_expired, ds_closed, ds_repopen


def _split_sdr(df: pd.DataFrame):
    """Split SDR dataframe into 5 buckets."""
    reply_col  = _find_col(df, "date replied", "replied")
    status_col = _find_col(df, "status", "company response", "response")
    due_col    = _find_col(df, "respond due", "due date")

    mask_replied = df[reply_col].notna() if reply_col else pd.Series(False, index=df.index)
    mask_closed  = pd.Series(False, index=df.index)
    if status_col:
        mask_closed = mask_replied & df[status_col].astype(str).str.upper().str.contains("CLOS|ACCEPT|APPROV", na=False)

    ds_issued  = df.copy()
    ds_closed  = df[mask_closed].copy() if status_col else df.iloc[0:0].copy()
    ds_repopen = df[mask_replied & ~mask_closed].copy()

    mask_open = ~mask_replied
    if due_col:
        due_dates = pd.to_datetime(df[due_col], errors="coerce")
        ds_valid   = df[mask_open & (due_dates.dt.date >= TODAY)].copy()
        ds_expired = df[mask_open & (due_dates.dt.date < TODAY)].copy()
        ds_valid["Days Remaining"]   = due_dates[ds_valid.index].dt.date.apply(
            lambda d: (d - TODAY).days if pd.notna(d) else "")
        ds_expired["Days Remaining"] = due_dates[ds_expired.index].dt.date.apply(
            lambda d: (TODAY - d).days if pd.notna(d) else "")
    else:
        ds_valid   = df[mask_open].copy()
        ds_expired = df.iloc[0:0].copy()

    return ds_issued, ds_valid, ds_expired, ds_closed, ds_repopen


# ── Excel builders ─────────────────────────────────────────────────────────

def _build_kpi_summary(wb: Workbook, title: str, header_color: str,
                        kpis: list[tuple]) -> None:
    """Add a Summary sheet with KPI boxes."""
    ws = wb.create_sheet("SUMMARY")
    ws.sheet_properties.tabColor = header_color

    # Title row
    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = title
    c.font  = Font(name="Calibri", bold=True, size=14, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=header_color)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # Date row
    ws.merge_cells("A2:F2")
    d = ws["A2"]
    d.value = f"Report Date: {datetime.today().strftime('%d %B %Y')}"
    d.font  = Font(name="Calibri", italic=True, size=10, color="595959")
    d.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # KPI boxes — two per row
    BOX_COLS = [(1, 2), (3, 4), (5, 6)]
    row = 4
    for i, (label, count, fg, bg) in enumerate(kpis):
        col_start, col_end = BOX_COLS[i % 3]
        if i % 3 == 0 and i > 0:
            row += 7

        ws.merge_cells(start_row=row,   start_column=col_start,
                       end_row=row+2,   end_column=col_end)
        ws.merge_cells(start_row=row+3, start_column=col_start,
                       end_row=row+5,   end_column=col_end)

        lbl_cell = ws.cell(row=row, column=col_start, value=label)
        lbl_cell.font      = Font(name="Calibri", bold=True, size=10, color=fg)
        lbl_cell.fill      = PatternFill("solid", fgColor=bg)
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)

        cnt_cell = ws.cell(row=row+3, column=col_start, value=count)
        cnt_cell.font      = Font(name="Calibri", bold=True, size=28, color=fg)
        cnt_cell.fill      = PatternFill("solid", fgColor=bg)
        cnt_cell.alignment = Alignment(horizontal="center", vertical="center")

        for r in range(row, row+6):
            for col in range(col_start, col_end+1):
                ws.cell(r, col).border = _border()

    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 18
    for r in range(3, row+8):
        ws.row_dimensions[r].height = 16


def _build_data_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame,
                       tab_color: str, extra_col: str | None = None) -> None:
    """Add a data sheet with formatted table."""
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = tab_color
    ws.freeze_panes = "A3"

    if df.empty:
        ws.cell(2, 1, "No records for this category").font = Font(italic=True, color="595959")
        return

    cols = list(df.columns)
    if extra_col and extra_col in cols:
        # Move extra col to end
        cols = [c for c in cols if c != extra_col] + [extra_col]

    # Header
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    h = ws.cell(1, 1, sheet_name.upper())
    h.font      = Font(name="Calibri", bold=True, size=11, color=WHITE)
    h.fill      = PatternFill("solid", fgColor=DARK_BLUE)
    h.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    for ci, col in enumerate(cols, 1):
        _hdr(ws, 2, ci, col, bg=tab_color)
        ws.column_dimensions[get_column_letter(ci)].width = max(12, min(35, len(str(col))+4))
    ws.row_dimensions[2].height = 30

    ALT = "EBF3FB"
    for ri, (_, row_data) in enumerate(df[cols].iterrows(), 3):
        bg = "FFFFFF" if ri % 2 == 0 else ALT
        for ci, col in enumerate(cols, 1):
            val = row_data[col]
            if pd.isna(val):
                val = ""
            elif isinstance(val, (datetime, date)):
                val = _fmt_date(val)
            _cell(ws, ri, ci, val, bg=bg)
        ws.row_dimensions[ri].height = 14

    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}2"


def _wb_to_bytes(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def generate_tq_sdr(raw: bytes) -> dict:
    """
    Accepts raw bytes of COMP5_-_TQ-SDR-CR-ARN-DCR_LOG.xlsx.
    Returns dict with keys: tqy_bytes, tqy_name, sdr_bytes, sdr_name, summary.
    """
    tqy_df = _read_tqy(raw)
    sdr_df = _read_sdr(raw)

    # --- TQY ---
    tqy_issued, tqy_valid, tqy_expired, tqy_closed, tqy_repopen = _split_tqy(tqy_df)

    tqy_kpis = [
        ("TQ ISSUED\n(by Sai)",              len(tqy_issued),  "1F3864", "BDD7EE"),
        ("TQ NOT REPLIED\n(Not Expired)",    len(tqy_valid),   "1F7A3C", "C6EFCE"),
        ("TQ NOT REPLIED\n(Expired)",        len(tqy_expired), "C00000", "FFCCCC"),
        ("TQ REPLIED\nCLOSED",               len(tqy_closed),  "595959", "E0E0E0"),
        ("TQ REPLIED\nOPEN",                 len(tqy_repopen), "7030A0", "E2CFEF"),
    ]
    wb_tqy = Workbook()
    wb_tqy.remove(wb_tqy.active)
    _build_kpi_summary(wb_tqy, "TQ – TECHNICAL QUERY", "1F3864", tqy_kpis)
    _build_data_sheet(wb_tqy, "TQ ISSUED (by Sai)",           tqy_issued,  "1F3864")
    _build_data_sheet(wb_tqy, "TQ NOT REPLIED (Not Expired)", tqy_valid,   "1F7A3C", "Days Remaining")
    _build_data_sheet(wb_tqy, "TQ NOT REPLIED (Expired)",     tqy_expired, "C00000", "Days Remaining")
    _build_data_sheet(wb_tqy, "TQ REPLIED CLOSED",            tqy_closed,  "595959")
    _build_data_sheet(wb_tqy, "TQ REPLIED OPEN",              tqy_repopen, "7030A0")

    tqy_name  = f"COMP5_TQY_Report_{DATE_STR}.xlsx"
    tqy_bytes = _wb_to_bytes(wb_tqy)

    # --- SDR ---
    sdr_issued, sdr_valid, sdr_expired, sdr_closed, sdr_repopen = _split_sdr(sdr_df)

    sdr_kpis = [
        ("SDR ISSUED\n(by Sai)",             len(sdr_issued),  "7B3F00", "FCE4D6"),
        ("SDR NOT REPLIED\n(Not Expired)",   len(sdr_valid),   "1F7A3C", "C6EFCE"),
        ("SDR NOT REPLIED\n(Expired)",       len(sdr_expired), "C00000", "FFCCCC"),
        ("SDR REPLIED\nCLOSED",              len(sdr_closed),  "595959", "E0E0E0"),
        ("SDR REPLIED\nOPEN",                len(sdr_repopen), "7030A0", "E2CFEF"),
    ]
    wb_sdr = Workbook()
    wb_sdr.remove(wb_sdr.active)
    _build_kpi_summary(wb_sdr, "SDR – SPECIFICATION DEVIATION REQUEST", "7B3F00", sdr_kpis)
    _build_data_sheet(wb_sdr, "SDR ISSUED (by Sai)",           sdr_issued,  "7B3F00")
    _build_data_sheet(wb_sdr, "SDR NOT REPLIED (Not Expired)", sdr_valid,   "1F7A3C", "Days Remaining")
    _build_data_sheet(wb_sdr, "SDR NOT REPLIED (Expired)",     sdr_expired, "C00000", "Days Remaining")
    _build_data_sheet(wb_sdr, "SDR REPLIED CLOSED",            sdr_closed,  "595959")
    _build_data_sheet(wb_sdr, "SDR REPLIED OPEN",              sdr_repopen, "7030A0")

    sdr_name  = f"COMP5_SDR_Report_{DATE_STR}.xlsx"
    sdr_bytes = _wb_to_bytes(wb_sdr)

    summary = {
        "tqy": {
            "issued":  len(tqy_issued),
            "valid":   len(tqy_valid),
            "expired": len(tqy_expired),
            "closed":  len(tqy_closed),
            "repopen": len(tqy_repopen),
        },
        "sdr": {
            "issued":  len(sdr_issued),
            "valid":   len(sdr_valid),
            "expired": len(sdr_expired),
            "closed":  len(sdr_closed),
            "repopen": len(sdr_repopen),
        },
    }

    return {
        "tqy_bytes": tqy_bytes,
        "tqy_name":  tqy_name,
        "sdr_bytes": sdr_bytes,
        "sdr_name":  sdr_name,
        "summary":   summary,
    }


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — WEEKLY COMP5 ISSUED DOCUMENTS REPORT
# ══════════════════════════════════════════════════════════════════════════════

COMP5_DISCIPLINES = [
    "Civil", "Structural", "Piping", "Mechanical", "Electrical",
    "Instrument", "Telecom", "Safety", "Process", "HVAC",
    "Vendor", "HSE", "QA/QC", "Planning",
]

COMP5_DOC_TYPES = [
    "IFC", "IFR", "IFI", "IFD", "AFC", "AFI", "AFR",
    "FOR APPROVAL", "FOR INFORMATION", "FOR REVIEW",
]


def _read_comp5(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(raw), sheet_name="Issued Documents", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    date_col = _find_col(df, "date issued", "issue date")
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    return df


def _get_week_range() -> tuple[date, date]:
    """Return Monday–Sunday of the current week."""
    today = datetime.today().date()
    monday = today - __import__("datetime").timedelta(days=today.weekday())
    sunday = monday + __import__("datetime").timedelta(days=6)
    return monday, sunday


def _build_comp5_summary(wb: Workbook, df_week: pd.DataFrame,
                          df_all: pd.DataFrame,
                          w_start: date, w_end: date) -> None:
    ws = wb.create_sheet("SUMMARY")
    ws.sheet_properties.tabColor = "1F3864"

    title = f"COMP5 – WEEKLY ISSUED DOCUMENTS REPORT"
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = title
    c.font  = Font(name="Calibri", bold=True, size=14, color=WHITE)
    c.fill  = PatternFill("solid", fgColor=DARK_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    week_str = f"Week: {w_start.strftime('%d %b')} – {w_end.strftime('%d %b %Y')}"
    ws.merge_cells("A2:H2")
    d = ws["A2"]
    d.value = week_str
    d.font  = Font(name="Calibri", italic=True, size=10, color="595959")
    d.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # KPI row
    kpis = [
        ("DOCS THIS WEEK", len(df_week), DARK_BLUE, LIGHT_BLUE),
        ("TOTAL ISSUED",   len(df_all),  MID_BLUE,  "DDEEFF"),
    ]
    row = 4
    for i, (label, count, fg, bg) in enumerate(kpis):
        col = i * 3 + 1
        ws.merge_cells(start_row=row,   start_column=col, end_row=row+2,   end_column=col+2)
        ws.merge_cells(start_row=row+3, start_column=col, end_row=row+5,   end_column=col+2)

        lc = ws.cell(row, col, label)
        lc.font      = Font(name="Calibri", bold=True, size=11, color=fg)
        lc.fill      = PatternFill("solid", fgColor=bg)
        lc.alignment = Alignment(horizontal="center", vertical="center")

        nc = ws.cell(row+3, col, count)
        nc.font      = Font(name="Calibri", bold=True, size=28, color=fg)
        nc.fill      = PatternFill("solid", fgColor=bg)
        nc.alignment = Alignment(horizontal="center", vertical="center")

        for r in range(row, row+6):
            for c2 in range(col, col+3):
                ws.cell(r, c2).border = _border()

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16


def _build_comp5_sheet(wb: Workbook, sheet_name: str,
                        df: pd.DataFrame, tab_color: str) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = tab_color
    ws.freeze_panes = "A3"

    if df.empty:
        ws.cell(2, 1, "No documents for this period").font = Font(italic=True, color="595959")
        return

    cols = list(df.columns)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    h = ws.cell(1, 1, sheet_name.upper())
    h.font      = Font(name="Calibri", bold=True, size=11, color=WHITE)
    h.fill      = PatternFill("solid", fgColor=DARK_BLUE)
    h.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    for ci, col in enumerate(cols, 1):
        _hdr(ws, 2, ci, col, bg=tab_color)
        ws.column_dimensions[get_column_letter(ci)].width = max(12, min(40, len(str(col))+4))
    ws.row_dimensions[2].height = 30

    ALT = "EBF3FB"
    for ri, (_, row_data) in enumerate(df.iterrows(), 3):
        bg = "FFFFFF" if ri % 2 == 0 else ALT
        for ci, col in enumerate(cols, 1):
            val = row_data[col]
            if pd.isna(val):
                val = ""
            elif isinstance(val, (datetime, date)):
                val = _fmt_date(val)
            _cell(ws, ri, ci, val, bg=bg)
        ws.row_dimensions[ri].height = 14

    ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}2"


def generate_comp5(raw: bytes) -> dict:
    """
    Accepts raw bytes of COMP5_REGISTER.xlsx (sheet: Issued Documents).
    Returns dict with keys: bytes, filename, summary.
    """
    df = _read_comp5(raw)
    date_col = _find_col(df, "date issued", "issue date")

    w_start, w_end = _get_week_range()

    if date_col:
        dates = df[date_col].dt.date
        df_week = df[(dates >= w_start) & (dates <= w_end)].copy()
    else:
        df_week = df.copy()  # fallback: all records

    df_all = df.copy()

    wb = Workbook()
    wb.remove(wb.active)

    _build_comp5_summary(wb, df_week, df_all, w_start, w_end)
    _build_comp5_sheet(wb, "THIS WEEK ISSUED", df_week, "1F3864")
    _build_comp5_sheet(wb, "ALL ISSUED DOCUMENTS", df_all, MID_BLUE)

    # Discipline breakdown tab
    disc_col = _find_col(df, "discipline", "disc")
    if disc_col and not df_week.empty:
        disc_counts = df_week[disc_col].value_counts().reset_index()
        disc_counts.columns = ["Discipline", "Count This Week"]
        _build_comp5_sheet(wb, "BY DISCIPLINE (This Week)", disc_counts, "375623")

    filename  = f"COMP5_Weekly_IssuedDocs_{DATE_STR}.xlsx"
    out_bytes = _wb_to_bytes(wb)

    summary = {
        "week":  f"{w_start.strftime('%d %b')} – {w_end.strftime('%d %b %Y')}",
        "count_week": len(df_week),
        "count_all":  len(df_all),
    }

    return {"bytes": out_bytes, "filename": filename, "summary": summary}
