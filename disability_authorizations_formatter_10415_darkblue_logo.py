import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import pytz
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ----------------------------
# Page Config
# ----------------------------
st.set_page_config(page_title="HCHSP — Disability Authorizations Formatter", layout="wide")

# ----------------------------
# Header (UI ONLY)
# ----------------------------
logo_path = Path("header_logo.png")
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown(
        """
        <h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Disability Authorizations</h1>
        <p style='text-align:center; font-size:16px; margin-top:0;'>
        Upload the <b>10415 Authorization</b> export and download the cleaned, formatted workbook.
        </p>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ----------------------------
# File Upload
# ----------------------------
up = st.file_uploader("Upload *10415*.xlsx", type=["xlsx"], key="qf")

# ----------------------------
# Styles & Constants
# ----------------------------
BLUE = "1F4E78"   # darker header blue
WHITE = "FFFFFF"
GREEN = "008000"  # valid dates font (green)
RED = "C00000"    # missing dates X (red)

THIN = Side(style="thin", color="000000")
MED  = Side(style="medium", color="000000")


def _rename_columns(cols):
    """Standardize column names for 10415 variants."""
    mapping = {}
    for c in cols:
        s = str(c).strip()
        if re.search(r"authorization:\s*regarding my child", s, re.I):
            mapping[c] = "Child Name"
        elif re.search(r"authorization:\s*date", s, re.I):
            mapping[c] = "Authorization Date"
        elif re.search(r"IEP/IFSP\s*Dis:Identified", s, re.I):
            mapping[c] = "Disability Identified"
        elif re.search(r"primary\s*disability", s, re.I):
            mapping[c] = "Primary Disability"
        elif re.search(r"\bcenter name\b|\bcenter\b", s, re.I):
            mapping[c] = "Center"
        elif re.search(r"\bclass name\b|\bclass\b", s, re.I):
            mapping[c] = "Class"
        elif re.search(r"\bparticipant pid\b|\bpid\b", s, re.I):
            mapping[c] = "PID"
        elif re.search(r"\bfirst name\b", s, re.I):
            mapping[c] = "First Name"
        elif re.search(r"\blast name\b", s, re.I):
            mapping[c] = "Last Name"
        else:
            mapping[c] = s
    return [mapping[c] for c in cols]


def _split_child_name(df: pd.DataFrame) -> pd.DataFrame:
    """If only 'Child Name' exists, split into First/Last."""
    if "Child Name" in df.columns:
        def split_name(full):
            if pd.isna(full):
                return pd.Series({"First Name": "", "Last Name": ""})
            parts = str(full).split()
            if len(parts) == 1:
                return pd.Series({"First Name": parts[0], "Last Name": ""})
            return pd.Series({"First Name": " ".join(parts[:-1]), "Last Name": parts[-1]})
        name_split = df["Child Name"].apply(split_name)
        for col in ["First Name", "Last Name"]:
            if col not in df.columns:
                df[col] = name_split[col]
    return df


def _autosize_columns(ws, header_row=3):
    """Autosize columns using header + data length (avoid merged cells)."""
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        hdr_val = ws.cell(row=header_row, column=col_idx).value
        max_len = len(str(hdr_val)) if hdr_val is not None else 0
        for r in range(header_row + 1, ws.max_row + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[letter].width = min(max_len + 2, 45)


def _build_workbook(df: pd.DataFrame, title_text: str, logo: Path | None) -> bytes:
    """Return formatted workbook bytes (dark blue header, centered title, logo in A1)."""
    # Normalize date column to MM/DD/YYYY strings
    if "Authorization Date" in df.columns:
        dt = pd.to_datetime(df["Authorization Date"], errors="coerce")
        df["Authorization Date"] = dt.dt.strftime("%m/%d/%Y").fillna("")

    wb = Workbook()
    ws = wb.active
    ws.title = "Authorizations"

    # Write header row at row 3
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=3, column=j, value=col)
        c.font = Font(bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center")

    # Write data starting row 4
    for i, row in enumerate(df.itertuples(index=False), start=4):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=val)

    # Title row (row 2)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=df.shape[1])
    tcell = ws.cell(row=2, column=1, value=title_text)
    tcell.font = Font(bold=True, size=14)
    tcell.alignment = Alignment(horizontal="center", vertical="center")

    # Logo (row 1, col A)
    if logo and logo.exists():
        try:
            img = XLImage(str(logo))
            img.anchor = "A1"
            ws.add_image(img)
            ws.row_dimensions[1].height = 60
        except Exception:
            pass  # if Pillow or image fails, still deliver workbook

    # Borders (grid)
    max_row, max_col = ws.max_row, ws.max_column
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            left   = MED if c == 1 else THIN
            right  = MED if c == max_col else THIN
            top    = MED if r == 2 else THIN
            bottom = MED if r == max_row else THIN
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    # Freeze below header & add filter
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = ws.dimensions

    # Authorization Date styling: red ✗ if blank, green if present
    if "Authorization Date" in df.columns:
        date_idx = list(df.columns).index("Authorization Date") + 1
        for r in range(4, ws.max_row + 1):
            cell = ws.cell(row=r, column=date_idx)
            val = cell.value
            if val in (None, ""):
                cell.value = "✗"
                cell.font = Font(bold=True, color=RED)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = Font(color=GREEN)

    _autosize_columns(ws, header_row=3)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def _detect_header_row(raw_df: pd.DataFrame) -> int:
    """Detect header row by content in first column."""
    first_col = raw_df.iloc[:, 0].astype(str).str.strip().str.lower()
    if first_col.str.contains(r"authorization:\s*regarding my child", na=False).any():
        return first_col.str.contains(r"authorization:\s*regarding my child", na=False).idxmax()
    if first_col.str.contains("participant pid", na=False).any():
        return first_col.str.contains("participant pid", na=False).idxmax()
    # fallback: densest row
    return raw_df.notna().sum(axis=1).idxmax()


# ----------------------------
# Main
# ----------------------------
if up:
    safe_name = getattr(up, "name", "") or ""
    if "10415" not in safe_name:
        st.error("Please upload the correct file: filename must include **10415**.")
        st.stop()

    raw = pd.read_excel(up, header=None)
    hdr_row = _detect_header_row(raw)

    df = pd.read_excel(up, header=hdr_row).dropna(how="all")
    df.columns = _rename_columns(df.columns)
    df = _split_child_name(df)

    # Column order: Center immediately after Last Name
    preferred = ["PID", "First Name", "Last Name", "Center", "Class",
                 "Authorization Date", "Disability Identified", "Primary Disability"]
    existing = [c for c in preferred if c in df.columns]
    df = df[existing]

    tz = pytz.timezone("America/Chicago")
    now_str = datetime.now(tz).strftime("%m/%d/%Y %I:%M %p")
    fixed_title = f"25-26 Disability Authorizations — Exported {now_str}"

    xlsx = _build_workbook(df, fixed_title, logo_path if logo_path.exists() else None)

    st.success("File processed successfully. Click below to download.")
    st.download_button(
        "⬇️ Download Disability Authorizations (.xlsx)",
        data=xlsx,
        file_name=f"HCHSP_DisabilityAuthorizations_{datetime.now(tz).strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload the **10415** export (.xlsx) to begin.")
