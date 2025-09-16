import streamlit as st
import pandas as pd
from io import BytesIO
import pytz
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image

# Utility: split child full name into first and last names (last word = last name)
def split_name(full_name):
    if pd.isna(full_name):
        return "", ""
    parts = str(full_name).strip().split()
    if len(parts) == 0:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return " ".join(parts[:-1]), parts[-1]

# Build formatted workbook
def build_workbook(df, logo_path=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Disability Authorizations"

    # Title
    title = "Hidalgo County Head Start — Disability Authorizations"
    ws.merge_cells("A1:H1")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Timestamp
    cst = pytz.timezone("America/Chicago")
    ts = datetime.now(cst).strftime("%m/%d/%Y %I:%M %p")
    ws.merge_cells("A2:H2")
    ws["A2"] = f"Generated: {ts}"
    ws["A2"].alignment = Alignment(horizontal="center")

    startrow = 4

    # Write headers
    for j, col in enumerate(df.columns, 1):
        cell = ws.cell(row=startrow, column=j, value=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    # Borders
    thin = Side(style="thin")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # Write data
    for i, row in df.iterrows():
        for j, val in enumerate(row, 1):
            cell = ws.cell(row=startrow + 1 + i, column=j, value=val)
            cell.border = border
            if "Date" in df.columns[j-1]:
                if pd.isna(val) or val == "":
                    cell.value = "✗"
                    cell.font = Font(color="FF0000")
                else:
                    cell.font = Font(color="006100")

    # Auto column width
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for c in col:
            if c.value:
                max_len = max(max_len, len(str(c.value)))
        ws.column_dimensions[col_letter].width = max_len + 2

    # Freeze
    ws.freeze_panes = f"A{startrow+1}"

    # Logo
    if logo_path:
        try:
            img = Image(logo_path)
            img.height = 60
            img.width = 120
            ws.add_image(img, "A1")
        except Exception:
            pass

    return wb

# Streamlit app
st.set_page_config(page_title="Disability Authorizations Formatter (10415)", layout="centered")
st.image("header_logo.png", width=160)
st.markdown(
    """
    <h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Disability Authorizations</h1>
    <p style='text-align:center; font-size:16px; margin-top:0;'>
    Upload the <b>10415 Authorization</b> export and download the cleaned, formatted workbook.
    </p>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("Upload 10415 Disability Authorizations Excel", type=["xlsx"])

if uploaded:
    if "10415" not in uploaded.name:
        st.error("❌ This is not a 10415 file. Please upload the correct export.")
    else:
        raw = pd.read_excel(uploaded, header=None)
        first_col = raw.iloc[:,0].astype(str).str.strip().str.lower()

        if first_col.str.contains("authorization: regarding my child", na=False).any():
            hdr_row = first_col.str.contains("authorization: regarding my child", na=False).idxmax()
        elif first_col.str.contains("participant pid", na=False).any():
            hdr_row = first_col.str.contains("participant pid", na=False).idxmax()
        else:
            hdr_row = raw.notna().sum(axis=1).idxmax()

        df = pd.read_excel(uploaded, header=hdr_row)

        # Normalize columns
        rename_map = {
            "25-26 Authorization: Regarding my child": "Child Name",
            "25-26 Authorization: Date": "Authorization Date",
            "IEP/IFSP Dis:Identified": "Disability Identified",
            "IEP/IFSP Dis:Primary Disability": "Primary Disability",
            "Participant PID": "PID",
            "Center": "Center"
        }
        df = df.rename(columns={c: rename_map.get(c,c) for c in df.columns})

        # Split child name
        if "Child Name" in df.columns:
            df["First Name"], df["Last Name"] = zip(*df["Child Name"].map(split_name))
            df.drop(columns=["Child Name"], inplace=True)

        # Reorder columns
        col_order = []
        if "PID" in df.columns:
            col_order.append("PID")
        if "Last Name" in df.columns:
            col_order.append("Last Name")
        if "Center" in df.columns:
            col_order.append("Center")
        if "First Name" in df.columns:
            col_order.append("First Name")
        for c in df.columns:
            if c not in col_order:
                col_order.append(c)
        df = df[col_order]

        # Build workbook
        wb = build_workbook(df, logo_path="header_logo.png")
        output = BytesIO()
        wb.save(output)

        st.success("✅ File formatted successfully!")
        st.download_button(
            "📥 Download Formatted Disability Authorizations",
            data=output.getvalue(),
            file_name=f"HCHSP_DisabilityAuthorizations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
