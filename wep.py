import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer



import streamlit as st

# -------- SECURITY --------
def check_password():
    st.title("🔐 Login Required")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "admin" and password == "matrix@123":
            st.session_state["authenticated"] = True
        else:
            st.error("Invalid username or password")

    return st.session_state.get("authenticated", False)


if not check_password():
    st.stop()
# --------------------------

# =========================
# Helpers (UNCHANGED)
# =========================
def format_inr_number(value):
    try:
        n = int(round(float(value)))
    except Exception:
        return ''
    s = str(abs(n))
    if len(s) <= 3:
        res = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.append(rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.append(rest)
        parts.reverse()
        res = ','.join(parts) + ',' + last3
    if n < 0:
        res = '-' + res
    return res

def normalize_col(col):
    return (
        str(col)
        .replace("\n", " ")
        .replace("\r", " ")
        .replace(".", "")
        .replace("_", " ")
        .strip()
        .lower()
    )

def find_column(df, candidates):
    norm_candidates = [normalize_col(c) for c in candidates]

    for col in df.columns:
        col_norm = normalize_col(col)
        if col_norm in norm_candidates:
            return col

    for col in df.columns:
        col_norm = normalize_col(col)
        for cand in norm_candidates:
            if cand in col_norm or col_norm in cand:
                return col

    return None


# =========================
# PDF Generator
# =========================
from io import BytesIO

def generate_pdf_from_excel(uploaded_file):

    df = pd.read_excel(uploaded_file, sheet_name=0)
    df.columns = df.columns.str.strip()
    report_name = uploaded_file.name.split(".")[0].upper()

    party_col = find_column(df, ["party's name", "party name", "customer name", "customer"])
    pending_col = find_column(df, ["pending amount", "pending_amount", "pending amt", "amount", "amount pending"])
    age_col = find_column(df, ["age of bill in days", "age of bill", "age_days", "age"])
    due_col = find_column(df, ["due days", "due_days", "due"])

    if not all([party_col, pending_col, age_col, due_col]):
        raise ValueError("Required columns not found in Excel")

    # Clean + convert
    df[party_col] = df[party_col].astype(str).str.strip()
    df[pending_col] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0)
    df[age_col] = pd.to_numeric(df[age_col], errors='coerce').fillna(0)
    df[due_col] = pd.to_numeric(df[due_col], errors='coerce').fillna(0)

    # Above 90
    above_90_df = df[df[age_col] > 89]
    above_90_sum = above_90_df.groupby(party_col, as_index=False)[pending_col].sum() \
        .rename(columns={pending_col: 'Above 90 days'})
    max_overdue = above_90_df.groupby(party_col, as_index=False)[age_col].max() \
        .rename(columns={age_col: 'Max Overdue days'})

    # Overdue
    overdue_df = df[(df[age_col] >= df[due_col]) & (df[age_col] <= 89)]
    overdue_sum = overdue_df.groupby(party_col, as_index=False)[pending_col].sum() \
        .rename(columns={pending_col: 'Overdue'})

    # Merge
    summary = pd.merge(above_90_sum, max_overdue, on=party_col, how='outer')
    summary = pd.merge(summary, overdue_sum, on=party_col, how='outer')

    for c in ['Above 90 days', 'Max Overdue days', 'Overdue']:
        summary[c] = pd.to_numeric(summary[c], errors='coerce').fillna(0)

    summary['Total'] = summary['Above 90 days'] + summary['Overdue']

    summary = summary[(summary['Above 90 days'] > 0) | (summary['Overdue'] > 0)]

    summary = summary.sort_values(
        by=['Above 90 days', 'Overdue'],
        ascending=[False, False]
    )

    final_df = summary[[party_col, 'Above 90 days', 'Max Overdue days', 'Overdue', 'Total']].copy()
    final_df.rename(columns={party_col: "Party's Name"}, inplace=True)
    final_df.insert(0, 'S.No.', range(1, len(final_df) + 1))

    # Formatting
    def disp_amt(x):
        return format_inr_number(x) if (pd.notnull(x) and float(x) != 0) else ''

    final_df['Above 90 days'] = final_df['Above 90 days'].apply(disp_amt)
    final_df['Overdue'] = final_df['Overdue'].apply(disp_amt)
    final_df['Total'] = final_df['Total'].apply(disp_amt)
    final_df['Max Overdue days'] = final_df['Max Overdue days'].apply(lambda x: int(x) if x != 0 else '')

    total_above_90 = format_inr_number(summary['Above 90 days'].sum())
    total_overdue  = format_inr_number(summary['Overdue'].sum())
    total_total    = format_inr_number(summary['Total'].sum())

    # PDF (MEMORY)
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=15*mm, leftMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm
    )

    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("MATRIX ELECTRICALS, COIMBATORE 641 012", styles['Title']))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(
        f"{report_name}OUTSTANDING ABOVE 90 DAYS & OVERDUE UPTO {datetime.now().strftime('%d.%m.%Y')}",
        styles['Heading2']
    ))
    
    elements.append(Spacer(1, 14))

    wrap_text = ParagraphStyle('wrap_text', fontSize=9, alignment=TA_LEFT)
    wrap_num  = ParagraphStyle('wrap_num',  fontSize=9, alignment=TA_RIGHT)
    wrap_hdr  = ParagraphStyle('wrap_hdr',  fontSize=9, alignment=TA_CENTER, fontName='Helvetica-Bold')

    data = [[Paragraph(c, wrap_hdr) for c in final_df.columns]]

    for _, row in final_df.iterrows():
        r = []
        for c in final_df.columns:
            val = row[c] if pd.notna(row[c]) else ''
            r.append(Paragraph(str(val), wrap_num if c in ['Above 90 days','Overdue','Total','Max Overdue days'] else wrap_text))
        data.append(r)

    data.append([
        Paragraph('', wrap_text),
        Paragraph('TOTAL', wrap_hdr),
        Paragraph(total_above_90, wrap_num),
        Paragraph('', wrap_num),
        Paragraph(total_overdue, wrap_num),
        Paragraph(total_total, wrap_num),
    ])

    table = Table(data, colWidths=[12*mm, 70*mm, 25*mm, 25*mm, 20*mm, 20*mm], repeatRows=1)

    ts = TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('BACKGROUND', (0,-1), (-1,-1), colors.whitesmoke),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
    ])

    above_idx = final_df.columns.get_loc('Above 90 days')
    overdue_idx = final_df.columns.get_loc('Overdue')

    for i, row in enumerate(data[1:-1], start=1):
        if row[above_idx].getPlainText().strip():
            ts.add('BACKGROUND', (above_idx, i), (above_idx, i), colors.red)
            ts.add('TEXTCOLOR', (above_idx, i), (above_idx, i), colors.white)
        if row[overdue_idx].getPlainText().strip():
            ts.add('BACKGROUND', (overdue_idx, i), (overdue_idx, i), colors.yellow)

    table.setStyle(ts)
    elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ================= STREAMLIT UI =================
import streamlit as st

st.set_page_config(page_title="Excel to PDF", layout="centered")

st.title("📄 Excel → Outstanding PDF Generator")

st.write("Upload Excel file to generate PDF")

uploaded_file = st.file_uploader(
    "Choose Excel file",
    type=["xls", "xlsx"]
)

if uploaded_file is not None:
    try:
        pdf_buffer = generate_pdf_from_excel(uploaded_file)

        st.success("PDF generated successfully")

        st.download_button(
            label="⬇ Download PDF",
            data=pdf_buffer,
            file_name=f"{uploaded_file.name.split('.')[0]}_output.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Error: {e}")

