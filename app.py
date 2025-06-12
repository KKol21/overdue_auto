import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

st.title("ERP Overdue Report Generator")

uploaded_file = st.file_uploader("Upload raw ERP Excel file", type=["xlsx"])
template_file = st.file_uploader("Upload template Excel file", type=["xlsx"])

if uploaded_file and template_file:
    # Load raw data
    df = pd.read_excel(uploaded_file, sheet_name='Open customer invoices')
    df['Due date'] = pd.to_datetime(df['Due date'], errors='coerce')

    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    df['Days late'] = (today - df['Due date']).dt.days
    df_filtered = df[(df['Days late'] >= 2) & ~df['Invoice'].astype(str).str.startswith('6')]

    # Load template
    wb = load_workbook(template_file)
    ws = wb['Open customer invoices']

    # Clear old A–J values (keep formulas in I+)
    data_cols = 7
    for row in ws.iter_rows(min_row=2, max_row=1000, min_col=1, max_col=data_cols):
        for cell in row:
            cell.value = None

    # Write filtered data
    for i, row in df_filtered.iterrows():
        for j, val in enumerate(row.values[:data_cols]):
            ws.cell(row=i + 2, column=j + 1, value=val)

    # Output to download
    output = BytesIO()
    wb.save(output)
    st.success("Report generated successfully!")
    st.download_button("Download report", data=output.getvalue(), file_name="daily_report.xlsx")
