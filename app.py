import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook

st.title("ERP Overdue Report Generator")

uploaded_file = st.file_uploader("Upload ERP export Excel file", type=["xlsx"])

if uploaded_file:
    # Step 1: Load raw ERP data
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df['Due date'] = pd.to_datetime(df['Due date'], errors='coerce')

    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    df['Days late'] = (today - df['Due date']).dt.days
    df_filtered = df[(df['Days late'] >= 2) & ~df['Invoice'].astype(str).str.startswith('6')]

    # Step 2: Load internal Excel template
    wb = load_workbook("template.xlsx")
    ws = wb["Open customer invoices"]

    data_cols = 7

    # Step 4: Write new data
    for i, row in df_filtered.iterrows():
        for j in range(data_cols):
            ws.cell(row=i + 2, column=j + 1, value=row.iloc[j])

    # Step 5: Output updated Excel file
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.success("Report generated successfully.")
    st.download_button("Download updated Excel file", data=output, file_name="daily_report.xlsx")
