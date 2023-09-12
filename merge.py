import streamlit as st
import pandas as pd
import xlsxwriter
from tempfile import NamedTemporaryFile

def read_excel_sheet(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name=None, engine='xlrd')

def merge_sheets(sheet1, sheet2, sheet3):
    sheet1['ID'] = sheet1['ID'].astype(str)
    sheet2['ID'] = sheet2['ID'].astype(str)
    sheet3['NAME'] = sheet3['NAME'].astype(str)
    merge1 = pd.merge(sheet1, sheet2, on='ID', how='outer')
    merge2 = pd.merge(merge1, sheet3, on='NAME', how='outer')
    return merge1, merge2

st.title("Excel Sheet Merger")
uploaded_file = st.file_uploader("Choose an Excel file", type=['xls', 'xlsx'])

if uploaded_file is not None:
    sheets = read_excel_sheet(uploaded_file)
    sheet1, sheet2, sheet3 = sheets['sheet1'], sheets['sheet2'], sheets['sheet3']
    st.write("### Sheet 1")
    st.write(sheet1.head())
    st.write("### Sheet 2")
    st.write(sheet2.head())
    st.write("### Sheet 3")
    st.write(sheet3.head())

    if st.button("Merge Sheets"):
        merge1, merge2 = merge_sheets(sheet1, sheet2, sheet3)
        st.write("### Merge1 Result")
        st.write(merge1.head())
        st.write("### Merge2 Result")
        st.write(merge2.head())
        
        merge2_filtered = merge2[merge2['ENROLLED'].str.upper() == 'ENROLLED']
        st.write("### Merged and Filtered Result (Merge2)")
        st.write(merge2_filtered.head())

        merge2_filtered.fillna("", inplace=True)

        tmpfile = NamedTemporaryFile(delete=False, suffix='.xlsx')
        with xlsxwriter.Workbook(tmpfile.name) as book:
            ws = book.add_worksheet('Merge2_filtered')
            for i, (index, row) in enumerate(merge2_filtered.iterrows()):
                for j, col_val in enumerate(row):
                    ws.write(i, j, col_val)
        
        with open(tmpfile.name, "rb") as f:
            bytes_data = f.read()

        st.download_button(
            "Download Excel File",
            bytes_data,
            "merged_filtered_result.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
