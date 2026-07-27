
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Page setup
st.set_page_config(page_title="Break Compliance Report", page_icon="📊", layout="centered")

# Title
st.markdown("# 📊 Break Compliance Report")
st.markdown("Upload attendance CSV → Get formatted report instantly!")
st.markdown("---")

# Date selector (default = yesterday)
yesterday = date.today() - timedelta(days=1)
report_date = st.date_input("Report Date", value=yesterday)

# File upload
uploaded_file = st.file_uploader("Upload your attendance CSV file", type=['csv'])

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    st.success(f"Loaded {len(df)} employees!")

    def safe_int(val):
        try:
            if pd.notna(val):
                return int(val)
            return 0
        except:
            return 0

    # Process
    df['Department'] = np.where(
        df['1st Break Status'].str.contains('Combined', na=False), 'IB', 'OB'
    )
    df['Total Break (min)'] = np.where(
        df['Department'] == 'IB',
        df['1st Break (min)'].fillna(0),
        df['1st Break (min)'].fillna(0) + df['2nd Break (min)'].fillna(0)
    )
    df['Break Flag'] = np.where(
        df['Total Break (min)'] >= 66, 'Excess Break',
        np.where(df['Total Break (min)'] <= 54, 'Less Break', 'OK')
    )

    excess = df[df['Break Flag']=='Excess Break'][['Employee ID','Employee Name','Total Break (min)','Department']].sort_values('Total Break (min)', ascending=False)
    less = df[df['Break Flag']=='Less Break'][['Employee ID','Employee Name','Total Break (min)','Department']].sort_values('Total Break (min)')

    # Show results on screen
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Exceptions", len(excess) + len(less))
    col2.metric("Excess (≥66 min)", len(excess))
    col3.metric("Less (≤54 min)", len(less))

    st.markdown("---")

    # Excess table
    st.markdown("### 🔴 Excess Break (≥66 min)")
    if len(excess) == 0:
        st.info("No exceptions found ✅")
    else:
        st.dataframe(excess.reset_index(drop=True), use_container_width=True)

    # Less table
    st.markdown("### 🟠 Less Break (≤54 min)")
    if len(less) == 0:
        st.info("No exceptions found ✅")
    else:
        st.dataframe(less.reset_index(drop=True), use_container_width=True)

    st.markdown("---")

    # Generate Excel download
    def generate_excel():
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        ws.sheet_view.showGridLines = False

        squid_ink = '232F3E'
        teal = '00BCD4'
        coral = 'FF6B6B'
        sunset = 'FFA726'
        snow = 'FAFAFA'
        ice_blue = 'E0F7FA'
        light_coral = 'FFEBEE'
        light_sunset = 'FFF3E0'
        white = 'FFFFFF'
        dark_text = '212121'

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 12

        # Header
        for r in range(1, 3):
            for c in range(1, 5):
                ws.cell(row=r, column=c).fill = PatternFill(start_color=squid_ink, end_color=squid_ink, fill_type='solid')
        ws.row_dimensions[1].height = 10
        ws.row_dimensions[2].height = 35
        ws.merge_cells('A2:D2')
        ws['A2'] = "BREAK COMPLIANCE REPORT"
        ws['A2'].font = Font(name='Calibri', size=16, bold=True, color=teal)
        ws['A2'].alignment = Alignment(vertical='center', horizontal='center')

        ws.row_dimensions[3].height = 4
        for c in range(1, 5):
            ws.cell(row=3, column=c).fill = PatternFill(start_color=teal, end_color=teal, fill_type='solid')

        ws.row_dimensions[4].height = 22
        ws.merge_cells('A4:D4')
        ws['A4'] = report_date.strftime("%A, %d %B %Y")
        ws['A4'].font = Font(name='Calibri', size=10, color='666666')
        ws['A4'].fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')
        ws['A4'].alignment = Alignment(vertical='center', horizontal='center')
        for c in range(1, 5):
            ws.cell(row=4, column=c).fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')

        ws.row_dimensions[5].height = 30
        ws['A5'] = len(excess) + len(less)
        ws['A5'].font = Font(name='Calibri', size=18, bold=True, color=squid_ink)
        ws['A5'].alignment = Alignment(horizontal='center', vertical='center')
        ws['A5'].fill = PatternFill(start_color=ice_blue, end_color=ice_blue, fill_type='solid')
        ws['B5'] = "exceptions"
        ws['B5'].font = Font(name='Calibri', size=9, color='888888')
        ws['B5'].alignment = Alignment(vertical='center')
        ws['B5'].fill = PatternFill(start_color=ice_blue, end_color=ice_blue, fill_type='solid')
        ws['C5'] = len(excess)
        ws['C5'].font = Font(name='Calibri', size=14, bold=True, color=coral)
        ws['C5'].alignment = Alignment(horizontal='center', vertical='center')
        ws['C5'].fill = PatternFill(start_color=light_coral, end_color=light_coral, fill_type='solid')
        ws['D5'] = len(less)
        ws['D5'].font = Font(name='Calibri', size=14, bold=True, color=sunset)
        ws['D5'].alignment = Alignment(horizontal='center', vertical='center')
        ws['D5'].fill = PatternFill(start_color=light_sunset, end_color=light_sunset, fill_type='solid')

        ws.row_dimensions[6].height = 14
        ws['C6'] = "excess"
        ws['C6'].font = Font(name='Calibri', size=8, color=coral)
        ws['C6'].alignment = Alignment(horizontal='center')
        ws['D6'] = "less"
        ws['D6'].font = Font(name='Calibri', size=8, color=sunset)
        ws['D6'].alignment = Alignment(horizontal='center')
        ws.row_dimensions[7].height = 8

        # Excess table
        row = 8
        ws.merge_cells(f'A{row}:D{row}')
        ws[f'A{row}'] = "EXCESS BREAK  >=66 min"
        ws[f'A{row}'].font = Font(name='Calibri', size=10, bold=True, color=white)
        ws[f'A{row}'].fill = PatternFill(start_color=coral, end_color=coral, fill_type='solid')
        ws[f'A{row}'].alignment = Alignment(vertical='center')
        for c in range(1, 5):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=coral, end_color=coral, fill_type='solid')
        ws.row_dimensions[row].height = 20

        row += 1
        for col, header in enumerate(['Employee ID', 'Name', 'Min', 'Dept'], 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(name='Calibri', size=8, bold=True, color=squid_ink)
            cell.fill = PatternFill(start_color=light_coral, end_color=light_coral, fill_type='solid')
        ws.row_dimensions[row].height = 16

        row += 1
        if len(excess) == 0:
            ws[f'A{row}'] = "No exceptions"
            ws[f'A{row}'].font = Font(name='Calibri', size=9, italic=True, color='AAAAAA')
            row += 1
        else:
            for i, (_, emp) in enumerate(excess.iterrows()):
                ws.row_dimensions[row].height = 17
                if i % 2 == 0:
                    for c in range(1, 5):
                        ws.cell(row=row, column=c).fill = PatternFill(start_color=light_coral, end_color=light_coral, fill_type='solid')
                ws.cell(row=row, column=1, value=safe_int(emp['Employee ID'])).font = Font(name='Calibri', size=9, color=dark_text)
                ws.cell(row=row, column=2, value=str(emp['Employee Name'])).font = Font(name='Calibri', size=9, color=dark_text)
                ws.cell(row=row, column=3, value=safe_int(emp['Total Break (min)'])).font = Font(name='Calibri', size=9, bold=True, color=coral)
                ws.cell(row=row, column=4, value=str(emp['Department'])).font = Font(name='Calibri', size=9, color=dark_text)
                row += 1

        ws.row_dimensions[row].height = 8
        row += 1

        # Less table
        ws.merge_cells(f'A{row}:D{row}')
        ws[f'A{row}'] = "LESS BREAK  <=54 min"
        ws[f'A{row}'].font = Font(name='Calibri', size=10, bold=True, color=white)
        ws[f'A{row}'].fill = PatternFill(start_color=sunset, end_color=sunset, fill_type='solid')
        ws[f'A{row}'].alignment = Alignment(vertical='center')
        for c in range(1, 5):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=sunset, end_color=sunset, fill_type='solid')
        ws.row_dimensions[row].height = 20

        row += 1
        for col, header in enumerate(['Employee ID', 'Name', 'Min', 'Dept'], 1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(name='Calibri', size=8, bold=True, color=squid_ink)
            cell.fill = PatternFill(start_color=light_sunset, end_color=light_sunset, fill_type='solid')
        ws.row_dimensions[row].height = 16

        row += 1
        if len(less) == 0:
            ws[f'A{row}'] = "No exceptions"
            ws[f'A{row}'].font = Font(name='Calibri', size=9, italic=True, color='AAAAAA')
            row += 1
        else:
            for i, (_, emp) in enumerate(less.iterrows()):
                ws.row_dimensions[row].height = 17
                if i % 2 == 0:
                    for c in range(1, 5):
                        ws.cell(row=row, column=c).fill = PatternFill(start_color=light_sunset, end_color=light_sunset, fill_type='solid')
                ws.cell(row=row, column=1, value=safe_int(emp['Employee ID'])).font = Font(name='Calibri', size=9, color=dark_text)
                ws.cell(row=row, column=2, value=str(emp['Employee Name'])).font = Font(name='Calibri', size=9, color=dark_text)
                ws.cell(row=row, column=3, value=safe_int(emp['Total Break (min)'])).font = Font(name='Calibri', size=9, bold=True, color=sunset)
                ws.cell(row=row, column=4, value=str(emp['Department'])).font = Font(name='Calibri', size=9, color=dark_text)
                row += 1

        for c in range(1, 5):
            ws.cell(row=row, column=c).border = Border(top=Side(style='medium', color=teal))

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # Download button
    excel_file = generate_excel()
    st.download_button(
        label="📥 Download Excel Report",
        data=excel_file,
        file_name=f"Break_Compliance_Report_{report_date.strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.markdown("### 📋 How to use:")
    st.markdown("1. Export attendance CSV from the dashboard")
    st.markdown("2. Upload it above")
    st.markdown("3. View results and download the formatted report!")
    st.markdown("")
    st.markdown("---")
    st.markdown("**Criteria:**")
    st.markdown("- **IB** = 1 break (Combined) | **OB** = 2 breaks")
    st.markdown("- **Excess** = ≥66 min | **Less** = ≤54 min")
    st.markdown("- **Shift:** DXB3 | 08:00 - 18:00")

