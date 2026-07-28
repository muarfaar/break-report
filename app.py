
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Break Compliance Report", page_icon="📊", layout="centered")

st.markdown("# 📊 Break Compliance Report")
st.markdown("Upload attendance CSV → Get formatted report instantly!")
st.markdown("---")

# === FILE UPLOADS ===
uploaded_file = st.file_uploader("📄 Upload Attendance CSV", type=['csv'], help="Required — your daily attendance export")
hc_file = st.file_uploader("📁 Upload HC Excel File (optional)", type=['xlsx'], help="Your DXB3 HC file — includes Roster + History sheets")

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    st.success(f"✅ Loaded {len(df)} employees!")

    today_str = date.today().strftime('%Y-%m-%d')

    def safe_int(val):
        try:
            if pd.notna(val):
                return int(val)
            return 0
        except:
            return 0

    # Load HC file (Roster + History)
    roster = pd.DataFrame(columns=['Psoft ID', 'Department'])
    history = pd.DataFrame(columns=['Employee ID', 'Employee Name', 'Department', 'Date', 'Flag'])

    if hc_file is not None:
        try:
            xl = pd.ExcelFile(hc_file)

            # Load Roster for department lookup
            if 'Roster' in xl.sheet_names:
                hc_df = pd.read_excel(xl, sheet_name='Roster')
                if 'Psoft ID' in hc_df.columns and 'Department' in hc_df.columns:
                    roster = hc_df[['Psoft ID', 'Department']].copy()
                    roster['Psoft ID'] = roster['Psoft ID'].astype(str)
                    st.success(f"👥 Roster: {len(roster)} employees loaded")

            # Load History for repeat tracking
            if 'History' in xl.sheet_names:
                history = pd.read_excel(xl, sheet_name='History')
                history['Employee ID'] = history['Employee ID'].astype(str)
                st.success(f"📋 History: {len(history)} past records loaded")
            else:
                st.info("📋 No 'History' sheet found — will create one in download")

        except Exception as e:
            st.warning(f"⚠️ Could not read HC file: {e}")

    # Merge department from roster
    df['Employee ID'] = df['Employee ID'].astype(str)
    if not roster.empty:
        df = df.merge(
            roster.rename(columns={'Psoft ID': 'Employee ID', 'Department': 'Dept'}),
            on='Employee ID',
            how='left'
        )
        df['Dept'] = df['Dept'].fillna('Unknown')
    else:
        df['Dept'] = 'Unknown'

    # Process break logic
    df['Break Type'] = np.where(
        df['1st Break Status'].str.contains('Combined', na=False), 'IB', 'OB'
    )
    df['Total Break (min)'] = np.where(
        df['Break Type'] == 'IB',
        df['1st Break (min)'].fillna(0),
        df['1st Break (min)'].fillna(0) + df['2nd Break (min)'].fillna(0)
    )

    # Separate missed punch (total = 0)
    missed = df[df['Total Break (min)'] == 0][['Employee ID', 'Employee Name', 'Dept', 'Total Break (min)']].sort_values('Employee Name')

    # Flag Excess/Less
    df_with_breaks = df[df['Total Break (min)'] > 0].copy()
    df_with_breaks['Break Flag'] = np.where(
        df_with_breaks['Total Break (min)'] >= 65, 'Excess Break',
        np.where(df_with_breaks['Total Break (min)'] <= 55, 'Less Break', 'OK')
    )

    excess = df_with_breaks[df_with_breaks['Break Flag']=='Excess Break'][['Employee ID','Employee Name','Dept','Total Break (min)']].sort_values('Total Break (min)', ascending=False)
    less = df_with_breaks[df_with_breaks['Break Flag']=='Less Break'][['Employee ID','Employee Name','Dept','Total Break (min)']].sort_values('Total Break (min)')

    # Repeat count function
    def get_repeat_count(employee_id, flag_type):
        if history.empty:
            return 0
        past = history[
            (history['Employee ID'] == str(employee_id)) &
            (history['Flag'] == flag_type) &
            (history['Date'] != today_str)
        ]
        return len(past)

    def add_repeat_info(data, flag_type):
        data = data.copy()
        data['Repeat'] = data['Employee ID'].apply(lambda x: get_repeat_count(x, flag_type))
        return data

    excess_display = add_repeat_info(excess, 'Excess')
    less_display = add_repeat_info(less, 'Less')
    missed_display = add_repeat_info(missed, 'Missed')

    # Show metrics
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Exceptions", len(excess) + len(less))
    col2.metric("Excess (≥65 min)", len(excess))
    col3.metric("Less (≤55 min)", len(less))
    col4.metric("Missed Punch", len(missed))

    st.markdown("---")

    # Display tables
    def show_table(title, emoji, data):
        st.markdown(f"### {emoji} {title}")
        if len(data) == 0:
            st.info("No exceptions found ✅")
        else:
            display = data.copy().reset_index(drop=True)
            display.columns = ['Employee ID', 'Employee Name', 'Department', 'Break (min)', 'Repeat']
            display['Repeat'] = display['Repeat'].apply(lambda x: f"⚠️ {x + 1}x" if x > 0 else "")
            st.dataframe(display, use_container_width=True)

    show_table("Excess Break (≥65 min)", "🔴", excess_display)
    show_table("Less Break (≤55 min)", "🟠", less_display)
    show_table("Missed Break Punch (0 min)", "⚫", missed_display)

    st.markdown("---")

    # Build updated history
    new_records = []
    for _, emp in excess.iterrows():
        new_records.append({'Employee ID': str(safe_int(emp['Employee ID'])), 'Employee Name': str(emp['Employee Name']), 'Department': str(emp['Dept']), 'Date': today_str, 'Flag': 'Excess'})
    for _, emp in less.iterrows():
        new_records.append({'Employee ID': str(safe_int(emp['Employee ID'])), 'Employee Name': str(emp['Employee Name']), 'Department': str(emp['Dept']), 'Date': today_str, 'Flag': 'Less'})
    for _, emp in missed.iterrows():
        new_records.append({'Employee ID': str(safe_int(emp['Employee ID'])), 'Employee Name': str(emp['Employee Name']), 'Department': str(emp['Dept']), 'Date': today_str, 'Flag': 'Missed'})

    # Merge with existing history (remove today's duplicates)
    if not history.empty:
        updated_history = history[history['Date'] != today_str].copy()
        updated_history = pd.concat([updated_history, pd.DataFrame(new_records)], ignore_index=True)
    else:
        updated_history = pd.DataFrame(new_records)

    # Generate Excel Report
    def generate_report_excel():
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        ws.sheet_view.showGridLines = False

        squid_ink = '232F3E'
        teal = '00BCD4'
        coral = 'FF6B6B'
        sunset = 'FFA726'
        charcoal = '424242'
        snow = 'FAFAFA'
        ice_blue = 'E0F7FA'
        light_coral = 'FFEBEE'
        light_sunset = 'FFF3E0'
        light_charcoal = 'F5F5F5'
        white = 'FFFFFF'
        dark_text = '212121'
        repeat_red = 'D32F2F'
        repeat_bg = 'FFCDD2'

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 12

        # Header
        for r in range(1, 3):
            for c in range(1, 6):
                ws.cell(row=r, column=c).fill = PatternFill(start_color=squid_ink, end_color=squid_ink, fill_type='solid')
        ws.row_dimensions[1].height = 10
        ws.row_dimensions[2].height = 35
        ws.merge_cells('A2:E2')
        ws['A2'] = "BREAK COMPLIANCE REPORT"
        ws['A2'].font = Font(name='Calibri', size=16, bold=True, color=teal)
        ws['A2'].alignment = Alignment(vertical='center', horizontal='center')

        ws.row_dimensions[3].height = 4
        for c in range(1, 6):
            ws.cell(row=3, column=c).fill = PatternFill(start_color=teal, end_color=teal, fill_type='solid')

        ws.row_dimensions[4].height = 22
        ws.merge_cells('A4:E4')
        ws['A4'] = date.today().strftime("%A, %d %B %Y")
        ws['A4'].font = Font(name='Calibri', size=10, color='666666')
        ws['A4'].fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')
        ws['A4'].alignment = Alignment(vertical='center', horizontal='center')
        for c in range(1, 6):
            ws.cell(row=4, column=c).fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')

        ws.row_dimensions[5].height = 30
        metrics = [
            (len(excess) + len(less), squid_ink, ice_blue),
            (len(excess), coral, light_coral),
            (len(less), sunset, light_sunset),
            (len(missed), charcoal, light_charcoal),
        ]
        for i, (val, font_color, bg_color) in enumerate(metrics, 1):
            cell = ws.cell(row=5, column=i, value=val)
            cell.font = Font(name='Calibri', size=14 if i > 1 else 18, bold=True, color=font_color)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        ws.cell(row=5, column=5).fill = PatternFill(start_color=snow, end_color=snow, fill_type='solid')

        ws.row_dimensions[6].height = 14
        labels = [("total", '888888'), ("excess", coral), ("less", sunset), ("missed", charcoal)]
        for i, (label, color) in enumerate(labels, 1):
            cell = ws.cell(row=6, column=i, value=label)
            cell.font = Font(name='Calibri', size=8, color=color)
            cell.alignment = Alignment(horizontal='center')
        ws.row_dimensions[7].height = 8

        def write_table(ws, row, title, header_color, light_color, data_df, flag_type):
            ws.merge_cells(f'A{row}:E{row}')
            ws[f'A{row}'] = title
            ws[f'A{row}'].font = Font(name='Calibri', size=10, bold=True, color=white)
            ws[f'A{row}'].fill = PatternFill(start_color=header_color, end_color=header_color, fill_type='solid')
            ws[f'A{row}'].alignment = Alignment(vertical='center')
            for c in range(1, 6):
                ws.cell(row=row, column=c).fill = PatternFill(start_color=header_color, end_color=header_color, fill_type='solid')
            ws.row_dimensions[row].height = 20
            row += 1

            for col, header in enumerate(['Employee ID', 'Name', 'Department', 'Min', 'Repeat'], 1):
                cell = ws.cell(row=row, column=col, value=header)
                cell.font = Font(name='Calibri', size=8, bold=True, color=squid_ink)
                cell.fill = PatternFill(start_color=light_color, end_color=light_color, fill_type='solid')
            ws.row_dimensions[row].height = 16
            row += 1

            if len(data_df) == 0:
                ws[f'A{row}'] = "No exceptions"
                ws[f'A{row}'].font = Font(name='Calibri', size=9, italic=True, color='AAAAAA')
                row += 1
            else:
                for i, (_, emp) in enumerate(data_df.iterrows()):
                    emp_id = safe_int(emp['Employee ID'])
                    repeat_count = get_repeat_count(emp['Employee ID'], flag_type)
                    is_repeat = repeat_count > 0

                    ws.row_dimensions[row].height = 17

                    if is_repeat:
                        for c in range(1, 6):
                            ws.cell(row=row, column=c).fill = PatternFill(start_color=repeat_bg, end_color=repeat_bg, fill_type='solid')
                        ws.cell(row=row, column=1, value=emp_id).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=2, value=str(emp['Employee Name'])).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=3, value=str(emp['Dept'])).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=4, value=safe_int(emp['Total Break (min)'])).font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                        ws.cell(row=row, column=5, value=f"⚠️ {repeat_count + 1}x").font = Font(name='Calibri', size=9, bold=True, color=repeat_red)
                    else:
                        if i % 2 == 0:
                            for c in range(1, 6):
                                ws.cell(row=row, column=c).fill = PatternFill(start_color=light_color, end_color=light_color, fill_type='solid')
                        ws.cell(row=row, column=1, value=emp_id).font = Font(name='Calibri', size=9, color=dark_text)
                        ws.cell(row=row, column=2, value=str(emp['Employee Name'])).font = Font(name='Calibri', size=9, color=dark_text)
                        ws.cell(row=row, column=3, value=str(emp['Dept'])).font = Font(name='Calibri', size=9, color=dark_text)
                        ws.cell(row=row, column=4, value=safe_int(emp['Total Break (min)'])).font = Font(name='Calibri', size=9, bold=True, color=header_color)
                        ws.cell(row=row, column=5, value="").font = Font(name='Calibri', size=9)
                    row += 1

            ws.row_dimensions[row].height = 8
            row += 1
            return row

        row = 8
        row = write_table(ws, row, "EXCESS BREAK  >=65 min", coral, light_coral, excess, 'Excess')
        row = write_table(ws, row, "LESS BREAK  <=55 min", sunset, light_sunset, less, 'Less')
        row = write_table(ws, row, "MISSED BREAK PUNCH  0 min", charcoal, light_charcoal, missed, 'Missed')

        for c in range(1, 6):
            ws.cell(row=row, column=c).border = Border(top=Side(style='medium', color=teal))
        row += 1
        ws.row_dimensions[row].height = 20
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "⚠️ Red highlighted rows = repeat offenders (flagged on previous days)"
        ws[f'A{row}'].font = Font(name='Calibri', size=8, italic=True, color='888888')

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # Generate updated HC file with new history
    def generate_updated_hc():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Re-write Roster sheet (keep original data)
            if hc_file is not None:
                hc_file.seek(0)
                xl = pd.ExcelFile(hc_file)
                for sheet_name in xl.sheet_names:
                    if sheet_name != 'History':
                        sheet_df = pd.read_excel(xl, sheet_name=sheet_name)
                        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Write updated History sheet
            updated_history.to_excel(writer, sheet_name='History', index=False)

        output.seek(0)
        return output

    # Download buttons
    st.markdown("### 📥 Downloads")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        excel_file = generate_report_excel()
        st.download_button(
            label="📥 Excel Report",
            data=excel_file,
            file_name=f"Break_Compliance_{date.today().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_dl2:
        if hc_file is not None:
            updated_hc = generate_updated_hc()
            st.download_button(
                label="📁 Updated HC File",
                data=updated_hc,
                file_name=f"DXB3_HC_{date.today().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Save this as your HC file — History sheet is updated!"
            )
        else:
            # Just give history as separate Excel
            hist_output = BytesIO()
            updated_history.to_excel(hist_output, index=False, sheet_name='History')
            hist_output.seek(0)
            st.download_button(
                label="📋 Download History",
                data=hist_output,
                file_name="history.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Add this as a 'History' sheet in your HC file"
            )

    st.markdown("---")
    st.caption("💡 **Tip:** Always download the updated HC file and use it next time — the History sheet grows daily!")

else:
    st.markdown("### 📋 How to use:")
    st.markdown("1. Upload your **attendance CSV** *(required)*")
    st.markdown("2. Upload your **HC Excel file** *(optional — for departments + repeat tracking)*")
    st.markdown("3. View results + download report!")
    st.markdown("")
    st.markdown("---")
    st.markdown("**Criteria:**")
    st.markdown("- **Excess** = ≥65 min | **Less** = ≤55 min")
    st.markdown("- **Repeat offenders** highlighted in red with count")
    st.markdown("- **Department** auto-matched from HC Roster sheet")


