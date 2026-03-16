import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Attendance Processor", page_icon="📊", layout="wide")

st.title("📊 Weekly Attendance Processor")
st.info("Upload your files. The app will generate the updated Master Sheet and an Absentee List.")

col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader("1. Upload Master Sheet", type=['xlsx'])
    if master_file:
        st.subheader("Master Sheet Preview")
        preview_master = pd.read_excel(master_file).head(5)
        st.dataframe(preview_master, use_container_width=True)

with col2:
    response_file = st.file_uploader("2. Upload Student Responses", type=['xlsx'])
    if response_file:
        st.subheader("Responses Preview")
        preview_resp = pd.read_excel(response_file).head(5)
        st.dataframe(preview_resp, use_container_width=True)

if master_file and response_file:
    st.divider()
    if st.button("🚀 Process Attendance", type="primary"):
        try:
            # 1. Load Data
            df_responses = pd.read_excel(response_file)
            response_numbers = {str(x).replace('.0', '').strip() for x in df_responses["StudentNumber"] if str(x).replace('.0', '').strip().isdigit()}

            wb = openpyxl.load_workbook(master_file)
            ws = wb.active
            col_map = {cell.value: cell.column for cell in ws[1] if cell.value is not None}
            
            # List to store absentee details for the second file
            absentees_list = []
            present_count = 0

            # 2. Process Rows
            for row in range(2, ws.max_row + 1):
                sid_val = ws.cell(row=row, column=col_map["StudentNumber"]).value
                name_val = ws.cell(row=row, column=col_map.get("StudentName", col_map["StudentNumber"])).value
                
                if sid_val is None: continue
                
                sid = str(sid_val).replace('.0', '').strip()

                if sid in response_numbers:
                    ws.cell(row=row, column=col_map["Status"]).value = "Present"
                    present_count += 1
                else:
                    ws.cell(row=row, column=col_map["Status"]).value = "Absent"
                    absentees_list.append({"StudentNumber": sid, "StudentName": name_val})

            # 3. Create Buffers for Downloads
            # Buffer 1: Updated Master
            master_buffer = BytesIO()
            wb.save(master_buffer)
            master_buffer.seek(0)
            
            # Buffer 2: Absentee List
            absentee_df = pd.DataFrame(absentees_list)
            absentee_buffer = BytesIO()
            with pd.ExcelWriter(absentee_buffer, engine='openpyxl') as writer:
                absentee_df.to_excel(writer, index=False, sheet_name='Absentees')
            absentee_buffer.seek(0)

            # 4. Results UI
            st.balloons()
            st.success(f"Processing Complete! {present_count} marked 'Present', {len(absentees_list)} marked 'Absent'.")
            
            timestamp = datetime.now().strftime("%Y-%m-%d")
            
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label="📥 Download Updated Master",
                    data=master_buffer,
                    file_name=f"Master_Updated_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with dl_col2:
                st.download_button(
                    label="⚠️ Download Absentee List",
                    data=absentee_buffer,
                    file_name=f"Absentees_Only_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"An error occurred: {e}")
