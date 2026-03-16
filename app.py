import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Attendance Processor", page_icon="📊", layout="wide")

st.title("📊 Weekly Attendance Processor")

# Layout with two columns for uploads
col1, col2 = st.columns(2)

with col1:
    # 1. Custom Font Size for Header
    st.markdown("### <h1 style='font-size: 24px;'>1. Upload Master Sheet</h1>", unsafe_allow_html=True)
    master_file = st.file_uploader("Select the master Excel file", type=['xlsx'], label_visibility="collapsed")
    
    if master_file:
        st.subheader("Master Sheet Preview")
        preview_master = pd.read_excel(master_file).head(5)
        st.dataframe(preview_master, use_container_width=True)

with col2:
    # 2. Custom Font Size for Header
    st.markdown("### <h1 style='font-size: 24px;'>2. Upload Student Responses</h1>", unsafe_allow_html=True)
    response_file = st.file_uploader("Select the student responses file", type=['xlsx'], label_visibility="collapsed")
    
    if response_file:
        st.subheader("Responses Preview")
        preview_resp = pd.read_excel(response_file).head(5)
        st.dataframe(preview_resp, use_container_width=True)

if master_file and response_file:
    st.divider()
    if st.button("🚀 Process Attendance", type="primary"):
        try:
            df_responses = pd.read_excel(response_file)
            response_numbers = {str(x).replace('.0', '').strip() for x in df_responses["StudentNumber"] if str(x).replace('.0', '').strip().isdigit()}

            wb = openpyxl.load_workbook(master_file)
            ws = wb.active
            col_map = {cell.value: cell.column for cell in ws[1] if cell.value is not None}
            
            absentees_list = []
            present_count = 0

            for row in range(2, ws.max_row + 1):
                sid_val = ws.cell(row=row, column=col_map["StudentNumber"]).value
                # Grabs name if exists, otherwise uses ID
                name_val = ws.cell(row=row, column=col_map.get("StudentName", col_map["StudentNumber"])).value
                
                if sid_val is None: continue
                sid = str(sid_val).replace('.0', '').strip()

                if sid in response_numbers:
                    ws.cell(row=row, column=col_map["Status"]).value = "Present"
                    present_count += 1
                else:
                    ws.cell(row=row, column=col_map["Status"]).value = "Absent"
                    absentees_list.append({"StudentNumber": sid, "StudentName": name_val})

            # Create Buffers
            master_buffer = BytesIO()
            wb.save(master_buffer)
            master_buffer.seek(0)
            
            absentee_df = pd.DataFrame(absentees_list)
            absentee_buffer = BytesIO()
            with pd.ExcelWriter(absentee_buffer, engine='openpyxl') as writer:
                absentee_df.to_excel(writer, index=False, sheet_name='Absentees')
            absentee_buffer.seek(0)

            st.balloons()
            st.success(f"Processing Complete! {present_count} Present, {len(absentees_list)} Absent.")
            
            timestamp = datetime.now().strftime("%Y-%m-%d")
            
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                # Updated Button Label
                st.download_button(
                    label="📥 Download Attendance status record",
                    data=master_buffer,
                    file_name=f"Attendance_Status_{timestamp}.xlsx",
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
        
