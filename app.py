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
    st.markdown("### <h1 style='font-size: 24px;'>1. Upload Master Sheet</h1>", unsafe_allow_html=True)
    master_file = st.file_uploader("Select the master Excel file", type=['xlsx'], label_visibility="collapsed")
    
    if master_file:
        st.subheader("Master Sheet Preview")
        preview_master = pd.read_excel(master_file).head(5)
        st.dataframe(preview_master, use_container_width=True)

with col2:
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
            
            # --- MODIFICATION: EXTRACT ID FROM EMAIL ---
            # We look for the 'Email' column. If the user named it 'Email Address', it checks for that too.
            email_col = next((c for c in df_responses.columns if 'Email' in c), None)
            
            if email_col:
                # Extract the part before the '@', clean whitespace/decimals
                response_numbers = {
                    str(email).split('@')[0].replace('.0', '').strip() 
                    for email in df_responses[email_col] 
                    if pd.notnull(email) and '@' in str(email)
                }
            else:
                # Fallback to StudentNumber if Email column isn't found
                response_numbers = {str(x).replace('.0', '').strip() for x in df_responses.get("StudentNumber", []) if pd.notnull(x)}
            # --------------------------------------------

            wb = openpyxl.load_workbook(master_file)
            ws = wb.active
            col_map = {cell.value: cell.column for cell in ws[1] if cell.value is not None}
            
            absentees_list = []
            present_count = 0

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
                    absentees_list.append({
                        "Index": row - 1, 
                        "StudentNumber": sid, 
                        "StudentName": name_val
                    })

            # Create Buffers
            master_buffer = BytesIO()
            wb.save(master_buffer)
            master_buffer.seek(0)
            
            absentee_df = pd.DataFrame(absentees_list)
            absentee_buffer = BytesIO()
            with pd.ExcelWriter(absentee_buffer, engine='openpyxl') as writer:
                absentee_df.to_excel(writer, index=False, sheet_name='Absentees')
            absentee_buffer.seek(0)

            # Results UI
            st.balloons()
            st.success(f"Processing Complete! {present_count} Present, {len(absentees_list)} Absent.")
            
            if not absentee_df.empty:
                st.subheader("📋 Absentee List Summary")
                st.dataframe(absentee_df, use_container_width=True, hide_index=True)
            else:
                st.info("Perfect attendance! No absentees found.")

            st.divider()
            timestamp = datetime.now().strftime("%Y-%m-%d")
            
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label="📥 Download Attendance status record",
                    data=master_buffer,
                    file_name=f"Attendance_Status_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with dl_col2:
                st.download_button(
                    label="⚠️ Download Absentee List (Excel)",
                    data=absentee_buffer,
                    file_name=f"Absentees_Only_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"An error occurred: {e}")

