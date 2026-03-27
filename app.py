import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Attendance Processor", page_icon="📊", layout="wide")

st.title("📊 Weekly Attendance Processor")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 1. Upload Master Sheet")
    master_file = st.file_uploader("Select the original Master Excel", type=['xlsx'])

with col2:
    st.markdown("### 2. Upload Student Responses")
    response_file = st.file_uploader("Select the Form Responses", type=['xlsx'])

if master_file and response_file:
    st.divider()
    if st.button("🚀 Process & Match by Email", type="primary"):
        try:
            # 1. Read Responses to get Emails
            df_responses = pd.read_excel(response_file)
            
            # Find the Email column in the responses
            resp_email_col = next((c for c in df_responses.columns if 'Email' in c), None)
            
            if not resp_email_col:
                st.error("Could not find an 'Email' column in the Response file.")
                st.stop()

            # Create a set of clean, lowercase emails for fast matching
            # We extract the ID part (before @) AND the full email to be safe
            present_emails = {
                str(email).strip().lower() 
                for email in df_responses[resp_email_col] 
                if pd.notnull(email)
            }

            # 2. Load Master Workbook (KEEPING ORIGINAL FORMAT)
            # We seek(0) to ensure we read from the start of the uploaded file
            master_file.seek(0)
            wb = openpyxl.load_workbook(master_file)
            ws = wb.active

            # 3. Map Columns in Master Sheet
            # We need 'Email' (to match) and 'Status' (to update)
            master_cols = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value is not None}
            
            # Identify the Email column in Master (might be 'Email', 'StudentEmail', etc.)
            m_email_key = next((k for k in master_cols.keys() if 'Email' in k), None)
            
            if not m_email_key or "Status" not in master_cols:
                st.error(f"Master Sheet must have an Email column and a 'Status' column. Found: {list(master_cols.keys())}")
                st.stop()

            present_count = 0
            total_students = 0

            # 4. Update the cells DIRECTLY without touching other formatting
            for row in range(2, ws.max_row + 1):
                email_cell = ws.cell(row=row, column=master_cols[m_email_key])
                if email_cell.value is None:
                    continue
                
                total_students += 1
                master_email = str(email_cell.value).strip().lower()

                # Match logic
                if master_email in present_emails:
                    ws.cell(row=row, column=master_cols["Status"]).value = "Present"
                    present_count += 1
                else:
                    ws.cell(row=row, column=master_cols["Status"]).value = "Absent"

            # 5. Save the Workbook to a Buffer
            # This method preserves all original styles, hidden columns, and formatting
            output_buffer = BytesIO()
            wb.save(output_buffer)
            processed_data = output_buffer.getvalue()

            st.balloons()
            st.success(f"Matched {present_count} students out of {total_students} via Email.")

            # 6. Download
            timestamp = datetime.now().strftime("%Y-%m-%d")
            st.download_button(
                label="📥 Download Corrected Master Sheet",
                data=processed_data,
                file_name=f"Attendance_Status_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Error: {str(e)}")
