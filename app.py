import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Attendance Processor", page_icon="📊", layout="wide")

st.title("📊 Weekly Attendance Processor")
st.info("Upload your files below. The app will validate the data and prepare a download link.")

# Layout with two columns for uploads
col1, col2 = st.columns(2)

with col1:
    master_file = st.file_uploader("1. Upload Master Sheet", type=['xlsx'])
    if master_file:
        st.subheader("Master Sheet Preview")
        # Load just the top few rows to keep it fast
        preview_master = pd.read_excel(master_file).head(5)
        st.dataframe(preview_master, use_container_width=True)

with col2:
    response_file = st.file_uploader("2. Upload Student Responses", type=['xlsx'])
    if response_file:
        st.subheader("Responses Preview")
        preview_resp = pd.read_excel(response_file).head(5)
        st.dataframe(preview_resp, use_container_width=True)

# Processing Section
if master_file and response_file:
    st.divider()
    if st.button("🚀 Process Attendance", type="primary"):
        try:
            # Re-read files for full processing
            df_responses = pd.read_excel(response_file)
            response_numbers = {str(x).replace('.0', '').strip() for x in df_responses["StudentNumber"] if str(x).replace('.0', '').strip().isdigit()}

            wb = openpyxl.load_workbook(master_file)
            ws = wb.active
            col_map = {cell.value: cell.column for cell in ws[1] if cell.value is not None}
            
            # Check if required columns exist
            required = ["StudentNumber", "Status"]
            if not all(col in col_map for col in required):
                st.error(f"Error: Master sheet must have columns named {required}")
            else:
                present_count = 0
                for row in range(2, ws.max_row + 1):
                    sid = str(ws.cell(row=row, column=col_map["StudentNumber"]).value).replace('.0', '').strip()
                    if sid in response_numbers:
                        ws.cell(row=row, column=col_map["Status"]).value = "Present"
                        present_count += 1
                    else:
                        ws.cell(row=row, column=col_map["Status"]).value = "Absent"

                # Prepare Download
                buffer = BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                st.balloons() # Fun visual effect
                st.success(f"Successfully processed! {present_count} students marked 'Present'.")
                
                timestamp = datetime.now().strftime("%Y-%m-%d")
                st.download_button(
                    label="📥 Download Updated Master Sheet",
                    data=buffer,
                    file_name=f"Attendance_Updated_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"An error occurred: {e}")
