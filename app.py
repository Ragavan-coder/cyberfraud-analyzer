import streamlit as st
import os
import uuid
from processor import process_pdf, save_consolidated_excel

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Cyber Fraud Analyzer",
    layout="centered"
)

st.title("Cyber Fraud PDF Analyzer")
st.write("Upload cybercrime complaint PDFs and download a structured Excel report.")

# =====================================================
# BACKEND STORAGE
# =====================================================
BACKEND_FOLDER = "uploaded_pdfs"
os.makedirs(BACKEND_FOLDER, exist_ok=True)

# =====================================================
# FILE UPLOAD
# =====================================================
uploaded_files = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} PDF(s) uploaded.")

    if st.button("Start Processing"):
        all_data = []

        with st.spinner("Processing PDFs..."):
            for uploaded_file in uploaded_files:
                try:
                    # -------------------------------------------------
                    # SAFE UNIQUE FILE SAVE
                    # -------------------------------------------------
                    unique_name = f"{uuid.uuid4().hex}_{uploaded_file.name}"
                    backend_path = os.path.join(BACKEND_FOLDER, unique_name)

                    with open(backend_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    # -------------------------------------------------
                    # PROCESS PDF
                    # -------------------------------------------------
                    pdf_data = process_pdf(backend_path)
                    all_data.append(pdf_data)

                except Exception as e:
                    st.error(f"Failed to process {uploaded_file.name}: {e}")

        if not all_data:
            st.error("No valid PDFs were processed.")
            st.stop()

        # =================================================
        # SAVE EXCEL
        # =================================================
        output_excel_path = os.path.join(
            BACKEND_FOLDER,
            "Consolidated_Report.xlsx"
        )

        save_consolidated_excel(all_data, output_excel_path)

        st.success("Processing complete.")

        # =================================================
        # DOWNLOAD
        # =================================================
        with open(output_excel_path, "rb") as f:
            st.download_button(
                label="Download Consolidated Excel",
                data=f,
                file_name="Consolidated_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
