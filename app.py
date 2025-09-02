import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import io
from docx2pdf import convert as docx2pdf_convert
import subprocess # Added subprocess import

from utils import (
    translate_docx,
    translate_pptx,
    translate_pdf,
    translate_pdf_ocr,
    log_activity,
    translate_text_chunk,
    detect_and_report_language,
)

# Set page config
st.set_page_config(page_title="PV-COE Translate-Pro", page_icon="🌐")

# Function to load local CSS
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# Load custom CSS
local_css("style.css")

# Custom title with symbol
st.markdown(
    '<h1 style="text-align: left;">🌐PV-COE Translate-Pro</h1>',
    unsafe_allow_html=True
)

# Initialize session state variables
if 'translated_text' not in st.session_state:
    st.session_state.translated_text = ""
if 'source_language_name' not in st.session_state:
    st.session_state.source_language_name = ""
if 'download_path' not in st.session_state:
    st.session_state.download_path = None

# Language selection dictionary
LANGUAGE_NAMES = {
    'en': 'English',
    'es': 'Spanish',
    'fr': 'French',
    'de': 'German',
    'zh-cn': 'Chinese (Simplified)',
    'ja': 'Japanese',
    'ko': 'Korean',
    'ar': 'Arabic',
    'ru': 'Russian',
    'pt': 'Portuguese',
    # ... (other languages as before)
}
NAMES_TO_CODES = {v: k for k, v in LANGUAGE_NAMES.items()}

# Create directories
os.makedirs("uploads", exist_ok=True)
os.makedirs("translated_files", exist_ok=True)

# --- Tabs ---
tab1, tab2, tab3 = st.tabs(["Text Translation", "File Translation", "Activity Logs"])

# --- Tab 1: Text Translation ---
with tab1:
    st.header("Translate Text Instantly")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Enter text to translate")
        text_to_translate = st.text_area("", height=250, key="text_to_translate_input")

    with col2:
        st.subheader("Translated Text")
        st.text_area("", st.session_state.translated_text, height=250, key="translated_text_area")

    if st.session_state.source_language_name:
        st.info(f"Detected source language: {st.session_state.source_language_name}")

    target_language_text = st.selectbox(
        "Select target language",
        list(LANGUAGE_NAMES.values()),
        key="text_translation_language"
    )

    if st.button("Translate Text"):
        if text_to_translate and target_language_text:
            detected_code = detect_and_report_language(text_to_translate)
            if detected_code == "multi":
                st.session_state.source_language_name = "Multi-Language detected"
            else:
                st.session_state.source_language_name = LANGUAGE_NAMES.get(detected_code, "Unknown")

            with st.spinner("Translating text..."):
                target_language_code = NAMES_TO_CODES[target_language_text]
                st.session_state.translated_text = translate_text_chunk(text_to_translate, target_language_code)
            st.rerun()
        else:
            st.warning("Please enter text and select a target language.")

# --- Tab 2: File Translation ---
def convert_docx_to_pdf(docx_path, pdf_path):
    docx2pdf_convert(docx_path, pdf_path)

def convert_pptx_to_pdf(pptx_path, pdf_path):
    try:
        # unoconv requires LibreOffice to be installed and accessible in the system's PATH.
        # It converts the PPTX file to PDF.
        subprocess.run(
            ["unoconv", "-f", "pdf", "-o", pdf_path, pptx_path],
            check=True,
            capture_output=True
        )
    except subprocess.CalledProcessError as e:
        st.error(f"Error converting PPTX to PDF: {e.stderr.decode()}")
        raise
    except FileNotFoundError:
        st.error("unoconv not found. Please install LibreOffice and unoconv.")
        raise

with tab2:
    st.header("Translate Your Documents")
    st.write("Upload your file, select a language, and get a translated PDF with a report as the first page.")

    uploaded_file = st.file_uploader(
        "Upload a document (.docx, .pptx, .pdf)", type=["docx", "pptx", "pdf"]
    )
    if uploaded_file and uploaded_file.name.lower().endswith(".pdf"):
        st.warning(
            "PDF translation will only work for text-based PDFs. "
            "Scanned/image PDFs will use OCR, and formatting/layout will be lost. "
            "For best results, upload original DOCX or PPTX files."
        )

    target_language_file = st.selectbox(
        "Select target language",
        list(LANGUAGE_NAMES.values()),
        key="file_translation_language"
    )

    # Prepare metadata for report
    app_name = "PV COE TRANSLATOR PRO"
    date = pd.Timestamp.now().strftime("%d%m%Y")
    time = pd.Timestamp.now().strftime("%H:%M:%S")
    source_lang = "English"
    destination_lang = target_language_file
    file_name = uploaded_file.name if uploaded_file else "No file"

    metadata_table = Table([
        ["App Name", app_name],
        ["Date", date],
        ["Time", time],
        ["Source Language", source_lang],
        ["Destination Language", destination_lang],
        ["File Name", file_name]
    ], style=TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 12)
    ]))

    def create_pdf_report(table):
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=letter)
        doc.build([table])
        buf.seek(0)
        return buf

    if st.button("Translate File"):
        if uploaded_file is not None and target_language_file:
            file_path = os.path.join("uploads", uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            file_extension = os.path.splitext(file_path)[1].lower()
            target_code = NAMES_TO_CODES[target_language_file]
            translated_path = None
            pdf_translated_path = None

            with st.spinner(f"Translating {uploaded_file.name}... This may take a moment."):
                try:
                    if file_extension == ".docx":
                        translated_path = translate_docx(file_path, target_code)
                        pdf_translated_path = os.path.join(
                            "translated_files",
                            os.path.splitext(os.path.basename(translated_path))[0] + ".pdf"
                        )
                        convert_docx_to_pdf(translated_path, pdf_translated_path)
                    elif file_extension == ".pptx":
                        translated_path = translate_pptx(file_path, target_code)
                        pdf_translated_path = os.path.join(
                            "translated_files",
                            os.path.splitext(os.path.basename(translated_path))[0] + ".pdf"
                        )
                        convert_pptx_to_pdf(translated_path, pdf_translated_path)
                    elif file_extension == ".pdf":
                        pdf_translated_path, text_found = translate_pdf(file_path, target_code)
                        if not text_found:
                            pdf_translated_path = translate_pdf_ocr(file_path, target_code)
                    else:
                        st.error("Unsupported file type.")
                        pdf_translated_path = None

                    if not pdf_translated_path or not os.path.exists(pdf_translated_path):
                        st.error("PDF conversion failed.")
                        st.session_state.download_path = None
                    else:
                        st.session_state.download_path = pdf_translated_path
                        log_activity(
                            "file_translation",
                            uploaded_file.name,
                            "auto",
                            target_language_file,
                        )
                except Exception as e:
                    st.error(f"An error occurred during translation: {e}")
                    st.session_state.download_path = None
            st.rerun()
        else:
            st.warning("Please upload a file and select a target language.")

    # Download logic after translation
    if st.session_state.download_path:
        st.success("Your file has been successfully translated!")

        # Create PDF report
        report_buffer = create_pdf_report(metadata_table)

        # Merge report and translated PDF
        try:
            report_reader = PdfReader(report_buffer)
            translated_reader = PdfReader(st.session_state.download_path)
            writer = PdfWriter()
            for page in report_reader.pages:
                writer.add_page(page)
            for page in translated_reader.pages:
                writer.add_page(page)
            merged_buffer = io.BytesIO()
            writer.write(merged_buffer)
            merged_buffer.seek(0)

            st.download_button(
                label="Download Translated PDF (with Report)",
                data=merged_buffer,
                file_name=f"translated_with_report_{os.path.basename(st.session_state.download_path)}",
            )
        except Exception as e:
            st.error(f"Error merging PDFs: {e}")

    # PDF Metadata Table and Generation (standalone report)
    st.subheader("Generate PDF Report")
    if st.button("Generate PDF"):
        report_buffer = create_pdf_report(metadata_table)
        st.download_button(
            label="Download PDF Report",
            data=report_buffer,
            file_name="report.pdf",
        )

# --- Tab 3: Activity Logs ---
with tab3:
    st.header("Your Translation History")
    if os.path.exists("user_log.csv"):
        log_df = pd.read_csv("user_log.csv")
        if "Username" in log_df.columns:
            log_df = log_df.drop(columns=["Username"])
        st.dataframe(log_df)
    else:
        st.info("No activity logs found.")