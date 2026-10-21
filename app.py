# app.py

import os
import shutil
import subprocess
import sys
import time 
from pathlib import Path
import streamlit as st

def reset_upload_dir(path: str):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path, exist_ok=True)

# UI
st.set_page_config(page_title="SOW Generator", layout="wide")

st.title("📄 SOW Generator")

SOW_files = st.file_uploader(
    "Upload Sample SOWs", type=["pdf"], accept_multiple_files=True
)

input_files = st.file_uploader(
    "Upload Input Files", type=["pdf", "ppt", "pptx", "eml"], accept_multiple_files=True
)

template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

# ---------------- GENERATE ----------------
if st.button("Generate SOW"):
    if not SOW_files or not input_files:
        st.error("Please upload at least one file in both sections.")
        st.stop()

    start_time = time.time()

    reset_upload_dir("uploads/SOW")
    reset_upload_dir("uploads/Input")
    reset_upload_dir("uploads/template")

    # Save SOW files
    for f in SOW_files:
        with open(Path("uploads/SOW") / f.name, "wb") as out:
            out.write(f.read())

    # Save Input files
    for f in input_files:
        with open(Path("uploads/Input") / f.name, "wb") as out:
            out.write(f.read())

    # Save Template file
    if template_file:
        with open(Path("uploads/template") / template_file.name, "wb") as out:
            out.write(template_file.read())

    # Run backend
    with st.spinner("Generating SOW using agentic pipeline..."):
        subprocess.run(
            [
                sys.executable,
                "main.py",
                "--SOW-dir",
                "uploads/SOW",
                "--Input-dir",
                "uploads/Input",
                "--output",
                "sow.docx",
            ]
        )
    
    end_time = time.time()
    total_time = end_time - start_time
    minutes, seconds = divmod(total_time, 60)

    # ---------------- DISPLAY & DOWNLOAD ----------------
    # Check if Markdown exists (for preview)
    if os.path.exists("sow.md"):
        st.success(f"✅ SOW Generated Successfully in {int(minutes)}m {int(seconds)}s ({total_time:.2f}s)")

        with open("sow.md", "r", encoding="utf-8") as f:
            sow_md = f.read()

        # Display content in UI
        st.markdown(sow_md)

        st.divider()
        st.subheader("⬇️ Download Generated SOW")

        # Check if Docx exists (for download)
        if os.path.exists("sow.docx"):
            with open("sow.docx", "rb") as f:
                st.download_button(
                    "Download Word (.docx)",
                    data=f,
                    file_name="sow.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

    else:
        st.error("❌ SOW generation failed")