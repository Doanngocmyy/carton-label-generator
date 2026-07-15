#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit web app for carton_label_generator.py

Lets someone with no Python installed:
  1. Upload the Packing List Excel (.xlsx)
  2. Click "Generate labels"
  3. Download the 4x6" label PDF and the audit CSV

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy: see PUBLISH_GUIDE.md in this folder.
"""

import io
import tempfile
import traceback
from contextlib import redirect_stdout
from pathlib import Path

import streamlit as st

from carton_label_generator import main as generate_labels

st.set_page_config(page_title="Carton Label Generator", page_icon="\U0001F4E6", layout="centered")

st.title("Carton Label Generator")
st.write(
    "Upload the Packing List Excel file. The app will generate scan-safe "
    "4x6 inch carton labels (PDF) plus an audit CSV."
)

uploaded_file = st.file_uploader("Packing List (.xlsx)", type=["xlsx"])

max_skus = 3  # fixed default, matches the original script

run = st.button("Generate labels", type="primary", disabled=uploaded_file is None)

if run and uploaded_file is not None:
    with st.spinner("Reading the workbook and generating labels..."):
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_dir_path = Path(tmp_dir)
            input_path = tmp_dir_path / uploaded_file.name
            input_path.write_bytes(uploaded_file.getvalue())

            output_dir = tmp_dir_path / "LABEL_OUTPUT"
            log_buffer = io.StringIO()

            try:
                with redirect_stdout(log_buffer):
                    output_pdf, output_csv = generate_labels(
                        input_file=input_path,
                        output_dir=output_dir,
                        max_skus_per_label=max_skus,
                    )
                pdf_bytes = Path(output_pdf).read_bytes()
                csv_bytes = Path(output_csv).read_bytes()
            except Exception as exc:  # noqa: BLE001
                st.error(f"Could not generate labels: {exc}")
                with st.expander("Technical details"):
                    st.code(traceback.format_exc())
                st.stop()

    st.success("Labels generated.")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "Download label PDF",
            data=pdf_bytes,
            file_name=Path(output_pdf).name,
            mime="application/pdf",
            type="primary",
        )
    with col2:
        st.download_button(
            "Download audit CSV",
            data=csv_bytes,
            file_name=Path(output_csv).name,
            mime="text/csv",
        )

    with st.expander("Validation summary / log"):
        st.text(log_buffer.getvalue())

    st.caption(
        "Print settings: 4 x 6 inch, Portrait, Actual Size / 100% "
        "(do NOT use Fit / Shrink / Scale to page)."
    )

st.divider()
st.caption(
    "Expected columns in the sheet: PO No., Packaging code, SKU#, BarCode/UPC, Quantity "
    "(English or bilingual header row is auto-detected)."
)
