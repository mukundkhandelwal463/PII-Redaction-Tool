import csv
import io
import json
import tempfile
import zipfile
from pathlib import Path
import streamlit as st

import redactor

st.set_page_config(
    page_title="Scaler AI Labs | PII Redaction Tool",
    page_icon="🛡️",
    layout="wide"
)

# Header Branding
st.markdown("### 🚀 Scaler AI Labs")
st.title("🛡️ Enterprise PII Redaction Tool")
st.markdown(
    "Detect and redact sensitive Personally Identifiable Information (PII) like names, emails, "
    "phone numbers, addresses, SSNs, credit card numbers, DOBs, and IP addresses."
)

st.sidebar.markdown("## 🚀 Scaler AI Labs")
st.sidebar.markdown("---")
st.sidebar.header("⚙️ Options")
option = st.sidebar.radio("Select Mode", ["Upload Document", "Evaluation & Metrics"])

if option == "Upload Document":
    st.header("📄 Upload Document for Redaction")
    uploaded_file = st.file_uploader("Upload a .txt or .docx file", type=["txt", "docx"])

    sample_text_button = st.button("Use Sample Ticket Log")
    input_text = ""
    file_name = "sample_ticket_log.txt"

    if sample_text_button:
        sample_path = Path("sample_ticket_log.txt")
        if sample_path.exists():
            input_text = sample_path.read_text(encoding="utf-8")
            st.success("Loaded sample_ticket_log.txt")

    elif uploaded_file is not None:
        file_name = uploaded_file.name
        if file_name.endswith(".docx"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name
            input_text = redactor.read_input_file(tmp_path)
            Path(tmp_path).unlink(missing_ok=True)
        else:
            input_text = uploaded_file.getvalue().decode("utf-8")

    if input_text:
        eda = redactor.get_basic_eda(input_text)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Lines", eda["total_lines"])
        col2.metric("Total Words", eda["total_words"])
        col3.metric("Total Characters", eda["total_characters"])

        found_items = redactor.find_pii(input_text)
        col4.metric("PII Spans Found", len(found_items))

        redacted_text, mapping = redactor.redact_text(input_text)

        st.subheader("🔍 PII Detection & Replacement Mapping Table")
        if mapping:
            st.dataframe(mapping, use_container_width=True)
        else:
            st.info("No PII detected in this document.")

        col_left, col_right = st.columns(2)
        with col_left:
            st.subheader("Original Text")
            st.text_area("Original", input_text, height=350)
        with col_right:
            st.subheader("Redacted Text")
            st.text_area("Redacted", redacted_text, height=350)

        # Generate Styled DOCX output
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
            tmp_out_path = tmp_out.name

        redactor.write_text_to_docx(redacted_text, tmp_out_path, len(mapping))
        docx_bytes = Path(tmp_out_path).read_bytes()
        Path(tmp_out_path).unlink(missing_ok=True)

        # Generate CSV Mapping string
        csv_buffer = io.StringIO()
        if mapping:
            writer = csv.DictWriter(csv_buffer, fieldnames=["type", "original", "replacement"])
            writer.writeheader()
            writer.writerows(mapping)
        csv_data = csv_buffer.getvalue()

        # Build ZIP archive of ALL deliverables
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr("redacted_output.docx", docx_bytes)
            zip_file.writestr("redacted_output.txt", redacted_text.encode("utf-8"))
            zip_file.writestr("redaction_mapping.json", json.dumps(mapping, indent=2).encode("utf-8"))
            zip_file.writestr("redaction_mapping.csv", csv_data.encode("utf-8"))
            zip_file.writestr("eda_summary.json", json.dumps(eda, indent=2).encode("utf-8"))
            if Path("evaluation_report.docx").exists():
                zip_file.writestr("evaluation_report.docx", Path("evaluation_report.docx").read_bytes())
            if Path("evaluation_report.md").exists():
                zip_file.writestr("evaluation_report.md", Path("evaluation_report.md").read_bytes())

        zip_data = zip_buffer.getvalue()

        st.markdown("---")
        st.subheader("📦 Download Deliverables & Processed Output")

        # Master ZIP Download Button
        st.download_button(
            label="📦 DOWNLOAD ALL DELIVERABLES (ZIP Archive)",
            data=zip_data,
            file_name="all_redacted_deliverables.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        st.markdown("##### Individual Download Links:")
        d_col1, d_col2, d_col3, d_col4 = st.columns(4)

        with d_col1:
            st.download_button(
                label="📄 Redacted DOCX (.docx)",
                data=docx_bytes,
                file_name="redacted_output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        with d_col2:
            st.download_button(
                label="📝 Redacted Text (.txt)",
                data=redacted_text.encode("utf-8"),
                file_name="redacted_output.txt",
                mime="text/plain",
                use_container_width=True
            )

        with d_col3:
            st.download_button(
                label="📊 Mapping CSV for Excel (.csv)",
                data=csv_data.encode("utf-8"),
                file_name="redaction_mapping.csv",
                mime="text/csv",
                use_container_width=True
            )

        with d_col4:
            st.download_button(
                label="🔗 Mapping JSON (.json)",
                data=json.dumps(mapping, indent=2).encode("utf-8"),
                file_name="redaction_mapping.json",
                mime="application/json",
                use_container_width=True
            )

elif option == "Evaluation & Metrics":
    st.header("📊 Model Evaluation & Benchmarks")
    test_file = Path("synthetic_eval_labeled_realistic_1200.txt")

    if not test_file.exists():
        redactor.make_large_test_file(test_file, 1200)

    report_path = Path("evaluation_report.md")
    accuracy, precision, recall = redactor.write_evaluation_report(test_file, report_path)

    col1, col2, col3 = st.columns(3)
    col1.metric("Accuracy", f"{accuracy:.2f}")
    col2.metric("Precision", f"{precision:.2f}")
    col3.metric("Recall", f"{recall:.2f}")

    if report_path.exists():
        st.markdown(report_path.read_text(encoding="utf-8"))

    if Path("evaluation_report.docx").exists():
        st.download_button(
            label="📥 Download Professional Evaluation Report (.docx)",
            data=Path("evaluation_report.docx").read_bytes(),
            file_name="evaluation_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
