import streamlit as st
import pypandoc
import tempfile
import os
import re
import time

st.set_page_config(
    page_title="Markdown to Word Converter",
    page_icon="📄",
    layout="centered"
)

# -----------------------------
# Styling
# -----------------------------
st.markdown("""
<style>
.main { padding-top: 2rem; }
.banner {
    background: linear-gradient(90deg, #1f4fd8, #4f46e5);
    padding: 1.2rem 1.5rem;
    border-radius: 12px;
    color: white;
    margin-bottom: 2rem;
}
.banner h1 { margin: 0; font-size: 24px; }
.banner p { margin: 0; opacity: 0.85; font-size: 14px; }
.card {
    background: #ffffff;
    padding: 1.5rem;
    border-radius: 12px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.05);
}
.footer {
    text-align: center;
    margin-top: 3rem;
    font-size: 12px;
    color: #888;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="banner">
    <h1>Markdown → Word Converter</h1>
    <p>Preserves JSON-LD schema automatically</p>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# JSON-LD Transformer
# -----------------------------
def transform_jsonld_scripts(markdown_text):
    pattern = re.compile(
        r"<script\s+type=['\"]application/ld\+json['\"]>\s*(.*?)\s*</script>",
        re.DOTALL | re.IGNORECASE
    )

    def replacer(match):
        json_content = match.group(1).strip()
        return f"\n```json\n{json_content}\n```\n"

    return pattern.sub(replacer, markdown_text)

# -----------------------------
# Main Card
# -----------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Markdown (.md) file", type=["md"])

if uploaded_file:

    if uploaded_file.size > 5 * 1024 * 1024:
        st.error("File exceeds 5MB limit.")
        st.stop()

    progress = st.progress(0)
    status_text = st.empty()

    try:
        # Step 1: Validation
        status_text.text("Validating file...")
        progress.progress(20)

        raw_markdown = uploaded_file.read().decode("utf-8")

        # Step 2: Transform JSON-LD
        status_text.text("Processing JSON-LD schema...")
        processed_markdown = transform_jsonld_scripts(raw_markdown)
        progress.progress(50)

        # Step 3: Temporary Save
        status_text.text("Preparing document...")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as tmp_md:
            tmp_md.write(processed_markdown.encode("utf-8"))
            tmp_md_path = tmp_md.name

        output_docx_path = tmp_md_path.replace(".md", ".docx")
        progress.progress(70)

        # Step 4: Conversion
        status_text.text("Converting to Word document...")
        pypandoc.convert_file(
            tmp_md_path,
            "docx",
            outputfile=output_docx_path,
            extra_args=["--standalone"]
        )
        progress.progress(90)

        # Step 5: Finalizing
        status_text.text("Finalizing...")
        progress.progress(100)

        with open(output_docx_path, "rb") as f:
            st.download_button(
                label="Download Converted Document",
                data=f,
                file_name="converted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        st.success("Conversion completed successfully.")

        os.remove(tmp_md_path)
        os.remove(output_docx_path)

        status_text.empty()

    except Exception as e:
        st.error(f"Conversion failed: {e}")
        progress.empty()
        status_text.empty()

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    Internal Utility • Structured Documentation Workflow • Version 1.1
</div>
""", unsafe_allow_html=True)