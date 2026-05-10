import streamlit as st
from pdf2docx import Converter
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os
import tempfile

st.title("📄 PDF ↔ Word 双向转换工具")
st.info("支持拖拽上传文件，自动识别格式并转换，转换完成后可直接下载。")

def docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    y = height - 40
    for para in doc.paragraphs:
        if para.text.strip() != "":
            c.drawString(40, y, para.text)
            y -= 15
            if y < 40:
                c.showPage()
                y = height - 40
    c.save()

uploaded_file = st.file_uploader("上传文件（.pdf 或 .docx）", type=["pdf", "docx"])

if uploaded_file is not None:
    file_name = uploaded_file.name
    file_ext = os.path.splitext(file_name)[1].lower()

    with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    output_path = None
    output_name = None

    if file_ext == ".pdf":
        st.write("正在将 PDF 转换为 Word...")
        output_name = os.path.splitext(file_name)[0] + ".docx"
        output_path = tempfile.mktemp(suffix=".docx")
        cv = Converter(tmp_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()

    elif file_ext == ".docx":
        st.write("正在将 Word 转换为 PDF...")
        output_name = os.path.splitext(file_name)[0] + ".pdf"
        output_path = tempfile.mktemp(suffix=".pdf")
        docx_to_pdf(tmp_path, output_path)

    if output_path and os.path.exists(output_path):
        with open(output_path, "rb") as f:
            st.download_button(
                label=f"下载转换后的 {os.path.splitext(output_name)[1][1:].upper()} 文件",
                data=f,
                file_name=output_name
            )
        os.unlink(output_path)
    os.unlink(tmp_path)