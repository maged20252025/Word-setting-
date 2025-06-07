
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

TEMPLATE_PATH = "دعوى استحقاق.docx"

st.title("📄 أداة تنسيق الدعاوى تلقائيًا")

option = st.radio("طريقة إدخال النص:", ["📥 رفع ملف Word", "✍️ لصق نص يدويًا"])

text_content = ""

if option == "📥 رفع ملف Word":
    uploaded_file = st.file_uploader("ارفع ملف الدعوى (Word)", type=["docx"])
    if uploaded_file:
        doc = Document(uploaded_file)
        text_content = "\n".join([p.text for p in doc.paragraphs])
elif option == "✍️ لصق نص يدويًا":
    text_content = st.text_area("ألصق نص الدعوى هنا", height=400)

if st.button("🔧 تنسيق وإنشاء الملف"):
    if text_content.strip() == "":
        st.error("يرجى لصق النص أو رفع ملف قبل التنسيق.")
    else:
        template_doc = Document(TEMPLATE_PATH)
        while len(template_doc.paragraphs):
            p = template_doc.paragraphs[0]
            p._element.getparent().remove(p._element)

        for line in text_content.strip().split('\n'):
            if line.strip():
                p = template_doc.add_paragraph()
                run = p.add_run(line.strip())
                run.font.name = 'Traditional Arabic'
                run.font.size = Pt(16)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        buffer = io.BytesIO()
        template_doc.save(buffer)
        buffer.seek(0)

        st.success("✅ تم إنشاء الملف بنجاح!")
        st.download_button("📥 تحميل الملف المنسق", data=buffer, file_name="دعوى_منسقة.docx")
