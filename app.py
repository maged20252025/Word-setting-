
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os

TEMPLATE_PATH = "دعوى استحقاق.docx"

st.set_page_config(page_title="منسق الدعاوى", layout="centered")
st.title("📄 منسق الدعاوى بنموذج موحد")

st.markdown("**اختر طريقة إدخال النص:**")
method = st.radio("", ["📤 رفع ملف Word", "✍️ لصق النص يدويًا"])

text = ""

if method == "📤 رفع ملف Word":
    uploaded = st.file_uploader("ارفع ملف Word", type=["docx"])
    if uploaded:
        doc = Document(uploaded)
        text = "\n".join([p.text for p in doc.paragraphs])
else:
    text = st.text_area("ألصق النص هنا", height=400)

if st.button("🔧 تنسيق النص وإنشاء الملف"):
    if not text.strip():
        st.error("يرجى إدخال أو رفع نص أولاً.")
    else:
        try:
            doc = Document(TEMPLATE_PATH)
            while len(doc.paragraphs):
                p = doc.paragraphs[0]
                p._element.getparent().remove(p._element)

            for line in text.strip().split('\n'):
                if line.strip():
                    para = doc.add_paragraph()
                    run = para.add_run(line.strip())
                    run.font.name = 'Traditional Arabic'
                    run.font.size = Pt(16)
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.success("✅ تم إنشاء الملف بنجاح!")
            st.download_button("📥 تحميل الملف المنسق", data=buffer, file_name="دعوى_منسقة.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        except Exception as e:
            st.error("حدث خطأ أثناء التنسيق:")
            st.exception(e)
