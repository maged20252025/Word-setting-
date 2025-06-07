
import streamlit as st
from docxtpl import DocxTemplate
import io

TEMPLATE_PATH = "template_docxtpl.docx"

st.set_page_config(page_title="منسق الدعاوى باحتراف", layout="centered")
st.title("📄 منسق الدعاوى باستخدام قالب احترافي")

uploaded_doc = st.file_uploader("📤 ارفع ملف الدعوى (Word غير منسق)", type=["docx"])

if uploaded_doc and st.button("🔧 تنسيق الملف"):
    try:
        from docx import Document
        raw_doc = Document(uploaded_doc)
        lines = [p.text.strip() for p in raw_doc.paragraphs if p.text.strip()]
        full_text = "\n\n".join(lines)

        tpl = DocxTemplate(TEMPLATE_PATH)
        tpl.render({ "المحتوى": full_text })

        buffer = io.BytesIO()
        tpl.save(buffer)
        buffer.seek(0)

        st.success("✅ تم تنسيق الملف بنجاح!")
        st.download_button("📥 تحميل الملف المنسق", data=buffer, file_name="دعوى_منسقة.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except Exception as e:
        st.error("حدث خطأ أثناء التنسيق:")
        st.exception(e)
