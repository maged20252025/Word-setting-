import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import os

st.set_page_config(page_title="منسق المستندات القانونية", layout="centered")
st.title("📄 منسق المستندات القانونية")
st.markdown("ارفع ملف Word غير منسق، واختر قالبًا لتطبيق تنسيقه على الملف.")

uploaded_file = st.file_uploader("📤 ارفع ملف غير منسق (.docx)", type=["docx"])

TEMPLATE_DIR = "templates"
template_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".docx")]
selected_template = st.selectbox("📑 اختر القالب:", template_files)

if uploaded_file and selected_template:
    if st.button("✅ تنسيق الملف"):
        try:
            # قراءة الملف غير المنسق
            original_doc = Document(uploaded_file)

            # قراءة القالب واستخراج أول فقرة
            template_path = os.path.join(TEMPLATE_DIR, selected_template)
            template_doc = Document(template_path)
            if not template_doc.paragraphs:
                st.error("❌ القالب لا يحتوي على فقرات.")
                st.stop()

            sample_para = template_doc.paragraphs[0]
            sample_run = sample_para.runs[0] if sample_para.runs else None

            # تجهيز مستند جديد
            new_doc = Document()

            for para in original_doc.paragraphs:
                if para.text.strip():
                    new_para = new_doc.add_paragraph()
                    new_run = new_para.add_run(para.text)

                    # تطبيق المحاذاة من القالب
                    new_para.alignment = sample_para.alignment

                    if sample_run:
                        font = new_run.font
                        sample_font = sample_run.font
                        font.name = sample_font.name
                        font.size = sample_font.size
                        font.bold = sample_font.bold
                        font.italic = sample_font.italic
                        font.underline = sample_font.underline
                        if sample_font.color and sample_font.color.rgb:
                            font.color.rgb = sample_font.color.rgb

            # حفظ الملف الجديد
            output = BytesIO()
            new_doc.save(output)
            output.seek(0)

            st.success("✅ تم تنسيق الملف بنجاح!")
            st.download_button("📥 تحميل الملف المنسق", output, file_name="ملف_منسق.docx")
        except Exception as e:
            st.error(f"حدث خطأ أثناء التنسيق: {str(e)}")
