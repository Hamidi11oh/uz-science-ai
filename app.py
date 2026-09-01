import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
import json

# Page Configuration
st.set_page_config(
    page_title="UZ SCIENCE AI - Akademik Tarjimon",
    page_icon="🎓",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 800; color: #0284c7; }
    .sub-title { font-size: 1rem; color: #64748b; margin-bottom: 20px; }
    .stDownloadButton button { width: 100%; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🎓 UZ SCIENCE AI</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Magistrant va tadqiqotchilar uchun ilmiy maqola va taqdimotlar tarjimoni hamda dissertatsiya assistenti</div>', unsafe_allow_html=True)

# Helper function to generate PDF
def create_pdf(title, translation_text, passport_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, leading=20, textColor='#0284c7', spaceAfter=10)
    h2_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontSize=12, leading=16, textColor='#1e293b', spaceBefore=12, spaceAfter=6)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=9.5, leading=14, textColor='#334155', spaceAfter=4)
    
    story = []
    clean_title = title.replace('<', '&lt;').replace('>', '&gt;')
    story.append(Paragraph(f"<b>UZ SCIENCE AI — {clean_title}</b>", title_style))
    story.append(Spacer(1, 8))
    
    story.append(Paragraph("<b>1. Akademik O‘zbekcha Tarjima</b>", h2_style))
    for p in translation_text.split('\n'):
        if p.strip():
            clean_p = p.replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(clean_p, body_style))
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>2. Magistr Ilmiy Pasporti & Tahlil</b>", h2_style))
    
    label_map = {
        "objective": "🎯 Tadqiqot maqsadi",
        "problem": "⚠️ Asosiy muammo",
        "methodology": "🔬 Metodologiya",
        "results": "📊 Asosiy natijalar",
        "novelty": "✨ Ilmiy yangilik",
        "limitations": "⏳ Cheklovlar",
        "cite": "📌 Dissertatsiyada iqtibos qilish",
        "idea": "💡 Yangi tadqiqot g'oyasi"
    }
    
    for k, v in passport_dict.items():
        label = label_map.get(k, k.capitalize())
        clean_val = str(v).replace('<', '&lt;').replace('>', '&gt;')
        story.append(Paragraph(f"<b>{label}:</b> {clean_val}", body_style))
        story.append(Spacer(1, 3))
        
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# Sidebar: Settings & API Key
with st.sidebar:
    st.header("⚙️ Sozlamalar")
    api_key = st.text_input(
        "Google Gemini API Kalit:",
        type="password",
        help="aistudio.google.com saytidan olingan bepul kalit"
    )
    st.markdown("💡 [Bepul API Kalit Olish (Google AI Studio)](https://aistudio.google.com/app/apikey)")
    
    st.divider()
    st.markdown("### 📋 Platforma Imkoniyatlari")
    st.markdown("✔️ **Formatlar:** PDF, PPTX (taqdimot), DOCX, TXT")
    st.markdown("✔️ **Formatlash:** Formulalar, LaTeX va iqtiboslar saqlanadi")
    st.markdown("✔️ **Magistr Pasporti:** SOTA natijalar, yangilik, thesis tavsiyalari")
    st.markdown("✔️ **Eksport:** Ham **PDF**, ham **Word (.docx)** yuklab olish")

# Main Tabs: File Upload vs Text Paste
tab_upload, tab_paste = st.tabs(["📁 Fayl Yuklash (PDF/PPTX/DOCX)", "✍️ Matnni Joylashtirish (Paste)"])

extracted_text = ""
file_name = "Ilmiy Hujjat"

with tab_upload:
    uploaded_file = st.file_uploader(
        "Ilmiy maqola, dissertatsiya yoki taqdimot faylini yuklang:",
        type=["pdf", "docx", "pptx", "txt"]
    )
    if uploaded_file is not None:
        file_name = uploaded_file.name
        suffix = file_name.split(".")[-1].lower()
        
        try:
            bytes_data = uploaded_file.read()
            if suffix == "pdf":
                doc = fitz.open(stream=bytes_data, filetype="pdf")
                pages = []
                for i, page in enumerate(doc):
                    t = page.get_text()
                    if t.strip():
                        pages.append(f"[Sahifa {i+1}]:\n" + t)
                extracted_text = "\n\n".join(pages)
            elif suffix == "docx":
                doc = Document(io.BytesIO(bytes_data))
                extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text])
            elif suffix == "pptx":
                from pptx import Presentation
                prs = Presentation(io.BytesIO(bytes_data))
                slides = []
                for idx, slide in enumerate(prs.slides):
                    slide_texts = []
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text:
                            slide_texts.append(shape.text)
                    slides.append(f"[Slayd {idx+1}]:\n" + "\n".join(slide_texts))
                extracted_text = "\n\n".join(slides)
            else:
                extracted_text = bytes_data.decode("utf-8", errors="ignore")
            
            if extracted_text.strip():
                st.success(f"✅ Fayl muvaffaqiyatli o'qildi: {len(extracted_text)} ta belgi")
            else:
                st.warning("⚠️ Fayldan matn ajratib bo'lmadi. Agar bu skaner qilingan rasm bo'lsa, 'Matnni Joylashtirish' bo'limidan foydalaning.")
        except Exception as e:
            st.error(f"Faylni o'qishda xatolik: {e}")

with tab_paste:
    direct_text = st.text_area("Maqola matnini to'g'ridan-to'g'ri shu yerga qo'ying:", height=180)
    if direct_text.strip():
        extracted_text = direct_text
        file_name = "Kiritilgan Maqola Matni"

# Action Button
if st.button("🚀 TARJIMA VA TAHLIL QILISH", type="primary", use_container_width=True):
    if not extracted_text.strip():
        st.warning("⚠️ Iltimos, oldin fayl yuklang yoki matn kiriting.")
    elif not api_key.strip():
        st.error("⚠️ Iltimos, chap tarafdagi menyuga Gemini API kalitingizni kiriting.")
    else:
        with st.spinner("Google Gemini AI maqolani tahlil qilmoqda va akademik o'zbek tiliga o'girmoqda..."):
            try:
                genai.configure(api_key=api_key)

                system_prompt = """
Siz oliy toifali akademik tarjimon va magistr ilmiy maslahatchisisiz.
Vazifangiz: Berilgan ilmiy maqola yoki akademik hujjatni ilmiy o'zbek tiliga tarjima qilish va magistr uchun tahlil pasportini tuzish.
Qoidalar:
1. So'zma-so'z xom tarjima qilmang, akademik ilmiy uslub va ravon tildan foydalaning.
2. Dasturlash kodlari, formulalar ($...$) va citationlarni ([1], [2]) o'zgartirmang.
3. Javobni FAQAT quyidagi toza JSON formatida qaytaring (hech qanday markdown ```json belgilarisiz):
{
  "translated_text": "To'liq o'zbekcha akademik tarjima...",
  "passport": {
    "objective": "Tadqiqot maqsadi...",
    "problem": "Asosiy ilmiy muammo...",
    "methodology": "Qo'llanilgan metodologiya...",
    "results": "Asosiy natijalar...",
    "novelty": "Asosiy ilmiy yangilik...",
    "limitations": "Talablar va cheklovlar...",
    "cite": "Magistr dissertatsiyasida qayerda va qanday foydalanish mumkin...",
    "idea": "Thesis uchun yangi tadqiqot g'oyasi..."
  }
}
"""
                # Free-tier prioritized models: Flash models are 100% free with unlimited quota
                free_tier_models = [
                    "gemini-1.5-flash",
                    "gemini-1.5-flash-latest",
                    "gemini-2.0-flash",
                    "gemini-1.5-flash-8b",
                    "gemini-2.0-flash-exp"
                ]

                response = None
                last_error = None

                for model_candidate in free_tier_models:
                    try:
                        m = genai.GenerativeModel(model_candidate)
                        response = m.generate_content(
                            system_prompt + "\n\nHujjat Matni:\n" + extracted_text[:25000],
                            generation_config={"temperature": 0.2}
                        )
                        if response and response.text:
                            break
                    except Exception as m_err:
                        last_error = m_err
                        continue

                if response is None or not response.text:
                    raise Exception(f"Barcha bepul Flash modellari sinab ko'rildi, oxirgi xato: {last_error}")

                raw_json = response.text.strip()
                if raw_json.startswith("```json"):
                    raw_json = raw_json[7:-3].strip()
                elif raw_json.startswith("```"):
                    raw_json = raw_json[3:-3].strip()

                data = json.loads(raw_json)
                st.session_state["result_data"] = data
                st.session_state["file_title"] = file_name
                st.success("✅ Tarjima va ilmiy tahlil muvaffaqiyatli yakunlandi!")

            except Exception as err:
                st.error(f"Xatolik yuz berdi: {err}")

# Results Display and Export
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    p = data.get("passport", {})
    t_text = data.get("translated_text", "")
    current_title = st.session_state.get("file_title", "Ilmiy Tarjima")

    st.divider()
    res_tab1, res_tab2 = st.tabs(["📑 Akademik Tarjima", "🎓 Magistr Pasporti & Tahlil"])

    with res_tab1:
        st.subheader("🇺🇿 O'zbekcha Akademik Tarjima")
        st.markdown(t_text)

        # Export Buttons: PDF and Word
        st.divider()
        st.markdown("### 📥 Hujjatni Eksport Qilish")
        col_pdf, col_docx = st.columns(2)
        
        with col_pdf:
            try:
                pdf_data = create_pdf(current_title, t_text, p)
                st.download_button(
                    label="📥 PDF Formatida Yuklab Olish",
                    data=pdf_data,
                    file_name=f"Tarjima_{current_title}.pdf",
                    mime="application/pdf",
                    type="primary"
                )
            except Exception as pdf_err:
                st.error(f"PDF tayyorlashda xatolik: {pdf_err}")

        with col_docx:
            try:
                doc = Document()
                doc.add_heading(current_title, level=1)
                doc.add_paragraph(t_text)
                doc.add_heading("Magistr Tahlil Pasporti", level=2)
                for k, v in p.items():
                    doc.add_paragraph(f"{k.capitalize()}: {v}")
                
                docx_io = io.BytesIO()
                doc.save(docx_io)
                docx_io.seek(0)

                st.download_button(
                    label="📄 Word (.docx) Formatida Yuklab Olish",
                    data=docx_io,
                    file_name=f"Tarjima_{current_title}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as docx_err:
                st.error(f"Word hujjati tayyorlashda xatolik: {docx_err}")

    with res_tab2:
        st.subheader("📌 Maqola Pasporti (Executive Summary)")
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"**🎯 Tadqiqot maqsadi:**\n\n{p.get('objective', '-')}")
            st.warning(f"**⚠️ Asosiy muammo:**\n\n{p.get('problem', '-')}")
            st.success(f"**🔬 Metodologiya:**\n\n{p.get('methodology', '-')}")
        with c2:
            st.info(f"**📊 Asosiy natijalar:**\n\n{p.get('results', '-')}")
            st.success(f"**✨ Ilmiy yangilik:**\n\n{p.get('novelty', '-')}")
            st.warning(f"**⏳ Cheklovlar:**\n\n{p.get('limitations', '-')}")

        st.divider()
        st.markdown("### 📚 Magistr Dissertatsiyasi Uchun Tavsiyalar")
        st.markdown(f"**📌 Qayerda iqtibos (cite) qilish kerak:** {p.get('cite', '-')}")
        st.markdown(f"**💡 Yangi ilmiy tadqiqot g'oyasi:** {p.get('idea', '-')}")
