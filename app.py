import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import json

# Page Configuration
st.set_page_config(
    page_title="UZ SCIENCE AI - Professional Ilmiy Platforma",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Sleek CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%);
        padding: 24px;
        border-radius: 16px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(2, 132, 199, 0.25);
    }
    
    .main-header h1 {
        color: white !important;
        font-size: 2.2rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
    }
    
    .main-header p {
        color: #e0f2fe;
        font-size: 1rem;
        margin-top: 6px;
        margin-bottom: 0;
    }
    
    .metric-badge {
        background: #f0f9ff;
        border: 1px solid #bae6fd;
        border-radius: 10px;
        padding: 10px 14px;
        display: inline-block;
        font-size: 0.85rem;
        color: #0369a1;
        font-weight: 600;
        margin-right: 8px;
    }
    
    .feature-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        margin-bottom: 16px;
    }
    
    .feature-card h4 {
        color: #0f172a;
        margin-top: 0;
        font-size: 1.05rem;
        font-weight: 700;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        background-color: #f1f5f9;
    }

    .stTabs [aria-selected="true"] {
        background-color: #0284c7 !important;
        color: white !important;
    }
    
    .stDownloadButton button {
        width: 100%;
        border-radius: 10px;
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# Top Banner
st.markdown("""
<div class="main-header">
    <h1>🎓 UZ SCIENCE AI — Professional Ilmiy Platforma</h1>
    <p>Ilmiy maqola va dissertatsiyalarni to‘liq akademik tarjima qilish, chuqur tahlil va magistrlik xulosalarini tayyorlash tizimi</p>
</div>
""", unsafe_allow_html=True)

# Helper function to generate PDF
def create_pdf(title, translation_text, summary_data, thesis_data, terms_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#0284c7'), spaceAfter=8)
    h1_style = ParagraphStyle('H1', parent=styles['Heading2'], fontSize=12, leading=16, textColor=colors.HexColor('#0f172a'), spaceBefore=12, spaceAfter=6)
    h2_style = ParagraphStyle('H2', parent=styles['Heading3'], fontSize=10, leading=14, textColor=colors.HexColor('#0369a1'), spaceBefore=8, spaceAfter=4)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=8.5, leading=12.5, textColor=colors.HexColor('#334155'), spaceAfter=3)
    bold_label = ParagraphStyle('BoldLabel', parent=styles['Normal'], fontSize=8.5, leading=12.5, textColor=colors.HexColor('#0f172a'))
    
    story = []
    clean_title = str(title).replace('<', '&lt;').replace('>', '&gt;')
    story.append(Paragraph(f"<b>UZ SCIENCE AI — {clean_title}</b>", title_style))
    story.append(Paragraph("<i>To‘liq akademik tarjima, ilmiy pasport va magistr dissertatsiyasi tahlili</i>", body_style))
    story.append(Spacer(1, 8))
    
    # 1. Chuqur Ilmiy Xulosa
    story.append(Paragraph("<b>1. Chuqur Ilmiy Xulosa va Tahlil (Executive Summary)</b>", h1_style))
    for k, v in summary_data.items():
        label = k.replace('_', ' ').capitalize()
        story.append(Paragraph(f"<b>• {label}:</b> {str(v).replace('<', '&lt;').replace('>', '&gt;')}", body_style))
    
    story.append(Spacer(1, 8))
    
    # 2. Magistr Yo'riqnomasi
    story.append(Paragraph("<b>2. Magistrlik Dissertatsiyasi Uchun Tavsiyalar</b>", h1_style))
    for k, v in thesis_data.items():
        label = k.replace('_', ' ').capitalize()
        story.append(Paragraph(f"<b>• {label}:</b> {str(v).replace('<', '&lt;').replace('>', '&gt;')}", body_style))
        
    story.append(Spacer(1, 8))

    # 3. Terminlar
    if terms_data:
        story.append(Paragraph("<b>3. Asosiy Ilmiy Terminlar</b>", h1_style))
        for item in terms_data:
            term_en = item.get("term_en", "")
            term_uz = item.get("term_uz", "")
            desc = item.get("desc", "")
            story.append(Paragraph(f"<b>• {term_en} → {term_uz}:</b> {desc}", body_style))
        story.append(Spacer(1, 8))

    # 4. To'liq Tarjima
    story.append(Paragraph("<b>4. To‘liq Akademik O‘zbekcha Tarjima</b>", h1_style))
    for p in translation_text.split('\n'):
        if p.strip():
            clean_p = p.replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(clean_p, body_style))
            
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# Sidebar: Settings & API Key
selected_model = "gemini-3.6-flash"

with st.sidebar:
    st.header("⚙️ Tizim Sozlamalari")
    api_key = st.text_input(
        "Google Gemini API Kalit:",
        type="password",
        help="aistudio.google.com saytidan olingan bepul kalit"
    )
    st.markdown("💡 [Bepul API Kalit Olish (Google AI Studio)](https://aistudio.google.com/app/apikey)")
    
    model_options = [
        "gemini-3.6-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-3.1-pro"
    ]
    
    if api_key.strip():
        try:
            genai.configure(api_key=api_key.strip())
            fetched_models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            if fetched_models:
                for fm in fetched_models:
                    if fm not in model_options:
                        model_options.append(fm)
        except Exception:
            pass

    selected_model = st.selectbox(
        "🤖 AI Modeli:",
        model_options,
        index=0,
        help="Gemini 3.6 Flash eng tezkor va bepul ilmiy modeldir"
    )

    st.divider()
    st.markdown("### 🌟 Yangi Imkoniyatlar")
    st.markdown("✅ **To‘liq Ilmiy Tarjima:** Bo‘limma-bo‘lim to‘liq tarjima")
    st.markdown("✅ **Chuqur Xulosa (Summary):** Muammo, taklif etilgan yechim va diqqat markazlari")
    st.markdown("✅ **Magistr Maslahatchisi:** Dissertatsiyada qo‘llash va iqtibos olish")
    st.markdown("✅ **Terminlar Glossariysi:** AI/CS atamalar izohi")
    st.markdown("✅ **2 xil Eksport:** Ham **PDF**, ham **Word (.docx)**")

# Main Section
tab_upload, tab_paste = st.tabs(["📁 Fayl Yuklash (PDF/PPTX/DOCX)", "✍️ Matnni Joylashtirish (Paste)"])

extracted_text = ""
file_name = "Ilmiy Hujjat"

with tab_upload:
    uploaded_file = st.file_uploader(
        "Ilmiy maqola, dissertatsiya yoki taqdimot faylini tanlang:",
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
                st.markdown(f"""
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; padding: 12px 16px; border-radius: 10px; color: #166534; font-size: 0.9rem; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span>✅ Fayl muvaffaqiyatli o‘qildi: <b>{len(extracted_text):,} ta belgi</b> ({file_name})</span>
                </div>
                """, unsafe_allow_html=True)
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
if st.button("🚀 MUKAMMAL TARJIMA VA CHUQUR TAHLIL QILISH", type="primary", use_container_width=True):
    if not extracted_text.strip():
        st.warning("⚠️ Iltimos, oldin fayl yuklang yoki matn kiriting.")
    elif not api_key.strip():
        st.error("⚠️ Iltimos, chap tarafdagi menyuga Gemini API kalitingizni kiriting.")
    else:
        with st.spinner(f"Google Gemini ({selected_model}) ilmiy maqolani to‘liq o‘rganmoqda va akademik tahlil qilmoqda..."):
            try:
                genai.configure(api_key=api_key.strip())
                model = genai.GenerativeModel(selected_model)

                system_prompt = """
Siz oliy toifali akademik tarjimon, ilmiy tahrirchi va magistrlik dissertatsiyalari bo'yicha ilmiy maslahatchisisiz.
Vazifangiz: Taqdim etilgan ilmiy maqola / akademik hujjatni quyidagi 4 ta asosiy bo'lim bo'yicha mukammal akademik o'zbek tilida tahlil qilib berish.

Qoidalar:
1. To'liq tarjima qismida (full_translation) hech narsani qisqartirmang, har bir bo'limni (Abstract, Intro, Methods, Results, Discussion, Conclusion) to'liq va ravon ilmiy tilda bering. Formulalar ($...$), kodlar va citationlarni to'liq saqlang.
2. Xulosa qismida (research_summary) quyidagi savollarga alohida chuqur javob bering:
   - core_problem: Maqolada ko'tarilgan asosiy ilmiy yoki amaliy muammo nima edi?
   - proposed_solution: Mualliflar qanday yangi yechim, model yoki metodologiya taklif qilishdi?
   - key_focus_areas: Ushbu maqolani o'qishda magistrant eng ko'p nimaga e'tibor berishi kerak? Qaysi qismlari eng muhim?
   - experimental_results: Qanday tajribalar o'tkazildi, qanday datasetlar va qanday SOTA ko'rsatkichlarga erishildi?
   - limitations: Tadqiqotning qanday cheklovlari va kamchiliklari mavjud?
3. Magistr tavsiyalari qismida (thesis_advisor) dissertatsiyaning qaysi bobida qanday iqtibos keltirish va ushbu maqoladan yangi tadqiqot g'oyasini olishni aniq ko'rsating.
4. Terminlar qismida (key_terms) kamida 4-6 ta asosiy atamani (term_en, term_uz, desc) bering.

Javobni FAQAT quyidagi JSON formatida qaytaring (hech qanday markdown ```json belgilarisiz):
{
  "full_translation": "To'liq, qisqartirilmagan akademik tarjima matni...",
  "research_summary": {
    "core_problem": "Asosiy muammo tavsifi...",
    "proposed_solution": "Taklif etilgan yechim va arxitektura...",
    "key_focus_areas": "Nimasiga alohida e'tibor berish kerak va qaysi qismlari muhim...",
    "experimental_results": "Asosiy tajribaviy natijalar va benchmarklar...",
    "limitations": "Tadqiqot cheklovlari..."
  },
  "thesis_advisor": {
    "where_to_cite": "Dissertatsiyaning qaysi bo'limida (Mavjud tadqiqotlar tahlili yoki Metodologiya) qanday iqtibos olish kerak...",
    "how_to_use_method": "Metodni o'z tadqiqotida qanday qo'llash mumkin...",
    "new_research_idea": "Ushbu maqola asosida magistr uchun yangi ilmiy g'oya / mavzu..."
  },
  "key_terms": [
    {"term_en": "Self-Attention", "term_uz": "O'z-o'ziga e'tibor", "desc": "Qisqa izoh..."}
  ]
}
"""
                response = model.generate_content(
                    system_prompt + "\n\nHujjat Matni:\n" + extracted_text[:28000],
                    generation_config={"temperature": 0.2}
                )

                raw_json = response.text.strip()
                if raw_json.startswith("```json"):
                    raw_json = raw_json[7:-3].strip()
                elif raw_json.startswith("```"):
                    raw_json = raw_json[3:-3].strip()

                data = json.loads(raw_json)
                st.session_state["result_data"] = data
                st.session_state["file_title"] = file_name
                st.success("✅ Mukammal tarjima va ilmiy tahlil to‘liq yakunlandi!")

            except Exception as err:
                st.error(f"Xatolik yuz berdi: {err}")

# Results Display & Tabs
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    summary = data.get("research_summary", {})
    thesis = data.get("thesis_advisor", {})
    terms = data.get("key_terms", [])
    full_trans = data.get("full_translation", "")
    current_title = st.session_state.get("file_title", "Ilmiy Maqola")

    st.divider()
    
    # Download Card Section
    st.markdown("### 📥 Tayyor Hujjatlarni Yuklab Olish")
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        try:
            pdf_bytes = create_pdf(current_title, full_trans, summary, thesis, terms)
            st.download_button(
                label="📥 To‘liq PDF Hisobotni Yuklab Olish",
                data=pdf_bytes,
                file_name=f"UZ_SCIENCE_AI_{current_title}.pdf",
                mime="application/pdf",
                type="primary"
            )
        except Exception as pdf_err:
            st.error(f"PDF tayyorlashda xatolik: {pdf_err}")

    with col_d2:
        try:
            doc = Document()
            doc.add_heading(current_title, level=1)
            
            doc.add_heading("1. Chuqur Ilmiy Xulosa", level=2)
            for k, v in summary.items():
                p = doc.add_paragraph()
                p.add_run(f"{k.replace('_', ' ').capitalize()}: ").bold = True
                p.add_run(str(v))
                
            doc.add_heading("2. Magistr Dissertatsiyasi Uchun Tavsiyalar", level=2)
            for k, v in thesis.items():
                p = doc.add_paragraph()
                p.add_run(f"{k.replace('_', ' ').capitalize()}: ").bold = True
                p.add_run(str(v))
                
            doc.add_heading("3. To‘liq Akademik Tarjima", level=2)
            doc.add_paragraph(full_trans)
            
            docx_io = io.BytesIO()
            doc.save(docx_io)
            docx_io.seek(0)

            st.download_button(
                label="📄 Tahrirlanadigan Word (.docx) Hujjatini Yuklab Olish",
                data=docx_io,
                file_name=f"UZ_SCIENCE_AI_{current_title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as docx_err:
            st.error(f"Word tayyorlashda xatolik: {docx_err}")

    st.markdown("<br>", unsafe_allow_html=True)

    # 4 Detailed Analysis Tabs
    res_tab1, res_tab2, res_tab3, res_tab4 = st.tabs([
        "🧠 Chuqur Ilmiy Xulosa (Summary)",
        "📄 To‘liq Akademik Tarjima",
        "🎓 Magistr Dissertatsiyasi Tavsiyalari",
        "📖 Ilmiy Terminlar Lug‘ati"
    ])

    # Tab 1: Deep Research Summary
    with res_tab1:
        st.subheader("📌 Maqolaning Asosiy Mazmuni va Diqqat Markazlari")
        
        # Key Focus Areas Callout
        st.markdown(f"""
        <div style="background: #eff6ff; border-left: 5px solid #0284c7; padding: 16px 20px; border-radius: 8px; margin-bottom: 20px;">
            <h4 style="color: #0369a1; margin: 0 0 6px 0; font-size: 1.1rem; font-weight: 700;">⚡ Nimasiga alohida e’tibor berish kerak? (Qaysi qismi muhim):</h4>
            <p style="color: #1e293b; font-size: 0.95rem; margin: 0; line-height: 1.6;">{summary.get('key_focus_areas', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.markdown(f"""
            <div class="feature-card">
                <h4>🎯 1. Ko‘tarilgan Asosiy Muammo</h4>
                <p style="color: #334155; font-size: 0.92rem; line-height: 1.55;">{summary.get('core_problem', '-')}</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="feature-card">
                <h4>🔬 3. Natijalar va Benchmarklar</h4>
                <p style="color: #334155; font-size: 0.92rem; line-height: 1.55;">{summary.get('experimental_results', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

        with col_s2:
            st.markdown(f"""
            <div class="feature-card">
                <h4>💡 2. Taklif Etilgan Yechim & Metod</h4>
                <p style="color: #334155; font-size: 0.92rem; line-height: 1.55;">{summary.get('proposed_solution', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
            <div class="feature-card">
                <h4>⏳ 4. Cheklovlar va Kamchiliklar</h4>
                <p style="color: #334155; font-size: 0.92rem; line-height: 1.55;">{summary.get('limitations', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

    # Tab 2: Full Translation
    with res_tab2:
        st.subheader("🇺🇿 Hech Qayeri Qisqartirilmagan To‘liq Akademik Tarjima")
        st.markdown(full_trans)

    # Tab 3: Master Thesis Advisor
    with res_tab3:
        st.subheader("📚 Magistr Dissertatsiyasi Uchun Yo‘riqnoma")
        
        st.markdown(f"""
        <div class="feature-card">
            <h4>📌 Qayerda va qanday iqtibos (citation) olish kerak?</h4>
            <p style="color: #334155; font-size: 0.95rem; line-height: 1.6;">{thesis.get('where_to_cite', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="feature-card">
            <h4>🔬 Ushbu metodologiyani o‘z tadqiqotingizda qanday qo‘llash mumkin?</h4>
            <p style="color: #334155; font-size: 0.95rem; line-height: 1.6;">{thesis.get('how_to_use_method', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="feature-card" style="border-left: 4px solid #f59e0b;">
            <h4 style="color: #b45309;">💡 Ushbu maqola asosida yangi dissertatsiya / maqola g‘oyasi:</h4>
            <p style="color: #334155; font-size: 0.95rem; line-height: 1.6; font-weight: 500;">{thesis.get('new_research_idea', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

    # Tab 4: Terminology Glossary
    with res_tab4:
        st.subheader("📖 Maqoladagi Asosiy Ilmiy Terminlar Lug‘ati")
        if terms:
            for item in terms:
                st.markdown(f"""
                <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px 16px; margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                        <span style="font-weight: 700; color: #0284c7; font-size: 0.95rem;">{item.get('term_en', '')}</span>
                        <span style="font-weight: 600; color: #0f172a; font-size: 0.9rem; background: #e0f2fe; padding: 2px 8px; border-radius: 6px;">🇺🇿 {item.get('term_uz', '')}</span>
                    </div>
                    <p style="margin: 0; font-size: 0.88rem; color: #64748b;">{item.get('desc', '')}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Terminlar ro‘yxati mavjud emas.")
