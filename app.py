import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import json
import os

# Page Configuration
st.set_page_config(
    page_title="UZ SCIENCE AI - Ilmiy Tarjima Platformasi",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom Exact-Match Tailwind/Clean CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* Background */
    .stApp {
        background-color: #f8fafc;
    }
    
    /* Container Width */
    .block-container {
        max-width: 960px !important;
        padding-top: 1.5rem !important;
        padding-bottom: 4rem !important;
    }
    
    /* Top Header Bar */
    .top-nav {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 12px 20px;
        margin-bottom: 28px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    
    .nav-brand {
        display: flex;
        align-items: center;
        gap: 12px;
    }
    
    .nav-logo {
        width: 40px;
        height: 40px;
        border-radius: 12px;
        background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%);
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 1.2rem;
        box-shadow: 0 4px 10px rgba(2, 132, 199, 0.25);
    }
    
    .badge-active {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        color: #166534;
        padding: 3px 10px;
        border-radius: 14px;
        font-size: 0.72rem;
        font-weight: 700;
        margin-left: 8px;
    }
    
    /* Hero Section */
    .hero-container {
        text-align: center;
        margin-bottom: 28px;
    }
    
    .hero-title {
        font-size: 2.35rem;
        font-weight: 800;
        letter-spacing: -0.8px;
        color: #0f172a;
        margin-bottom: 10px;
    }
    
    .hero-highlight {
        background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .hero-subtitle {
        font-size: 0.98rem;
        color: #64748b;
        max-width: 620px;
        margin: 0 auto;
        line-height: 1.55;
    }
    
    /* Floating Main Card */
    .main-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 32px;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.03), 0 8px 10px -6px rgba(0, 0, 0, 0.02);
        margin-bottom: 24px;
    }
    
    /* Options Box Grid */
    .option-box {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 14px;
        font-size: 0.85rem;
    }
    
    .option-title {
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
        font-size: 0.88rem;
    }
    
    /* Result Section Cards */
    .res-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    
    .res-card h4 {
        color: #0f172a;
        margin-top: 0;
        margin-bottom: 10px;
        font-size: 1.05rem;
        font-weight: 700;
    }
    
    .focus-banner {
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        border: 1px solid #bfdbfe;
        border-left: 6px solid #0284c7;
        padding: 20px 24px;
        border-radius: 14px;
        margin-bottom: 20px;
    }
    
    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 8px 18px;
        font-weight: 600;
        font-size: 0.88rem;
        color: #64748b;
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
    }

    .stTabs [aria-selected="true"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border-color: #0284c7 !important;
    }
    
    /* Primary Button */
    .stButton button {
        background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%) !important;
        color: white !important;
        border: none !important;
        padding: 14px 28px !important;
        border-radius: 14px !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        box-shadow: 0 10px 20px -3px rgba(2, 132, 199, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    
    .stButton button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 14px 25px -3px rgba(2, 132, 199, 0.45) !important;
    }
    
    .footer-text {
        text-align: center;
        font-size: 0.78rem;
        color: #94a3b8;
        margin-top: 36px;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to generate PDF
def create_pdf(title, translation_text, summary_data, thesis_data, terms_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#0284c7'), spaceAfter=8)
    h1_style = ParagraphStyle('H1', parent=styles['Heading2'], fontSize=12, leading=16, textColor=colors.HexColor('#0f172a'), spaceBefore=12, spaceAfter=6)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=8.5, leading=12.5, textColor=colors.HexColor('#334155'), spaceAfter=3)
    
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

# Automatic Master API Key Resolution (from Streamlit Secrets or Environment)
MASTER_KEY = ""
if "GEMINI_API_KEY" in st.secrets:
    MASTER_KEY = st.secrets["GEMINI_API_KEY"]
elif os.environ.get("GEMINI_API_KEY"):
    MASTER_KEY = os.environ.get("GEMINI_API_KEY")

# Top Navigation Bar
status_badge = '<span class="badge-active"><span style="width:6px; height:6px; border-radius:50%; background:#22c55e;"></span> Gemini AI Faol</span>' if MASTER_KEY else '<span class="badge-active" style="background:#fffbeb; border-color:#fde68a; color:#b45309;">API Kalit kiritilmagan</span>'

st.markdown(f"""
<div class="top-nav">
    <div class="nav-brand">
        <div class="nav-logo">✨</div>
        <div>
            <div style="font-weight: 800; font-size: 1.15rem; color: #0f172a; display: flex; align-items: center;">
                UZ SCIENCE AI {status_badge}
            </div>
            <div style="font-size: 0.76rem; color: #64748b;">Magistrant va tadqiqotchilar uchun ilmiy tarjima platformasi</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Hero Section
st.markdown("""
<div class="hero-container">
    <div class="hero-title">Ilmiy maqolalarni <span class="hero-highlight">o‘zbek tiliga</span> tarjima qiling</div>
    <div class="hero-subtitle">PDF, DOCX, PPTX yoki matnni to‘g‘ridan-to‘g‘ri joylashtiring. Akademik uslub, formulalar va terminologiya to‘liq saqlanadi.</div>
</div>
""", unsafe_allow_html=True)

# Sidebar (Only for optional personal API key if needed)
with st.sidebar:
    st.header("⚙️ Sozlamalar")
    if MASTER_KEY:
        st.success("✅ Umumiy server API kaliti faol.")
        api_key = MASTER_KEY
    else:
        api_key = st.text_input(
            "Google Gemini API Kalit:",
            type="password",
            help="aistudio.google.com/app/apikey sahifasidan olingan kalit"
        )
        st.markdown("👉 [Bepul API Kalit Olish](https://aistudio.google.com/app/apikey)")

# Main Floating White Card
st.markdown('<div class="main-card">', unsafe_allow_html=True)

# Input Mode Tabs
tab_upload, tab_paste = st.tabs(["📄 Fayl Yuklash (PDF/DOCX/PPTX)", "✍️ Matnni Qo‘yish (Paste)"])

extracted_text = ""
file_name = "Ilmiy Hujjat"

with tab_upload:
    uploaded_file = st.file_uploader(
        "PDF yoki boshqa faylni shu yerga tashlang (PDF • DOCX • PPTX • TXT):",
        type=["pdf", "docx", "pptx", "txt"],
        label_visibility="collapsed"
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
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; padding: 10px 16px; border-radius: 10px; color: #166534; font-size: 0.88rem; font-weight: 600; display: flex; align-items: center; justify-content: space-between; margin-top: 10px;">
                    <span>📄 <b>{file_name}</b> muvaffaqiyatli yuklandi</span>
                    <span style="background: #dcfce7; padding: 2px 8px; border-radius: 6px; font-size: 0.78rem;">{len(extracted_text):,} belgi</span>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ Fayldan matn ajratib bo'lmadi. Agar bu skaner qilingan rasm bo'lsa, 'Matnni Qo‘yish' bo'limidan foydalaning.")
        except Exception as e:
            st.error(f"Faylni o'qishda xatolik: {e}")

with tab_paste:
    direct_text = st.text_area(
        "Maqola matnini joylashtiring:",
        height=180,
        placeholder="Abstract, Introduction yoki butun maqola matnini Ctrl+V qilib bu yerga tashlang...",
        label_visibility="collapsed"
    )
    if direct_text.strip():
        extracted_text = direct_text
        file_name = "Kiritilgan Maqola Matni"

# Presets & Settings Info Grid (Exact match with design)
st.markdown("<br>", unsafe_allow_html=True)
col_opt1, col_opt2 = st.columns(2)

with col_opt1:
    st.markdown("""
    <div class="option-box">
        <div class="option-title">🌐 Til sozlamalari</div>
        <div style="display: flex; gap: 8px;">
            <div style="flex:1; background:#ffffff; border:1px solid #cbd5e1; padding:6px 10px; border-radius:8px; font-size:0.78rem; color:#475569;">
                <b>Asl til:</b> 🔄 Avtomatik
            </div>
            <div style="flex:1; background:#f0f9ff; border:1px solid #bae6fd; padding:6px 10px; border-radius:8px; font-size:0.78rem; color:#0369a1;">
                <b>Tarjima:</b> 🇺🇿 O‘zbek tili
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_opt2:
    st.markdown("""
    <div class="option-box">
        <div class="option-title">🛡️ Akademik rejim qoidalari</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 6px; font-size: 0.76rem; color: #334155;">
            <div>☑️ Ilmiy uslub</div>
            <div>☑️ Terminlar bazasi</div>
            <div>☑️ Formulalar & Kod</div>
            <div>☑️ Iqtiboslar (Citations)</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# Big Action Button
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 TARJIMA VA TAHLIL QILISH", type="primary", use_container_width=True):
    if not extracted_text.strip():
        st.warning("⚠️ Iltimos, oldin fayl yuklang yoki matn kiriting.")
    elif not api_key:
        st.error("⚠️ Tizimda API kalit topilmadi. Iltimos, chap tarafdagi menyuga Gemini API kalitingizni kiriting.")
    else:
        with st.spinner("✨ Google Gemini 3.6 Flash ilmiy maqolani to‘liq o‘rganmoqda va akademik tahlil qilmoqda..."):
            try:
                genai.configure(api_key=api_key.strip())
                model = genai.GenerativeModel("gemini-3.6-flash")

                system_prompt = """
Siz oliy toifali akademik tarjimon, ilmiy tahrirchi va magistrlik dissertatsiyalari bo'yicha ilmiy maslahatchisisiz.
Vazifangiz: Taqdim etilgan ilmiy maqola / akademik hujjatni quyidagi 4 ta asosiy bo'lim bo'yicha mukammal akademik o'zbek tilida tahlil qilib berish.

Qoidalar:
1. To'liq tarjima qismida (full_translation) hech narsani qisqartirmang, har bir bo'limni (Abstract, Introduction, Methodology, Results, Discussion, Conclusion) to'liq va ravon ilmiy tilda bering. Formulalar ($...$), kodlar va citationlarni to'liq saqlang.
2. Xulosa qismida (research_summary) quyidagi savollarga alohida chuqur javob bering:
   - core_problem: Maqolada ko'tarilgan asosiy ilmiy yoki amaliy muammo nima edi?
   - proposed_solution: Mualliflar qanday yangi yechim, model yoki metodologiya taklif qilishdi?
   - key_focus_areas: Ushbu maqolani o'qishda magistrant eng ko'p nimaga e'tibor berishi kerak? Qaysi qismlari eng muhim?
   - experimental_results: Qanday tajribalar o'tkazildi, qanday datasetlar va qanday SOTA ko'rsatkichlarga erishildi?
   - limitations: Tadqiqotning qanday cheklovlari va kamchiliklari mavjud?
3. Magistr tavsiyalari qismida (thesis_advisor) dissertatsiyaning qaysi bobida qanday iqtibos keltirish va ushbu maqoladan yangi tadqiqot g'oyasini olishni aniq ko'rsating.
4. Terminlar qismida (key_terms) kamida 4-6 ta asosiy atamani (term_en, term_uz, desc) bering.

Javobni FAQAT quyidagi toza JSON formatida qaytaring (hech qanday markdown ```json belgilarisiz):
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
    "where_to_cite": "Dissertatsiyaning qaysi bo'limida qanday iqtibos olish kerak...",
    "how_to_use_method": "Metodni o'z tadqiqotida qanday qo'llash mumkin...",
    "new_research_idea": "Ushbu maqola asosida magistr uchun yangi ilmiy g'oya..."
  },
  "key_terms": [
    {"term_en": "Self-Attention", "term_uz": "O'z-o'ziga e'tibor mexanizmi", "desc": "Qisqa izoh..."}
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
                st.success("✨ Mukammal tarjima va ilmiy tahlil to‘liq yakunlandi!")

            except Exception as err:
                st.error(f"Xatolik yuz berdi: {err}")

st.markdown('</div>', unsafe_allow_html=True)

# Results Section
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    summary = data.get("research_summary", {})
    thesis = data.get("thesis_advisor", {})
    terms = data.get("key_terms", [])
    full_trans = data.get("full_translation", "")
    current_title = st.session_state.get("file_title", "Ilmiy Maqola")

    # Download Buttons Card
    st.markdown(f"""
    <div class="res-card" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; background: linear-gradient(135deg, #f0f9ff 0%, #ffffff 100%); border-color: #bae6fd;">
        <div>
            <div style="font-weight: 800; font-size: 1.15rem; color: #0369a1;">📄 {current_title}</div>
            <div style="font-size: 0.82rem; color: #64748b;">Akademik tarjima va tahliliy hisobot tayyor</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        try:
            pdf_bytes = create_pdf(current_title, full_trans, summary, thesis, terms)
            st.download_button(
                label="📥 PDF Formatida Yuklab Olish",
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
                label="📄 Word (.docx) Formatida Yuklab Olish",
                data=docx_io,
                file_name=f"UZ_SCIENCE_AI_{current_title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as docx_err:
            st.error(f"Word tayyorlashda xatolik: {docx_err}")

    st.markdown("<br>", unsafe_allow_html=True)

    # 4 Output Tabs
    res_tab1, res_tab2, res_tab3, res_tab4 = st.tabs([
        "🧠 Chuqur Ilmiy Xulosa (Summary)",
        "📄 To‘liq Akademik Tarjima",
        "🎓 Magistr Dissertatsiyasi Tavsiyalari",
        "📖 Ilmiy Terminlar Lug‘ati"
    ])

    # Tab 1: Deep Research Summary
    with res_tab1:
        st.markdown(f"""
        <div class="focus-banner">
            <div style="font-size: 0.78rem; font-weight: 800; text-transform: uppercase; color: #0284c7; letter-spacing: 0.5px; margin-bottom: 4px;">⚡ DIQQAT MARKAZI • ENG MUHIM QISM</div>
            <h4 style="color: #0c4a6e; margin: 0 0 6px 0; font-size: 1.1rem; font-weight: 800;">Ushbu maqolada nimasiga alohida e’tibor berish kerak?</h4>
            <p style="color: #1e293b; font-size: 0.94rem; margin: 0; line-height: 1.6;">{summary.get('key_focus_areas', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.markdown(f"""
            <div class="res-card">
                <h4>🎯 1. Ko‘tarilgan Asosiy Muammo</h4>
                <p style="color: #334155; font-size: 0.9rem; line-height: 1.55; margin: 0;">{summary.get('core_problem', '-')}</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="res-card">
                <h4>📊 3. Natijalar va Benchmarklar</h4>
                <p style="color: #334155; font-size: 0.9rem; line-height: 1.55; margin: 0;">{summary.get('experimental_results', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

        with col_s2:
            st.markdown(f"""
            <div class="res-card">
                <h4>💡 2. Taklif Etilgan Yechim & Metod</h4>
                <p style="color: #334155; font-size: 0.9rem; line-height: 1.55; margin: 0;">{summary.get('proposed_solution', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
            <div class="res-card">
                <h4>⏳ 4. Cheklovlar va Kamchiliklar</h4>
                <p style="color: #334155; font-size: 0.9rem; line-height: 1.55; margin: 0;">{summary.get('limitations', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

    # Tab 2: Full Translation
    with res_tab2:
        st.markdown(f"""
        <div class="res-card">
            <h4>🇺🇿 Hech Qayeri Qisqartirilmagan To‘liq Akademik Tarjima</h4>
            <div style="color: #1e293b; font-size: 0.94rem; line-height: 1.7; white-space: pre-wrap;">{full_trans}</div>
        </div>
        """, unsafe_allow_html=True)

    # Tab 3: Master Thesis Advisor
    with res_tab3:
        st.markdown(f"""
        <div class="res-card">
            <h4>📌 Qayerda va qanday iqtibos (citation) olish kerak?</h4>
            <p style="color: #334155; font-size: 0.92rem; line-height: 1.6; margin: 0;">{thesis.get('where_to_cite', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="res-card">
            <h4>🔬 Metodologiyani o‘z tadqiqotingizda qanday qo‘llash mumkin?</h4>
            <p style="color: #334155; font-size: 0.92rem; line-height: 1.6; margin: 0;">{thesis.get('how_to_use_method', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="res-card" style="background: #fffbeb; border-color: #fde68a; border-left: 5px solid #f59e0b;">
            <h4 style="color: #92400e;">💡 Ushbu maqola asosida yangi dissertatsiya / maqola g‘oyasi:</h4>
            <p style="color: #78350f; font-size: 0.92rem; line-height: 1.6; font-weight: 500; margin: 0;">{thesis.get('new_research_idea', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

    # Tab 4: Terminology Glossary
    with res_tab4:
        if terms:
            for item in terms:
                st.markdown(f"""
                <div class="res-card" style="padding: 14px 18px; margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                        <span style="font-weight: 800; color: #0284c7; font-size: 0.95rem;">{item.get('term_en', '')}</span>
                        <span style="font-weight: 700; color: #0f172a; font-size: 0.84rem; background: #e0f2fe; padding: 2px 8px; border-radius: 6px;">🇺🇿 {item.get('term_uz', '')}</span>
                    </div>
                    <p style="margin: 0; font-size: 0.86rem; color: #475569; line-height: 1.45;">{item.get('desc', '')}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Terminlar ro‘yxati mavjud emas.")

# Footer
st.markdown("""
<div class="footer-text">
    © 2026 UZ SCIENCE AI — Magistrant va ilmiy izlanuvchilar uchun universal platforma.
</div>
""", unsafe_allow_html=True)
