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
    page_title="UZ SCIENCE AI | Akademik Tarjimon & Tahlil",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom Ultra-Modern SaaS CSS (Linear / Vercel style)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* Main Background & Clean Spacing */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 1200px !important;
    }
    
    /* Sleek Top Navbar Banner */
    .hero-banner {
        background: radial-gradient(100% 100% at 50% 0%, rgba(2, 132, 199, 0.15) 0%, rgba(248, 250, 252, 0) 100%), #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 32px 28px;
        text-align: center;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.03);
    }
    
    .hero-title {
        font-size: 2.3rem;
        font-weight: 800;
        letter-spacing: -0.8px;
        background: linear-gradient(135deg, #0f172a 0%, #0284c7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 8px;
    }
    
    .hero-sub {
        font-size: 1rem;
        color: #64748b;
        max-width: 650px;
        margin: 0 auto;
        line-height: 1.5;
    }
    
    .status-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        color: #166534;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-bottom: 12px;
    }
    
    /* Premium Modern Card */
    .saas-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02), 0 6px 16px -4px rgba(0,0,0,0.04);
        transition: all 0.2s ease;
    }
    
    .saas-card:hover {
        border-color: #cbd5e1;
        box-shadow: 0 10px 25px -5px rgba(0,0,0,0.06);
    }
    
    .card-heading {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .highlight-box {
        background: #f8fafc;
        border-left: 4px solid #0284c7;
        padding: 16px;
        border-radius: 0 12px 12px 0;
        margin-bottom: 14px;
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
        font-size: 0.9rem;
        color: #64748b;
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
    }

    .stTabs [aria-selected="true"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border-color: #0284c7 !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.25);
    }
    
    /* Big Action Button */
    .stButton button {
        background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%) !important;
        color: white !important;
        border: none !important;
        padding: 14px 28px !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        letter-spacing: 0.2px !important;
        box-shadow: 0 8px 20px -4px rgba(2, 132, 199, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    
    .stButton button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 12px 25px -4px rgba(2, 132, 199, 0.45) !important;
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
    story.append(Paragraph("<b>1. Chuqur Ilmiy Xulosa va Tahlil</b>", h1_style))
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

# Automatic Master API Key Resolution (No user input required if configured in Secrets or Environment)
MASTER_KEY = ""
if "GEMINI_API_KEY" in st.secrets:
    MASTER_KEY = st.secrets["GEMINI_API_KEY"]
elif os.environ.get("GEMINI_API_KEY"):
    MASTER_KEY = os.environ.get("GEMINI_API_KEY")

# Hero Banner
status_text = "🟢 Tizim tayyor • Gemini 3.6 Flash faol" if MASTER_KEY else "⚪ Shaxsiy API kalit kiritish mumkin"
st.markdown(f"""
<div class="hero-banner">
    <div class="status-pill">{status_text}</div>
    <div class="hero-title">UZ SCIENCE AI</div>
    <div class="hero-sub">Xorijiy ilmiy maqolalar, dissertatsiyalar va taqdimotlarni mazmunini buzmasdan akademik o‘zbek tiliga o‘giruvchi hamda dissertatsiya uchun chuqur tahlil qiluvchi aqlli tizim.</div>
</div>
""", unsafe_allow_html=True)

# Sidebar (Only for optional personal API key or custom overrides)
with st.sidebar:
    st.header("⚙️ Sozlamalar")
    if MASTER_KEY:
        st.success("✅ Umumiy server API kaliti faol. Saytdan barcha bepul foydalanishi mumkin.")
        api_key = MASTER_KEY
    else:
        st.info("💡 Agar sayt egasi umumiy kalit kiritmagan bo'lsa, o'z kalitingizni kiriting:")
        api_key = st.text_input(
            "Google Gemini API Kalit:",
            type="password",
            help="aistudio.google.com/app/apikey sahifasidan olingan kalit"
        )
        st.markdown("👉 [Bepul API Kalit Olish](https://aistudio.google.com/app/apikey)")

    selected_model = "gemini-3.6-flash"
    st.divider()
    st.markdown("### 📋 Platforma Imkoniyatlari")
    st.markdown("✔️ **Formatlar:** PDF, PPTX (taqdimot), DOCX, TXT")
    st.markdown("✔️ **Formatlash:** Formulalar, LaTeX va iqtiboslar saqlanadi")
    st.markdown("✔️ **Magistr Pasporti:** SOTA natijalar, yangilik, thesis tavsiyalari")
    st.markdown("✔️ **Eksport:** Ham **PDF**, ham **Word (.docx)** yuklab olish")

# Main Input Section
tab_upload, tab_paste = st.tabs(["📁 Fayl Yuklash (PDF • PPTX • DOCX)", "✍️ Matnni Joylashtirish (Paste)"])

extracted_text = ""
file_name = "Ilmiy Hujjat"

with tab_upload:
    uploaded_file = st.file_uploader(
        "Ilmiy maqola, dissertatsiya yoki taqdimot faylini tanlang:",
        type=["pdf", "docx", "pptx", "txt"],
        help="PDF, Word yoki PowerPoint fayllarni to'g'ridan-to'g'ri yuklashingiz mumkin"
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
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; padding: 12px 18px; border-radius: 12px; color: #166534; font-size: 0.92rem; font-weight: 600; display: flex; align-items: center; justify-content: space-between;">
                    <span>📄 <b>{file_name}</b> muvaffaqiyatli o‘qildi</span>
                    <span style="background: #dcfce7; padding: 2px 8px; border-radius: 6px; font-size: 0.8rem;">{len(extracted_text):,} belgi</span>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ Fayldan matn ajratib bo'lmadi. Agar bu skaner qilingan rasm bo'lsa, 'Matnni Joylashtirish' bo'limidan foydalaning.")
        except Exception as e:
            st.error(f"Faylni o'qishda xatolik: {e}")

with tab_paste:
    direct_text = st.text_area("Maqola matnini to'g'ridan-to'g'ri shu yerga qo'ying:", height=180, placeholder="Abstract, Introduction yoki butun maqola matnini Ctrl+V qilib tashlang...")
    if direct_text.strip():
        extracted_text = direct_text
        file_name = "Kiritilgan Maqola Matni"

# Big Action Button
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 MUKAMMAL TARJIMA VA CHUQUR TAHLIL QILISH", type="primary", use_container_width=True):
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

# Results Dashboard
if "result_data" in st.session_state:
    data = st.session_state["result_data"]
    summary = data.get("research_summary", {})
    thesis = data.get("thesis_advisor", {})
    terms = data.get("key_terms", [])
    full_trans = data.get("full_translation", "")
    current_title = st.session_state.get("file_title", "Ilmiy Maqola")

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Download Card Section
    st.markdown(f"""
    <div class="saas-card" style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border: 1px solid #cbd5e1;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <div>
                <h3 style="margin: 0; color: #0f172a; font-size: 1.25rem; font-weight: 800;">📄 {current_title}</h3>
                <p style="margin: 4px 0 0 0; color: #64748b; font-size: 0.88rem;">Akademik tahlil va tarjima hujjati to‘liq shakllantirildi</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

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
        # Key Focus Areas Hero Card
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); border: 1px solid #bfdbfe; border-left: 6px solid #0284c7; padding: 20px 24px; border-radius: 14px; margin-bottom: 24px;">
            <div style="font-size: 0.8rem; font-weight: 800; text-transform: uppercase; color: #0284c7; letter-spacing: 0.5px; margin-bottom: 6px;">⚡ DIQQAT MARKAZI • ENG MUHIM QISM</div>
            <h4 style="color: #0c4a6e; margin: 0 0 8px 0; font-size: 1.15rem; font-weight: 800;">Ushbu maqolada nimasiga alohida e’tibor berish kerak?</h4>
            <p style="color: #1e293b; font-size: 0.96rem; margin: 0; line-height: 1.65;">{summary.get('key_focus_areas', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            st.markdown(f"""
            <div class="saas-card">
                <div class="card-heading">🎯 1. Ko‘tarilgan Asosiy Muammo</div>
                <p style="color: #334155; font-size: 0.93rem; line-height: 1.6; margin: 0;">{summary.get('core_problem', '-')}</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="saas-card">
                <div class="card-heading">📊 3. Natijalar va Benchmarklar</div>
                <p style="color: #334155; font-size: 0.93rem; line-height: 1.6; margin: 0;">{summary.get('experimental_results', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

        with col_s2:
            st.markdown(f"""
            <div class="saas-card">
                <div class="card-heading">💡 2. Taklif Etilgan Yechim & Metod</div>
                <p style="color: #334155; font-size: 0.93rem; line-height: 1.6; margin: 0;">{summary.get('proposed_solution', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
            <div class="saas-card">
                <div class="card-heading">⏳ 4. Cheklovlar va Kamchiliklar</div>
                <p style="color: #334155; font-size: 0.93rem; line-height: 1.6; margin: 0;">{summary.get('limitations', '-')}</p>
            </div>
            """, unsafe_allow_html=True)

    # Tab 2: Full Translation
    with res_tab2:
        st.markdown(f"""
        <div class="saas-card">
            <div class="card-heading">🇺🇿 Hech Qayeri Qisqartirilmagan To‘liq Akademik Tarjima</div>
            <div style="color: #1e293b; font-size: 0.95rem; line-height: 1.7; white-space: pre-wrap;">{full_trans}</div>
        </div>
        """, unsafe_allow_html=True)

    # Tab 3: Master Thesis Advisor
    with res_tab3:
        st.markdown(f"""
        <div class="saas-card">
            <div class="card-heading">📌 Qayerda va qanday iqtibos (citation) olish kerak?</div>
            <p style="color: #334155; font-size: 0.95rem; line-height: 1.65; margin: 0;">{thesis.get('where_to_cite', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="saas-card">
            <div class="card-heading">🔬 Metodologiyani o‘z tadqiqotingizda qanday qo‘llash mumkin?</div>
            <p style="color: #334155; font-size: 0.95rem; line-height: 1.65; margin: 0;">{thesis.get('how_to_use_method', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="saas-card" style="background: #fffbeb; border: 1px solid #fef3c7; border-left: 5px solid #f59e0b;">
            <div class="card-heading" style="color: #92400e;">💡 Ushbu maqola asosida yangi dissertatsiya / maqola g‘oyasi:</div>
            <p style="color: #78350f; font-size: 0.95rem; line-height: 1.65; font-weight: 500; margin: 0;">{thesis.get('new_research_idea', '-')}</p>
        </div>
        """, unsafe_allow_html=True)

    # Tab 4: Terminology Glossary
    with res_tab4:
        if terms:
            for item in terms:
                st.markdown(f"""
                <div class="saas-card" style="padding: 16px 20px; margin-bottom: 12px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                        <span style="font-weight: 800; color: #0284c7; font-size: 1rem;">{item.get('term_en', '')}</span>
                        <span style="font-weight: 700; color: #0f172a; font-size: 0.88rem; background: #e0f2fe; padding: 3px 10px; border-radius: 8px;">🇺🇿 {item.get('term_uz', '')}</span>
                    </div>
                    <p style="margin: 0; font-size: 0.9rem; color: #475569; line-height: 1.5;">{item.get('desc', '')}</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Terminlar ro‘yxati mavjud emas.")
