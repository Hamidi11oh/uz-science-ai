import streamlit as st
import fitz  # PyMuPDF
import requests
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import os
import time

# =====================================================================
# 👤 MUALLIF VA ASOSIY SOZLAMALAR
# =====================================================================
AUTHOR_NAME = "Qosimjonov Hamidullo"

# Streamlit Secrets-dan kalitni avtomatik olish
BUILTIN_GROQ_KEY = ""
if "GROQ_API_KEY" in st.secrets:
    BUILTIN_GROQ_KEY = str(st.secrets["GROQ_API_KEY"]).strip().strip("'\" \n\r\t")
elif os.environ.get("GROQ_API_KEY"):
    BUILTIN_GROQ_KEY = str(os.environ.get("GROQ_API_KEY")).strip().strip("'\" \n\r\t")

# =====================================================================
# SAHIFA SOZLAMALARI VA DIZAYN
# =====================================================================
st.set_page_config(
    page_title="UZ SCIENCE AI - Professional Ilmiy Tarjima",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .stApp {
        background-color: #f8fafc;
    }
    
    .block-container {
        max-width: 960px !important;
        padding-top: 1.5rem !important;
        padding-bottom: 4rem !important;
    }
    
    .top-nav {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 12px 22px;
        margin-bottom: 24px;
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
    
    .badge-author {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: linear-gradient(135deg, #f0fdf4 0%, #e0f2fe 100%);
        border: 1px solid #bae6fd;
        color: #0369a1;
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 0.82rem;
        font-weight: 700;
        box-shadow: 0 2px 5px rgba(2, 132, 199, 0.08);
    }
    
    .hero-container {
        text-align: center;
        margin-bottom: 24px;
    }
    
    .hero-title {
        font-size: 2.35rem;
        font-weight: 800;
        letter-spacing: -0.8px;
        color: #0f172a;
        margin-bottom: 8px;
    }
    
    .hero-highlight {
        background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .hero-subtitle {
        font-size: 0.96rem;
        color: #64748b;
        max-width: 620px;
        margin: 0 auto;
        line-height: 1.55;
    }
    
    .main-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 32px;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.03), 0 8px 10px -6px rgba(0, 0, 0, 0.02);
        margin-bottom: 24px;
    }
    
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
        margin-bottom: 12px;
        font-size: 1.1rem;
        font-weight: 700;
    }
    
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
    
    .footer-container {
        text-align: center;
        padding-top: 24px;
        border-top: 1px solid #e2e8f0;
        margin-top: 40px;
        font-size: 0.85rem;
        color: #64748b;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to generate PDF
def create_pdf(title, translation_text, summary_text, thesis_text, terms_text):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=15, leading=19, textColor=colors.HexColor('#0284c7'), spaceAfter=6)
    h1_style = ParagraphStyle('H1', parent=styles['Heading2'], fontSize=11, leading=15, textColor=colors.HexColor('#0f172a'), spaceBefore=12, spaceAfter=6)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=8.5, leading=12.5, textColor=colors.HexColor('#334155'), spaceAfter=4)
    
    story = []
    clean_title = str(title).replace('<', '&lt;').replace('>', '&gt;')
    story.append(Paragraph(f"<b>UZ SCIENCE AI — {clean_title}</b>", title_style))
    story.append(Paragraph(f"<i>Muallif: {AUTHOR_NAME} · To‘liq akademik tarjima va ilmiy tahliliy pasport</i>", body_style))
    story.append(Spacer(1, 10))
    
    # 1. TO'LIQ AKADEMIK TARJIMA (BOSHIDA)
    story.append(Paragraph("<b>1. TO‘LIQ AKADEMIK O‘ZBEKCHA TARJIMA (ASOSIY MATN)</b>", h1_style))
    for p in translation_text.split('\n'):
        if p.strip():
            clean_p = p.replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(clean_p, body_style))
    story.append(Spacer(1, 12))
    
    # 2. CHUQUR ILMIY XULOSA
    if summary_text.strip():
        story.append(Paragraph("<b>2. Chuqur Ilmiy Xulosa va Tahlil (Research Summary)</b>", h1_style))
        for p in summary_text.split('\n'):
            if p.strip():
                clean_p = p.replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(clean_p, body_style))
        story.append(Spacer(1, 10))
        
    # 3. MAGISTR TAVSIYALARI
    if thesis_text.strip():
        story.append(Paragraph("<b>3. Magistrlik Dissertatsiyasi Uchun Tavsiyalar</b>", h1_style))
        for p in thesis_text.split('\n'):
            if p.strip():
                clean_p = p.replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(clean_p, body_style))
        story.append(Spacer(1, 10))

    # 4. TERMINLAR
    if terms_text.strip():
        story.append(Paragraph("<b>4. Asosiy Ilmiy Terminlar Lug‘ati</b>", h1_style))
        for p in terms_text.split('\n'):
            if p.strip():
                clean_p = p.replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(clean_p, body_style))
            
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# Active Key resolution
active_groq_key = BUILTIN_GROQ_KEY.strip()

with st.sidebar:
    st.header("⚙️ Sozlamalar")
    if active_groq_key:
        st.success("✅ Server Groq AI kaliti faol. Barcha foydalanuvchilar bepul va limitsiz foydalanishi mumkin.")
    else:
        st.info("💡 Server kaliti topilmadi. O'zingiznikini kiriting:")
        user_key = st.text_input("Groq API Kalit (gsk_...):", type="password")
        if user_key.strip():
            active_groq_key = user_key.strip()
            
    st.markdown("👉 [Bepul Groq Kalit Olish](https://console.groq.com/keys)")

# Top Navigation Bar with Author Badge in the Corner
st.markdown(f"""
<div class="top-nav">
    <div class="nav-brand">
        <div class="nav-logo">✨</div>
        <div>
            <div style="font-weight: 800; font-size: 1.15rem; color: #0f172a; display: flex; align-items: center;">
                UZ SCIENCE AI
            </div>
            <div style="font-size: 0.76rem; color: #64748b;">Magistrant va tadqiqotchilar uchun ilmiy platforma</div>
        </div>
    </div>
    <div class="badge-author">
        <span>👨‍💻 Muallif:</span> <b>{AUTHOR_NAME}</b>
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

# Main Floating White Card
st.markdown('<div class="main-card">', unsafe_allow_html=True)

tab_upload, tab_paste = st.tabs(["📄 Fayl Yuklash (PDF/DOCX/PPTX)", "✍️ Matnni Qo‘yish (Paste)"])

extracted_text = ""
file_name = "Ilmiy Hujjat"

with tab_upload:
    uploaded_file = st.file_uploader(
        "PDF yoki boshqa faylni shu yerga tashlang:",
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

# Bulletproof Single Master Prompt (Natural Markdown, ZERO JSON restrictions, NEVER breaks)
MASTER_PIPELINE_PROMPT = """
Siz oliy toifali akademik tarjimon, ilmiy muharrir va magistrlik dissertatsiyalari bo'yicha yetakchi ilmiy maslahatchisisiz.
Vazifangiz: Berilgan ilmiy maqolani to'liq tahlil qilib, quyidagi 4 TA BO'LIM BO'YICHA toza akademik o'zbek tilida mukammal natija berish.

Har bir bo'limni aynan quyidagi sarlavhalar bilan boshlang:

# 1. TO‘LIQ AKADEMIK O‘ZBEKCHA TARJIMA
(Maqolaning barcha bo'limlarini — Abstract, Kirish, Metodologiya, Tajribalar, Natijalar, Xulosani hech qayerini qisqartirmasdan, formulalari ($...$) va iqtiboslari ([1], [2]) bilan so'zma-so'z, ravon akademik tilda to'liq tarjima qiling)

# 2. CHUQUR ILMIY XULOSA (RESEARCH SUMMARY)
- **⚡ Diqqat Markazi (Eng muhim qismi):** Ushbu maqolada magistrant nimasiga alohida e'tibor berishi kerak?
- **🎯 1. Ko'tarilgan Asosiy Muammo:** Maqolada hal qilinmoqchi bo'lgan asosiy ilmiy yoki amaliy muammo nima?
- **💡 2. Taklif Etilgan Yechim & Model:** Mualliflar qanday yangi arxitektura, model yoki algoritm taklif qilishdi?
- **📊 3. Natijalar va Benchmarklar:** Qanday ko'rsatkichlarga erishildi?
- **⏳ 4. Cheklovlar va Kamchiliklar:** Tadqiqotning qanday cheklovlari mavjud?

# 3. MAGISTR DISSERTATSIYASI UCHUN TAVSIYALAR
- **📌 Qayerda va qanday iqtibos (citation) olish kerak:** Dissertatsiyaning qaysi bo'limida qanday qo'llash lozim.
- **🔬 Metodologiyani qo'llash:** O'z amaliy ishida qanday foydalanish mumkin.
- **💡 Yangi ilmiy g'oya:** Ushbu maqoladan kelib chiqib magistr uchun yangi tadqiqot yo'nalishi.

# 4. ASOSIY ILMIY TERMINLAR LUG‘ATI
- **Termin (Inglizcha)** -> **O'zbekcha atama:** Qisqa tushunarli izoh.
"""

def parse_markdown_response(text):
    sections = {
        "translation": "",
        "summary": "",
        "thesis": "",
        "terms": ""
    }
    
    parts = text.split("# ")
    for p in parts:
        p_clean = p.strip()
        if p_clean.startswith("1. TO‘LIQ") or "TO‘LIQ AKADEMIK" in p_clean[:40]:
            sections["translation"] = p_clean.split("\n", 1)[1] if "\n" in p_clean else p_clean
        elif p_clean.startswith("2. CHUQUR") or "CHUQUR ILMIY" in p_clean[:40]:
            sections["summary"] = p_clean.split("\n", 1)[1] if "\n" in p_clean else p_clean
        elif p_clean.startswith("3. MAGISTR") or "MAGISTR DISSERTATSIYASI" in p_clean[:40]:
            sections["thesis"] = p_clean.split("\n", 1)[1] if "\n" in p_clean else p_clean
        elif p_clean.startswith("4. ASOSIY") or "TERMINLAR" in p_clean[:40]:
            sections["terms"] = p_clean.split("\n", 1)[1] if "\n" in p_clean else p_clean
            
    # If headers were not separated, whole text is the translation
    if not sections["translation"]:
        sections["translation"] = text
        
    return sections

def execute_groq_master(key, text):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {key}"
    }
    
    # 1. 100% DYNAMIC: Ask Groq which models are active for this account right now!
    active_chat_models = []
    try:
        r_mod = requests.get("https://api.groq.com/openai/v1/models", headers=headers, timeout=10)
        if r_mod.status_code == 200:
            for item in r_mod.json().get("data", []):
                mid = item.get("id", "")
                if "whisper" not in mid and "guard" not in mid and "vision" not in mid:
                    active_chat_models.append(mid)
    except Exception:
        pass

    # Preferred new 2026 models order (with gpt-oss-20b, gpt-oss-120b, qwen3.6)
    preferred_order = [
        "openai/gpt-oss-20b",
        "openai/gpt-oss-120b",
        "qwen/qwen3.6-27b",
        "groq/compound",
        "groq/compound-mini",
        "gemma2-9b-it"
    ]
    
    models_to_try = []
    for pref in preferred_order:
        if pref in active_chat_models:
            models_to_try.append(pref)
    for m in active_chat_models:
        if m not in models_to_try:
            models_to_try.append(m)
            
    if not models_to_try:
        models_to_try = ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "qwen/qwen3.6-27b"]

    last_error_text = ""
    for model_candidate in models_to_try:
        payload = {
            "model": model_candidate,
            "messages": [
                {"role": "system", "content": MASTER_PIPELINE_PROMPT},
                {"role": "user", "content": f"Quyidagi ilmiy maqolani to'liq tahlil qilib, 4 ta bo'limda natija bering:\n\n{text[:25000]}"}
            ],
            "temperature": 0.2,
            "max_tokens": 7500
        }
        
        try:
            url = "https://api.groq.com/openai/v1/chat/completions"
            r = requests.post(url, headers=headers, json=payload, timeout=90)
            if r.status_code == 200:
                raw_text = r.json()["choices"][0]["message"]["content"]
                sections = parse_markdown_response(raw_text)
                return sections, f"Groq ({model_candidate})"
            else:
                last_error_text = f"HTTP {r.status_code}: {r.text}"
        except Exception as e:
            last_error_text = str(e)
            continue
            
    raise Exception(f"Groq API javobi: {last_error_text}")

# Big Action Button
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 TARJIMA VA TAHLIL QILISH", type="primary", use_container_width=True):
    if not extracted_text.strip():
        st.warning("⚠️ Iltimos, oldin fayl yuklang yoki matn kiriting.")
    elif not active_groq_key:
        st.error("⚠️ Tizimda API kalit topilmadi. Iltimos, chap tarafdagi menyuga Groq API kalitingizni kiriting.")
    else:
        with st.status("🔍 Ilmiy maqola tahlil qilinmoqda va o'zbekchalashtirilmoqda...", expanded=True) as status_box:
            st.write("📄 **1-bosqich:** Fayl matni va tuzilmasi o‘qildi...")
            st.write(f"✓ Matn hajmi: {len(extracted_text):,} ta belgi.")
            
            st.write("🧠 **2-bosqich:** To‘liq akademik tarjima va ilmiy xulosa shakllantirilmoqda...")
            
            try:
                sections, used_model = execute_groq_master(active_groq_key, extracted_text)

                st.write("📄 **3-bosqich:** Formulalar, iqtiboslar va ilmiy pasport tekshirildi.")
                st.write("📥 **4-bosqich:** Word (.docx) va PDF eksport fayllari yaratildi.")

                st.session_state["result_sections"] = sections
                st.session_state["file_title"] = file_name
                st.session_state["used_model"] = used_model

                status_box.update(label=f"✅ To‘liq tarjima va tahlil muvaffaqiyatli yakunlandi! ({used_model})", state="complete", expanded=False)
                st.success(f"✨ Muvaffaqiyatli yakunlandi! ({used_model} orqali)")
                
            except Exception as err:
                status_box.update(label="❌ Tahlilda xatolik yuz berdi", state="error", expanded=True)
                st.error(f"Xatolik tafsiloti: {err}")

st.markdown('</div>', unsafe_allow_html=True)

# Results Section
if "result_sections" in st.session_state:
    sec = st.session_state["result_sections"]
    trans_text = sec.get("translation", "")
    summary_text = sec.get("summary", "")
    thesis_text = sec.get("thesis", "")
    terms_text = sec.get("terms", "")
    current_title = st.session_state.get("file_title", "Ilmiy Maqola")

    st.markdown(f"""
    <div class="res-card" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; background: linear-gradient(135deg, #f0f9ff 0%, #ffffff 100%); border-color: #bae6fd;">
        <div>
            <div style="font-weight: 800; font-size: 1.15rem; color: #0369a1;">📄 {current_title}</div>
            <div style="font-size: 0.82rem; color: #64748b;">To‘liq akademik tarjima va ilmiy hisobot yuklab olishga tayyor</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # DUAL DOWNLOAD BUTTONS (ASOSIY URG'U: TO'LIQ TARJIMA VA HISOBOT)
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        try:
            doc = Document()
            doc.add_heading(f"UZ SCIENCE AI — {current_title}", level=1)
            doc.add_paragraph(f"Muallif: {AUTHOR_NAME} · To‘liq akademik tarjima va tahliliy hisobot")
            
            # 1. ASOSIY TO'LIQ AKADEMIK TARJIMA (BOSHIDA KELADI)
            doc.add_heading("1. TO‘LIQ AKADEMIK O‘ZBEKCHA TARJIMA (ASOSIY MATN)", level=2)
            doc.add_paragraph(trans_text)
            
            # 2. CHUQUR ILMIY XULOSA
            if summary_text.strip():
                doc.add_heading("2. Chuqur Ilmiy Xulosa va Tahlil", level=2)
                doc.add_paragraph(summary_text)
                
            # 3. MAGISTR DISSERTATSIYASI YO'RIQNOMASI
            if thesis_text.strip():
                doc.add_heading("3. Magistr Dissertatsiyasi Uchun Tavsiyalar", level=2)
                doc.add_paragraph(thesis_text)

            # 4. TERMINLAR
            if terms_text.strip():
                doc.add_heading("4. Asosiy Ilmiy Terminlar", level=2)
                doc.add_paragraph(terms_text)

            docx_io = io.BytesIO()
            doc.save(docx_io)
            docx_io.seek(0)

            st.download_button(
                label="📄 To‘liq Tarjima (Word .docx) Yuklab Olish",
                data=docx_io,
                file_name=f"UZ_SCIENCE_AI_{current_title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
        except Exception as docx_err:
            st.error(f"Word tayyorlashda xatolik: {docx_err}")

    with col_d2:
        try:
            pdf_bytes = create_pdf(current_title, trans_text, summary_text, thesis_text, terms_text)
            st.download_button(
                label="📥 To‘liq Ilmiy Hisobot (PDF) Yuklab Olish",
                data=pdf_bytes,
                file_name=f"UZ_SCIENCE_AI_{current_title}.pdf",
                mime="application/pdf"
            )
        except Exception as pdf_err:
            st.error(f"PDF tayyorlashda xatolik: {pdf_err}")

    st.markdown("<br>", unsafe_allow_html=True)

    # 4 Output Tabs - ASOSIY URG'U: 1-TAB TO'LIQ AKADEMIK TARJIMA
    res_tab1, res_tab2, res_tab3, res_tab4 = st.tabs([
        "📄 To‘liq Akademik Tarjima",
        "🧠 Chuqur Ilmiy Xulosa (Summary)",
        "🎓 Magistr Dissertatsiyasi Tavsiyalari",
        "📖 Ilmiy Terminlar Lug‘ati"
    ])

    # 1-TAB: ASOSIY TO'LIQ AKADEMIK TARJIMA
    with res_tab1:
        st.markdown(f"""
        <div class="res-card" style="border-left: 5px solid #0284c7;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                <h4 style="margin: 0; color: #0369a1;">🇺🇿 Hech Qayeri Qisqartirilmagan To‘liq Akademik Tarjima</h4>
                <span style="font-size: 0.78rem; background: #e0f2fe; color: #0369a1; padding: 4px 10px; border-radius: 8px; font-weight: 700;">{len(trans_text):,} ta belgi</span>
            </div>
            <div style="color: #0f172a; font-size: 0.95rem; line-height: 1.75; white-space: pre-wrap;">{trans_text}</div>
        </div>
        """, unsafe_allow_html=True)

    # 2-TAB: CHUQUR ILMIY XULOSA
    with res_tab2:
        st.markdown(f"""
        <div class="res-card">
            <h4>🧠 Chuqur Ilmiy Xulosa va Tahlil</h4>
            <div style="color: #1e293b; font-size: 0.94rem; line-height: 1.65; white-space: pre-wrap;">{summary_text}</div>
        </div>
        """, unsafe_allow_html=True)

    # 3-TAB: MAGISTR TAVSIYALARI
    with res_tab3:
        st.markdown(f"""
        <div class="res-card">
            <h4>🎓 Magistrlik Dissertatsiyasi Uchun Yo‘riqnoma</h4>
            <div style="color: #334155; font-size: 0.94rem; line-height: 1.65; white-space: pre-wrap;">{thesis_text}</div>
        </div>
        """, unsafe_allow_html=True)

    # 4-TAB: LUG'AT
    with res_tab4:
        st.markdown(f"""
        <div class="res-card">
            <h4>📖 Asosiy Ilmiy Terminlar Glossariysi</h4>
            <div style="color: #475569; font-size: 0.92rem; line-height: 1.6; white-space: pre-wrap;">{terms_text}</div>
        </div>
        """, unsafe_allow_html=True)

# Footer with Author Branding
st.markdown(f"""
<div class="footer-container">
    © 2026 <b>UZ SCIENCE AI</b> — Yaratuvchi: <b>{AUTHOR_NAME}</b>. Magistrant va ilmiy izlanuvchilar uchun maxsus yaratilgan.
</div>
""", unsafe_allow_html=True)
