import streamlit as st
import pandas as pd
from datetime import datetime
import json
import io
import os
import time
import re
import base64
import docx
import pdfplumber
import google.generativeai as genai
from PIL import Image

# --- Хуудасны тохиргоо ---
st.set_page_config(page_title="Өр нэхэмжлэх удирдлага", layout="wide", page_icon="⚖️")

# Uploads фолдер үүсгэх
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# --- Орчин үеийн өнгө үзэмж (CSS) + Утасны (Mobile) тохиргоо ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    h1, h2, h3 { color: #1f3a5f; }
    .stTabs [data-baseweb="tab-list"] { gap: 5px; background-color: #ffffff; padding: 5px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] { padding: 8px 12px; border-radius: 8px; background-color: #f0f2f6; color: #555; font-weight: bold; font-size: 14px; }
    .stTabs [aria-selected="true"] { background-color: #1f3a5f !important; color: white !important; }
    .stButton>button { background: linear-gradient(90deg, #1f3a5f 0%, #2d5985 100%); color: white; border: none; border-radius: 8px; padding: 10px 24px; font-weight: bold; width: 100%; }
    .stButton>button:hover { background: linear-gradient(90deg, #2d5985 0%, #1f3a5f 100%); color: white; }
    
    .metric-card { background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%); padding: 15px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); text-align: center; border-top: 4px solid #1f3a5f; margin-bottom: 15px; }
    .metric-num { font-size: 28px; font-weight: 800; color: #1f3a5f; }
    .metric-label { font-size: 13px; color: #666; font-weight: 500; margin-top: 5px; }
    
    .client-card { background: white; padding: 15px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); margin-bottom: 20px; border-left: 6px solid #1f3a5f; }
    .client-name { font-size: 18px; font-weight: bold; color: #1f3a5f; margin-bottom: 10px; }
    .status-badge { padding: 5px 10px; border-radius: 20px; font-size: 11px; font-weight: bold; color: white; float: right; }
    .info-row { font-size: 14px; color: #555; margin-top: 8px; }
    .info-label { font-weight: bold; color: #333; }
    .note-box { background: #f8f9fa; padding: 10px; border-radius: 8px; margin-top: 10px; font-size: 13px; color: #666; border: 1px solid #e9ecef; }
    .img-container { margin-top: 15px; border-radius: 8px; overflow: hidden; border: 1px solid #e9ecef; }
    .img-container img { width: 100%; display: block; }

    /* Утсанд зориулсан (Mobile Responsive) тохиргоо */
    @media (max-width: 768px) {
        .metric-num { font-size: 22px; }
        .metric-label { font-size: 11px; }
        .client-name { font-size: 16px; }
        .stTabs [data-baseweb="tab"] { font-size: 12px; padding: 8px 6px; width: 100%; text-align: center; }
        [data-testid="stHorizontalBlock"] { flex-direction: column !important; }
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚖️ Шүүх нэхэмжлэх болон Эвлэрүүлэн зуучлалын систем")
st.markdown("##### Google Gemini AI дэмжлэгтэй веб апп")

STATUS_OPTIONS = [
    "Шүүхэд өгсөн", 
    "Эвлэрүүлэн зуучлалд өгсөн", 
    "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа", 
    "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон",
    "Захирамж гарсан", 
    "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн", 
    "Гүйцэтгэх хуудас гарсан шүүхийн шийдвэрт шилжүүлсэн", 
    "Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа", 
    "Өр төлбөр дууссан"
]

STATUS_COLORS = {
    "Шүүхэд өгсөн": "#1f3a5f",
    "Эвлэрүүлэн зуучлалд өгсөн": "#8e44ad",
    "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа": "#2980b9",
    "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон": "#7f8c8d",
    "Захирамж гарсан": "#f39c12",
    "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн": "#d35400",
    "Гүйцэтгэх хуудас гарсан шүүхийн шийдвэрт шилжүүлсэн": "#c0392b",
    "Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа": "#e74c3c",
    "Өр төлбөр дууссан": "#27ae60"
}

DATA_FILE = "court_data.csv"
API_KEY_FILE = "api_key.txt"
required_cols = ["№", "Зээлдэгч", "Хариуцсан ажилтан", "Шүүхэд өгсөн огноо", "Эвлэрүүлэнд өгсөн огноо", "Захирамж гарсан огноо", "Одоогийн төлөв", "Тэмдэглэл", "Файлын нэр"]

if 'df_court' not in st.session_state:
    if os.path.exists(DATA_FILE):
        try:
            st.session_state.df_court = pd.read_csv(DATA_FILE)
            for col in required_cols:
                if col not in st.session_state.df_court.columns:
                    st.session_state.df_court[col] = ""
        except Exception:
            st.session_state.df_court = pd.DataFrame(columns=required_cols)
    else:
        st.session_state.df_court = pd.DataFrame(columns=required_cols)

def save_data():
    st.session_state.df_court.to_csv(DATA_FILE, index=False)

# --- Sidebar ---
st.sidebar.header("⚙️ Тохиргоо")

if 'api_key' not in st.session_state:
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, 'r') as f:
            st.session_state.api_key = f.read().strip()
    else:
        st.session_state.api_key = ""

api_key = st.sidebar.text_input("Google Gemini API Key оруулна уу", value=st.session_state.api_key, type="password", help="aistudio.google.com сайтад бүртгүүлж үнэгүй key авна уу. Оруулсны дараа Enter дарна уу.")
if api_key != st.session_state.api_key:
    st.session_state.api_key = api_key
    if api_key:
        with open(API_KEY_FILE, 'w') as f:
            f.write(api_key)

st.sidebar.markdown("---")
st.sidebar.header("📁 Файл оруулах/Татах")
up_excel = st.sidebar.file_uploader("Excel файлаас өгөгдөл татах", type=['xlsx'])
if up_excel:
    try:
        df_import = pd.read_excel(up_excel, sheet_name="Шүүх нэхэмжлэл")
        st.session_state.df_court = df_import
        save_data()
        st.sidebar.success("Excel амжилттай уншигдлаа!")
    except Exception as e:
        st.sidebar.error(f"Алдаа: {e}")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Шүүх нэхэмжлэл')
        df_med = df[df["Одоогийн төлөв"].isin(["Эвлэрүүлэн зуучлалд өгсөн", "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа", "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон"])].copy()
        df_med.to_excel(writer, index=False, sheet_name='Эвлэрүүлэн зуучлал')
        dashboard_data = {
            "Үзүүлэлт": ["Нийт хэрэг", "Шүүхэд өгсөн", "Захирамж гарсан", "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн", "Гүйцэтгэлд явж байгаа", "Өр дууссан", "Эвлэрүүлэнд өгсөн", "Эвлэрүүлэх захирамж дагуу төлж байгаа", "Дуусгавар болсон"],
            "Тоо": [
                len(df), len(df[df["Одоогийн төлөв"] == "Шүүхэд өгсөн"]), len(df[df["Одоогийн төлөв"] == "Захирамж гарсан"]),
                len(df[df["Одоогийн төлөв"] == "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн"]), len(df[df["Одоогийн төлөв"] == "Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа"]),
                len(df[df["Одоогийн төлөв"] == "Өр төлбөр дууссан"]), len(df[df["Одоогийн төлөв"] == "Эвлэрүүлэн зуучлалд өгсөн"]), len(df[df["Одоогийн төлөв"] == "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа"]),
                len(df[df["Одоогийн төлөв"] == "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон"])
            ]
        }
        pd.DataFrame(dashboard_data).to_excel(writer, index=False, sheet_name='Dashboard')
    return output.getvalue()

if not st.session_state.df_court.empty:
    st.sidebar.download_button(label="📥 Excel-ээ татах", data=to_excel(st.session_state.df_court), file_name='Шүүх_нэхэмжлэл_бүртгэл.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.sidebar.markdown("---")
if st.sidebar.button("🗑️ Бүх бүртгэлийг устгах", use_container_width=True):
    st.session_state.df_court = pd.DataFrame(columns=required_cols)
    if os.path.exists(DATA_FILE):
        os.remove(DATA_FILE)
    st.rerun()

# --- AI унших функц (Хязгааргүй хүлээлттэй) ---
def extract_info_from_file(file_obj, key):
    if not key: 
        st.error("⚠️ Зүүн талын цэснээс Google Gemini API Key оруулна уу!")
        return None, None, None, None, None, None, None, None
    try:
        genai.configure(api_key=key)
        model = genai.GenerativeModel('gemini-flash-latest')
        
        file_text = ""
        prompt = """Энэхүү баримт бичгийн зураг эсвэл текстийг маш нарийнаар шинжилж, дараах мэдээллийг татаад зөвхөн JSON формат буцаа:
        1. "doc_type": Баримтын төрөл ("Шүүхийн нэхэмжлэл", "Эвлэрүүлэн зуучлалын өргөдөл", эсвэл "Гүйцэтгэх захирамж").
        2. "name": Өрнийн эзэн буюу ЗЭЭЛДЭГЧИЙН нэр (Овог Нэр). Хариуцагч, төлөөлөгчийн нэрийг бичиж болохгүй!
        3. "officer": Баримт бичиг дээр "Итгэмжлэгдсэн төлөөлөгч", "Хуульч", "Гүйцэтгэлийн ажилтан" гэх зэргээр бичигдсэн хүний нэр. Хэрэв олдоогүй бол "null" гэж бич.
        4. "status_hint": ЗУРАГ ДЭЭРХ АГУУЛГААР ЭНЭ ХЭРГИЙН ОДООГИЙН ТӨЛӨВ ЮУ БАЙГААГ НАРИЙН ШИНЖИЛЖ ТОГТОО. Зөвхөн эдгээрээс аль нэгийг сонго: 
        - "Шүүхэд өгсөн" (Хэрэв зураг дээр шүүхэд өгсөн тухай, эсвэл шүүхийн нэхэмжлэл гарсан тухай бичигдсэн бол)
        - "Эвлэрүүлэн зуучлалд өгсөн" (Хэрэв эвлэрүүлэн зуучлалд өгсөн бол)
        - "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон" (Хэрэв эвлэрүүлэх гэсэн ч ирээгүй, гэрээ байгуулаагүй дуусгавар болсон бол)
        - "Захирамж гарсан" (Хэрэв захирамж гарсан бол)
        - "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн" (Хэрэв "гүйцэтгэх хуудас олгож өгнө үү", "гүйцэтгэх захирамж гаргаж" гэх мэт үг байвал)
        Анхаар: Зөвхөн нэг үг байна гээд төлвийг нь буруу тогтоохгүй байх. Зургийн бүхэл бүтэн агуулгыг уншиж тогтоо.
        5. "court_date": Шүүхэд өгсөн эсвэл нэхэмжлэх гарсан огноо (YYYY-MM-DD форматад). Өөр огноо бүү бич.
        6. "mediation_date": Эвлэрүүлэнд өгсөн огноо (YYYY-MM-DD форматад, үгүй бол null).
        7. "order_date": Захирамж гарсан огноо (YYYY-MM-DD форматад, үгүй бол null).
        8. "summary": Баримт бичгийн гол агуулга, нэхэмжилсэн зүйл, шаардсан дүн зэрэгийг товч тодорхой 1-3 өгүүлбэрээр монгол хэл дээр бич. ХЭРЭВ ХЭРЭГ ДУУСГАВАР БОЛСОН БОЛ ЯМАР ШАЛТГААНААР ДУУСГАВАР БОЛСНЫГ ЗААВАЛ ЭНД БИЧНЭ ҮҮ.
        Зөвхөн JSON буцаа."""

        if file_obj.name.endswith('.docx'):
            doc = docx.Document(file_obj)
            file_text = "\n".join([para.text for para in doc.paragraphs])
            response = model.generate_content(prompt + "\n\nБаримтын текст:\n" + file_text)
        elif file_obj.name.endswith('.pdf'):
            with pdfplumber.open(file_obj) as pdf:
                for page in pdf.pages: file_text += page.extract_text() + "\n"
            response = model.generate_content(prompt + "\n\nБаримтын текст:\n" + file_text)
        elif file_obj.name.endswith(('.png', '.jpg', '.jpeg', '.heic', '.webp')):
            img = Image.open(file_obj)
            max_size = (1024, 1024)
            img.thumbnail(max_size)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            response = model.generate_content([prompt, img])
        else:
            return None, None, None, None, None, None, None, None

        result = response.text.strip()
        if result.startswith("```json"): result = result.replace("```json", "").replace("```", "").strip()
            
        json_match = re.search(r'\{.*\}', result, re.DOTALL)
        if json_match:
            result = json_match.group(0)
        else:
            return None, None, None, None, None, None, None, None

        data = json.loads(result)
        doc_type = data.get("doc_type", "Шүүхийн нэхэмжлэл")
        name = data.get("name")
        officer = data.get("officer") if data.get("officer") and data.get("officer") != "null" else "Тодорхой бус"
        status_hint = data.get("status_hint", "Шүүхэд өгсөн")
        
        c_date = data.get("court_date")
        m_date = data.get("mediation_date")
        o_date = data.get("order_date")
        summary = data.get("summary", "")
        
        return doc_type, name, officer, status_hint, c_date, m_date, o_date, summary
            
    except Exception as e:
        st.error(f"Алдаа ({file_obj.name}): {e}")
        return None, None, None, None, None, None, None, None

# --- TAB ҮҮСГЭХ ХЭСЭГ ---
tab1, tab2, tab3 = st.tabs(["📊 Хяналтын самбар", "🤖 Шинэ бүртгэл", "👥 Харилцагчид"])

with tab1:
    df = st.session_state.df_court
    cols = st.columns(4)
    def get_count(status_name):
        if df.empty or "Одоогийн төлөв" not in df.columns: return 0
        return len(df[df["Одоогийн төлөв"] == status_name])

    with cols[0]:
        st.markdown(f'<div class="metric-card"><div class="metric-num">{len(df)}</div><div class="metric-label">Нийт хэрэг</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Шүүхэд өгсөн")}</div><div class="metric-label">Шүүхэд өгсөн</div></div>', unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Захирамж гарсан")}</div><div class="metric-label">Захирамж гарсан</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Эвлэрүүлэн зуучлалд өгсөн")}</div><div class="metric-label">Эвлэрүүлэнд өгсөн</div></div>', unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Гүйцэтгэх хуудас бичүүлэх гэж өгсөн")}</div><div class="metric-label">Гүйцэтгэх хуудас бичүүлэх</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа")}</div><div class="metric-label">Гүйцэтгэлд явж байгаа</div></div>', unsafe_allow_html=True)
    with cols[3]:
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон")}</div><div class="metric-label">Дуусгавар болсон</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Өр төлбөр дууссан")}</div><div class="metric-label">Өр дууссан</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    
    # Хариуцсан ажилтнаар ангилсан Дашбоард
    st.subheader("👨‍💼 Хариуцсан ажилтнаар ангилсан бүртгэл")
    if not df.empty and "Хариуцсан ажилтан" in df.columns and "Одоогийн төлөв" in df.columns:
        officer_df = df.copy()
        officer_df["Хариуцсан ажилтан"] = officer_df["Хариуцсан ажилтан"].replace("", "Тодорхой бус").fillna("Тодорхой бус")
        
        pivot_df = officer_df.pivot_table(
            index="Хариуцсан ажилтан", 
            columns="Одоогийн төлөв", 
            aggfunc='size', 
            fill_value=0
        ).reset_index()
        
        pivot_df['Нийт хэрэг'] = pivot_df.drop(columns=['Хариуцсан ажилтан']).sum(axis=1)
        st.dataframe(pivot_df, use_container_width=True, hide_index=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### 📊 Хариуцсан ажилтны хэрэгүүдийн төлөв (Ямар ямар хэрэгтэй вэ?)")
        # Статус бүрийн өнгөөр ялгасан давхар баганан график (Stacked bar chart)
        chart_data = pivot_df.set_index("Хариуцсан ажилтан").drop(columns=['Нийт хэрэг'])
        if not chart_data.empty:
            st.bar_chart(chart_data)
    else:
        st.info("ℹ️ Хариуцсан ажилтнаар ангилсан өгөгдөл хоосон байна.")

with tab2:
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("📄 Олон файл зэрэг оруулах")
        uploaded_files = st.file_uploader("Олон файл оруулна уу (Word, PDF, Зураг)", type=['png', 'jpg', 'jpeg', 'pdf', 'docx'], accept_multiple_files=True)
        if st.button("🤖 Бүх файлыг AI-аар уншуулах", use_container_width=True):
            if uploaded_files:
                if not api_key: st.error("⚠️ Зүүн талын цэснээс Google Gemini API Key оруулна уу!")
                else:
                    progress_bar = st.progress(0); success_count = 0
                    for i, file_obj in enumerate(uploaded_files):
                        with st.spinner(f"Уншиж байна: {file_obj.name} (Хүлээж байна...)"):
                            doc_type, name, officer, status_hint, c_date, m_date, o_date, summary = extract_info_from_file(file_obj, api_key)
                            if name:
                                try: court_date = datetime.strptime(c_date, "%Y-%m-%d").date() if c_date and c_date != "null" else ""
                                except: court_date = ""
                                try: mediation_date = datetime.strptime(m_date, "%Y-%m-%d").date() if m_date and m_date != "null" else ""
                                except: mediation_date = ""
                                try: order_date = datetime.strptime(o_date, "%Y-%m-%d").date() if o_date and o_date != "null" else ""
                                except: order_date = ""
                                
                                if status_hint not in STATUS_OPTIONS:
                                    current_status = "Эвлэрүүлэн зуучлалд өгсөн" if "эвлэр" in status_hint.lower() else "Шүүхэд өгсөн"
                                else:
                                    current_status = status_hint

                                new_id = len(st.session_state.df_court) + 1
                                
                                safe_name = "".join(c for c in str(name) if c.isalnum() or c in (' ', '.', '_')).rstrip()
                                saved_filename = f"{new_id}_{safe_name}_{file_obj.name}"
                                file_path = os.path.join(UPLOAD_DIR, saved_filename)
                                with open(file_path, "wb") as f:
                                    f.write(file_obj.getbuffer())

                                new_data = {
                                    "№": new_id, "Зээлдэгч": name, "Хариуцсан ажилтан": officer,
                                    "Шүүхэд өгсөн огноо": court_date.strftime("%Y-%m-%d") if court_date else "",
                                    "Эвлэрүүлэнд өгсөн огноо": mediation_date.strftime("%Y-%m-%d") if mediation_date else "",
                                    "Захирамж гарсан огноо": order_date.strftime("%Y-%m-%d") if order_date else "",
                                    "Одоогийн төлөв": current_status, "Тэмдэглэл": summary if summary else "",
                                    "Файлын нэр": saved_filename
                                }
                                st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([new_data])], ignore_index=True)
                                save_data(); success_count += 1
                        progress_bar.progress((i + 1) / len(uploaded_files))
                    if success_count > 0:
                        st.success(f"✅ {success_count} ширхэг файл амжилттай уншигдаж бүртгэгдлээ!")
                    else:
                        st.error("⚠️ Файл уншигдсангүй. API Key-ээ шалгана уу.")
            else: st.warning("Эхлээд файлуудаа оруулна уу.")

    with col2:
        st.subheader("✍️ Гар аргаар нэг бүрчлэн бүртгэх")
        with st.form("burtgeh_form"):
            name = st.text_input("Зээлдэгчийн нэр", value=st.session_state.get('temp_name', ''))
            c_col, m_col, o_col = st.columns(3)
            with c_col: court_date = st.date_input("Шүүхэд өгсөн огноо", value=st.session_state.get('temp_c_date', datetime.now()))
            with m_col: mediation_date = st.date_input("Эвлэрүүлэнд өгсөн огноо", value=st.session_state.get('temp_m_date', datetime.now()))
            with o_col: order_date = st.date_input("Захирамж гарсан огноо (Хоосон орхиж болно)", value=st.session_state.get('temp_o_date', None))
            status = st.selectbox("Одоогийн төлөв / Дараагийн хийх ажил", STATUS_OPTIONS, index=0)
            officer = st.text_input("Хариуцсан ажилтан", value="Б.Адъяабазар")
            note = st.text_area("Тэмдэглэл (Өргөдлийн агуулга)", value=st.session_state.get('temp_note', ''))
            submitted = st.form_submit_button("Бүртгэл хадгалах", use_container_width=True)
            if submitted:
                if name:
                    new_id = len(st.session_state.df_court) + 1
                    new_data = {
                        "№": new_id, "Зээлдэгч": name, "Хариуцсан ажилтан": officer,
                        "Шүүхэд өгсөн огноо": court_date.strftime("%Y-%m-%d"), "Эвлэрүүлэнд өгсөн огноо": mediation_date.strftime("%Y-%m-%d"),
                        "Захирамж гарсан огноо": order_date.strftime("%Y-%m-%d") if order_date else "", "Одоогийн төлөв": status, "Тэмдэглэл": note,
                        "Файлын нэр": ""
                    }
                    st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([new_data])], ignore_index=True)
                    save_data()
                    st.success(f"✅ {name} амжилттай бүртгэгдлээ!")
                else: st.error("Зээлдэгчийн нэр хоосон байна!")

with tab3:
    st.subheader("👥 Бүртгэлтэй харилцагчид")
    view_mode = st.radio("Харах хэлбэр", ["Карт хэлбэрээр (Гоё)", "Хүснэгтээр (Засварлах)"], horizontal=True)
    
    if not st.session_state.df_court.empty:
        if view_mode == "Хүснэгтээр (Засварлах)":
            display_df = st.session_state.df_court.fillna("").copy()
            if "Устгах" not in display_df.columns:
                display_df.insert(0, "Устгах", False)
                
            edited_df = st.data_editor(
                display_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "Устгах": st.column_config.CheckboxColumn("Устгах", default=False),
                    "Одоогийн төлөв": st.column_config.SelectboxColumn("Одоогийн төлөв", help="Харилцагчийн үе шатыг сонгоно уу", options=STATUS_OPTIONS, required=True)
                }
            )
            
            rows_to_delete = edited_df[edited_df["Устгах"] == True]
            if not rows_to_delete.empty:
                st.warning(f"⚠️ {len(rows_to_delete)} харилцагч устгахад бэлэн боллоо.")
                if st.button(f"❌ Сонгосон {len(rows_to_delete)} харилцагчийг устгах", use_container_width=True):
                    st.session_state.df_court = edited_df[edited_df["Устгах"] == False].drop(columns=["Устгах"]).reset_index(drop=True)
                    save_data()
                    st.rerun()
            else:
                st.session_state.df_court = edited_df.drop(columns=["Устгах"]).reset_index(drop=True)
                save_data()
        else:
            # Карт хэлбэрээр хурдан гаргах (Хуудаслалт - Pagination)
            PAGE_SIZE = 8
            total_items = len(st.session_state.df_court)
            total_pages = (total_items + PAGE_SIZE - 1) // PAGE_SIZE
            
            if 'card_page' not in st.session_state:
                st.session_state.card_page = 0
            
            if st.session_state.card_page >= total_pages:
                st.session_state.card_page = 0
                
            start_idx = st.session_state.card_page * PAGE_SIZE
            end_idx = min(start_idx + PAGE_SIZE, total_items)
            
            st.write(f"Нийт {total_items} харилцагчийн {start_idx+1}-{end_idx} харуулж байна (Хуудас {st.session_state.card_page + 1}/{total_pages})")
            
            cols = st.columns(2)
            current_page_df = st.session_state.df_court.iloc[start_idx:end_idx].fillna("")
            
            display_idx = 0
            for idx, row in current_page_df.iterrows():
                status = row["Одоогийн төлөв"]
                color = STATUS_COLORS.get(status, "#1f3a5f")
                
                image_html = ""
                if "Файлын нэр" in row and row["Файлын нэр"]:
                    file_path = os.path.join(UPLOAD_DIR, row["Файлын нэр"])
                    if os.path.exists(file_path):
                        if row["Файлын нэр"].lower().endswith(('.png', '.jpg', '.jpeg', '.heic', '.webp')):
                            with open(file_path, "rb") as img_file:
                                img_b64 = base64.b64encode(img_file.read()).decode()
                            image_html = f'<div class="img-container"><img src="data:image/jpeg;base64,{img_b64}" alt="Баримт"></div>'
                        else:
                            image_html = f'<div class="note-box">📎 Файл: {row["Файлын нэр"]}</div>'

                card_html = f"""
                <div class="client-card" style="border-left-color: {color};">
                    <div class="client-name">{row['Зээлдэгч']} 
                        <span class="status-badge" style="background-color: {color};">{status}</span>
                    </div>
                    <div class="info-row"><span class="info-label">👤 Хариуцсан:</span> {row['Хариуцсан ажилтан']}</div>
                    <div class="info-row"><span class="info-label">📅 Шүүхэд өгсөн:</span> {row['Шүүхэд өгсөн огноо']}</div>
                    <div class="info-row"><span class="info-label">🤝 Эвлэрүүлэнд өгсөн:</span> {row['Эвлэрүүлэнд өгсөн огноо']}</div>
                    <div class="info-row"><span class="info-label">⚖️ Захирамж гарсан:</span> {row['Захирамж гарсан огноо']}</div>
                    {f'<div class="note-box">📝 {row["Тэмдэглэл"]}</div>' if row['Тэмдэглэл'] else ''}
                    {image_html}
                </div>
                """
                with cols[display_idx % 2]:
                    st.markdown(card_html, unsafe_allow_html=True)
                display_idx += 1
                
            # Хуудас солих товчнууд
            st.markdown("---")
            col1, col2, col3 = st.columns([1, 2, 1])
            with col1:
                if st.button("⬅️ Өмнөх хуудас", use_container_width=True, disabled=(st.session_state.card_page == 0)):
                    st.session_state.card_page -= 1
                    st.rerun()
            with col3:
                if st.button("Дараагийн хуудас ➡️", use_container_width=True, disabled=(st.session_state.card_page >= total_pages - 1)):
                    st.session_state.card_page += 1
                    st.rerun()
    else:
        st.warning("Бүртгэл хоосон байна. Шинэ нэхэмжлэл бүртгэнэ үү.")
