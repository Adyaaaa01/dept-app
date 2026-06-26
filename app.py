import streamlit as st
import pandas as pd
from datetime import datetime
import json
import io
import os
import time
import re
import base64
import requests
import docx
import pdfplumber
import google.generativeai as genai
from PIL import Image

# --- Хуудасны тохиргоо ---
st.set_page_config(page_title="Өр нэхэмжлэх удирдлага", layout="wide", page_icon="⚖️")

UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# --- Нууц үгс (Secrets) унших ---
API_KEY = st.secrets.get("GEMINI_API_KEY", "")
GH_TOKEN = st.secrets.get("GH_TOKEN", "")
GH_OWNER = st.secrets.get("GH_OWNER", "")
GH_REPO = st.secrets.get("GH_REPO", "")

# --- Орчин үеийн өнгө үзэмж (CSS) ---
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
st.markdown("##### Мөнхөө хадгалагдах, AI дэмжлэгтэй систем (Сервер унтарсан ч өгөгдөл хадгалагдана)")

STATUS_OPTIONS = [
    "Шүүхэд өгсөн", "Эвлэрүүлэн зуучлалд өгсөн", "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа", 
    "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон", "Захирамж гарсан", "Шүүхийн шийдвэрийн зардалын гэрээ", 
    "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн", "Гүйцэтгэх хуудас гарсан шүүхийн шийдвэрт шилжүүлсэн", 
    "Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа", "Өр төлбөр дууссан"
]

STATUS_COLORS = {
    "Шүүхэд өгсөн": "#1f3a5f", "Эвлэрүүлэн зуучлалд өгсөн": "#8e44ad", "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа": "#2980b9",
    "Эвлэрүүлэнд өгсөн ч хэрэг дуусгавар болсон": "#7f8c8d", "Захирамж гарсан": "#f39c12", "Шүүхийн шийдвэрийн зардалын гэрээ": "#16a085",
    "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн": "#d35400", "Гүйцэтгэх хуудас гарсан шүүхийн шийдвэрт шилжүүлсэн": "#c0392b",
    "Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа": "#e74c3c", "Өр төлбөр дууссан": "#27ae60"
}

DATA_FILE = "court_data.csv"
required_cols = ["№", "Зээлдэгч", "Хариуцсан ажилтан", "Шүүхэд өгсөн огноо", "Эвлэрүүлэнд өгсөн огноо", "Захирамж гарсан огноо", "Одоогийн төлөв", "Тэмдэглэл", "Файлын нэр"]

# --- GitHub Sync Функцүүд ---
def sync_to_github(df, path="court_data.csv"):
    if not GH_TOKEN or not GH_OWNER or not GH_REPO: return False
    url = f"https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/contents/{path}"
    headers = {"Authorization": f"token {GH_TOKEN}", "Accept": "application/vnd.github.v3+json"}
    r = requests.get(url, headers=headers)
    sha = r.json().get("sha") if r.status_code == 200 else None
    content = df.to_csv(index=False).encode('utf-8')
    content_b64 = base64.b64encode(content).decode('utf-8')
    data = {"message": "Auto backup", "content": content_b64, "sha": sha}
    r = requests.put(url, headers=headers, json=data)
    return r.status_code in [200, 201]

def load_from_github(path="court_data.csv"):
    url = f"https://raw.githubusercontent.com/{GH_OWNER}/{GH_REPO}/main/{path}"
    try:
        return pd.read_csv(url)
    except:
        return None

# --- Өгөгдөл ачаалах ---
if 'df_court' not in st.session_state:
    gh_df = load_from_github() if GH_TOKEN else None
    if gh_df is not None and not gh_df.empty:
        st.session_state.df_court = gh_df
        for col in required_cols:
            if col not in st.session_state.df_court.columns:
                st.session_state.df_court[col] = ""
    else:
        st.session_state.df_court = pd.DataFrame(columns=required_cols)

def save_data():
    st.session_state.df_court.to_csv(DATA_FILE, index=False)
    if GH_TOKEN:
        with st.spinner("☁️ Өгөгдөл хадгалж байна..."):
            if sync_to_github(st.session_state.df_court):
                st.sidebar.success("☁️ Нөөцлөгдлөө!")
            else:
                st.sidebar.error("⚠️ Нөөцлөхөд алдаа!")

# --- Sidebar ---
st.sidebar.header("⚙️ Тохиргоо")
if API_KEY:
    st.sidebar.success("✅ Gemini API Key холбогдсон.")
else:
    st.sidebar.error("⚠️ Secrets хэсэгт GEMINI_API_KEY оруулна уу!")
    
if GH_TOKEN:
    st.sidebar.success("✅ GitHub нөөцлөлт идэвхтэй.")
else:
    st.sidebar.error("⚠️ Secrets хэсэгт GH_TOKEN оруулна уу!")

# --- ШИНЭ: AI Модель сонгох ---
st.sidebar.markdown("---")
st.sidebar.header("🤖 AI Модель сонгох")
selected_model = st.sidebar.selectbox(
    "Ашиглах моделийг сонгоно уу:",
    ["gemini-flash-latest", "gemini-2.0-flash", "gemini-2.5-flash", "gemini-1.5-flash"],
    index=0,
    help="Хэрэв нэг модель алдаа өгвөл өөр модель сонгоорой."
)

st.sidebar.markdown("---")
st.sidebar.header("📁 Excel оруулах/Татах")
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
    return output.getvalue()

if not st.session_state.df_court.empty:
    st.sidebar.download_button(label="📥 Excel-ээ татах", data=to_excel(st.session_state.df_court), file_name='Шүүх_нэхэмжлэл_бүртгэл.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.sidebar.markdown("---")
if st.sidebar.button("🗑️ Бүх бүртгэлийг устгах", use_container_width=True):
    st.session_state.df_court = pd.DataFrame(columns=required_cols)
    save_data()
    st.rerun()

# --- AI унших функц ---
def generate_with_retry(model, prompt_parts, max_retries=3):
    for attempt in range(max_retries):
        try:
            return model.generate_content(prompt_parts, request_options={"timeout": 120})
        except Exception as e:
            if "429" in str(e) and attempt < max_retries - 1:
                st.warning(f"⏳ AI-н хязгаар хэтэрсэн тул 60 секунд хүлээж байна... ({attempt+1}/{max_retries-1})")
                time.sleep(60)
            else:
                raise e

def extract_info_from_file(file_obj):
    if not API_KEY:
        st.error("⚠️ API Key оруулаагүй байна!")
        return None
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel(selected_model) # Хэрэглэгчийн сонгосон модель
        
        file_text = ""
        prompt = """Энэхүү баримт бичгийн зураг эсвэл текстийг шинжилж, зөвхөн JSON формат буцаа:
        1. "doc_type": Баримтын төрөл
        2. "name": Зээлдэгчийн нэр
        3. "officer": Хариуцсан ажилтан
        4. "status_hint": Төлөв (${STATUS_OPTIONS} эдгээрээс нэгийг сонго)
        5. "court_date": YYYY-MM-DD
        6. "order_date": YYYY-MM-DD
        7. "summary": Гол агуулга. Зөвхөн JSON."""
        
        if file_obj.name.endswith('.docx'):
            doc = docx.Document(file_obj)
            file_text = "\n".join([para.text for para in doc.paragraphs])
            response = generate_with_retry(model, prompt + "\n\nТекст:\n" + file_text)
        elif file_obj.name.endswith('.pdf'):
            with pdfplumber.open(file_obj) as pdf:
                for page in pdf.pages: file_text += page.extract_text() + "\n"
            response = generate_with_retry(model, prompt + "\n\nТекст:\n" + file_text)
        elif file_obj.name.endswith(('.png', '.jpg', '.jpeg')):
            img = Image.open(file_obj)
            if img.mode != 'RGB': img = img.convert('RGB')
            response = generate_with_retry(model, [prompt, img])
        else:
            return None

        result = response.text.strip()
        if result.startswith("```json"): result = result.replace("```json", "").replace("```", "").strip()
        json_match = re.search(r'\{.*\}', result, re.DOTALL)
        if json_match:
            data = json.loads(json_match.group(0))
            return data.get("doc_type"), data.get("name"), data.get("officer"), data.get("status_hint"), data.get("court_date"), None, data.get("order_date"), data.get("summary", "")
        return None
    except Exception as e:
        st.error(f"Алдаа: {e}")
        return None

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
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Шүүхийн шийдвэрийн зардалын гэрээ")}</div><div class="metric-label">Зардалын гэрээ</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Гүйцэтгэх хуудас бичүүлэх гэж өгсөн")}</div><div class="metric-label">Гүйцэтгэх хуудас бичүүлэх</div></div>', unsafe_allow_html=True)
    with cols[3]:
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа")}</div><div class="metric-label">Гүйцэтгэлд явж байгаа</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Өр төлбөр дууссан")}</div><div class="metric-label">Өр дууссан</div></div>', unsafe_allow_html=True)

with tab2:
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("📄 Олон файл зэрэг оруулах")
        uploaded_files = st.file_uploader("Олон файл оруулна уу", type=['png', 'jpg', 'jpeg', 'pdf', 'docx'], accept_multiple_files=True)
        if st.button("🤖 Бүх файлыг AI-аар уншуулах", use_container_width=True):
            if uploaded_files:
                progress_bar = st.progress(0); success_count = 0
                for i, file_obj in enumerate(uploaded_files):
                    with st.spinner(f"Уншиж байна: {file_obj.name}..."):
                        info = extract_info_from_file(file_obj)
                        if info and info[1]:
                            new_id = len(st.session_state.df_court) + 1
                            existing_names = st.session_state.df_court["Зээлдэгч"].astype(str).tolist()
                            count = sum(1 for n in existing_names if n.startswith(info[1]))
                            display_name = f"{info[1]} ({count + 1})" if count > 0 else info[1]
                            
                            safe_name = "".join(c for c in str(info[1]) if c.isalnum() or c in (' ', '.', '_')).rstrip()
                            saved_filename = f"{new_id}_{safe_name}_{file_obj.name}"
                            with open(os.path.join(UPLOAD_DIR, saved_filename), "wb") as f:
                                f.write(file_obj.getbuffer())

                            st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([{
                                "№": new_id, "Зээлдэгч": display_name, "Хариуцсан ажилтан": info[2] if info[2] and info[2] != "null" else "Тодорхой бус",
                                "Шүүхэд өгсөн огноо": info[4] if info[4] and info[4] != "null" else "",
                                "Эвлэрүүлэнд өгсөн огноо": "",
                                "Захирамж гарсан огноо": info[6] if info[6] and info[6] != "null" else "",
                                "Одоогийн төлөв": info[3] if info[3] in STATUS_OPTIONS else "Шүүхэд өгсөн",
                                "Тэмдэглэл": info[7] if info[7] else "", "Файлын нэр": saved_filename
                            }])], ignore_index=True)
                            save_data(); success_count += 1
                    progress_bar.progress((i + 1) / len(uploaded_files))
                if success_count > 0:
                    st.success(f"✅ {success_count} ширхэг файл амжилттай уншигдлаа!")

    with col2:
        st.subheader("✍️ Гар аргаар нэг бүрчлэн бүртгэх")
        with st.form("burtgeh_form"):
            name = st.text_input("Зээлдэгчийн нэр")
            c_col, o_col = st.columns(2)
            with c_col: court_date = st.date_input("Шүүхэд өгсөн огноо", datetime.now())
            with o_col: order_date = st.date_input("Захирамж гарсан огноо", None)
            status = st.selectbox("Одоогийн төлөв", STATUS_OPTIONS, index=0)
            officer = st.text_input("Хариуцсан ажилтан", "Б.Адъяабазар")
            note = st.text_area("Тэмдэглэл")
            if st.form_submit_button("Бүртгэл хадгалах", use_container_width=True):
                if name:
                    new_id = len(st.session_state.df_court) + 1
                    st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([{
                        "№": new_id, "Зээлдэгч": name, "Хариуцсан ажилтан": officer,
                        "Шүүхэд өгсөн огноо": court_date.strftime("%Y-%m-%d"), "Эвлэрүүлэнд өгсөн огноо": "",
                        "Захирамж гарсан огноо": order_date.strftime("%Y-%m-%d") if order_date else "",
                        "Одоогийн төлөв": status, "Тэмдэглэл": note, "Файлын нэр": ""
                    }])], ignore_index=True)
                    save_data()
                    st.success(f"✅ {name} амжилттай бүртгэгдлээ!")

with tab3:
    st.subheader("👥 Бүртгэлтэй харилцагчид")
    col_search1, col_search2 = st.columns([3, 2])
    with col_search1:
        search_query = st.text_input("🔍 Нэр эсвэл тэмдэглэлээр хайх", "")
    with col_search2:
        filter_status = st.selectbox("Төлөвөөр шүүх", ["Бүгд"] + STATUS_OPTIONS)
        
    df_display = st.session_state.df_court.copy().fillna("")
    if search_query:
        df_display = df_display[df_display['Зээлдэгч'].str.contains(search_query, case=False) | df_display['Тэмдэглэл'].str.contains(search_query, case=False)]
    if filter_status != "Бүгд":
       styled_df = df_display.style.map(color_status, subset=['Одоогийн төлөв'])     
    if not df_display.empty:
        def color_status(val):
            return f'background-color: {STATUS_COLORS.get(val, "#1f3a5f")}; color: white; font-weight: bold;'
            styled_df = df_display.style.map(color_status, subset=['Одоогийн төлөв'])
        st.dataframe(styled_df, use_container_width=True, hide_index=True)
    else:
        st.warning("Хайлтын үр дүн хоосон байна.")
