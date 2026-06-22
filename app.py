import streamlit as st
import pandas as pd
from datetime import datetime
import json
import io
import docx
import pdfplumber
import google.generativeai as genai
from PIL import Image

# --- Хуудасны тохиргоо ---
st.set_page_config(page_title="Өр нэхэмжлэх удирдлага", layout="wide", page_icon="⚖️")

# CSS төрх тохируулах
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { border-radius: 20px; border: 1px solid #4CAF50; color: #4CAF50; }
    .stButton>button:hover { background-color: #4CAF50; color: white; }
    .alert-box { padding: 15px; border-radius: 10px; margin-bottom: 10px; color: white; font-weight: bold; }
    .danger { background-color: #dc3545; }
    .warning { background-color: #ffc107; color: black; }
    .metric-card { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; }
    .metric-num { font-size: 24px; font-weight: bold; color: #4CAF50; }
    .metric-label { font-size: 14px; color: #555; }
    </style>
""", unsafe_allow_html=True)

st.title("⚖️ Шүүх нэхэмжлэх болон Эвлэрүүлэн зуучлалын систем")
st.markdown("##### Google Gemini AI дэмжлэгтэй хөнгөн бөгөөд хэрэглэхэд хялбар веб апп")

# --- Үе шатын төлөвүүд ---
STATUS_OPTIONS = [
    "Шүүхэд өгсөн", 
    "Эвлэрүүлэн зуучлалд өгсөн", 
    "Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа", 
    "Захирамж гарсан", 
    "Гүйцэтгэх хуудас бичүүлэх гэж өгсөн", 
    "Гүйцэтгэх хуудас гарсан шүүхийн шийдвэрт шилжүүлсэн", 
    "Шүүхийн шийдвэр гүйцэтгэх ажиллагаанд явж байгаа", 
    "Өр төлбөр дууссан"
]

import os
# --- Session State үүсгэх болон өгөгдөл хадгалах ---
DATA_FILE = "court_data.csv"

if 'df_court' not in st.session_state:
    # Хэрэв өмнө хадгалсан файл байвал түүнийг уншина
    if os.path.exists(DATA_FILE):
        st.session_state.df_court = pd.read_csv(DATA_FILE)
        # Хуучин өгөгдөлд багана дутуу байвал нэмж өгөх
        if 'Одоогийн төлөв' not in st.session_state.df_court.columns:
            st.session_state.df_court['Одоогийн төлөв'] = ''
    else:
        # Анх удаа ажиллуулж байгаа бол хоосон хүснэгт үүсгэх
        st.session_state.df_court = pd.DataFrame(columns=[
            "№", "Зээлдэгч", "Хариуцсан ажилтан", "Шүүхэд өгсөн огноо", 
            "Захирамж гарсан огноо", "Одоогийн төлөв", "Тэмдэглэл"
        ])

# Өгөгдлийг файлд хадгалах функц
def save_data():
    st.session_state.df_court.to_csv(DATA_FILE, index=False)

# --- Sidebar ---
st.sidebar.header("⚙️ Тохиргоо")
api_key = st.sidebar.text_input("Google Gemini API Key оруулна уу", type="password", help="aistudio.google.com сайтад бүртгүүлж үнэгүй key авна уу.")

st.sidebar.markdown("---")
st.sidebar.header("📁 Файл оруулах/Татах")
up_excel = st.sidebar.file_uploader("Excel файлаас өгөгдөл татах", type=['xlsx'])
if up_excel:
    try:
        df_import = pd.read_excel(up_excel)
        st.session_state.df_court = df_import
        st.sidebar.success("Excel амжилттай уншигдлаа!")
    except Exception as e:
        st.sidebar.error(f"Алдаа: {e}")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Бүртгэл')
    processed_data = output.getvalue()
    return processed_data

if not st.session_state.df_court.empty:
    st.sidebar.download_button(
        label="📥 Excel-ээ татах",
        data=to_excel(st.session_state.df_court),
        file_name='Шүүх_нэхэмжлэл_бүртгэл.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# --- Dashboard (Хяналтын самбар) ---
st.header("📊 Хяналтын самбар")
df = st.session_state.df_court
cols = st.columns(4)

# Самбарын тоонуудыг бодож гаргах
def get_count(status_name):
    if df.empty: return 0
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
    st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Эвлэрүүлэн зуучлалын захирамж дагуу төлж байгаа")}</div><div class="metric-label">Эвлэрүүлэх захирамж дагуу төлж байгаа</div></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric-card"><div class="metric-num">{get_count("Өр төлбөр дууссан")}</div><div class="metric-label">Өр дууссан</div></div>', unsafe_allow_html=True)

st.markdown("---")

# --- 1. AI ашиглан Зураг/Word/PDF уншуулах (Gemini) ---
def extract_info_from_file(file_obj, key):
    if not key:
        st.error("⚠️ Зүүн талын цэснээс Google Gemini API Key оруулна уу!")
        return None, None, None
    
    try:
        genai.configure(api_key=key)
        model = genai.GenerativeModel('gemini-flash-latest')
        
        file_text = ""
        prompt = """Энэхүү баримтаас дараах мэдээллийг татаад зөвхөн JSON формат буцаа:
        1. "name": Зээлдэгчийн буюу хариуцагчийн нэр (Овог нэр)
        2. "court_date": Шүүхэд шилжүүлсэн эсвэл өгсөн огноо (YYYY-MM-DD форматад)
        3. "order_date": Захирамж гарсан огноо (Олдоогүй бол null гэж бичнэ үү) (YYYY-MM-DD форматад)
        Бусад тайлбаргүй зөвхөн JSON буцаа."""

        # Word файл унших
        if file_obj.name.endswith('.docx'):
            doc = docx.Document(file_obj)
            file_text = "\n".join([para.text for para in doc.paragraphs])
            response = model.generate_content(prompt + "\n\nБаримтын текст:\n" + file_text)
            
        # PDF файл унших
        elif file_obj.name.endswith('.pdf'):
            with pdfplumber.open(file_obj) as pdf:
                for page in pdf.pages:
                    file_text += page.extract_text() + "\n"
            response = model.generate_content(prompt + "\n\nБаримтын текст:\n" + file_text)
            
        # Зураг унших
        elif file_obj.name.endswith(('.png', '.jpg', '.jpeg')):
            img = Image.open(file_obj)
            response = model.generate_content([prompt, img])
            
        else:
            st.error("Дэмжигдээгүй файлын төрөл байна.")
            return None, None, None

        # Gemini-н хариуг цэвэрлэх
        result = response.text.strip()
        if result.startswith("```json"):
            result = result.replace("```json", "").replace("```", "").strip()
            
        data = json.loads(result)
        return data.get("name"), data.get("court_date"), data.get("order_date")
            
    except Exception as e:
        st.error(f"AI уншихад алдаа гарлаа: {e}")
        return None, None, None

# --- 2. Шинэ хэрэг бүртгэх хэсэг ---
st.header("📄 Шинэ нэхэмжлэл / Захирамж бүртгэх")
col1, col2 = st.columns([1, 2])

with col1:
    uploaded_file = st.file_uploader("Захирамж / Нэхэмжлэл оруулна уу (Word, PDF, Зураг)", type=['png', 'jpg', 'jpeg', 'pdf', 'docx'])
    if st.button("🤖 AI-аар файл уншуулах", use_container_width=True):
        if uploaded_file:
            with st.spinner("Gemini AI файл уншиж байна. Түр хүлээнэ үү..."):
                name, c_date, o_date = extract_info_from_file(uploaded_file, api_key)
                if name:
                    st.session_state['temp_name'] = name
                    try:
                        if c_date: st.session_state['temp_c_date'] = datetime.strptime(c_date, "%Y-%m-%d").date()
                        else: st.session_state['temp_c_date'] = datetime.now().date()
                    except: st.session_state['temp_c_date'] = datetime.now().date()
                    
                    try:
                        if o_date and o_date != "null": st.session_state['temp_o_date'] = datetime.strptime(o_date, "%Y-%m-%d").date()
                        else: st.session_state['temp_o_date'] = None
                    except: st.session_state['temp_o_date'] = None
                    
                    st.success("✅ Амжилттай уншилаа! Баруун талд шалгаад баталгаажуулна уу.")
        else:
            st.warning("Эхлээд файл оруулна уу.")

with col2:
    with st.form("burtgeh_form"):
        st.subheader("Бүртгэлийн мэдээлэл баталгаажуулах")
        name = st.text_input("Зээлдэгчийн нэр", value=st.session_state.get('temp_name', ''))
        
        c_col, o_col = st.columns(2)
        with c_col:
            court_date = st.date_input("Шүүхэд өгсөн огноо", value=st.session_state.get('temp_c_date', datetime.now()))
        with o_col:
            # Захирамж гараагүй бол хоосон орхих боломжтой болгох
            order_date = st.date_input("Захирамж гарсан огноо (Хоосон орхиж болно)", value=st.session_state.get('temp_o_date', None))
        
        status = st.selectbox("Одоогийн төлөв / Дараагийн хийх ажил", STATUS_OPTIONS, index=0)
        officer = st.text_input("Хариуцсан ажилтан", value="Б.Адъяабазар")
        note = st.text_input("Тэмдэглэл")
        
        submitted = st.form_submit_button("Бүртгэл хадгалах", use_container_width=True)
        if submitted:
            if name:
                new_id = len(st.session_state.df_court) + 1
    'if' st.button("Бүртгэх"):
    'if' name:
        new_data = {
            ...
        }
    "Зээлдэгч": name,
    "Шүүх": court_name,
    "Шийдвэрийн дугаар": decision_no,
    "Шүүхэд өгсөн огноо": court_date.strftime("%Y-%m-%d"),
    "Захирамж гарсан огноо": order_date.strftime("%Y-%m-%d") if order_date else "",
    "Одоогийн төлөв": status,
    "Тэмдэглэл": note
        '}'

st.session_state.df_court = pd.concat(
    [st.session_state.df_court, pd.DataFrame([new_data])],
    ignore_index=True
)

save_data()
st.success(f"✅ {name} амжилттай бүртгэгдлээ!")
                # Түр зуурын өгөгдөл цэвэрлэх
                for key in ['temp_name', 'temp_c_date', 'temp_o_date']:
                    if key in st.session_state: del st.session_state[key]
            else:
                st.error("Зээлдэгчийн нэр хоосон байна!")

# --- 3. AI Зөвлөх: Хугацаа шалгаж мэдэгдэл өгөх (Захирамж гарсан огноогоор) ---
st.header("🚨 AI Зөвлөх - Хугацааны мэдэгдэл")
today = datetime.now().date()
alerts = []

if not st.session_state.df_court.empty:
    for idx, row in st.session_state.df_court.iterrows():
        if pd.notna(row["Захирамж гарсан огноо"]) and row["Захирамж гарсан огноо"] != "":
            try:
                exp_date = pd.to_datetime(row["Захирамж гарсан огноо"]).date()
                days_left = (exp_date - today).days
                
                if days_left < 0:
                    alerts.append(f"danger|🚨 <b>{row['Зээлдэгч']}</b>-ийн захирамжийн хугацаа <b>{-days_left} хоногийн өмнө</b> дууссан! Яаралтай очиж уулзах шаардлагатай.")
                elif days_left <= 7:
                    alerts.append(f"warning|⏰ <b>{row['Зээлдэгч']}</b>-ийн захирамжийн хугацаа <b>{days_left} хоног</b> үлдлээ. Очиж уулзах бэлтгэл хийгээрэй.")
            except:
                pass

if alerts:
    for alert in alerts:
        css_class, msg = alert.split("|", 1)
        st.markdown(f'<div class="alert-box {css_class}">{msg}</div>', unsafe_allow_html=True)
else:
    st.info("ℹ️ Хугацаа дуусаж байгаа захирамж алга байна.")

# --- 4. Бүртгэлийн жагсаалт (Шууд засварлах боломжтой) ---
st.header("📋 Бүртгэлийн жагсаалт (Төлвийг шууд өөрчилж болно)")
if not st.session_state.df_court.empty:
    # Data editor ашиглан хүснэгтийг засварлах боломжтой болгоно
    edited_df = st.data_editor(
        st.session_state.df_court,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Одоогийн төлөв": st.column_config.SelectboxColumn(
                "Одоогийн төлөв",
                help="Харилцагчийн үе шатыг сонгоно уу",
                options=STATUS_OPTIONS,
                required=True
            )
        }
    )
    # Хэрвээ хэрэглэгч засварласан бол session-д хадгалах
    st.session_state.df_court = edited_df
else:
    st.warning("Бүртгэл хоосон байна. Шинэ нэхэмжлэл бүртгэнэ үү.")
