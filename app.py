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
    </style>
""", unsafe_allow_html=True)

st.title("⚖️ Шүүх нэхэмжлэх болон Эвлэрүүлэн зуучлалын систем")
st.markdown("##### Google Gemini AI дэмжлэгтэй хөнгөн бөгөөд хэрэглэхэд хялбар веб апп")

# --- Session State үүсгэх ---
if 'df_court' not in st.session_state:
    st.session_state.df_court = pd.DataFrame(columns=[
        "№", "Хавтас/Зээлдэгч", "Ажиллагааны төрөл", "Хариуцсан ажилтан", 
        "Хүлээж авсан огноо", "Захирамж гарсан огноо", "Төлөв", "Тэмдэглэл"
    ])

# --- Sidebar ---
st.sidebar.header("⚙️ Тохиргоо")
api_key = st.sidebar.text_input("Google Gemini API Key оруулна уу", type="password", help="aistudio.google.com сайтад бүртгүүлж үнэгүй key авна уу.")

st.sidebar.markdown("---")
st.sidebar.header("📁 Файл оруулах/Татах")
up_excel = st.sidebar.file_uploader("Excel файлаас өгөгдөл татах", type=['xlsx'])
if up_excel:
    try:
        df_import = pd.read_excel(up_excel, sheet_name="Шүүх нэхэмжлэл")
        st.session_state.df_court = df_import
        st.sidebar.success("Excel амжилттай уншигдлаа!")
    except Exception as e:
        st.sidebar.error(f"Алдаа: {e}")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Шүүх нэхэмжлэл')
    processed_data = output.getvalue()
    return processed_data

if not st.session_state.df_court.empty:
    st.sidebar.download_button(
        label="📥 Excel-ээ татах",
        data=to_excel(st.session_state.df_court),
        file_name='Шүүх_нэхэмжлэл_бүртгэл.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# --- 1. AI ашиглан Зураг/Word/PDF уншуулах (Gemini) ---
def extract_info_from_file(file_obj, key):
    if not key:
        st.error("⚠️ Зүүн талын цэснээс Google Gemini API Key оруулна уу!")
        return None, None
    
    try:
        genai.configure(api_key=key)
        
        # Боломжтой моделийг автоматаар сонгох (зураг бол vision, текст бол pro)
        available_models = [m.name for m in genai.list_models()]
        
        # Зураг бол gemini-1.5-pro ашиглана
        if file_obj.name.endswith(('.png', '.jpg', '.jpeg')):
            model_name = 'gemini-1.5-pro'
            if model_name not in available_models and 'models/gemini-1.0-pro-vision' in available_models:
                model_name = 'gemini-1.0-pro-vision'
        else:
            # Текст (Word, PDF) бол gemini-pro ашиглана
            model_name = 'gemini-pro'
            if model_name not in available_models and 'models/gemini-1.0-pro' in available_models:
                model_name = 'gemini-1.0-pro'

        model = genai.GenerativeModel(model_name)
        
        file_text = ""
        prompt = """Энэхүү баримтаас дараах мэдээллийг татаад зөвхөн JSON формат буцаа:
        1. "name": Зээлдэгчийн буюу хариуцагчийн нэр (Овог нэр)
        2. "date": Захирамж гарсан эсвэл хугацаа дуусах огноо (YYYY-MM-DD форматад)
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
            return None, None

        # Gemini-н хариуг цэвэрлэх
        result = response.text.strip()
        if result.startswith("```json"):
            result = result.replace("```json", "").replace("```", "").strip()
            
        data = json.loads(result)
        return data.get("name"), data.get("date")
            
    except Exception as e:
        # Хэрэв модель олдохгүй бол боломжтой моделийн жагсаалтыг харуулна
        if "404" in str(e):
            try:
                models_list = [m.name for m in genai.list_models()]
                st.error(f"Модель олдсонгүй. Таны API дээр боломжтой моделиуд: {', '.join(models_list)}")
            except:
                st.error("API Key буруу эсвэл холболт хийгдсэнгүй.")
        else:
            st.error(f"AI уншихад алдаа гарлаа: {e}")
        return None, None

# --- 2. Шинэ хэрэг бүртгэх хэсэг ---
st.header("📄 Шинэ нэхэмжлэл / Захирамж бүртгэх")
col1, col2 = st.columns([1, 2])

with col1:
    uploaded_file = st.file_uploader("Захирамж / Нэхэмжлэл оруулна уу (Word, PDF, Зураг)", type=['png', 'jpg', 'jpeg', 'pdf', 'docx'])
    if st.button("🤖 AI-аар файл уншуулах", use_container_width=True):
        if uploaded_file:
            with st.spinner("Gemini AI файл уншиж байна. Түр хүлээнэ үү..."):
                name, date_str = extract_info_from_file(uploaded_file, api_key)
                if name and date_str:
                    st.session_state['temp_name'] = name
                    try:
                        st.session_state['temp_date'] = datetime.strptime(date_str, "%Y-%m-%d").date()
                    except:
                        st.session_state['temp_date'] = datetime.now().date()
                    st.success("✅ Амжилттай уншилаа! Баруун талд шалгаад баталгаажуулна уу.")
        else:
            st.warning("Эхлээд файл оруулна уу.")

with col2:
    with st.form("burtgeh_form"):
        st.subheader("Бүртгэлийн мэдээлэл баталгаажуулах")
        name = st.text_input("Зээлдэгчийн нэр", value=st.session_state.get('temp_name', ''))
        
        default_date = st.session_state.get('temp_date', datetime.now())
        date_obj = st.date_input("Захирамж гарсан / Дуусах огноо", value=default_date)
        
        next_action = st.selectbox("Дараагийн хийх ажил", ["Очиж уулзах", "Утсаар мэдэгдэх", "Шийдвэр хүлээх", "Гүйцэтгэх хуудас бичүүлэх"])
        officer = st.text_input("Хариуцсан ажилтан", value="Б.Адъяабазар")
        
        submitted = st.form_submit_button("Бүртгэл хадгалах", use_container_width=True)
        if submitted:
            if name:
                new_id = len(st.session_state.df_court) + 1
                new_data = {
                    "№": new_id,
                    "Хавтас/Зээлдэгч": name,
                    "Ажиллагааны төрөл": "Шүүхийн нэхэмжлэх",
                    "Хариуцсан ажилтан": officer,
                    "Хүлээж авсан огноо": datetime.now().strftime("%Y-%m-%d"),
                    "Захирамж гарсан огноо": date_obj.strftime("%Y-%m-%d"),
                    "Төлөв": "Шүүхэд өгсөн",
                    "Тэмдэглэл": next_action
                }
                st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([new_data])], ignore_index=True)
                st.success(f"✅ {name} амжилттай бүртгэгдлээ!")
                if 'temp_name' in st.session_state: del st.session_state['temp_name']
                if 'temp_date' in st.session_state: del st.session_state['temp_date']
            else:
                st.error("Зээлдэгчийн нэр хоосон байна!")

# --- 3. AI Зөвлөх: Хугацаа шалгаж мэдэгдэл өгөх ---
st.header("🚨 AI Зөвлөх - Хугацааны мэдэгдэл")
today = datetime.now().date()
alerts = []

if not st.session_state.df_court.empty:
    for idx, row in st.session_state.df_court.iterrows():
        if pd.notna(row["Захирамж гарсан огноо"]):
            try:
                exp_date = pd.to_datetime(row["Захирамж гарсан огноо"]).date()
                days_left = (exp_date - today).days
                
                if days_left < 0:
                    alerts.append(f"danger|🚨 <b>{row['Хавтас/Зээлдэгч']}</b>-ийн захирамжийн хугацаа <b>{-days_left} хоногийн өмнө</b> дууссан! Яаралтай очиж уулзах шаардлагатай.")
                elif days_left <= 7:
                    alerts.append(f"warning|⏰ <b>{row['Хавтас/Зээлдэгч']}</b>-ийн захирамжийн хугацаа <b>{days_left} хоног</b> үлдлээ. Очиж уулзах бэлтгэл хийгээрэй.")
            except:
                pass

if alerts:
    for alert in alerts:
        css_class, msg = alert.split("|", 1)
        st.markdown(f'<div class="alert-box {css_class}">{msg}</div>', unsafe_allow_html=True)
else:
    st.info("ℹ️ Хугацаа дуусаж байгаа нэхэмжлэл алга байна.")

# --- 4. Бүртгэлийн жагсаалт ---
st.header("📊 Бүртгэлийн жагсаалт")
if not st.session_state.df_court.empty:
    st.data_editor(
        st.session_state.df_court, 
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic"
    )
else:
    st.warning("Бүртгэл хоосон байна. Шинэ нэхэмжлэл бүртгэнэ үү.")
