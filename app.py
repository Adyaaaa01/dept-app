import streamlit as st
import pandas as pd
from datetime import datetime
import json
import io
import os
import time
import docx
import pdfplumber
from groq import Groq
import base64
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
    .metric-card { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; margin-bottom: 10px; }
    .metric-num { font-size: 24px; font-weight: bold; color: #4CAF50; }
    .metric-label { font-size: 14px; color: #555; }
    </style>
""", unsafe_allow_html=True)

st.title("⚖️ Шүүх нэхэмжлэх болон Эвлэрүүлэн зуучлалын систем")
st.markdown("##### Groq AI (Llama) дэмжлэгтэй хөнгөн бөгөөд хэрэглэхэд хялбар веб апп")

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

# --- Session State үүсгэх болон өгөгдөл хадгалах ---
DATA_FILE = "court_data.csv"

if 'df_court' not in st.session_state:
    if os.path.exists(DATA_FILE):
        try:
            st.session_state.df_court = pd.read_csv(DATA_FILE)
            required_cols = ["№", "Зээлдэгч", "Хариуцсан ажилтан", "Шүүхэд өгсөн огноо", "Захирамж гарсан огноо", "Одоогийн төлөв", "Тэмдэглэл"]
            for col in required_cols:
                if col not in st.session_state.df_court.columns:
                    st.session_state.df_court[col] = ""
        except Exception:
            st.session_state.df_court = pd.DataFrame(columns=required_cols)
    else:
        st.session_state.df_court = pd.DataFrame(columns=[
            "№", "Зээлдэгч", "Хариуцсан ажилтан", "Шүүхэд өгсөн огноо", 
            "Захирамж гарсан огноо", "Одоогийн төлөв", "Тэмдэглэл"
        ])

def save_data():
    st.session_state.df_court.to_csv(DATA_FILE, index=False)

# --- Sidebar ---
st.sidebar.header("⚙️ Тохиргоо")
api_key = st.sidebar.text_input("Groq API Key оруулна уу", type="password", help="console.groq.com сайтад бүртгүүлж үнэгүй key авна уу.")

st.sidebar.markdown("---")
st.sidebar.header("📁 Файл оруулах/Татах")
up_excel = st.sidebar.file_uploader("Excel файлаас өгөгдөл татах", type=['xlsx'])
if up_excel:
    try:
        df_import = pd.read_excel(up_excel)
        st.session_state.df_court = df_import
        save_data()
        st.sidebar.success("Excel амжилттай уншигдлаа!")
    except Exception as e:
        st.sidebar.error(f"Алдаа: {e}")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Бүртгэл')
    return output.getvalue()

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

def get_count(status_name):
    if df.empty or "Одоогийн төлөв" not in df.columns: 
        return 0
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

# --- 1. AI ашиглан Олон файл уншуулах (Groq - Llama) ---
def extract_info_from_file(file_obj, key):
    if not key:
        return None, None, None, None
    
    try:
        client = Groq(api_key=key)
        
        file_text = ""
        prompt = """Энэхүү баримтаас дараах мэдээллийг татаад зөвхөн JSON формат буцаа:
        1. "name": Зээлдэгчийн буюу хариуцагчийн нэр (Овог нэр)
        2. "court_date": Шүүхэд шилжүүлсэн эсвэл өгсөн огноо (YYYY-MM-DD форматад)
        3. "order_date": Захирамж гарсан огноо (Олдоогүй бол null гэж бичнэ үү) (YYYY-MM-DD форматад)
        4. "summary": Баримт бичгийн гол агуулга, нэхэмжилсэн зүйл, шаардсан дүн зэрэгийг товч тодорхой 1-3 өгүүлбэрээр бич.
        Бусад тайлбаргүй зөвхөн JSON буцаа."""

        if file_obj.name.endswith('.docx'):
            doc = docx.Document(file_obj)
            file_text = "\n".join([para.text for para in doc.paragraphs])
            
            completion = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": "Та гүйцэтгэлийн мэргэжилтэн юм."},
                    {"role": "user", "content": prompt + "\n\nБаримтын текст:\n" + file_text}
                ],
                response_format={"type": "json_object"}
            )
            result = completion.choices[0].message.content
            
        elif file_obj.name.endswith('.pdf'):
            with pdfplumber.open(file_obj) as pdf:
                for page in pdf.pages:
                    file_text += page.extract_text() + "\n"
                    
            completion = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": "Та гүйцэтгэлийн мэргэжилтэн юм."},
                    {"role": "user", "content": prompt + "\n\nБаримтын текст:\n" + file_text}
                ],
                response_format={"type": "json_object"}
            )
            result = completion.choices[0].message.content
            
        elif file_obj.name.endswith(('.png', '.jpg', '.jpeg')):
            img = Image.open(file_obj)
            # Зургийг base64 болгох
            buf = io.BytesIO()
            img.save(buf, format='JPEG')
            img_b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
            
            completion = client.chat.completions.create(
                model="llama-3.2-90b-vision-preview",
                messages=[
                    {"role": "system", "content": "Та гүйцэтгэлийн мэргэжилтэн юм."},
                    {"role": "user", "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}}
                    ]}
                ]
            )
            result = completion.choices[0].message.content
            # JSON-ыг цэвэрлэх
            if result.startswith("```json"):
                result = result.replace("```json", "").replace("```", "").strip()
            
        else:
            return None, None, None, None

        data = json.loads(result)
        return data.get("name"), data.get("court_date"), data.get("order_date"), data.get("summary", "")
            
    except Exception as e:
        return None, None, None, None

# --- 2. Олон файл зэрэг бүртгэх хэсэг ---
st.header("📄 Шинэ нэхэмжлэл / Захирамж бүртгэх (Олон файл зэрэг оруулах)")
col1, col2 = st.columns([1, 2])

with col1:
    uploaded_files = st.file_uploader("Олон файл оруулна уу (Word, PDF, Зураг)", type=['png', 'jpg', 'jpeg', 'pdf', 'docx'], accept_multiple_files=True)
    
    if st.button("🤖 Бүх файлыг AI-аар уншуулах", use_container_width=True):
        if uploaded_files:
            if not api_key:
                st.error("⚠️ Зүүн талын цэснээс Groq API Key оруулна уу!")
            else:
                progress_bar = st.progress(0)
                success_count = 0
                
                for i, file_obj in enumerate(uploaded_files):
                    with st.spinner(f"Уншиж байна: {file_obj.name}..."):
                        name, c_date, o_date, summary = extract_info_from_file(file_obj, api_key)
                        
                        if name:
                            try:
                                court_date = datetime.strptime(c_date, "%Y-%m-%d").date() if c_date else datetime.now().date()
                            except:
                                court_date = datetime.now().date()
                                
                            try:
                                order_date = datetime.strptime(o_date, "%Y-%m-%d").date() if o_date and o_date != "null" else None
                            except:
                                order_date = None
                            
                            new_id = len(st.session_state.df_court) + 1
                            new_data = {
                                "№": new_id,
                                "Зээлдэгч": name,
                                "Хариуцсан ажилтан": "Б.Адъяабазар",
                                "Шүүхэд өгсөн огноо": court_date.strftime("%Y-%m-%d"),
                                "Захирамж гарсан огноо": order_date.strftime("%Y-%m-%d") if order_date else "",
                                "Одоогийн төлөв": "Шүүхэд өгсөн",
                                "Тэмдэглэл": summary if summary else ""
                            }
                            st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([new_data])], ignore_index=True)
                            save_data()
                            success_count += 1
                        else:
                            st.warning(f"Алдаа: {file_obj.name} файлыг уншиж чадсангүй.")
                    
                    progress_bar.progress((i + 1) / len(uploaded_files))
                
                st.success(f"✅ {success_count} ширхэг файл амжилттай уншигдаж бүртгэгдлээ!")
        else:
            st.warning("Эхлээд файлуудаа оруулна уу.")

with col2:
    st.subheader("Гар аргаар нэг бүрчлэн бүртгэх (Хэрэв AI алдаа өгвөл)")
    with st.form("burtgeh_form"):
        name = st.text_input("Зээлдэгчийн нэр", value=st.session_state.get('temp_name', ''))
        
        c_col, o_col = st.columns(2)
        with c_col:
            court_date = st.date_input("Шүүхэд өгсөн огноо", value=st.session_state.get('temp_c_date', datetime.now()))
        with o_col:
            order_date = st.date_input("Захирамж гарсан огноо (Хоосон орхиж болно)", value=st.session_state.get('temp_o_date', None))
        
        status = st.selectbox("Одоогийн төлөв / Дараагийн хийх ажил", STATUS_OPTIONS, index=0)
        officer = st.text_input("Хариуцсан ажилтан", value="Б.Адъяабазар")
        note = st.text_area("Тэмдэглэл (Өргөдлийн агуулга)", value=st.session_state.get('temp_note', ''))
        
        submitted = st.form_submit_button("Бүртгэл хадгалах", use_container_width=True)
        if submitted:
            if name:
                new_id = len(st.session_state.df_court) + 1
                new_data = {
                    "№": new_id,
                    "Зээлдэгч": name,
                    "Хариуцсан ажилтан": officer,
                    "Шүүхэд өгсөн огноо": court_date.strftime("%Y-%m-%d"),
                    "Захирамж гарсан огноо": order_date.strftime("%Y-%m-%d") if order_date else "",
                    "Одоогийн төлөв": status,
                    "Тэмдэглэл": note
                }
                st.session_state.df_court = pd.concat([st.session_state.df_court, pd.DataFrame([new_data])], ignore_index=True)
                save_data()
                st.success(f"✅ {name} амжилттай бүртгэгдлээ!")
                for key in ['temp_name', 'temp_c_date', 'temp_o_date', 'temp_note']:
                    if key in st.session_state: del st.session_state[key]
            else:
                st.error("Зээлдэгчийн нэр хоосон байна!")

# --- 3. AI Зөвлөх: Хугацаа шалгаж мэдэгдэл өгөх (Захирамж гарсан огноогоор) ---
st.header("🚨 AI Зөвлөх - Хугацааны мэдэгдэл")
today = datetime.now().date()
alerts = []

if not st.session_state.df_court.empty:
    for idx, row in st.session_state.df_court.iterrows():
        if pd.notna(row["Захирамж гарсан огноо"]) and str(row["Захирамж гарсан огноо"]) != "":
            try:
                exp_date = pd.to_datetime(row["Захирамж гарсан огноо"]).date()
                days_left = (exp_date - today).days
                
                if days_left < 0:
                    alerts.append(f"danger|🚨 <b>{row['Зээлдэгч']}</b>-ийн захирамжийн хугацаа <b>{-days_left} хоногийн өмнө</b> дууссан! Яаралтай очиж уулзах шаардлагатай.")
                elif days_left <= 7:
                    alerts.append(f"warning|⏰ <b>{row['Зээлдэгч']}</b>-ийн захирамжийн хугацаа <b>{days_left} хоног</b> үлдлээ. Очиж уулзах бэлтгэл хийгээрэй.")
            except:
                pass
