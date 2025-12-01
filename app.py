import streamlit as st

st.set_page_config(
    page_title="Permohonan Geser Meter",
    page_icon="assets/logo_pln.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

import os
from PIL import Image
from datetime import datetime

try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    def now_jakarta():
        return datetime.now()

BASE_DIR = os.path.dirname(__file__)
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
SIDEBAR_DIR = os.path.join(BASE_DIR, "sidebar")

st.markdown("""
<style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }
    
    h1 {
        margin-top: 0 !important;
        margin-bottom: 1rem !important;
    }
    
    h2, h3 {
        margin-top: 1rem !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: #1e3a5f;
    }
    
    [data-testid="stSidebar"] * {
        color: #ffffff !important;
    }
    
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #2c5282 !important;
    }
    
    [data-testid="stSidebar"] [data-baseweb="select"] span {
        color: #ffffff !important;
    }
    
    [data-testid="stSidebar"] a {
        color: #87ceeb !important;
    }
    
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1rem;
        padding-bottom: 0.5rem;
    }
    
    [data-testid="stSidebar"] .element-container {
        margin-bottom: 0.5rem;
    }
    
    [data-testid="stSidebar"] hr {
        margin-top: 0.8rem;
        margin-bottom: 0.8rem;
        border-color: #4a6fa5;
    }
    
    [data-testid="stSidebar"] button[kind="secondary"] {
        background-color: #2c5282 !important;
        color: #ffffff !important;
        border: 1px solid #4a6fa5 !important;
    }
    
    [data-testid="stSidebar"] button[kind="secondary"]:hover {
        background-color: #3d6aa3 !important;
        border-color: #5b8bc9 !important;
    }
    
    [data-testid="stSidebar"] button[kind="secondary"] p {
        color: #ffffff !important;
    }
</style>
""", unsafe_allow_html=True)

LOGO_PATH = os.path.join(ASSETS_DIR, "logo_pln.png")
if os.path.exists(LOGO_PATH):
    img = Image.open(LOGO_PATH)
    col1, col2 = st.sidebar.columns([1, 2])
    with col1:
        st.image(img, width=70)
    with col2:
        st.markdown(
            "<div style='padding-top:5px;'>"
            "<p style='margin:0; line-height:1.2; color:#ffd700; font-size:17px; font-weight:bold;'>PLN ULP</p>"
            "<p style='margin:0; line-height:1.2; color:#ffd700; font-size:17px; font-weight:bold;'>DINOYO</p>"
            "<p style='margin:3px 0 0 0; line-height:1.2; color:#87ceeb; font-size:12px; font-style:italic;'>Dashboard Petugas</p>"
            "</div>", 
            unsafe_allow_html=True
        )
else:
    st.sidebar.warning("Logo tidak ditemukan.")

st.sidebar.markdown(
    "<a href='https://maps.app.goo.gl/CnhdCBrhz3mihieL9' "
    "style='text-decoration:none; color:#87ceeb !important; font-size:12px; display:block; margin-top:8px;' "
    "target='_blank'>Jl. Pandan No.15, Gading Kasri, Klojen</a>",
    unsafe_allow_html=True
)

st.sidebar.markdown("<hr>", unsafe_allow_html=True)

st.sidebar.markdown(
    "<p style='font-size:14px; font-weight:bold; margin-bottom:8px; color:#ffd700;'>Menu Navigasi</p>", 
    unsafe_allow_html=True
)

pages = {
    "Proses": "Proses",
    "Eksekusi": "Eksekusi",
}

choice = st.sidebar.selectbox(
    "Pilih Menu", 
    list(pages.keys()),
    index=0,
    label_visibility="collapsed"
)

st.sidebar.markdown("<hr>", unsafe_allow_html=True)

if st.sidebar.button("Reload Data Harga", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

page_module = pages.get(choice)

if page_module:
    try:
        import importlib.util
        import sys
        
        module_path = os.path.join(SIDEBAR_DIR, f"{page_module}.py")
        
        if os.path.exists(module_path):
            spec = importlib.util.spec_from_file_location(page_module, module_path)
            if spec is not None and spec.loader is not None:
                module = importlib.util.module_from_spec(spec)
                sys.modules[page_module] = module
                spec.loader.exec_module(module)
            else:
                st.error(f"Gagal membuat ModuleSpec atau loader untuk {page_module}.py")
        else:
            st.error(f"File {page_module}.py tidak ditemukan di folder sidebar/")
            
    except Exception as e:
        st.error(f"Gagal memuat halaman: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
else:
    st.error("Halaman tidak ditemukan.")

st.sidebar.markdown("<hr>", unsafe_allow_html=True)
st.sidebar.markdown(
    "<p style='color:#ffd700; font-size:14px; margin-bottom:5px; font-weight:bold;'>Informasi Akses</p>",
    unsafe_allow_html=True
)
st.sidebar.markdown(
    f"<p style='color:#ffffff; font-size:13px; margin:3px 0;'>Tanggal: {now_jakarta().strftime('%d %B %Y')}</p>"
    f"<p style='color:#ffffff; font-size:13px; margin:3px 0;'>Waktu: {now_jakarta().strftime('%H:%M:%S WIB')}</p>",
    unsafe_allow_html=True
)

st.sidebar.markdown("<hr>", unsafe_allow_html=True)
UNI_LOGO = os.path.join(ASSETS_DIR, "Logo_Universitas_Brawijaya.svg.png")
if os.path.exists(UNI_LOGO):
    c1, c2 = st.sidebar.columns([1, 3])
    with c1:
        st.image(UNI_LOGO, width=40)
    with c2:
        st.markdown(
            "<p style='color:#ffffff; font-size:11px; margin:0; line-height:1.4;'>"
            "<b>Developed by</b><br>Universitas Brawijaya</p>", 
            unsafe_allow_html=True
        )
else:
    st.sidebar.markdown(
        "<p style='color:#ffffff; font-size:11px;'>Developed by Universitas Brawijaya</p>",
        unsafe_allow_html=True
    )