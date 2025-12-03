import streamlit as st
import os
import sys
import importlib.util
from PIL import Image
from datetime import datetime

# Set up the initial page configuration and layout settings.
st.set_page_config(
    page_title="Permohonan Geser Meter",
    page_icon="assets/logo_pln.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize the user authentication state if not already set.
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_login():
    # Validate the provided credentials against the stored secrets.
    user_input = st.session_state.get("input_username", "")
    pass_input = st.session_state.get("input_password", "")
    
    try:
        correct_user = st.secrets["auth"]["username"]
        correct_pass = st.secrets["auth"]["password"]
        
        if user_input == correct_user and pass_input == correct_pass:
            st.session_state.authenticated = True
        else:
            st.error("Username atau Password tidak valid.")
    except Exception:
        st.error("Konfigurasi autentikasi belum diatur dalam secrets.")

def logout():
    # Clear the session state and reload the application.
    st.session_state.authenticated = False
    st.rerun()

# Render the login form if the user is not authenticated.
if not st.session_state.authenticated:
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center; color: #1e3a5f;'>Login Petugas</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>PLN ULP DINOYO</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.text_input("Username", key="input_username")
            st.text_input("Password", type="password", key="input_password")
            st.form_submit_button("Masuk", type="primary", use_container_width=True, on_click=check_login)
            
    st.stop()

# Create a helper function to get the current time in Jakarta.
try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    def now_jakarta():
        return datetime.now()

# Define the base paths for assets and module directories.
BASE_DIR = os.path.dirname(__file__)
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
SIDEBAR_DIR = os.path.join(BASE_DIR, "sidebar")

# Inject custom CSS to style the sidebar and hide specific elements.
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
        margin-bottom: 0.5rem !important;
    }
    
    /* Sidebar Background & Text */
    [data-testid="stSidebar"] {
        background-color: #1e3a5f;
    }
    [data-testid="stSidebar"] * {
        color: #ffffff !important;
    }
    
    /* Adjust Sidebar Top Padding to align Logo with Collapse Button (<<) */
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 0rem;
    }
    
    /* Sidebar Selectbox */
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: #2c5282 !important;
        border-color: #4a6fa5 !important;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] span {
        color: #ffffff !important;
    }
    
    /* Sidebar Links */
    [data-testid="stSidebar"] a {
        color: #87ceeb !important;
    }
    
    /* Sidebar Divider */
    [data-testid="stSidebar"] hr {
        margin-top: 0.8rem;
        margin-bottom: 0.8rem;
        border-color: #4a6fa5;
    }

    /* Style for the secondary button in the sidebar. */
    [data-testid="stSidebar"] button[kind="secondary"] {
        background-color: #2c5282 !important;
        color: #ffffff !important;
        border: 1px solid #4a6fa5 !important;
        transition: all 0.3s ease;
        margin-top: 5px;
    }
    [data-testid="stSidebar"] button[kind="secondary"]:hover {
        background-color: #3d6aa3 !important;
        border-color: #ffd700 !important;
        color: #ffffff !important;
    }

    /* Style for the primary button in the sidebar. */
    [data-testid="stSidebar"] button[kind="primary"] {
        background-color: #c53030 !important;
        color: #ffffff !important;
        border: 1px solid #e53e3e !important;
        margin-top: 5px;
    }
    [data-testid="stSidebar"] button[kind="primary"]:hover {
        background-color: #e53e3e !important;
        border-color: #ffd700 !important;
    }
    
    /* Ensure text inside buttons remains white. */
    [data-testid="stSidebar"] button p {
        color: #ffffff !important;
    }

</style>
""", unsafe_allow_html=True)

# Display the agency logo and identity in the sidebar.
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

# Render a link to the office location on Google Maps.
st.sidebar.markdown(
    "<a href='https://maps.app.goo.gl/CnhdCBrhz3mihieL9' "
    "style='text-decoration:none; color:#87ceeb !important; font-size:12px; display:block; margin-top:8px;' "
    "target='_blank'>Jl. Pandan No.15, Gading Kasri, Klojen</a>",
    unsafe_allow_html=True
)

st.sidebar.markdown("<hr>", unsafe_allow_html=True)

# Create a selection box for navigating between application pages.
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

# Render the stacked action buttons for data reload and logout.
st.sidebar.markdown("<p style='font-size:12px; font-weight:bold; margin-bottom:5px; color:#87ceeb;'>Kontrol Aplikasi</p>", unsafe_allow_html=True)

if st.sidebar.button("Reload Data Harga", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

if st.sidebar.button("Logout", use_container_width=True, type="primary", on_click=logout):
    pass

# Dynamically load and execute the selected page module.
page_module = pages.get(choice)

if page_module:
    try:
        module_path = os.path.join(SIDEBAR_DIR, f"{page_module}.py")
        
        if os.path.exists(module_path):
            spec = importlib.util.spec_from_file_location(page_module, module_path)
            if spec is not None and spec.loader is not None:
                module = importlib.util.module_from_spec(spec)
                sys.modules[page_module] = module
                spec.loader.exec_module(module)
            else:
                st.error(f"Gagal memuat modul {page_module}.py")
        else:
            st.error(f"File {page_module}.py tidak ditemukan")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memuat halaman: {str(e)}")
else:
    st.error("Halaman tidak valid.")

st.sidebar.markdown("<hr>", unsafe_allow_html=True)

# Show the current server date and time information.
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

# Render the footer containing the university logo and credits.
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