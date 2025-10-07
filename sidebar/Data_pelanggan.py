import streamlit as st
import pandas as pd
import altair as alt
from auth import get_gspread_client

st.set_page_config(page_title="Data dari Google Sheets", layout="wide")

# Ambil dari secrets
try:
    SPREADSHEET_ID = str(st.secrets["SHEET_ID"])
    GID = str(st.secrets["SHEET_GID"])  # simpan sebagai string untuk perbandingan
except Exception as e:
    st.error(f"Konfigurasi secrets tidak lengkap: {e}")
    st.stop()

def load_sheet_by_gid(spreadsheet_id, gid):
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    target = None
    for ws in sh.worksheets():
        if str(ws.id) == str(gid):
            target = ws
            break
    if target is None:
        target = sh.sheet1
    return target

# Cache 3 menit agar tidak fetch berulang saat rerun
@st.cache_data(ttl=180, show_spinner=False)
def fetch_df(spreadsheet_id, gid) -> pd.DataFrame:
    ws = load_sheet_by_gid(spreadsheet_id, gid)
    data = ws.get_all_records()
    return pd.DataFrame(data)

try:
    df = fetch_df(SPREADSHEET_ID, GID)
except Exception as e:
    st.error(f"Gagal mengambil data dari Google Sheets: {e}")
    df = pd.DataFrame()

if not df.empty:
    # --- Tambahkan blok aman Arrow di bawah ini ---
    df = df.fillna("")
    # khusus kolom yang bermasalah
    if "Tarif / Daya" in df.columns:
        df["Tarif / Daya"] = df["Tarif / Daya"].astype(str)

    # (opsional kuat) jadikan semua kolom object -> string supaya aman di Arrow
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].astype(str)
    # --- sampai sini ---

    if "Foto KTP" in df.columns:
        def ktp_link(v):
            v = str(v).strip()
            if v and v.lower() not in ["nan", "none", ""]:
                return f'<a href="{v}" target="_blank">📷 Lihat KTP</a>'
            return ""
        df["Foto KTP"] = df["Foto KTP"].apply(ktp_link)

    if "Tarif / Daya" in df.columns:
        daya_count = df["Tarif / Daya"].value_counts().reset_index()
        daya_count.columns = ["Tarif / Daya", "Jumlah Pengguna"]

        st.subheader("📈 Jumlah Pengguna berdasarkan Daya")
        chart = (
            alt.Chart(data=daya_count)
            .mark_bar()
            .encode(
                x=alt.X("Tarif / Daya:N", title="Tarif / Daya"),
                y=alt.Y("Jumlah Pengguna:Q", title="Jumlah Pengguna"),
                tooltip=["Tarif / Daya", "Jumlah Pengguna"]
            )
        )
        st.altair_chart(chart, use_container_width=True)
    else:
        st.warning("Kolom 'Tarif / Daya' tidak ditemukan dalam data.")

    st.subheader("📊 Data dari Google Sheets")
    st.write(df.to_html(escape=False, index=False), unsafe_allow_html=True)
else:
    st.info("Belum ada data untuk ditampilkan.")