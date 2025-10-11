import os
import sys
import traceback
from typing import Optional, Callable
from datetime import datetime

import streamlit as st
import pandas as pd
from auth import get_gspread_client

# Timezone helper
try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    from datetime import timedelta
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

# Safe import of export module
THIS_DIR = os.path.dirname(__file__)
if THIS_DIR and THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

export_rekap_to_sheet: Optional[Callable] = None
HAVE_EXPORT = False
import_error_msg = None
try:
    import export_rekap_sheets as _export_mod
    export_rekap_to_sheet = getattr(_export_mod, "export_rekap_to_sheet", None)
    HAVE_EXPORT = callable(export_rekap_to_sheet)
except Exception:
    import_error_msg = traceback.format_exc()
    export_rekap_to_sheet = None
    HAVE_EXPORT = False

# Konfigurasi Google Sheet dari secrets
try:
    SPREADSHEET_ID = str(st.secrets["SHEET_ID"])
    GID = str(st.secrets["SHEET_GID"])
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

@st.cache_data(ttl=180, show_spinner=False)
def fetch_pelanggan_df(spreadsheet_id: str, gid: str) -> pd.DataFrame:
    ws = load_sheet_by_gid(spreadsheet_id, gid)
    data = ws.get_all_records()
    df = pd.DataFrame(data).fillna("")
    return df

# Load data pelanggan (cached)
df_sheets = fetch_pelanggan_df(SPREADSHEET_ID, GID)

# Siapkan mapping ID -> Nama
id_to_name = {}
if not df_sheets.empty and "ID Pelanggan" in df_sheets.columns:
    if "Nama" in df_sheets.columns:
        id_to_name = {
            str(row["ID Pelanggan"]): str(row.get("Nama", "-"))
            for _, row in df_sheets.iterrows()
            if str(row.get("ID Pelanggan", "")).strip() != ""
        }
    else:
        id_to_name = {
            str(row["ID Pelanggan"]): "-"
            for _, row in df_sheets.iterrows()
            if str(row.get("ID Pelanggan", "")).strip() != ""
        }

# Harga VENDOR (base price)
harga_vendor = {
    "Jasa Kegiatan Geser APP": 93000,
    "Jasa Kegiatan Geser Perubahan Situasi SR": 79000,
    "Service wedge clamp 2/4 x 6/10 mm": 3990,
    "Strainhook / ekor babi": 8000,
    "Imundex klem": 454,
    "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover": 11999,
    "Paku Beton": 74,
    "Pole Bracket 3-9\"": 36823,
    "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover": 29400,
    "Segel Plastik": 1754,
    "Twisted Cable 2 x 10 mm² - Al": 4339,
    "Asuransi": 0,
    "Twisted Cable 2x10 mm² - Al": 0,
}

# Harga PELANGGAN (1.11x dari vendor)
harga_pelanggan = {
    "Jasa Kegiatan Geser APP": 103230,
    "Jasa Kegiatan Geser Perubahan Situasi SR": 87690,
    "Service wedge clamp 2/4 x 6/10 mm": 4428.90,
    "Strainhook / ekor babi": 8880.00,
    "Imundex klem": 503.94,
    "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover": 13318.89,
    "Paku Beton": 82.14,
    "Pole Bracket 3-9\"": 40873.53,
    "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover": 32634.00,
    "Segel Plastik": 1946.94,
    "Twisted Cable 2 x 10 mm² - Al": 4816.29,
    "Asuransi": 0,
    "Twisted Cable 2x10 mm² - Al": 0,
}

# Data barang dengan harga PELANGGAN (untuk preview di website)
data_barang = [
    {"nama": "Jasa Kegiatan Geser APP", "SAT": "PLG", "harga": 103230},
    {"nama": "Jasa Kegiatan Geser Perubahan Situasi SR", "SAT": "PLG", "harga": 87690},
    {"nama": "Service wedge clamp 2/4 x 6/10 mm", "SAT": "B", "harga": 4428.90},
    {"nama": "Strainhook / ekor babi", "SAT": "B", "harga": 8880.00},
    {"nama": "Imundex klem", "SAT": "B", "harga": 503.94},
    {"nama": "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover", "SAT": "B", "harga": 13318.89},
    {"nama": "Paku Beton", "SAT": "B", "harga": 82.14},
    {"nama": "Pole Bracket 3-9\"", "SAT": "B", "harga": 40873.53},
    {"nama": "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover", "SAT": "B", "harga": 32634.00},
]
data_barang_tambahan = [
    {"nama": "Segel Plastik", "SAT": "B", "harga": 1946.94},
    {"nama": "Twisted Cable 2 x 10 mm² - Al", "SAT": "M", "harga": 4816.29},
    {"nama": "Asuransi", "harga": 0},
    {"nama": "Twisted Cable 2x10 mm² - Al", "SAT": "B", "harga": 0},
]
semua_barang = data_barang + [{"nama": "---- PEMBATAS ----", "SAT": "", "harga": 0}] + data_barang_tambahan

# Layout Streamlit
st.title("📋 Daftar Barang & Input Petugas")

# Filter: Tanggal + Search ID/Nama
st.subheader("🔎 Filter & Pilih Pelanggan")

# Konversi Timestamp ke Date
if "Timestamp" in df_sheets.columns:
    try:
        df_sheets["Date"] = pd.to_datetime(
            df_sheets["Timestamp"], 
            format="%d/%m/%Y %H:%M:%S",
            errors='coerce'
        ).dt.date
    except Exception:
        df_sheets["Date"] = pd.to_datetime(
            df_sheets["Timestamp"], 
            errors='coerce'
        ).dt.date

col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    if "Date" in df_sheets.columns:
        available_dates = df_sheets["Date"].dropna().unique()
        available_dates = sorted([d for d in available_dates if d], reverse=True)
        
        date_options = ["Semua Tanggal"] + [str(d) for d in available_dates]
        
        selected_date = st.selectbox(
            "📅 Filter Tanggal:",
            date_options,
            key="filter_date"
        )
    else:
        selected_date = "Semua Tanggal"
        st.info("Kolom Timestamp tidak ditemukan")

with col_filter2:
    search_text = st.text_input(
        "🔍 Cari IDPEL/Nama Pelanggan:",
        placeholder="Contoh: 513130665162 atau Sofia",
        key="filter_search"
    )

# Apply filters
df_filtered = df_sheets.copy()

if selected_date != "Semua Tanggal" and "Date" in df_sheets.columns:
    df_filtered = df_filtered[df_filtered["Date"].astype(str) == selected_date]

if search_text.strip():
    search_lower = search_text.strip().lower()
    mask_id = df_filtered["ID Pelanggan"].astype(str).str.lower().str.contains(search_lower, na=False)
    
    if "Nama" in df_filtered.columns:
        mask_nama = df_filtered["Nama"].astype(str).str.lower().str.contains(search_lower, na=False)
        df_filtered = df_filtered[mask_id | mask_nama]
    else:
        df_filtered = df_filtered[mask_id]

# Buat dropdown dari hasil filter
filtered_options = ["- Pilih ID -"]
if not df_filtered.empty:
    for _, row in df_filtered.iterrows():
        pid = str(row["ID Pelanggan"]).strip()
        pnama = str(row.get("Nama", "-")).strip()
        if pid:
            filtered_options.append(f"{pid} ({pnama})")
    
    result_count = len(filtered_options) - 1
    if result_count > 0:
        st.info(f"✅ Ditemukan **{result_count}** pelanggan yang sesuai filter")
    else:
        st.warning("⚠️ Tidak ada pelanggan yang cocok dengan filter. Coba ubah filter.")
else:
    st.warning("⚠️ Tidak ada pelanggan yang cocok dengan filter. Coba ubah filter.")

# Dropdown final
if len(filtered_options) > 1:
    pilihan_dropdown = st.selectbox(
        "🔑 Pilih ID Pelanggan:",
        filtered_options,
        key="select_idpel"
    )
else:
    pilihan_dropdown = "- Pilih ID -"
    st.info("💡 Silakan gunakan filter di atas untuk mencari pelanggan")

def extract_id(opt: str) -> str:
    if not opt or opt == "- Pilih ID -":
        return ""
    if " (" in opt:
        return opt.split(" (", 1)[0].strip()
    return opt.strip()

idpel_selected = extract_id(pilihan_dropdown)

# Layout 2 kolom: Data Pelanggan & Input Barang
col1, col2 = st.columns(2)

with col1:
    nama = "-"
    lokasi = "-"
    pekerjaan = ""
    ulp = ""
    no_spk = ""
    vendor = ""

    if idpel_selected:
        st.subheader("👤 Data Pelanggan Terpilih")
        df_selected = df_sheets[df_sheets["ID Pelanggan"].astype(str) == idpel_selected]
        if not df_selected.empty:
            first_row = df_selected.iloc[0]
            nama = str(first_row.get("Nama", "-"))
            lokasi = str(first_row.get("Alamat kWH Meter", "-"))
        else:
            nama = id_to_name.get(idpel_selected, "-")

        st.markdown(f"**NAMA:** {nama}")
        st.markdown(f"**LOKASI PEKERJAAN:** {lokasi}")

        pekerjaan = st.text_input("📌 Pekerjaan", key="pekerjaan_input")
        ulp = st.text_input("🏢 ULP", key="ulp_input")
        no_spk = st.text_input("📄 No SPK", key="no_spk_input")
        vendor = st.text_input("🏗 Vendor Pelaksana", key="vendor_input")
    else:
        st.info("Silakan pilih ID Pelanggan untuk melihat detail.")

# Input barang
barang_dipilih = []
with col2:
    st.subheader("🛠 Input Kuantitas Barang")
    with st.form("form_barang"):
        for idx, barang in enumerate(semua_barang):
            if str(barang.get("nama", "")).startswith("----"):
                st.markdown("---")
                continue

            key_name = f"qty_{idx}"
            sat_label = barang.get("SAT", "")
            qty = st.number_input(
                f"{barang.get('nama', 'Item')} ({sat_label})",
                min_value=0,
                step=1,
                key=key_name
            )
            if qty and qty > 0:
                harga = float(barang.get("harga", 0) or 0)
                total = qty * harga
                barang_dipilih.append({
                    "Rincian": barang.get("nama", ""),
                    "SAT": sat_label,
                    "Vol": int(qty),
                    "Harga Satuan Material": harga,
                    "Harga Total": total
                })
        submitted = st.form_submit_button("Hitung Rekap")

# Simpan hasil di session_state
if submitted:
    st.session_state["barang_dipilih"] = barang_dipilih
    st.session_state["show_preview"] = False
barang_dipilih = st.session_state.get("barang_dipilih", barang_dipilih)

# Rekapitulasi
st.subheader("📦 Rekapitulasi")
df_pilih = pd.DataFrame(barang_dipilih) if barang_dipilih else pd.DataFrame()

if not df_pilih.empty:
    st.markdown(f"**PEKERJAAN:** {pekerjaan or '-'}")
    st.markdown(f"**NAMA:** {nama} ({idpel_selected})")
    st.markdown(f"**LOKASI PEKERJAAN:** {lokasi}")
    st.markdown(f"**ULP:** {ulp or '-'}")
    st.markdown(f"**NO SPK:** {no_spk or '-'}")
    st.markdown(f"**VENDOR PELAKSANA:** {vendor or '-'}")

    st.write("---")
    st.dataframe(df_pilih, use_container_width=True)

    subtotal = df_pilih["Harga Total"].sum()
    ppn = subtotal * 0.11
    total_biaya = subtotal + ppn

    st.write(f"💰 **Subtotal:** Rp {subtotal:,.2f}")
    st.write(f"💸 **PPN (11%):** Rp {ppn:,.2f}")
    st.success(f"🏷 **TOTAL BIAYA SETELAH PPN: Rp {total_biaya:,.2f}**")
else:
    st.info("Belum ada barang yang dipilih (isi kuantitas > 0).")

# Tombol Export & Preview
st.markdown("---")
st.subheader("📤 Export Rekap ke Google Sheets")

if st.button("📥 Export ke Google Sheets", type="primary"):
    if not idpel_selected:
        st.error("⚠️ Silakan pilih ID Pelanggan terlebih dahulu!")
    elif df_pilih.empty:
        st.error("⚠️ Belum ada barang yang dipilih!")
    else:
        st.session_state["show_preview"] = True

# Preview Section
if st.session_state.get("show_preview", False):
    st.markdown("---")
    st.markdown("### 📋 Preview Rekap")
    
    # Prepare data untuk preview
    id_display = idpel_selected if idpel_selected else ""
    nama_dengan_id = f"{nama} ({id_display})" if id_display else f"{nama}"
    
    # Calculate for Vendor & Pelanggan
    df_preview_vendor = df_pilih.copy()
    df_preview_pelanggan = df_pilih.copy()
    
    # Update harga untuk preview vendor
    for i in range(len(df_preview_vendor)):
        item_name = df_preview_vendor.iloc[i]["Rincian"]
        qty = df_preview_vendor.iloc[i]["Vol"]
        harga_v = harga_vendor.get(item_name, 0)
        df_preview_vendor.loc[df_preview_vendor.index[i], "Harga Satuan Material"] = harga_v
        df_preview_vendor.loc[df_preview_vendor.index[i], "Harga Total"] = qty * harga_v
    
    subtotal_vendor = df_preview_vendor["Harga Total"].sum()
    ppn_vendor = subtotal_vendor * 0.11
    total_vendor = subtotal_vendor + ppn_vendor
    
    subtotal_pelanggan = df_pilih["Harga Total"].sum()
    ppn_pelanggan = subtotal_pelanggan * 0.11
    total_pelanggan = subtotal_pelanggan + ppn_pelanggan
    
    # Tabs
    tab1, tab2 = st.tabs(["📦 VENDOR", "👥 PELANGGAN"])
    
    with tab1:
        st.markdown("#### REKAP HARGA PEKERJAAN - VENDOR")
        st.markdown(f"**PEKERJAAN:** {pekerjaan or '-'}")
        st.markdown(f"**NAMA:** {nama_dengan_id}")
        st.markdown(f"**LOKASI:** {lokasi}")
        st.markdown(f"**ULP:** {ulp or '-'}")
        st.markdown(f"**NO SPK:** {no_spk or '-'}")
        st.markdown(f"**VENDOR PELAKSANA:** {vendor or '-'}")
        st.write("---")
        st.dataframe(df_preview_vendor[["Rincian", "SAT", "Vol", "Harga Satuan Material", "Harga Total"]], use_container_width=True)
        st.write(f"💰 **Subtotal:** Rp {subtotal_vendor:,.2f}")
        st.write(f"💸 **PPN (11%):** Rp {ppn_vendor:,.2f}")
        st.success(f"🏷 **TOTAL BIAYA: Rp {total_vendor:,.2f}**")
    
    with tab2:
        st.markdown("#### REKAP HARGA PEKERJAAN - PELANGGAN")
        st.markdown(f"**PEKERJAAN:** {pekerjaan or '-'}")
        st.markdown(f"**NAMA:** {nama_dengan_id}")
        st.markdown(f"**LOKASI:** {lokasi}")
        st.markdown(f"**ULP:** {ulp or '-'}")
        st.markdown(f"**NO SPK:** {no_spk or '-'}")
        st.markdown(f"**VENDOR PELAKSANA:** {vendor or '-'}")
        st.write("---")
        st.dataframe(df_pilih[["Rincian", "SAT", "Vol", "Harga Satuan Material", "Harga Total"]], use_container_width=True)
        st.write(f"💰 **Subtotal:** Rp {subtotal_pelanggan:,.2f}")
        st.write(f"💸 **PPN (11%):** Rp {ppn_pelanggan:,.2f}")
        st.success(f"🏷 **TOTAL BIAYA: Rp {total_pelanggan:,.2f}**")
    
    # Action buttons
    st.write("---")
    col_btn1, col_btn2 = st.columns([1, 2])
    
    with col_btn1:
        if st.button("🚫 Batal", use_container_width=True):
            st.session_state["show_preview"] = False
            st.rerun()
    
    with col_btn2:
        if st.button("✅ Konfirmasi & Export", type="primary", use_container_width=True):
            meta = {
                "Pekerjaan": pekerjaan or "-",
                "Nama": nama_dengan_id or "-",
                "Lokasi": lokasi or "-",
                "ULP": ulp or "-",
                "No SPK": no_spk or "-",
                "Vendor": vendor or "-"
            }
            
            now = now_jakarta().strftime("%Y%m%d_%H%M")
            safe_name = str(nama).replace("/", "-").replace("\\", "-")
            title_vendor = f"REKAP {safe_name} - {now}_Vendor"
            title_pelanggan = f"REKAP {safe_name} - {now}_Pelanggan"
            
            with st.spinner("Menulis dua rekapan (Vendor & Pelanggan) ke Google Sheets..."):
                try:
                    from export_rekap_sheets import export_rekap_pair
                    pair_info = export_rekap_pair(
                        spreadsheet_id=SPREADSHEET_ID,
                        base_sheet_title_vendor=title_vendor,
                        base_sheet_title_pelanggan=title_pelanggan,
                        meta=meta,
                        df_pilih=df_pilih,
                        idpel=idpel_selected,
                        gid=GID,
                    )
                    
                    st.session_state["show_preview"] = False
                    
                    st.success(
                        f"✅ Berhasil membuat: **{pair_info['vendor']['sheet_title']}** dan "
                        f"**{pair_info['pelanggan']['sheet_title']}**"
                    )
                    
                    survey_result = pair_info.get("survey_result", {})
                    if survey_result.get("success", False):
                        st.info(f"📅 {survey_result.get('message', 'Tanggal Survey berhasil diperbarui')}")
                    else:
                        st.warning(f"⚠️ Tanggal Survey gagal diperbarui: {survey_result.get('message', 'Unknown error')}")
                    
                    st.balloons()
                except Exception as e:
                    st.error(f"❌ Gagal mengekspor: {e}")
                    import traceback
                    st.error(traceback.format_exc())