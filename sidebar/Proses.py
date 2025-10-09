import os
import sys
import traceback
from typing import Optional, Callable
from datetime import datetime

import streamlit as st
import pandas as pd
from auth import get_gspread_client

# ==============================
#  TIMEZONE HELPER - NEW
# ==============================
try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    # Fallback jika zoneinfo tidak tersedia
    from datetime import timedelta
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

# --------------------------
# Safe import of export module (dari folder 'sidebar')
# --------------------------
THIS_DIR = os.path.dirname(__file__)
if THIS_DIR and THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

export_rekap_to_sheet: Optional[Callable] = None
HAVE_EXPORT = False
import_error_msg = None
try:
    import export_rekap_sheets as _export_mod  # type: ignore
    export_rekap_to_sheet = getattr(_export_mod, "export_rekap_to_sheet", None)
    HAVE_EXPORT = callable(export_rekap_to_sheet)
except Exception:
    import_error_msg = traceback.format_exc()
    export_rekap_to_sheet = None
    HAVE_EXPORT = False

# === Konfigurasi Google Sheet dari secrets ===
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

# -------- Cache data Google Sheets supaya hemat kuota --------
@st.cache_data(ttl=180, show_spinner=False)
def fetch_pelanggan_df(spreadsheet_id: str, gid: str) -> pd.DataFrame:
    """
    Ambil data pelanggan sekali, cache 3 menit.
    Menggunakan get_all_records agar aman ke struktur kolom saat ini.
    """
    ws = load_sheet_by_gid(spreadsheet_id, gid)
    data = ws.get_all_records()
    df = pd.DataFrame(data).fillna("")
    return df

# Load data pelanggan (cached)
df_sheets = fetch_pelanggan_df(SPREADSHEET_ID, GID)

# Siapkan mapping ID -> Nama (untuk backward compatibility jika perlu)
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

# Data barang dengan label SAT (B/M/PLG) - FINAL UPDATE
data_barang = [
    {"nama": "Jasa Kegiatan Geser APP", "SAT": "PLG", "harga": 103230},
    {"nama": "Jasa Kegiatan Geser Perubahan Situasi SR", "SAT": "PLG", "harga": 87690},
    {"nama": "Service wedge clamp 2/4 x 6/10 mm", "SAT": "B", "harga": 4429},  # ✅ DENGAN SPASI
    {"nama": "Strainhook / ekor babi", "SAT": "B", "harga": 8880},  # ✅ huruf kecil "ekor"
    {"nama": "Imundex klem", "SAT": "B", "harga": 504},  # ✅ huruf kecil "klem"
    {"nama": "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover", "SAT": "B", "harga": 13319},
    {"nama": "Paku Beton", "SAT": "B", "harga": 82},
    {"nama": "Pole Bracket 3-9\"", "SAT": "B", "harga": 40874},
    {"nama": "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover", "SAT": "B", "harga": 32634},
]
data_barang_tambahan = [
    {"nama": "Segel Plastik", "SAT": "B", "harga": 1947},
    {"nama": "Twisted Cable 2 x 10 mm² - Al", "SAT": "M", "harga": 4816},  # ✅ DENGAN SPASI, strip biasa
    {"nama": "Asuransi", "harga": 0},
    {"nama": "Twisted Cable 2x10 mm² - Al", "SAT": "B", "harga": 0},  # ✅ TANPA SPASI
]
semua_barang = data_barang + [{"nama": "---- PEMBATAS ----", "SAT": "", "harga": 0}] + data_barang_tambahan

# === Layout Streamlit ===
st.title("📋 Daftar Barang & Input Petugas")

# ============================================================
# FILTER: TANGGAL + SEARCH ID/NAMA
# ============================================================
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
        # Fallback jika format berbeda
        df_sheets["Date"] = pd.to_datetime(
            df_sheets["Timestamp"], 
            errors='coerce'
        ).dt.date

col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    # Filter Tanggal
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
    # Search ID atau Nama
    search_text = st.text_input(
        "🔍 Cari IDPEL/Nama Pelanggan:",
        placeholder="Contoh: 513130665162 atau Sofia",
        key="filter_search"
    )

# Apply filters
df_filtered = df_sheets.copy()

# Filter berdasarkan tanggal
if selected_date != "Semua Tanggal" and "Date" in df_sheets.columns:
    df_filtered = df_filtered[df_filtered["Date"].astype(str) == selected_date]

# Filter berdasarkan search text (ID atau Nama)
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
    
    # Info jumlah hasil
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

# Extract ID dari pilihan
def extract_id(opt: str) -> str:
    if not opt or opt == "- Pilih ID -":
        return ""
    if " (" in opt:
        return opt.split(" (", 1)[0].strip()
    return opt.strip()

idpel_selected = extract_id(pilihan_dropdown)

# ============================================================
# LAYOUT 2 KOLOM: DATA PELANGGAN & INPUT BARANG
# ============================================================
col1, col2 = st.columns(2)

with col1:
    # Inisialisasi default
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

# ------- Input barang (gunakan form agar tidak rerun tiap klik) -------
barang_dipilih = []
with col2:
    st.subheader("🛠 Input Kuantitas Barang")
    with st.form("form_barang"):
        for idx, barang in enumerate(semua_barang):
            if str(barang.get("nama", "")).startswith("----"):
                st.markdown("---")
                continue

            key_name = f"qty_{idx}"
            sat_label = barang.get("SAT", "")  # aman bila tidak ada
            qty = st.number_input(
                f"{barang.get('nama', 'Item')} ({sat_label})",
                min_value=0,
                step=1,
                key=key_name
            )
            if qty and qty > 0:
                harga = int(barang.get("harga", 0) or 0)
                total = int(qty) * harga
                barang_dipilih.append({
                    "Rincian": barang.get("nama", ""),
                    "SAT": sat_label,
                    "Vol": int(qty),
                    "Harga Satuan Material": harga,
                    "Harga Total": total
                })
        submitted = st.form_submit_button("Hitung Rekap")

# Simpan hasil terakhir di session_state agar tidak hilang saat rerun lain
if submitted:
    st.session_state["barang_dipilih"] = barang_dipilih
barang_dipilih = st.session_state.get("barang_dipilih", barang_dipilih)

# ========================
# Rekapitulasi
# ========================
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

    st.write(f"💰 **Subtotal:** Rp {subtotal:,.0f}")
    st.write(f"💸 **PPN (11%):** Rp {ppn:,.0f}")
    st.success(f"🏷 **TOTAL BIAYA SETELAH PPN: Rp {total_biaya:,.0f}**")
else:
    st.info("Belum ada barang yang dipilih (isi kuantitas > 0).")

# ========================
# Tombol Export Selalu Ada (dua sheet: Vendor & Pelanggan)
# ========================
st.markdown("---")
st.subheader("📤 Export Rekap ke Google Sheets")

# ✅ FIX: Gunakan now_jakarta() untuk timestamp yang benar
now = now_jakarta().strftime("%Y%m%d_%H%M")
safe_name = str(nama).replace("/", "-").replace("\\", "-")

title_vendor    = f"REKAP {safe_name} - {now}_Vendor"
title_pelanggan = f"REKAP {safe_name} - {now}_Pelanggan"

id_display = idpel_selected if idpel_selected else ""
nama_dengan_id = f"{nama} ({id_display})" if id_display else f"{nama}"

if st.button("📥 Export ke Google Sheets"):
    if not idpel_selected:
        st.error("⚠️ Silakan pilih ID Pelanggan terlebih dahulu!")
    elif df_pilih.empty:
        st.error("⚠️ Belum ada barang yang dipilih!")
    else:
        meta = {
            "Pekerjaan": pekerjaan or "-",
            "Nama": nama_dengan_id or "-",
            "Lokasi": lokasi or "-",
            "ULP": ulp or "-",
            "No SPK": no_spk or "-",
            "Vendor": vendor or "-"
        }
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
                
                st.success(
                    f"✅ Berhasil membuat: **{pair_info['vendor']['sheet_title']}** dan "
                    f"**{pair_info['pelanggan']['sheet_title']}**"
                )
                st.write(
                    f"Subtotal: Rp {pair_info['vendor'].get('subtotal', 0):,} · "
                    f"PPN: Rp {pair_info['vendor'].get('ppn', 0):,} · "
                    f"Total: Rp {pair_info['vendor'].get('total_after_ppn', 0):,}"
                )
                
                # Detail feedback update Tanggal Survey
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