import os
import sys
import traceback
from typing import Optional, Callable, Dict
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
    MASTER_HARGA_SHEET = str(st.secrets.get("MASTER_HARGA_SHEET", "Harga"))
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

def load_sheet_by_name(spreadsheet_id, sheet_name):
    """Load worksheet by name"""
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(sheet_name)
    except Exception:
        return None

@st.cache_data(ttl=180, show_spinner=False)
def fetch_pelanggan_df(spreadsheet_id: str, gid: str) -> pd.DataFrame:
    ws = load_sheet_by_gid(spreadsheet_id, gid)
    data = ws.get_all_records()
    df = pd.DataFrame(data).fillna("")
    return df

@st.cache_data(ttl=300, show_spinner=False)
def fetch_master_harga(spreadsheet_id: str, sheet_name: str):
    """
    Fetch harga dari Google Sheets.
    Returns: (harga_vendor, harga_pelanggan, is_from_sheets)
    """
    harga_vendor_fallback = {
        "Jasa Kegiatan Geser APP": 93000.0,
        "Jasa Kegiatan Geser Perubahan Situasi SR": 79000.0,
        "Service wedge clamp 2/4 x 6/10 mm": 3990.0,
        "Strainhook / ekor babi": 8000.0,
        "Imundex klem": 454.0,
        "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover": 11999.0,
        "Paku Beton": 74.0,
        "Pole Bracket 3-9\"": 36823.0,
        "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover": 29400.0,
        "Segel Plastik": 1754.0,
        "Twisted Cable 2 x 10 mm² - Al": 4339.0,
        "Asuransi": 0.0,
        "Twisted Cable 2x10 mm² - Al": 0.0,
    }
    
    harga_pelanggan_fallback = {
        "Jasa Kegiatan Geser APP": 103230.0,
        "Jasa Kegiatan Geser Perubahan Situasi SR": 87690.0,
        "Service wedge clamp 2/4 x 6/10 mm": 4428.9,
        "Strainhook / ekor babi": 8880.0,
        "Imundex klem": 503.94,
        "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover": 13318.89,
        "Paku Beton": 82.14,
        "Pole Bracket 3-9\"": 40873.53,
        "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover": 32634.0,
        "Segel Plastik": 1946.94,
        "Twisted Cable 2 x 10 mm² - Al": 4816.29,
        "Asuransi": 0.0,
        "Twisted Cable 2x10 mm² - Al": 0.0,
    }
    
    try:
        ws = load_sheet_by_name(spreadsheet_id, sheet_name)
        if ws is None:
            return harga_vendor_fallback, harga_pelanggan_fallback, False
        
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        
        # Validasi kolom yang diperlukan
        required_cols = ["Nama Barang", "Harga Vendor", "Harga Pelanggan"]
        if not all(col in df.columns for col in required_cols):
            return harga_vendor_fallback, harga_pelanggan_fallback, False
        
        # Build dictionaries
        harga_vendor = {}
        harga_pelanggan = {}
        
        for _, row in df.iterrows():
            nama = str(row["Nama Barang"]).strip()
            if not nama:
                continue
            
            try:
                harga_v = float(row["Harga Vendor"]) if pd.notna(row["Harga Vendor"]) else 0.0
                harga_p = float(row["Harga Pelanggan"]) if pd.notna(row["Harga Pelanggan"]) else 0.0
                harga_vendor[nama] = harga_v
                harga_pelanggan[nama] = harga_p
            except (ValueError, TypeError):
                continue
        
        # Jika berhasil load data, return
        if harga_vendor and harga_pelanggan:
            return harga_vendor, harga_pelanggan, True
        
        # Jika kosong, fallback
        return harga_vendor_fallback, harga_pelanggan_fallback, False
        
    except Exception:
        # Jika error, fallback
        return harga_vendor_fallback, harga_pelanggan_fallback, False

# Load data pelanggan (cached)
df_sheets = fetch_pelanggan_df(SPREADSHEET_ID, GID)

# Load harga dari sheets
harga_vendor, harga_pelanggan, is_from_sheets = fetch_master_harga(SPREADSHEET_ID, MASTER_HARGA_SHEET)

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

# Build data_barang dari harga_pelanggan
data_barang = []
data_barang_tambahan = []

# Mapping SAT untuk setiap barang
sat_mapping = {
    "Jasa Kegiatan Geser APP": "PLG",
    "Jasa Kegiatan Geser Perubahan Situasi SR": "PLG",
    "Service wedge clamp 2/4 x 6/10 mm": "B",
    "Strainhook / ekor babi": "B",
    "Imundex klem": "B",
    "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover": "B",
    "Paku Beton": "B",
    "Pole Bracket 3-9\"": "B",
    "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover": "B",
    "Segel Plastik": "B",
    "Twisted Cable 2 x 10 mm² - Al": "M",
    "Asuransi": "I",
    "Twisted Cable 2x10 mm² - Al": "B",
}

# Urutan barang utama (9 items)
main_items = [
    "Jasa Kegiatan Geser APP",
    "Jasa Kegiatan Geser Perubahan Situasi SR",
    "Service wedge clamp 2/4 x 6/10 mm",
    "Strainhook / ekor babi",
    "Imundex klem",
    "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover",
    "Paku Beton",
    "Pole Bracket 3-9\"",
    "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover",
]

# Barang tambahan (4 items)
additional_items = [
    "Segel Plastik",
    "Twisted Cable 2 x 10 mm² - Al",
    "Asuransi",
    "Twisted Cable 2x10 mm² - Al",
]

# Build data_barang
for nama in main_items:
    harga = harga_pelanggan.get(nama, 0)
    sat = sat_mapping.get(nama, "")
    data_barang.append({"nama": nama, "SAT": sat, "harga": harga})

for nama in additional_items:
    harga = harga_pelanggan.get(nama, 0)
    sat = sat_mapping.get(nama, "")
    data_barang_tambahan.append({"nama": nama, "SAT": sat, "harga": harga})

semua_barang = data_barang + [{"nama": "---- PEMBATAS ----", "SAT": "", "harga": 0}] + data_barang_tambahan

# Dialog untuk preview
@st.dialog("Preview Rekap", width="large")
def show_preview_dialog(barang_dipilih, nama, idpel_selected, lokasi, pekerjaan, ulp, no_spk, vendor):
    if not barang_dipilih:
        st.warning("Tidak ada barang yang dipilih.")
        return
    
    df_pilih = pd.DataFrame(barang_dipilih)
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
    tab1, tab2 = st.tabs(["VENDOR", "PELANGGAN"])
    
    with tab1:
        st.markdown("#### REKAP HARGA PEKERJAAN - VENDOR")
        st.markdown(f"**PEKERJAAN:** {pekerjaan or '-'}")
        st.markdown(f"**NAMA:** {nama_dengan_id}")
        st.markdown(f"**LOKASI:** {lokasi}")
        st.markdown(f"**ULP:** {ulp or '-'}")
        st.markdown(f"**NO SPK:** {no_spk or '-'}")
        st.markdown(f"**VENDOR PELAKSANA:** {vendor or '-'}")
        st.write("---")
        st.dataframe(df_preview_vendor[["Rincian", "SAT", "Vol", "Harga Satuan Material", "Harga Total"]], use_container_width=True, hide_index=True)
        st.write(f"**Subtotal:** Rp {subtotal_vendor:,.2f}")
        st.write(f"**PPN (11%):** Rp {ppn_vendor:,.2f}")
        st.success(f"**TOTAL BIAYA: Rp {total_vendor:,.2f}**")
    
    with tab2:
        st.markdown("#### REKAP HARGA PEKERJAAN - PELANGGAN")
        st.markdown(f"**PEKERJAAN:** {pekerjaan or '-'}")
        st.markdown(f"**NAMA:** {nama_dengan_id}")
        st.markdown(f"**LOKASI:** {lokasi}")
        st.markdown(f"**ULP:** {ulp or '-'}")
        st.markdown(f"**NO SPK:** {no_spk or '-'}")
        st.markdown(f"**VENDOR PELAKSANA:** {vendor or '-'}")
        st.write("---")
        st.dataframe(df_pilih[["Rincian", "SAT", "Vol", "Harga Satuan Material", "Harga Total"]], use_container_width=True, hide_index=True)
        st.write(f"**Subtotal:** Rp {subtotal_pelanggan:,.2f}")
        st.write(f"**PPN (11%):** Rp {ppn_pelanggan:,.2f}")
        st.success(f"**TOTAL BIAYA: Rp {total_pelanggan:,.2f}**")
    
    # Action buttons
    st.write("---")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
    
    with col_btn1:
        if st.button("Batal", use_container_width=True, key="btn_cancel"):
            st.rerun()
    
    with col_btn3:
        if st.button("Konfirmasi & Export", type="primary", use_container_width=True, key="btn_export"):
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
            
            with st.spinner("Menulis data rekap ke Google Sheets..."):
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
                        f"Berhasil membuat rekap: **{pair_info['vendor']['sheet_title']}** dan "
                        f"**{pair_info['pelanggan']['sheet_title']}**"
                    )
                    
                    survey_result = pair_info.get("survey_result", {})
                    if survey_result.get("success", False):
                        st.info(f"Tanggal Survey: {survey_result.get('message', 'Berhasil diperbarui')}")
                    else:
                        st.warning(f"Tanggal Survey gagal diperbarui: {survey_result.get('message', 'Unknown error')}")
                    
                    st.balloons()
                    
                    # Wait a bit then close dialog
                    import time
                    time.sleep(2)
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal mengekspor data: {e}")
                    import traceback
                    st.error(traceback.format_exc())

# Layout Streamlit
st.title("Daftar Barang & Input Petugas")

# Info sumber harga
if is_from_sheets:
    st.info(f"Harga berhasil dimuat dari sheet '{MASTER_HARGA_SHEET}'. Data akan diperbarui otomatis setiap 5 menit.")
else:
    st.warning(f"Harga menggunakan data fallback (hardcoded). Pastikan sheet '{MASTER_HARGA_SHEET}' tersedia dengan kolom: Nama Barang, Harga Vendor, Harga Pelanggan.")

# Filter: Tanggal + Search ID/Nama
st.subheader("Filter & Pilih Pelanggan")

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
            "Filter Tanggal:",
            date_options,
            key="filter_date"
        )
    else:
        selected_date = "Semua Tanggal"
        st.info("Kolom Timestamp tidak ditemukan")

with col_filter2:
    search_text = st.text_input(
        "Cari ID Pelanggan atau Nama:",
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
filtered_options = ["- Pilih ID Pelanggan -"]
if not df_filtered.empty:
    for _, row in df_filtered.iterrows():
        pid = str(row["ID Pelanggan"]).strip()
        pnama = str(row.get("Nama", "-")).strip()
        if pid:
            filtered_options.append(f"{pid} ({pnama})")
    
    result_count = len(filtered_options) - 1
    if result_count > 0:
        st.info(f"Ditemukan {result_count} pelanggan yang sesuai filter")
    else:
        st.warning("Tidak ada pelanggan yang cocok dengan filter. Silakan ubah filter.")
else:
    st.warning("Tidak ada pelanggan yang cocok dengan filter. Silakan ubah filter.")

# Dropdown final
if len(filtered_options) > 1:
    pilihan_dropdown = st.selectbox(
        "Pilih ID Pelanggan:",
        filtered_options,
        key="select_idpel"
    )
else:
    pilihan_dropdown = "- Pilih ID Pelanggan -"
    st.info("Silakan gunakan filter di atas untuk mencari pelanggan")

def extract_id(opt: str) -> str:
    if not opt or opt == "- Pilih ID Pelanggan -":
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
        st.subheader("Data Pelanggan Terpilih")
        df_selected = df_sheets[df_sheets["ID Pelanggan"].astype(str) == idpel_selected]
        if not df_selected.empty:
            first_row = df_selected.iloc[0]
            nama = str(first_row.get("Nama", "-"))
            lokasi = str(first_row.get("Alamat kWH Meter", "-"))
        else:
            nama = id_to_name.get(idpel_selected, "-")

        st.markdown(f"**NAMA:** {nama}")
        st.markdown(f"**LOKASI PEKERJAAN:** {lokasi}")

        pekerjaan = st.text_input("Pekerjaan", key="pekerjaan_input")
        ulp = st.text_input("ULP", key="ulp_input")
        no_spk = st.text_input("No SPK", key="no_spk_input")
        vendor = st.text_input("Vendor Pelaksana", key="vendor_input")
    else:
        st.info("Silakan pilih ID Pelanggan untuk melihat detail.")

# Input barang
barang_dipilih = []
with col2:
    st.subheader("Input Kuantitas Barang")
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

# Tombol Export
st.markdown("---")

if st.button("Export ke Google Sheets", type="primary", use_container_width=True):
    if not idpel_selected:
        st.error("Silakan pilih ID Pelanggan terlebih dahulu.")
    elif not barang_dipilih:
        st.error("Belum ada barang yang dipilih.")
    else:
        show_preview_dialog(barang_dipilih, nama, idpel_selected, lokasi, pekerjaan, ulp, no_spk, vendor)