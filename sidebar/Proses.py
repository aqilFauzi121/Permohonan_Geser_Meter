import os
import sys
import time
import traceback
from datetime import datetime, date

import streamlit as st
import pandas as pd
from auth import get_gspread_client, get_or_create_folder, upload_file_to_drive

try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    from datetime import timedelta
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

THIS_DIR = os.path.dirname(__file__)
if THIS_DIR and THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)

try:
    import draft_manager  # type: ignore
    HAVE_DRAFT_MANAGER = True
except Exception as e:
    draft_manager = None  # type: ignore
    HAVE_DRAFT_MANAGER = False
    st.error(f"Error importing draft_manager: {e}")

try:
    import export_rekap_sheets as _export_mod
    export_rekap_to_sheet = getattr(_export_mod, "export_rekap_to_sheet", None)
    HAVE_EXPORT = callable(export_rekap_to_sheet)
except Exception:
    export_rekap_to_sheet = None
    HAVE_EXPORT = False

try:
    SPREADSHEET_ID = str(st.secrets["SHEET_ID"])
    GID = str(st.secrets["SHEET_GID"])
    MASTER_HARGA_SHEET = str(st.secrets.get("MASTER_HARGA_SHEET", "Harga"))
    DRIVE_FOLDER_SURVEY = str(st.secrets.get("DRIVE_FOLDER_SURVEY", ""))
    
    if not DRIVE_FOLDER_SURVEY:
        st.error("DRIVE_FOLDER_SURVEY tidak diset di secrets!")
        st.stop()
        
except Exception as e:
    st.error(f"Konfigurasi secrets tidak lengkap: {e}")
    st.stop()


def format_rupiah(nilai):
    if pd.isna(nilai) or nilai == 0:
        return "0"
    
    nilai = float(nilai)
    is_negative = nilai < 0
    nilai = abs(nilai)
    
    bagian_bulat = int(nilai)
    bagian_desimal = nilai - bagian_bulat
    
    bulat_str = f"{bagian_bulat:,}".replace(",", ".")
    
    if bagian_desimal > 0.001:
        desimal_str = f"{bagian_desimal:.2f}".split(".")[1]
        hasil = f"{bulat_str},{desimal_str}"
    else:
        hasil = bulat_str
    
    if is_negative:
        hasil = f"-{hasil}"
    
    return hasil


def parse_number_from_sheets(value):
    if pd.isna(value) or value == "" or value is None:
        return 0.0
    
    if isinstance(value, (int, float)):
        return float(value)
    
    value_str = str(value).strip()
    
    if not value_str:
        return 0.0
    
    value_str = value_str.replace(".", "")
    value_str = value_str.replace(",", ".")
    
    try:
        return float(value_str)
    except (ValueError, TypeError):
        return 0.0


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
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(sheet_name)
    except Exception:
        return None


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_pelanggan_df(spreadsheet_id: str, gid: str) -> pd.DataFrame:
    ws = load_sheet_by_gid(spreadsheet_id, gid)
    data = ws.get_all_records()
    df = pd.DataFrame(data).fillna("")
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_master_harga(spreadsheet_id: str, sheet_name: str):
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
    }
    
    try:
        ws = load_sheet_by_name(spreadsheet_id, sheet_name)
        if ws is None:
            return harga_vendor_fallback, harga_pelanggan_fallback, False
        
        all_values = ws.get_all_values()
        
        if len(all_values) < 2:
            return harga_vendor_fallback, harga_pelanggan_fallback, False
        
        headers = all_values[0]
        
        nama_col_idx = None
        vendor_col_idx = None
        pelanggan_col_idx = None
        
        for idx, h in enumerate(headers):
            h_lower = str(h).strip().lower()
            if "nama barang" in h_lower or "nama" in h_lower:
                nama_col_idx = idx
            elif "vendor" in h_lower and "harga" in h_lower:
                vendor_col_idx = idx
            elif "pelanggan" in h_lower and "harga" in h_lower:
                pelanggan_col_idx = idx
        
        if nama_col_idx is None or vendor_col_idx is None or pelanggan_col_idx is None:
            return harga_vendor_fallback, harga_pelanggan_fallback, False
        
        harga_vendor = {}
        harga_pelanggan = {}
        
        for row in all_values[1:]:
            if len(row) <= max(nama_col_idx, vendor_col_idx, pelanggan_col_idx):
                continue
            
            nama = str(row[nama_col_idx]).strip()
            if not nama:
                continue
            
            harga_v = parse_number_from_sheets(row[vendor_col_idx])
            harga_p = parse_number_from_sheets(row[pelanggan_col_idx])
            
            harga_vendor[nama] = harga_v
            harga_pelanggan[nama] = harga_p
        
        if harga_vendor and harga_pelanggan:
            return harga_vendor, harga_pelanggan, True
        
        return harga_vendor_fallback, harga_pelanggan_fallback, False
        
    except Exception:
        return harga_vendor_fallback, harga_pelanggan_fallback, False


if "pelanggan_loaded" not in st.session_state:
    st.session_state.pelanggan_loaded = False

if "harga_loaded" not in st.session_state:
    st.session_state.harga_loaded = False

if not st.session_state.pelanggan_loaded:
    df_sheets = fetch_pelanggan_df(SPREADSHEET_ID, GID)
    st.session_state.df_sheets = df_sheets
    st.session_state.pelanggan_loaded = True
else:
    df_sheets = st.session_state.df_sheets

if not st.session_state.harga_loaded:
    harga_vendor, harga_pelanggan, is_from_sheets = fetch_master_harga(SPREADSHEET_ID, MASTER_HARGA_SHEET)
    st.session_state.harga_vendor = harga_vendor
    st.session_state.harga_pelanggan = harga_pelanggan
    st.session_state.is_from_sheets = is_from_sheets
    st.session_state.harga_loaded = True
else:
    harga_vendor = st.session_state.harga_vendor
    harga_pelanggan = st.session_state.harga_pelanggan
    is_from_sheets = st.session_state.is_from_sheets

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
}

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

additional_items = [
    "Segel Plastik",
    "Twisted Cable 2 x 10 mm² - Al",
]

data_barang = []
for nama in main_items:
    harga = harga_pelanggan.get(nama, 0)
    sat = sat_mapping.get(nama, "")
    data_barang.append({"nama": nama, "SAT": sat, "harga": harga})

data_barang_tambahan = []
for nama in additional_items:
    harga = harga_pelanggan.get(nama, 0)
    sat = sat_mapping.get(nama, "")
    data_barang_tambahan.append({"nama": nama, "SAT": sat, "harga": harga})

semua_barang = data_barang + [{"nama": "---- PEMBATAS ----", "SAT": "", "harga": 0}] + data_barang_tambahan


@st.dialog("Preview Rekap", width="large")
def show_preview_dialog(barang_dipilih, meta_data):
    if not barang_dipilih:
        st.warning("Tidak ada barang yang dipilih.")
        return
    
    df_pilih = pd.DataFrame(barang_dipilih)
    
    df_preview_vendor = df_pilih.copy()
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
    
    tab1, tab2 = st.tabs(["VENDOR", "PELANGGAN"])
    
    with tab1:
        st.markdown("#### REKAP HARGA PEKERJAAN - VENDOR")
        st.markdown(f"**PEKERJAAN:** {meta_data.get('Pekerjaan', '-')}")
        st.markdown(f"**NAMA:** {meta_data.get('Nama', '-')}")
        st.markdown(f"**LOKASI:** {meta_data.get('Lokasi', '-')}")
        st.markdown(f"**ULP:** {meta_data.get('ULP', '-')}")
        st.markdown(f"**NO SPK:** {meta_data.get('No SPK', '-')}")
        st.markdown(f"**VENDOR PELAKSANA:** {meta_data.get('Vendor', '-')}")
        st.write("---")
        
        df_vendor_display = df_preview_vendor.copy()
        df_vendor_display["Harga Satuan Material"] = df_vendor_display["Harga Satuan Material"].apply(format_rupiah)
        df_vendor_display["Harga Total"] = df_preview_vendor["Harga Total"].apply(format_rupiah)
        
        st.dataframe(df_vendor_display[["Rincian", "SAT", "Vol", "Harga Satuan Material", "Harga Total"]], 
                    use_container_width=True, hide_index=True)
        st.write(f"**Subtotal:** Rp {format_rupiah(subtotal_vendor)}")
        st.write(f"**PPN (11%):** Rp {format_rupiah(ppn_vendor)}")
        st.success(f"**TOTAL BIAYA: Rp {format_rupiah(total_vendor)}**")
    
    with tab2:
        st.markdown("#### REKAP HARGA PEKERJAAN - PELANGGAN")
        st.markdown(f"**PEKERJAAN:** {meta_data.get('Pekerjaan', '-')}")
        st.markdown(f"**NAMA:** {meta_data.get('Nama', '-')}")
        st.markdown(f"**LOKASI:** {meta_data.get('Lokasi', '-')}")
        st.markdown(f"**ULP:** {meta_data.get('ULP', '-')}")
        st.markdown(f"**NO SPK:** {meta_data.get('No SPK', '-')}")
        st.markdown(f"**VENDOR PELAKSANA:** {meta_data.get('Vendor', '-')}")
        st.write("---")
        
        df_pelanggan_display = df_pilih.copy()
        df_pelanggan_display["Harga Satuan Material"] = df_pilih["Harga Satuan Material"].apply(format_rupiah)
        df_pelanggan_display["Harga Total"] = df_pilih["Harga Total"].apply(format_rupiah)
        
        st.dataframe(df_pelanggan_display[["Rincian", "SAT", "Vol", "Harga Satuan Material", "Harga Total"]], 
                    use_container_width=True, hide_index=True)
        st.write(f"**Subtotal:** Rp {format_rupiah(subtotal_pelanggan)}")
        st.write(f"**PPN (11%):** Rp {format_rupiah(ppn_pelanggan)}")
        st.success(f"**TOTAL BIAYA: Rp {format_rupiah(total_pelanggan)}**")
    
    st.write("---")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
    
    with col_btn1:
        if st.button("Batal", use_container_width=True, key="btn_cancel_export"):
            st.rerun()
    
    with col_btn3:
        if st.button("Konfirmasi & Export", type="primary", use_container_width=True, key="btn_confirm_export"):
            nama_full = meta_data.get('Nama', '-')
            if " (" in nama_full:
                nama_only = nama_full.split(" (")[0].strip()
            else:
                nama_only = nama_full
            
            now = now_jakarta().strftime("%Y%m%d_%H%M")
            safe_name = str(nama_only).replace("/", "-").replace("\\", "-")
            title_vendor = f"REKAP {safe_name} - {now}_Vendor"
            title_pelanggan = f"REKAP {safe_name} - {now}_Pelanggan"
            
            with st.spinner("Menulis data rekap ke Google Sheets..."):
                try:
                    from export_rekap_sheets import export_rekap_pair
                    
                    pair_info = export_rekap_pair(
                        spreadsheet_id=SPREADSHEET_ID,
                        base_sheet_title_vendor=title_vendor,
                        base_sheet_title_pelanggan=title_pelanggan,
                        meta=meta_data,
                        df_pilih=df_pilih,
                        idpel=None,
                        gid=None,
                        update_survey=False
                    )
                    
                    st.success(
                        f"Berhasil membuat rekap: **{pair_info['vendor']['sheet_title']}** dan "
                        f"**{pair_info['pelanggan']['sheet_title']}**"
                    )
                    
                    st.balloons()
                    time.sleep(2)
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Gagal mengekspor data: {e}")
                    st.error(traceback.format_exc())


@st.dialog("Edit Data Survey", width="large")
def show_edit_dialog(idpel: str):
    if not HAVE_DRAFT_MANAGER or draft_manager is None:
        st.error("Draft manager tidak tersedia")
        return
    
    draft = draft_manager.load_single_draft(SPREADSHEET_ID, idpel)
    
    if not draft["found"]:
        st.error("Data tidak ditemukan")
        return
    
    data = draft["data"]
    
    st.markdown(f"### Edit: {data.get('nama', '')} ({idpel})")
    st.markdown(f"**Lokasi:** {data.get('lokasi', '')}")
    st.markdown(f"**Tersimpan:** {data.get('tanggal_save', '')}")
    st.markdown("---")
    
    with st.form("form_edit"):
        pekerjaan = st.text_input("Pekerjaan", value=data.get('pekerjaan', ''), key="edit_pekerjaan")
        ulp = st.text_input("ULP", value=data.get('ulp', ''), key="edit_ulp")
        no_spk = st.text_input("No SPK", value=data.get('no_spk', ''), key="edit_no_spk")
        vendor = st.text_input("Vendor Pelaksana", value=data.get('vendor', ''), key="edit_vendor")
        
        st.markdown("---")
        st.subheader("Edit Kuantitas Barang")
        
        existing_qty = {}
        for item in data.get('barang', []):
            existing_qty[item['Rincian']] = item['Vol']
        
        new_quantities = {}
        for idx, barang in enumerate(semua_barang):
            if "----" in barang['nama']:
                st.markdown("---")
                continue
            
            default_val = existing_qty.get(barang['nama'], 0)
            qty = st.number_input(
                f"{barang['nama']} ({barang['SAT']})",
                min_value=0,
                value=int(default_val),
                key=f"edit_qty_{idx}_{idpel}"
            )
            new_quantities[barang['nama']] = qty
        
        submitted = st.form_submit_button("Simpan Perubahan", use_container_width=True, type="primary")
    
    if submitted:
        barang_updated = []
        for nama, qty in new_quantities.items():
            if qty > 0:
                harga = harga_pelanggan.get(nama, 0)
                barang_updated.append({
                    "Rincian": nama,
                    "SAT": sat_mapping.get(nama, ""),
                    "Vol": int(qty),
                    "Harga Satuan Material": harga,
                    "Harga Total": qty * harga
                })
        
        if not barang_updated:
            st.error("Minimal 1 barang harus diisi")
        else:
            with st.spinner("Menyimpan perubahan..."):
                if draft_manager is not None:
                    # Keep existing foto survey
                    foto_survey_links = data.get('foto_survey', [])
                    
                    result = draft_manager.save_draft_survey(
                        spreadsheet_id=SPREADSHEET_ID,
                        idpel=idpel,
                        nama=str(data.get('nama', '')),
                        lokasi=str(data.get('lokasi', '')),
                        pekerjaan=pekerjaan or '',
                        ulp=ulp or '',
                        no_spk=no_spk or '',
                        vendor=vendor or '',
                        barang_data=barang_updated,
                        foto_survey_links=foto_survey_links
                    )
                    
                    if result["success"]:
                        st.success(result['message'])
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(result['message'])
                else:
                    st.error("Draft manager tidak tersedia")


@st.dialog("Lihat Foto Survey", width="large")
def show_foto_survey_dialog(foto_list, nama, idpel):
    st.markdown(f"### Foto Survey: {nama} ({idpel})")
    st.markdown(f"**Jumlah Foto:** {len(foto_list)}")
    st.markdown("---")
    
    if not foto_list:
        st.info("Belum ada foto survey yang diupload.")
        return
    
    # Display in grid (2 columns)
    cols_per_row = 2
    for i in range(0, len(foto_list), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            idx = i + j
            if idx < len(foto_list):
                foto = foto_list[idx]
                with cols[j]:
                    st.markdown(f"**[{foto.get('name', 'Foto')}]({foto.get('link', '#')})**")
                    if foto.get('link'):
                        st.markdown(
                            f'<a href="{foto["link"]}" target="_blank">'
                            f'<img src="{foto["link"]}" style="width:100%; border-radius:8px; border:1px solid #ddd;"></a>',
                            unsafe_allow_html=True
                        )


st.title("Daftar Barang & Input Petugas")

if is_from_sheets:
    st.info(f"Harga dimuat dari sheet '{MASTER_HARGA_SHEET}'")
else:
    st.warning(f"Harga menggunakan data fallback. Pastikan sheet '{MASTER_HARGA_SHEET}' tersedia.")

st.subheader("Filter & Pilih Pelanggan")

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

barang_dipilih = []
with col2:
    st.subheader("Input Kuantitas Barang")
    
    quantities = {}
    for idx, barang in enumerate(semua_barang):
        if str(barang.get("nama", "")).startswith("----"):
            st.markdown("---")
            continue

        key_name = f"qty_{idx}"
        sat_label = barang.get("SAT", "")
        
        default_value = st.session_state.get(key_name, 0)
        
        qty = st.number_input(
            f"{barang.get('nama', 'Item')} ({sat_label})",
            min_value=0,
            step=1,
            value=default_value,
            key=key_name
        )
        quantities[idx] = qty

for idx, barang in enumerate(semua_barang):
    if str(barang.get("nama", "")).startswith("----"):
        continue
    
    qty = quantities.get(idx, 0)
    if qty and qty > 0:
        harga = float(barang.get("harga", 0) or 0)
        total = qty * harga
        barang_dipilih.append({
            "Rincian": barang.get("nama", ""),
            "SAT": sat_mapping.get(barang.get("nama", ""), ""),
            "Vol": int(qty),
            "Harga Satuan Material": harga,
            "Harga Total": total
        })

st.markdown("---")

# Initialize variables
uploaded_foto_survey = None
tanggal_survey_input = date.today()

# Upload Foto Survey Section
if idpel_selected:
    st.subheader("Upload Foto Survey")
    
    tanggal_survey_input = st.date_input(
        "Tanggal Survey:",
        value=date.today(),
        key="tanggal_survey_date",
        format="DD/MM/YYYY"
    )
    
    uploaded_foto_survey = st.file_uploader(
        "Upload Foto Dokumentasi Survey (JPG/PNG, minimal 1 foto, maksimal 5 foto):",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="upload_foto_survey",
        help="Wajib upload minimal 1 foto survey"
    )
    
    if uploaded_foto_survey:
        if len(uploaded_foto_survey) > 5:
            st.error("Maksimal 5 foto yang dapat diupload!")
        else:
            st.success(f"**Jumlah foto terpilih:** {len(uploaded_foto_survey)}")
            cols = st.columns(min(len(uploaded_foto_survey), 4))
            for idx, file in enumerate(uploaded_foto_survey):
                with cols[idx % 4]:
                    st.image(file, caption=file.name, width=150)

st.markdown("---")

if st.button("Simpan", type="primary", use_container_width=True):
    if not idpel_selected:
        st.error("Silakan pilih ID Pelanggan terlebih dahulu.")
    elif not barang_dipilih:
        st.error("Minimal 1 barang harus diisi (quantity > 0).")
    elif not uploaded_foto_survey or len(uploaded_foto_survey) == 0:
        st.error("Minimal 1 foto survey harus diupload!")
    elif len(uploaded_foto_survey) > 5:
        st.error("Maksimal 5 foto yang dapat diupload!")
    elif not HAVE_DRAFT_MANAGER or draft_manager is None:
        st.error("Draft manager tidak tersedia.")
    else:
        with st.spinner("Mengupload foto survey dan menyimpan data..."):
            try:
                # Format tanggal untuk filename
                tanggal_prefix = tanggal_survey_input.strftime("%d%m%Y")
                
                # Create/get subfolder di Drive
                subfolder_id = get_or_create_folder(DRIVE_FOLDER_SURVEY, idpel_selected)
                
                # Upload foto survey
                foto_survey_links = []
                for idx, file in enumerate(uploaded_foto_survey, 1):
                    ext = file.name.split(".")[-1]
                    # Format: IDPEL_DDMMYYYY_NAMA_survey_01.ext
                    filename = f"{idpel_selected}_{tanggal_prefix}_{nama.replace(' ', '_')}_survey_{idx:02d}.{ext}"
                    
                    result = upload_file_to_drive(
                        file_content=file.read(),
                        filename=filename,
                        folder_id=subfolder_id,
                        mime_type=file.type
                    )
                    
                    foto_survey_links.append({
                        "name": filename,
                        "link": result.get("webViewLink", "")
                    })
                
                # Save draft with foto survey
                result = draft_manager.save_draft_survey(
                    spreadsheet_id=SPREADSHEET_ID,
                    idpel=idpel_selected,
                    nama=str(nama),
                    lokasi=str(lokasi),
                    pekerjaan=pekerjaan or '',
                    ulp=ulp or '',
                    no_spk=no_spk or '',
                    vendor=vendor or '',
                    barang_data=barang_dipilih,
                    foto_survey_links=foto_survey_links
                )
                
                if not result["success"]:
                    st.error(result['message'])
                else:
                    if result["is_new"]:
                        from export_rekap_sheets import update_tanggal_survey_if_empty
                        
                        survey_result = update_tanggal_survey_if_empty(
                            SPREADSHEET_ID, GID, idpel_selected
                        )
                        
                        if survey_result["success"]:
                            st.success(result['message'])
                            st.info(survey_result['message'])
                            st.success(f"Berhasil upload {len(foto_survey_links)} foto survey!")
                        elif survey_result.get("already_filled", False):
                            st.success(result['message'])
                            st.info(survey_result['message'])
                            st.success(f"Berhasil upload {len(foto_survey_links)} foto survey!")
                        else:
                            st.success(result['message'])
                            st.warning(f"Tanggal Survey: {survey_result['message']}")
                            st.success(f"Berhasil upload {len(foto_survey_links)} foto survey!")
                    else:
                        st.success(f"{result['message']} (Data diperbarui)")
                        st.success(f"Berhasil upload {len(foto_survey_links)} foto survey!")
                    
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                    
            except Exception as e:
                st.error(f"Terjadi kesalahan: {str(e)}")
                st.error(traceback.format_exc())

st.markdown("---")

st.markdown("## Data Survey yang Sudah Tersimpan")

if not HAVE_DRAFT_MANAGER or draft_manager is None:
    st.warning("Draft manager tidak tersedia. Tidak bisa menampilkan data tersimpan.")
else:
    col_reload, col_search = st.columns([1, 3])
    
    with col_reload:
        if st.button("Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    
    with col_search:
        search_draft = st.text_input(
            "Cari ID Pelanggan",
            placeholder="Ketik ID Pelanggan untuk mencari...",
            key="search_draft_id",
            label_visibility="collapsed"
        )
    
    df_drafts = draft_manager.load_all_drafts(SPREADSHEET_ID)
    
    if df_drafts.empty:
        st.info("Belum ada data survey yang tersimpan. Silakan input dan simpan data pelanggan terlebih dahulu.")
    else:
        df_filtered_drafts = df_drafts.copy()
        
        if search_draft.strip():
            search_lower = search_draft.strip().lower()
            df_filtered_drafts = df_filtered_drafts[
                df_filtered_drafts["ID Pelanggan"].astype(str).str.lower().str.contains(search_lower, na=False)
            ]
        
        total_drafts = len(df_filtered_drafts)
        
        if total_drafts == 0:
            st.warning(f"Tidak ada data yang cocok dengan pencarian '{search_draft}'")
        else:
            items_per_page = 5
            total_pages = (total_drafts + items_per_page - 1) // items_per_page
            
            if "current_page" not in st.session_state:
                st.session_state.current_page = 1
            
            if st.session_state.current_page > total_pages:
                st.session_state.current_page = total_pages
            
            start_idx = (st.session_state.current_page - 1) * items_per_page
            end_idx = min(start_idx + items_per_page, total_drafts)
            
            st.markdown(f"**Terdapat {total_drafts} pelanggan** | Menampilkan: {start_idx + 1}-{end_idx}")
            st.markdown("---")
            
            df_page = df_filtered_drafts.iloc[start_idx:end_idx]
            
            for idx, row in df_page.iterrows():
                idpel_draft = str(row['ID Pelanggan'])
                nama_draft = str(row['Nama'])
                lokasi_draft = str(row['Lokasi'])
                tanggal_save = str(row['Tanggal Save'])
                jumlah_item = row['Jumlah_Item']
                jumlah_foto = row.get('Jumlah_Foto', 0)
                
                with st.expander(f"{idpel_draft} ({nama_draft})", expanded=False):
                    col_info1, col_info2 = st.columns(2)
                    
                    with col_info1:
                        st.markdown(f"**Lokasi:** {lokasi_draft}")
                        st.markdown(f"**Tersimpan:** {tanggal_save}")
                    
                    with col_info2:
                        st.markdown(f"**Jumlah Item:** {jumlah_item} barang")
                        st.markdown(f"**Jumlah Foto:** {jumlah_foto} foto")
                        
                        if st.checkbox("Lihat Detail Barang", key=f"detail_{idpel_draft}"):
                            pass
                    
                    if st.session_state.get(f"detail_{idpel_draft}", False):
                        barang_list = row['Barang_List']
                        if barang_list:
                            df_barang = pd.DataFrame(barang_list)
                            if 'Harga Satuan Material' in df_barang.columns:
                                df_barang['Harga Satuan Material'] = df_barang['Harga Satuan Material'].apply(format_rupiah)
                            if 'Harga Total' in df_barang.columns:
                                df_barang['Harga Total'] = df_barang['Harga Total'].apply(format_rupiah)
                            
                            st.dataframe(
                                df_barang[['Rincian', 'SAT', 'Vol', 'Harga Satuan Material', 'Harga Total']],
                                use_container_width=True,
                                hide_index=True
                            )
                    
                    st.markdown("---")
                    
                    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
                    
                    with col_btn1:
                        if st.button("Export", key=f"export_{idpel_draft}", use_container_width=True, type="primary"):
                            if draft_manager is not None:
                                draft_full = draft_manager.load_single_draft(SPREADSHEET_ID, idpel_draft)
                                
                                if draft_full["found"]:
                                    data_full = draft_full["data"]
                                    
                                    nama_with_id = f"{data_full.get('nama', '-')} ({idpel_draft})"
                                    meta = {
                                        "Pekerjaan": data_full.get('pekerjaan', '-'),
                                        "Nama": nama_with_id,
                                        "Lokasi": data_full.get('lokasi', '-'),
                                        "ULP": data_full.get('ulp', '-'),
                                        "No SPK": data_full.get('no_spk', '-'),
                                        "Vendor": data_full.get('vendor', '-')
                                    }
                                    
                                    show_preview_dialog(data_full.get('barang', []), meta)
                                else:
                                    st.error("Data tidak ditemukan")
                            else:
                                st.error("Draft manager tidak tersedia")
                    
                    with col_btn2:
                        if st.button("Edit", key=f"edit_{idpel_draft}", use_container_width=True):
                            show_edit_dialog(idpel_draft)
                    
                    with col_btn3:
                        if jumlah_foto > 0:
                            if st.button("Lihat Foto Survey", key=f"foto_{idpel_draft}", use_container_width=True):
                                foto_list = row['Foto_List']
                                show_foto_survey_dialog(foto_list, nama_draft, idpel_draft)
                        else:
                            st.button("Lihat Foto Survey", key=f"foto_{idpel_draft}", use_container_width=True, disabled=True)
                    
                    with col_btn4:
                        if f"confirm_delete_{idpel_draft}" not in st.session_state:
                            st.session_state[f"confirm_delete_{idpel_draft}"] = False
                        
                        if not st.session_state[f"confirm_delete_{idpel_draft}"]:
                            if st.button("Hapus", key=f"delete_{idpel_draft}", use_container_width=True):
                                st.session_state[f"confirm_delete_{idpel_draft}"] = True
                                st.rerun()
                        else:
                            col_del1, col_del2 = st.columns(2)
                            
                            with col_del1:
                                if st.button("Batal", key=f"cancel_delete_{idpel_draft}", use_container_width=True):
                                    st.session_state[f"confirm_delete_{idpel_draft}"] = False
                                    st.rerun()
                            
                            with col_del2:
                                if st.button("Yakin?", key=f"confirm_delete_yes_{idpel_draft}", 
                                           use_container_width=True, type="primary"):
                                    with st.spinner("Menghapus data..."):
                                        if draft_manager is not None:
                                            result = draft_manager.delete_draft_survey(SPREADSHEET_ID, idpel_draft)
                                            
                                            if result["success"]:
                                                st.success(result['message'])
                                                st.cache_data.clear()
                                                if f"confirm_delete_{idpel_draft}" in st.session_state:
                                                    del st.session_state[f"confirm_delete_{idpel_draft}"]
                                                time.sleep(1)
                                                st.rerun()
                                            else:
                                                st.error(result['message'])
                                        else:
                                            st.error("Draft manager tidak tersedia")
            
            st.markdown("---")
            
            col_prev, col_info, col_next = st.columns([1, 2, 1])
            
            with col_prev:
                if st.session_state.current_page > 1:
                    if st.button("◀ Sebelumnya", use_container_width=True):
                        st.session_state.current_page -= 1
                        st.rerun()
            
            with col_info:
                st.markdown(
                    f"<div style='text-align: center; padding: 8px;'>"
                    f"<strong>Halaman {st.session_state.current_page} dari {total_pages}</strong>"
                    f"</div>",
                    unsafe_allow_html=True
                )
            
            with col_next:
                if st.session_state.current_page < total_pages:
                    if st.button("Selanjutnya ▶", use_container_width=True):
                        st.session_state.current_page += 1
                        st.rerun()