import streamlit as st
import pandas as pd
from datetime import datetime, date
from auth import get_gspread_client, get_or_create_folder, upload_file_to_drive

# === Konfigurasi ===
try:
    SPREADSHEET_ID = str(st.secrets["SHEET_ID"])
    GID = str(st.secrets["SHEET_GID"])
    DRIVE_FOLDER_ID = str(st.secrets["DRIVE_FOLDER_EKSEKUSI"])  # Folder "Foto Eksekusi"
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
    return pd.DataFrame(data).fillna("")

def update_tanggal_eksekusi(spreadsheet_id: str, gid: str, idpel: str, tanggal: str) -> dict:
    """
    Update kolom TanggalEksekusi untuk IDPEL tertentu (row terakhir)
    """
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(spreadsheet_id)
        
        target_ws = None
        for ws in sh.worksheets():
            if str(ws.id) == str(gid):
                target_ws = ws
                break
        
        if target_ws is None:
            return {"success": False, "message": "Worksheet tidak ditemukan"}
        
        header = target_ws.row_values(1)
        
        # Cari kolom TanggalEksekusi
        eksekusi_col = None
        for idx, col_name in enumerate(header):
            if "tanggaleksekusi" in str(col_name).strip().lower().replace(" ", ""):
                eksekusi_col = idx + 1
                break
        
        if eksekusi_col is None:
            return {"success": False, "message": "Kolom TanggalEksekusi tidak ditemukan"}
        
        # Cari kolom ID Pelanggan
        id_col = None
        for idx, col_name in enumerate(header):
            if "id pelanggan" in str(col_name).strip().lower():
                id_col = idx + 1
                break
        
        if id_col is None:
            return {"success": False, "message": "Kolom ID Pelanggan tidak ditemukan"}
        
        # Cari row terakhir yang match
        id_values = target_ws.col_values(id_col)
        matched_row = None
        
        for i in reversed(range(1, len(id_values))):
            if str(id_values[i]).strip() == str(idpel).strip():
                matched_row = i + 1
                break
        
        if matched_row is None:
            return {"success": False, "message": f"ID Pelanggan {idpel} tidak ditemukan"}
        
        # Update cell
        target_ws.update_cell(matched_row, eksekusi_col, tanggal)
        
        return {"success": True, "message": f"Berhasil update row {matched_row}"}
        
    except Exception as e:
        return {"success": False, "message": f"Error: {str(e)}"}

# === UI ===
st.title("📸 Upload Dokumentasi Eksekusi")

df_sheets = fetch_pelanggan_df(SPREADSHEET_ID, GID)

# Filter & Pilih Pelanggan
st.subheader("🔎 Pilih Pelanggan")

col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    search_id = st.text_input(
        "🔍 Cari ID Pelanggan:",
        placeholder="Contoh: 513130665162",
        key="search_id_eksekusi"
    )

with col_filter2:
    search_nama = st.text_input(
        "👤 Cari Nama:",
        placeholder="Contoh: Sofia",
        key="search_nama_eksekusi"
    )

# Apply filter
df_filtered = df_sheets.copy()

if search_id.strip():
    df_filtered = df_filtered[
        df_filtered["ID Pelanggan"].astype(str).str.contains(search_id.strip(), case=False, na=False)
    ]

if search_nama.strip():
    if "Nama" in df_filtered.columns:
        df_filtered = df_filtered[
            df_filtered["Nama"].astype(str).str.contains(search_nama.strip(), case=False, na=False)
        ]

# Dropdown
filtered_options = ["- Pilih ID Pelanggan -"]
if not df_filtered.empty:
    for _, row in df_filtered.iterrows():
        pid = str(row["ID Pelanggan"]).strip()
        pnama = str(row.get("Nama", "-")).strip()
        if pid:
            filtered_options.append(f"{pid} ({pnama})")

pilihan = st.selectbox(
    "🔑 Pilih ID Pelanggan:",
    filtered_options,
    key="select_idpel_eksekusi"
)

# Extract ID
def extract_id(opt: str) -> str:
    if not opt or opt == "- Pilih ID Pelanggan -":
        return ""
    if " (" in opt:
        return opt.split(" (", 1)[0].strip()
    return opt.strip()

idpel_selected = extract_id(pilihan)

if idpel_selected:
    st.success(f"✅ Terpilih: {pilihan}")
    
    # Data pelanggan
    df_selected = df_sheets[df_sheets["ID Pelanggan"].astype(str) == idpel_selected]
    if not df_selected.empty:
        nama = str(df_selected.iloc[0].get("Nama", "-"))
        alamat = str(df_selected.iloc[0].get("Alamat kWH Meter", "-"))
        
        st.markdown(f"**Nama:** {nama}")
        st.markdown(f"**Alamat:** {alamat}")
    
    st.markdown("---")
    
    # Form Eksekusi
    st.subheader("📋 Input Data Eksekusi")
    
    with st.form("form_eksekusi"):
        # Tanggal Eksekusi
        tanggal_eksekusi = st.date_input(
            "📅 Tanggal Eksekusi:",
            value=date.today(),
            key="tanggal_eksekusi_input"
        )
        
        # Upload Foto
        uploaded_files = st.file_uploader(
            "📸 Upload Foto Dokumentasi (JPG/PNG, minimal 1 foto):",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="upload_foto_eksekusi"
        )
        
        # Validasi preview
        if uploaded_files:
            st.write(f"**Jumlah foto:** {len(uploaded_files)}")
            cols = st.columns(min(len(uploaded_files), 4))
            for idx, file in enumerate(uploaded_files):
                with cols[idx % 4]:
                    st.image(file, caption=file.name, width=150)
        
        submitted = st.form_submit_button("📤 Submit Data Eksekusi")
    
    # Proses submit
    if submitted:
        # Validasi
        if not uploaded_files:
            st.error("⚠️ Minimal 1 foto harus diupload!")
        else:
            with st.spinner("Mengupload foto ke Google Drive dan update data..."):
                try:
                    # Format tanggal dd/mm/yyyy
                    tanggal_str = tanggal_eksekusi.strftime("%d/%m/%Y")
                    
                    # 1. Buat/ambil subfolder untuk IDPEL
                    subfolder_id = get_or_create_folder(DRIVE_FOLDER_ID, idpel_selected)
                    
                    # 2. Upload semua foto
                    uploaded_links = []
                    for file in uploaded_files:
                        # Generate filename dengan timestamp
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        ext = file.name.split(".")[-1]
                        filename = f"{idpel_selected}_{timestamp}.{ext}"
                        
                        # Upload
                        result = upload_file_to_drive(
                            file_content=file.read(),
                            filename=filename,
                            folder_id=subfolder_id,
                            mime_type=file.type
                        )
                        
                        uploaded_links.append({
                            "name": filename,
                            "link": result.get("webViewLink", "")
                        })
                    
                    # 3. Update TanggalEksekusi di Google Sheets
                    update_result = update_tanggal_eksekusi(
                        SPREADSHEET_ID,
                        GID,
                        idpel_selected,
                        tanggal_str
                    )
                    
                    if update_result["success"]:
                        st.success(f"✅ Berhasil upload {len(uploaded_files)} foto dan update tanggal eksekusi!")
                        st.info(f"📅 Tanggal Eksekusi: {tanggal_str}")
                        st.info(f"📁 Foto tersimpan di: Foto Eksekusi/{idpel_selected}/")
                        
                        # Detail foto yang diupload
                        with st.expander("📋 Detail Foto yang Diupload"):
                            for item in uploaded_links:
                                st.write(f"- [{item['name']}]({item['link']})")
                        
                        st.balloons()
                    else:
                        st.error(f"❌ Upload foto berhasil, tapi gagal update sheets: {update_result['message']}")
                    
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())

# === DEBUG: Test Akses Folder ===
st.markdown("---")
st.subheader("🔧 Debug Mode")

if st.button("🧪 Test Akses Folder Drive"):
    try:
        from auth import get_drive_service
        service = get_drive_service()
        
        # Test 1: Cek folder eksekusi accessible
        st.write(f"**Testing Folder ID:** `{DRIVE_FOLDER_EKSEKUSI}`")
        
        folder = service.files().get(
            fileId=DRIVE_FOLDER_EKSEKUSI,
            fields='id, name, owners, capabilities',
            supportsAllDrives=True
        ).execute()
        
        st.success(f"✅ Folder ditemukan: **{folder.get('name')}**")
        st.json({
            "Folder ID": folder.get('id'),
            "Folder Name": folder.get('name'),
            "Can Edit": folder.get('capabilities', {}).get('canEdit', False),
            "Can Add Children": folder.get('capabilities', {}).get('canAddChildren', False)
        })
        
        # Test 2: List isi folder
        results = service.files().list(
            q=f"'{DRIVE_FOLDER_EKSEKUSI}' in parents and trashed=false",
            fields='files(id, name, mimeType)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        files = results.get('files', [])
        st.info(f"📁 Jumlah item di folder: **{len(files)}**")
        
        if files:
            st.write("**Isi folder:**")
            for f in files[:5]:  # Show max 5
                st.write(f"- {f['name']} (`{f['mimeType']}`)")
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("💡 Silakan pilih ID Pelanggan untuk melanjutkan")
