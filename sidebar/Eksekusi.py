import streamlit as st
import pandas as pd
from datetime import datetime, date
from auth import get_gspread_client, get_or_create_folder, upload_file_to_drive, get_drive_service

try:
    SPREADSHEET_ID = str(st.secrets["SHEET_ID"])
    GID = str(st.secrets["SHEET_GID"])
    DRIVE_FOLDER_EKSEKUSI = str(st.secrets.get("DRIVE_FOLDER_EKSEKUSI", ""))
    
    if not DRIVE_FOLDER_EKSEKUSI:
        st.error("DRIVE_FOLDER_EKSEKUSI tidak diset di secrets!")
        st.stop()
        
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
        
        eksekusi_col = None
        for idx, col_name in enumerate(header):
            if "tanggaleksekusi" in str(col_name).strip().lower().replace(" ", ""):
                eksekusi_col = idx + 1
                break
        
        if eksekusi_col is None:
            return {"success": False, "message": "Kolom TanggalEksekusi tidak ditemukan"}
        
        id_col = None
        for idx, col_name in enumerate(header):
            if "id pelanggan" in str(col_name).strip().lower():
                id_col = idx + 1
                break
        
        if id_col is None:
            return {"success": False, "message": "Kolom ID Pelanggan tidak ditemukan"}
        
        id_values = target_ws.col_values(id_col)
        matched_row = None
        
        for i in reversed(range(1, len(id_values))):
            if str(id_values[i]).strip() == str(idpel).strip():
                matched_row = i + 1
                break
        
        if matched_row is None:
            return {"success": False, "message": f"ID Pelanggan {idpel} tidak ditemukan"}
        
        target_ws.update_cell(matched_row, eksekusi_col, tanggal)
        
        return {"success": True, "message": f"Berhasil update row {matched_row}"}
        
    except Exception as e:
        return {"success": False, "message": f"Error: {str(e)}"}

def get_eksekusi_photos(idpel: str):
    try:
        if 'drive_service' in st.session_state:
            del st.session_state['drive_service']
        
        service = get_drive_service()
        
        query = f"name='{idpel}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name, parents)',
            pageSize=100
        ).execute()
        
        folders = results.get('files', [])
        
        target_folder_id = None
        for folder in folders:
            parents = folder.get('parents', [])
            if DRIVE_FOLDER_EKSEKUSI in parents:
                target_folder_id = folder['id']
                break
        
        if not target_folder_id:
            return []
        
        query_files = f"'{target_folder_id}' in parents and trashed=false and (mimeType='image/jpeg' or mimeType='image/png')"
        
        files_results = service.files().list(
            q=query_files,
            spaces='drive',
            fields='files(id, name, webViewLink)',
            orderBy='name',
            pageSize=100
        ).execute()
        
        files = files_results.get('files', [])
        
        photo_links = []
        for file in files:
            photo_links.append({
                "name": file['name'],
                "link": file.get('webViewLink', ''),
                "id": file['id']
            })
        
        return photo_links
        
    except Exception as e:
        error_msg = str(e).lower()
        if 'invalid_grant' in error_msg or 'bad request' in error_msg:
            return None
        return []

def get_eksekusi_photos_service_account(idpel: str):
    try:
        from google.oauth2.service_account import Credentials as SACredentials
        from googleapiclient.discovery import build
        
        sa_info = dict(st.secrets["service_account"])
        pk = sa_info.get("private_key", "")
        if "\\n" in pk:
            sa_info["private_key"] = pk.replace("\\n", "\n")
        
        scopes = ["https://www.googleapis.com/auth/drive.readonly"]
        creds = SACredentials.from_service_account_info(sa_info, scopes=scopes)
        service = build('drive', 'v3', credentials=creds)
        
        query = f"name='{idpel}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name, parents)',
            pageSize=100
        ).execute()
        
        folders = results.get('files', [])
        
        target_folder_id = None
        for folder in folders:
            parents = folder.get('parents', [])
            if DRIVE_FOLDER_EKSEKUSI in parents:
                target_folder_id = folder['id']
                break
        
        if not target_folder_id:
            return []
        
        query_files = f"'{target_folder_id}' in parents and trashed=false and (mimeType='image/jpeg' or mimeType='image/png')"
        
        files_results = service.files().list(
            q=query_files,
            spaces='drive',
            fields='files(id, name, webViewLink)',
            orderBy='name',
            pageSize=100
        ).execute()
        
        files = files_results.get('files', [])
        
        photo_links = []
        for file in files:
            photo_links.append({
                "name": file['name'],
                "link": file.get('webViewLink', ''),
                "id": file['id']
            })
        
        return photo_links
        
    except Exception:
        return []

@st.dialog("Lihat Foto Eksekusi", width="large")
def show_foto_eksekusi_dialog(foto_list, nama, idpel):
    st.markdown(f"### Foto Eksekusi: {nama} ({idpel})")
    st.markdown(f"**Jumlah Foto:** {len(foto_list)}")
    st.markdown("---")
    
    if not foto_list:
        st.info("Belum ada foto eksekusi yang diupload.")
        return
    
    def get_drive_image_url(file_id):
        try:
            return f"https://drive.google.com/thumbnail?id={file_id}&sz=w800"
        except Exception:
            return None
    
    cols_per_row = 2
    for i in range(0, len(foto_list), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            idx = i + j
            if idx < len(foto_list):
                foto = foto_list[idx]
                with cols[j]:
                    foto_name = foto.get('name', f'Foto {idx+1}')
                    foto_link = foto.get('link', '#')
                    foto_id = foto.get('id', '')
                    
                    st.markdown(f"**[{foto_name}]({foto_link})**")
                    
                    if foto_id:
                        try:
                            direct_url = get_drive_image_url(foto_id)
                            if direct_url:
                                st.image(direct_url, use_container_width=True)
                            else:
                                st.warning(f"[Klik di sini untuk melihat foto]({foto_link})")
                        except Exception:
                            st.warning(f"[Klik di sini untuk melihat foto]({foto_link})")
                    
                    st.markdown("---")

st.title("Upload Dokumentasi Eksekusi")

df_sheets = fetch_pelanggan_df(SPREADSHEET_ID, GID)

st.subheader("Pilih Pelanggan")

if "TanggalEksekusi" in df_sheets.columns:
    df_sheets_available = df_sheets[
        (df_sheets["TanggalEksekusi"].isna()) | 
        (df_sheets["TanggalEksekusi"].astype(str).str.strip() == "") |
        (df_sheets["TanggalEksekusi"].astype(str).str.lower() == "nan")
    ].copy()
    
    filtered_count = len(df_sheets_available)
    total_count = len(df_sheets)
    
    st.info(f"Menampilkan {filtered_count} pelanggan yang belum eksekusi (dari total {total_count} pelanggan)")
else:
    df_sheets_available = df_sheets.copy()
    st.warning("Kolom 'TanggalEksekusi' tidak ditemukan di spreadsheet")

col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    search_id = st.text_input(
        "Cari ID Pelanggan:",
        placeholder="Contoh: 513130665162",
        key="search_id_eksekusi"
    )

with col_filter2:
    search_nama = st.text_input(
        "Cari Nama:",
        placeholder="Contoh: Sofia",
        key="search_nama_eksekusi"
    )

df_filtered = df_sheets_available.copy()

if search_id.strip():
    df_filtered = df_filtered[
        df_filtered["ID Pelanggan"].astype(str).str.contains(search_id.strip(), case=False, na=False)
    ]

if search_nama.strip():
    if "Nama" in df_filtered.columns:
        df_filtered = df_filtered[
            df_filtered["Nama"].astype(str).str.contains(search_nama.strip(), case=False, na=False)
        ]

filtered_options = ["- Pilih ID Pelanggan -"]
if not df_filtered.empty:
    for _, row in df_filtered.iterrows():
        pid = str(row["ID Pelanggan"]).strip()
        pnama = str(row.get("Nama", "-")).strip()
        if pid:
            filtered_options.append(f"{pid} ({pnama})")
    
    result_count = len(filtered_options) - 1
    if result_count > 0:
        st.success(f"Ditemukan {result_count} pelanggan yang belum eksekusi")
    else:
        st.warning("Tidak ditemukan pelanggan dengan kata kunci pencarian tersebut")
else:
    st.warning("Belum ada pelanggan yang perlu eksekusi")

if len(filtered_options) > 1:
    pilihan = st.selectbox(
        "Pilih ID Pelanggan:",
        filtered_options,
        key="select_idpel_eksekusi"
    )
else:
    pilihan = "- Pilih ID Pelanggan -"
    st.info("Silakan gunakan pencarian untuk menemukan pelanggan lain")

def extract_id(opt: str) -> str:
    if not opt or opt == "- Pilih ID Pelanggan -":
        return ""
    if " (" in opt:
        return opt.split(" (", 1)[0].strip()
    return opt.strip()

idpel_selected = extract_id(pilihan)

if idpel_selected:
    st.success(f"Terpilih: {pilihan}")
    
    df_selected = df_sheets[df_sheets["ID Pelanggan"].astype(str) == idpel_selected]
    if not df_selected.empty:
        nama = str(df_selected.iloc[0].get("Nama", "-"))
        alamat = str(df_selected.iloc[0].get("Alamat kWH Meter", "-"))
        
        st.markdown(f"**Nama:** {nama}")
        st.markdown(f"**Alamat:** {alamat}")
    
    st.markdown("---")
    
    st.subheader("Input Data Eksekusi")
    
    with st.form("form_eksekusi"):
        tanggal_eksekusi = st.date_input(
            "Tanggal Eksekusi:",
            value=date.today(),
            key="tanggal_eksekusi_input",
            format="DD/MM/YYYY"
        )
        
        uploaded_files = st.file_uploader(
            "Upload Foto Dokumentasi (JPG/PNG, minimal 1 foto):",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="upload_foto_eksekusi"
        )
        
        if uploaded_files:
            st.write(f"**Jumlah foto:** {len(uploaded_files)}")
            cols = st.columns(min(len(uploaded_files), 4))
            for idx, file in enumerate(uploaded_files):
                with cols[idx % 4]:
                    st.image(file, caption=file.name, width=150)
        
        submitted = st.form_submit_button("Submit Data Eksekusi")
    
    if submitted:
        if not uploaded_files:
            st.error("Minimal 1 foto harus diupload.")
        else:
            with st.spinner("Mengupload foto ke Google Drive dan update data..."):
                try:
                    tanggal_str = tanggal_eksekusi.strftime("%d/%m/%Y")
                    
                    tanggal_prefix = tanggal_eksekusi.strftime("%d%m%Y")
                    
                    df_selected = df_sheets[df_sheets["ID Pelanggan"].astype(str) == idpel_selected]
                    if not df_selected.empty:
                        nama = str(df_selected.iloc[0].get("Nama", "-"))
                    else:
                        nama = "-"
                    
                    subfolder_id = get_or_create_folder(DRIVE_FOLDER_EKSEKUSI, idpel_selected)
                    
                    uploaded_links = []
                    for idx, file in enumerate(uploaded_files, 1):
                        ext = file.name.split(".")[-1]
                        filename = f"{idpel_selected}_{tanggal_prefix}_{nama.replace(' ', '_')}_{idx:02d}.{ext}"
                        
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
                    
                    update_result = update_tanggal_eksekusi(
                        SPREADSHEET_ID,
                        GID,
                        idpel_selected,
                        tanggal_str
                    )
                    
                    if update_result["success"]:
                        st.success(f"Berhasil upload {len(uploaded_files)} foto dan update tanggal eksekusi.")
                        st.info(f"Tanggal Eksekusi: {tanggal_str}")
                        st.info(f"Foto tersimpan di: Foto Eksekusi/{idpel_selected}/")
                        
                        with st.expander("Detail Foto yang Diupload"):
                            for item in uploaded_links:
                                st.write(f"- [{item['name']}]({item['link']})")
                        
                        st.balloons()
                        st.cache_data.clear()
                    else:
                        st.error(f"Upload foto berhasil, tapi gagal update sheets: {update_result['message']}")
                    
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())

st.markdown("---")

st.markdown("## History Eksekusi yang Sudah Diinputkan")

col_refresh_history, col_search_history = st.columns([1, 3])

with col_refresh_history:
    if st.button("Refresh Data", use_container_width=True, key="refresh_history"):
        st.cache_data.clear()
        st.rerun()

with col_search_history:
    search_history = st.text_input(
        "Cari ID Pelanggan atau Nama",
        placeholder="Ketik untuk mencari...",
        key="search_history",
        label_visibility="collapsed"
    )

if "TanggalEksekusi" in df_sheets.columns:
    df_history = df_sheets[
        (df_sheets["TanggalEksekusi"].notna()) & 
        (df_sheets["TanggalEksekusi"].astype(str).str.strip() != "") &
        (df_sheets["TanggalEksekusi"].astype(str).str.lower() != "nan")
    ].copy()
    
    if "Timestamp" in df_history.columns:
        try:
            df_history["Date_Eksekusi"] = pd.to_datetime(
                df_history["TanggalEksekusi"],
                format="%d/%m/%Y",
                errors='coerce'
            )
            
            mask_invalid = df_history["Date_Eksekusi"].isna()
            if mask_invalid.any():
                df_history.loc[mask_invalid, "Date_Eksekusi"] = pd.to_datetime(
                    df_history.loc[mask_invalid, "TanggalEksekusi"],
                    errors='coerce'
                )
        except Exception:
            df_history["Date_Eksekusi"] = pd.NaT
    
    col_filter_date = st.columns(1)[0]
    
    with col_filter_date:
        if "Date_Eksekusi" in df_history.columns:
            date_series = df_history["Date_Eksekusi"].dropna()
            available_dates_history = date_series.apply(lambda x: x.date() if hasattr(x, 'date') else None).unique()
            available_dates_history = sorted([d for d in available_dates_history if d is not None], reverse=True)
            
            date_options_history = ["Semua Tanggal"] + [str(d) for d in available_dates_history]
            
            selected_date_history = st.selectbox(
                "Filter Tanggal Eksekusi:",
                date_options_history,
                key="filter_date_history"
            )
        else:
            selected_date_history = "Semua Tanggal"
    
    df_history_filtered = df_history.copy()
    
    if selected_date_history != "Semua Tanggal" and "Date_Eksekusi" in df_history.columns:
        df_history_filtered = df_history_filtered[
            df_history_filtered["Date_Eksekusi"].apply(lambda x: str(x.date()) if pd.notna(x) else "") == selected_date_history
        ]
    
    if search_history.strip():
        search_lower = search_history.strip().lower()
        mask_id = df_history_filtered["ID Pelanggan"].astype(str).str.lower().str.contains(search_lower, na=False)
        
        if "Nama" in df_history_filtered.columns:
            mask_nama = df_history_filtered["Nama"].astype(str).str.lower().str.contains(search_lower, na=False)
            df_history_filtered = df_history_filtered[mask_id | mask_nama]
        else:
            df_history_filtered = df_history_filtered[mask_id]
    
    total_history = len(df_history_filtered)
    
    if total_history == 0:
        if search_history.strip() or selected_date_history != "Semua Tanggal":
            st.warning("Tidak ada data yang cocok dengan filter")
        else:
            st.info("Belum ada data eksekusi yang tersimpan")
    else:
        items_per_page = 5
        total_pages = (total_history + items_per_page - 1) // items_per_page
        
        if "current_page_history" not in st.session_state:
            st.session_state.current_page_history = 1
        
        if st.session_state.current_page_history > total_pages:
            st.session_state.current_page_history = total_pages
        
        start_idx = (st.session_state.current_page_history - 1) * items_per_page
        end_idx = min(start_idx + items_per_page, total_history)
        
        st.markdown(f"**Terdapat {total_history} pelanggan** | Menampilkan: {start_idx + 1}-{end_idx}")
        st.markdown("---")
        
        df_page = df_history_filtered.iloc[start_idx:end_idx]
        
        for idx, row in df_page.iterrows():
            idpel_history = str(row['ID Pelanggan'])
            nama_history = str(row.get('Nama', '-'))
            alamat_history = str(row.get('Alamat kWH Meter', '-'))
            tanggal_eksekusi_history = str(row.get('TanggalEksekusi', '-'))
            
            with st.expander(f"{idpel_history} ({nama_history})", expanded=False):
                col_info1, col_info2 = st.columns(2)
                
                with col_info1:
                    st.markdown(f"**Nama:** {nama_history}")
                    st.markdown(f"**Alamat:** {alamat_history}")
                
                with col_info2:
                    st.markdown(f"**Tanggal Eksekusi:** {tanggal_eksekusi_history}")
                    
                    foto_list = []
                    jumlah_foto = 0
                    foto_error = False
                    
                    try:
                        foto_result = get_eksekusi_photos(idpel_history)
                        if foto_result is None:
                            foto_result = get_eksekusi_photos_service_account(idpel_history)
                            if foto_result:
                                foto_list = foto_result
                                jumlah_foto = len(foto_list)
                            else:
                                foto_error = True
                                st.caption("⚠️ Tidak dapat memuat foto")
                        else:
                            foto_list = foto_result
                            jumlah_foto = len(foto_list)
                    except Exception as e:
                        foto_error = True
                        st.caption(f"⚠️ Error loading photos")
                    
                    if not foto_error:
                        st.markdown(f"**Jumlah Foto:** {jumlah_foto} foto")
                
                st.markdown("---")
                
                col_btn1, col_btn2 = st.columns([1, 3])
                
                with col_btn1:
                    if jumlah_foto > 0:
                        if st.button("Lihat Foto", key=f"view_foto_{idpel_history}", use_container_width=True, type="primary"):
                            show_foto_eksekusi_dialog(foto_list, nama_history, idpel_history)
                    else:
                        st.button("Lihat Foto", key=f"view_foto_{idpel_history}", 
                                use_container_width=True, disabled=True, help="Belum ada foto")
                
                with col_btn2:
                    if jumlah_foto > 0:
                        folder_link = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_EKSEKUSI}"
                        st.markdown(
                            f"[Buka Folder di Google Drive ↗]({folder_link})",
                            unsafe_allow_html=True
                        )
        
        st.markdown("---")
        
        col_prev, col_info, col_next = st.columns([1, 2, 1])
        
        with col_prev:
            if st.session_state.current_page_history > 1:
                if st.button("Sebelumnya", use_container_width=True, key="prev_history"):
                    st.session_state.current_page_history -= 1
                    st.rerun()
        
        with col_info:
            st.markdown(
                f"<div style='text-align: center; padding: 8px;'>"
                f"<strong>Halaman {st.session_state.current_page_history} dari {total_pages}</strong>"
                f"</div>",
                unsafe_allow_html=True
            )
        
        with col_next:
            if st.session_state.current_page_history < total_pages:
                if st.button("Selanjutnya", use_container_width=True, key="next_history"):
                    st.session_state.current_page_history += 1
                    st.rerun()
else:
    st.warning("Kolom 'TanggalEksekusi' tidak ditemukan di spreadsheet")