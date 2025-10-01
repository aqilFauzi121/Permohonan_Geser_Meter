from google.oauth2.service_account import Credentials
from functools import lru_cache
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import gspread
import streamlit as st
import io

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@lru_cache(maxsize=1)
def get_gspread_client():
    sa_info = dict(st.secrets["service_account"])
    pk = sa_info.get("private_key", "")
    if "\\n" in pk:
        sa_info["private_key"] = pk.replace("\\n", "\n")
    creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return gspread.authorize(creds)

@lru_cache(maxsize=1)
def get_drive_service():
    """Get Google Drive API service"""
    sa_info = dict(st.secrets["service_account"])
    pk = sa_info.get("private_key", "")
    if "\\n" in pk:
        sa_info["private_key"] = pk.replace("\\n", "\n")
    creds = Credentials.from_service_account_info(
        sa_info, 
        scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)

def get_or_create_folder(parent_folder_id: str, folder_name: str) -> str:
    """
    Cari atau buat folder di Google Drive
    Returns: folder_id
    """
    service = get_drive_service()
    
    # Cek apakah folder sudah ada
    query = f"name='{folder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    
    try:
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        items = results.get('files', [])
        
        if items:
            return items[0]['id']
    except Exception as e:
        # Jika error saat list (mungkin permission issue), coba buat folder baru
        pass
    
    # Buat folder baru jika belum ada
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    
    try:
        folder = service.files().create(
            body=file_metadata,
            fields='id',
            supportsAllDrives=True
        ).execute()
        
        return folder.get('id')
    except Exception as e:
        raise Exception(f"Gagal membuat folder {folder_name}: {str(e)}")

def upload_file_to_drive(file_content, filename: str, folder_id: str, mime_type: str) -> dict:
    """
    Upload file ke Google Drive - FIXED VERSION untuk personal account
    
    Strategi:
    1. Coba upload langsung ke folder (jika berhasil, selesai)
    2. Jika gagal quota, upload tanpa parent dulu
    3. Lalu copy/move ke folder target yang sudah di-share
    
    Returns: {'id': file_id, 'name': filename, 'webViewLink': url}
    """
    service = get_drive_service()
    
    # METODE 1: Upload langsung ke folder target (optimal)
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    
    fh = io.BytesIO(file_content)
    media = MediaIoBaseUpload(
        fh,
        mimetype=mime_type,
        resumable=True
    )
    
    try:
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name, webViewLink',
            supportsAllDrives=True
        ).execute()
        
        return file
        
    except Exception as upload_error:
        error_msg = str(upload_error).lower()
        
        # Cek apakah error adalah storage quota
        if 'storage' in error_msg or 'quota' in error_msg:
            # METODE 2: Upload tanpa parent, lalu move ke folder
            return _upload_via_move(service, file_content, filename, folder_id, mime_type)
        else:
            # Error lain, raise exception
            raise upload_error

def _upload_via_move(service, file_content, filename: str, target_folder_id: str, mime_type: str) -> dict:
    """
    Workaround: Upload tanpa parent folder dulu, lalu move ke folder target
    """
    # Step 1: Upload ke root Drive (tanpa parent)
    file_metadata = {
        'name': filename
        # Sengaja tidak ada 'parents'
    }
    
    fh = io.BytesIO(file_content)
    media = MediaIoBaseUpload(
        fh,
        mimetype=mime_type,
        resumable=True
    )
    
    temp_file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, parents'
    ).execute()
    
    temp_file_id = temp_file['id']
    
    # Step 2: Move file dari root ke folder target
    # Get previous parents
    previous_parents = ",".join(temp_file.get('parents', []))
    
    # Move file
    file = service.files().update(
        fileId=temp_file_id,
        addParents=target_folder_id,
        removeParents=previous_parents,
        fields='id, name, webViewLink',
        supportsAllDrives=True
    ).execute()
    
    return file
