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
    except Exception:
        pass
    
    # Buat folder baru
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    
    folder = service.files().create(
        body=file_metadata,
        fields='id',
        supportsAllDrives=True
    ).execute()
    
    return folder.get('id')

def upload_file_to_drive(file_content, filename: str, folder_id: str, mime_type: str) -> dict:
    """
    Upload file ke Google Drive - Fixed untuk Service Account
    
    Strategy: Upload langsung dengan parent folder yang sudah di-share
    """
    service = get_drive_service()
    
    # Validasi: Cek apakah Service Account punya akses ke folder
    try:
        folder_check = service.files().get(
            fileId=folder_id,
            fields='capabilities',
            supportsAllDrives=True
        ).execute()
        
        capabilities = folder_check.get('capabilities', {})
        if not capabilities.get('canAddChildren', False):
            raise PermissionError(
                f"Service Account tidak punya permission 'Editor' di folder {folder_id}. "
                "Pastikan folder sudah di-share dengan email Service Account sebagai Editor."
            )
    except Exception as e:
        if "403" in str(e):
            raise PermissionError(
                f"Folder {folder_id} tidak dapat diakses. "
                "Pastikan folder sudah di-share dengan Service Account."
            )
        raise
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    
    fh = io.BytesIO(file_content)
    media = MediaIoBaseUpload(
        fh,
        mimetype=mime_type,
        resumable=True,
        chunksize=1024*1024  # 1MB chunks
    )
    
    # Upload dengan parameter yang benar
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, webViewLink',
        supportsAllDrives=True
    ).execute()
    
    return file
