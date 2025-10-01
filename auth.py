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
    Cari atau buat folder di Google Drive + SET PERMISSION untuk Service Account
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
    
    folder_id = folder.get('id')
    
    # CRITICAL: Set permission untuk Service Account agar bisa menulis
    try:
        # Ambil email Service Account dari credentials
        sa_info = dict(st.secrets["service_account"])
        sa_email = sa_info.get("client_email")
        
        if sa_email:
            permission = {
                'type': 'user',
                'role': 'writer',  # atau 'writer' jika tidak perlu full control
                'emailAddress': sa_email
            }
            
            service.permissions().create(
                fileId=folder_id,
                body=permission,
                fields='id',
                supportsAllDrives=True
            ).execute()
    except Exception as e:
        # Log error tapi jangan stop proses
        import traceback
        st.warning(f"Warning: Gagal set permission otomatis: {e}")
    
    return folder_id


def upload_file_to_drive(file_content, filename: str, folder_id: str, mime_type: str) -> dict:
    """
    Upload file ke Google Drive - Simplified version
    """
    service = get_drive_service()
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    
    fh = io.BytesIO(file_content)
    media = MediaIoBaseUpload(
        fh,
        mimetype=mime_type,
        resumable=False  # Ubah jadi False untuk file kecil
    )
    
    try:
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name, webViewLink',
            supportsAllDrives=True
        ).execute()
        
        return file
    except Exception as e:
        # Jika gagal, coba dengan strategy lain (upload ke parent langsung)
        if "403" in str(e) or "quota" in str(e).lower():
            raise PermissionError(
                f"Service Account tidak punya permission menulis ke folder {folder_id}. "
                f"Error: {str(e)}"
            )
        raise
