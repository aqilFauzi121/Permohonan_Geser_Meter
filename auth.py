from google.oauth2.service_account import Credentials
from functools import lru_cache
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload 
import gspread, streamlit as st
import io

SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

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
    results = service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)'
    ).execute()
    
    items = results.get('files', [])
    
    if items:
        return items[0]['id']
    
    # Buat folder baru jika belum ada
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    
    folder = service.files().create(
        body=file_metadata,
        fields='id'
    ).execute()
    
    return folder.get('id')

def upload_file_to_drive(file_content, filename: str, folder_id: str, mime_type: str) -> dict:
    """
    Upload file ke Google Drive
    Returns: {'id': file_id, 'name': filename, 'webViewLink': url}
    """
    service = get_drive_service()
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    
    # Create file in memory dengan benar
    fh = io.BytesIO(file_content)
    
    # Gunakan MediaIoBaseUpload untuk BytesIO object
    from googleapiclient.http import MediaIoBaseUpload
    
    media = MediaIoBaseUpload(
        fh,
        mimetype=mime_type,
        resumable=True
    )
    
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, webViewLink'
    ).execute()
    
    return file
