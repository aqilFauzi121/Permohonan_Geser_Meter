from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials as SACredentials
from functools import lru_cache
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import gspread
import streamlit as st
import io

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]

@lru_cache(maxsize=1)
def get_gspread_client():
    """Service Account untuk Sheets"""
    sa_info = dict(st.secrets["service_account"])
    pk = sa_info.get("private_key", "")
    if "\\n" in pk:
        sa_info["private_key"] = pk.replace("\\n", "\n")
    creds = SACredentials.from_service_account_info(sa_info, scopes=SCOPES)
    return gspread.authorize(creds)

def get_drive_service():
    """OAuth credentials untuk Drive dengan auto-refresh"""
    
    # Cache di session state
    if 'drive_service' in st.session_state:
        return st.session_state['drive_service']
    
    if "oauth_token" not in st.secrets:
        st.error("OAuth token belum di-setup di secrets!")
        st.stop()
    
    token_data = dict(st.secrets["oauth_token"])
    
    creds = Credentials(
        token=token_data.get("access_token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=token_data.get("token_uri"),
        client_id=token_data.get("client_id"),
        client_secret=token_data.get("client_secret"),
        scopes=SCOPES
    )
    
    # Auto refresh jika expired
    from google.auth.transport.requests import Request
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
    
    service = build('drive', 'v3', credentials=creds)
    st.session_state['drive_service'] = service
    
    return service

def get_or_create_folder(parent_folder_id: str, folder_name: str) -> str:
    service = get_drive_service()
    
    query = f"name='{folder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    
    try:
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)'
        ).execute()
        
        items = results.get('files', [])
        if items:
            return items[0]['id']
    except Exception:
        pass
    
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
    service = get_drive_service()
    
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    
    fh = io.BytesIO(file_content)
    media = MediaIoBaseUpload(fh, mimetype=mime_type, resumable=True)
    
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, webViewLink'
    ).execute()
    
    return file