from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials as SACredentials
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from functools import lru_cache
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import gspread
import streamlit as st
import io
import pickle
import os
from datetime import datetime, timedelta

# Scopes
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]

# Path untuk menyimpan token OAuth (persistent)
TOKEN_PICKLE = "token.pickle"

# ============================================================
# MODE: Service Account untuk Sheets, OAuth untuk Drive
# ============================================================

@lru_cache(maxsize=1)
def get_gspread_client():
    """Gunakan Service Account untuk Google Sheets (read-only)"""
    sa_info = dict(st.secrets["service_account"])
    pk = sa_info.get("private_key", "")
    if "\\n" in pk:
        sa_info["private_key"] = pk.replace("\\n", "\n")
    creds = SACredentials.from_service_account_info(sa_info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_oauth_credentials():
    """
    Dapatkan OAuth credentials untuk Drive upload.
    Login sekali, token disimpan persistent.
    """
    creds = None
    
    # Cek apakah token sudah ada di file pickle
    if os.path.exists(TOKEN_PICKLE):
        with open(TOKEN_PICKLE, 'rb') as token:
            creds = pickle.load(token)
    
    # Jika token expired atau tidak ada, refresh atau login ulang
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                st.error(f"Token expired dan gagal refresh: {e}")
                creds = None
        
        if not creds:
            # Perlu login ulang
            st.warning("Anda perlu login dengan Google untuk upload file ke Drive.")
            
            # Setup OAuth flow
            flow = Flow.from_client_config(
                {
                    "web": {
                        "client_id": st.secrets["oauth"]["client_id"],
                        "client_secret": st.secrets["oauth"]["client_secret"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                        "redirect_uris": [st.secrets["oauth"]["redirect_uri"]]
                    }
                },
                scopes=SCOPES,
                redirect_uri=st.secrets["oauth"]["redirect_uri"]
            )
            
            auth_url, _ = flow.authorization_url(prompt='consent')
            
            st.markdown(f"[Click di sini untuk Login dengan Google]({auth_url})")
            st.info("Setelah login, salin **seluruh URL** dari browser dan paste di bawah:")
            
            auth_response = st.text_input("Paste URL setelah login:", key="oauth_url")
            
            if auth_response:
                try:
                    flow.fetch_token(authorization_response=auth_response)
                    creds = flow.credentials
                    
                    # Simpan token ke file pickle
                    with open(TOKEN_PICKLE, 'wb') as token:
                        pickle.dump(creds, token)
                    
                    st.success("Login berhasil! Token tersimpan. Refresh halaman untuk melanjutkan.")
                    st.balloons()
                    st.stop()
                    
                except Exception as e:
                    st.error(f"Gagal login: {e}")
                    st.stop()
            else:
                st.stop()
    
    return creds


@lru_cache(maxsize=1)
def get_drive_service():
    """Get Google Drive API service dengan OAuth"""
    creds = get_oauth_credentials()
    return build('drive', 'v3', credentials=creds)


def get_or_create_folder(parent_folder_id: str, folder_name: str) -> str:
    """
    Cari atau buat folder di Google Drive (menggunakan OAuth user credentials)
    Returns: folder_id
    """
    service = get_drive_service()
    
    # Cek apakah folder sudah ada
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
    
    # Buat folder baru
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
    Upload file ke Google Drive menggunakan OAuth credentials
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
        resumable=True
    )
    
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, webViewLink'
    ).execute()
    
    return file
