import boto3
from botocore.client import Config
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from functools import lru_cache
import gspread
import streamlit as st

# Scopes permissions for Google Sheets and Drive
SCOPES_SHEETS = ["https://www.googleapis.com/auth/spreadsheets"]
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive.readonly"]

@lru_cache(maxsize=1)
def get_gspread_client():
    # Initialize gspread client using service account for Sheets access
    sa_info = dict(st.secrets["service_account"])
    pk = sa_info.get("private_key", "")
    if "\\n" in pk:
        sa_info["private_key"] = pk.replace("\\n", "\n")
    creds = SACredentials.from_service_account_info(sa_info, scopes=SCOPES_SHEETS)
    return gspread.authorize(creds)

def get_drive_service():
    # Initialize Google Drive service using Service Account (Read-Only)
    try:
        sa_info = dict(st.secrets["service_account"])
        pk = sa_info.get("private_key", "")
        if "\\n" in pk:
            sa_info["private_key"] = pk.replace("\\n", "\n")
        
        creds = SACredentials.from_service_account_info(sa_info, scopes=SCOPES_DRIVE)
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception:
        return None

def get_r2_client():
    # Initialize boto3 client for Cloudflare R2
    r2_config = st.secrets["CLOUDFLARE_R2"]
    
    s3_client = boto3.client(
        's3',
        endpoint_url=r2_config["S3_ENDPOINT"],
        aws_access_key_id=r2_config["ACCESS_KEY_ID"],
        aws_secret_access_key=r2_config["SECRET_ACCESS_KEY"],
        config=Config(signature_version='s3v4'),
        region_name='auto'
    )
    return s3_client

def upload_to_r2(file_content: bytes, object_key: str, content_type: str = 'image/jpeg') -> dict:
    # Upload file to Cloudflare R2 bucket
    try:
        r2_config = st.secrets["CLOUDFLARE_R2"]
        bucket_name = r2_config["BUCKET_NAME"]
        public_domain = r2_config.get("PUBLIC_DOMAIN", "")
        
        s3_client = get_r2_client()
        
        s3_client.put_object(
            Bucket=bucket_name,
            Key=object_key,
            Body=file_content,
            ContentType=content_type
        )
        
        public_url = f"{public_domain}/{object_key}"
        
        return {
            "success": True,
            "object_key": object_key,
            "public_url": public_url,
            "bucket": bucket_name
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

def list_r2_objects(prefix: str) -> list:
    # List objects in R2 bucket with specific prefix
    try:
        r2_config = st.secrets["CLOUDFLARE_R2"]
        bucket_name = r2_config["BUCKET_NAME"]
        public_domain = r2_config.get("PUBLIC_DOMAIN", "")
        
        s3_client = get_r2_client()
        
        response = s3_client.list_objects_v2(
            Bucket=bucket_name,
            Prefix=prefix
        )
        
        objects = []
        if 'Contents' in response:
            for obj in response['Contents']:
                object_key = obj['Key']
                if object_key.lower().endswith(('.jpg', '.jpeg', '.png')):
                    public_url = f"{public_domain}/{object_key}"
                    filename = object_key.split('/')[-1]
                    
                    objects.append({
                        "key": object_key,
                        "name": filename,
                        "url": public_url,
                        "size": obj.get('Size', 0),
                        "last_modified": obj.get('LastModified'),
                        "source": "r2"
                    })
        return objects
    except Exception:
        return []

def get_r2_public_url(object_key: str) -> str:
    # Generate public URL for R2 object
    r2_config = st.secrets["CLOUDFLARE_R2"]
    public_domain = r2_config.get("PUBLIC_DOMAIN", "")
    return f"{public_domain}/{object_key}"

def delete_r2_object(object_key: str) -> dict:
    # Delete object from R2 bucket
    try:
        r2_config = st.secrets["CLOUDFLARE_R2"]
        bucket_name = r2_config["BUCKET_NAME"]
        
        s3_client = get_r2_client()
        
        s3_client.delete_object(
            Bucket=bucket_name,
            Key=object_key
        )
        return {
            "success": True, 
            "message": f"Object {object_key} berhasil dihapus"
        }
    except Exception as e:
        return {
            "success": False, 
            "error": str(e)
        }

def list_drive_photos(folder_id: str, idpel: str) -> list:
    # List photos from legacy Google Drive folder
    try:
        service = get_drive_service()
        if service is None:
            return []
        
        # Search for folder with customer ID name
        query_folder = f"name='{idpel}' and '{folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        
        results_folder = service.files().list(
            q=query_folder,
            spaces='drive',
            fields='files(id, name)',
            pageSize=10
        ).execute()
        
        folders = results_folder.get('files', [])
        if not folders:
            return []
        
        target_folder_id = folders[0]['id']
        
        # List all image files in that folder
        query_files = f"'{target_folder_id}' in parents and trashed=false and (mimeType='image/jpeg' or mimeType='image/png')"
        
        results_files = service.files().list(
            q=query_files,
            spaces='drive',
            fields='files(id, name, webViewLink)',
            orderBy='name',
            pageSize=100
        ).execute()
        
        files = results_files.get('files', [])
        
        photo_links = []
        for file in files:
            photo_links.append({
                "name": file['name'],
                "link": file.get('webViewLink', ''),
                "id": file['id'],
                "source": "drive"
            })
        
        return photo_links
    except Exception:
        return []

def detect_photo_source(url: str) -> str:
    # Helper to detect if URL is from Drive or R2
    if not url or url == '#':
        return 'unknown'
    
    url_lower = url.lower()
    if 'drive.google.com' in url_lower or 'googleapis.com' in url_lower:
        return 'drive'
    if 'r2.dev' in url_lower or 'r2.cloudflarestorage.com' in url_lower:
        return 'r2'
    return 'unknown'

def get_drive_thumbnail_url(drive_url: str) -> str:
    # Convert standard Drive URL to direct thumbnail URL
    try:
        if '/file/d/' in drive_url:
            file_id = drive_url.split('/file/d/')[1].split('/')[0]
        elif 'id=' in drive_url:
            file_id = drive_url.split('id=')[1].split('&')[0]
        else:
            return drive_url
        
        return f"https://drive.google.com/thumbnail?id={file_id}&sz=w800"
    except Exception:
        return drive_url