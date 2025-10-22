import json
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
from auth import get_gspread_client
import gspread

try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

try:
    DRAFT_SHEET_NAME = str(st.secrets.get("DRAFT_SHEET_NAME", "_DRAFT"))
except Exception:
    DRAFT_SHEET_NAME = "_DRAFT"


def get_or_create_draft_sheet(spreadsheet_id: str) -> gspread.Worksheet:
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    
    try:
        ws = sh.worksheet(DRAFT_SHEET_NAME)
        return ws
    
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=DRAFT_SHEET_NAME, rows=1000, cols=10)
        
        headers = [
            "ID Pelanggan",
            "Nama",
            "Lokasi",
            "Pekerjaan",
            "ULP",
            "No SPK",
            "Vendor Pelaksana",
            "Data Barang (JSON)",
            "Tanggal Save",
            "Status Tanggal Survey"
        ]
        ws.update('A1:J1', [headers])
        
        ws.format('A1:J1', {
            "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.6},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER"
        })
        
        try:
            sh.batch_update({
                "requests": [{
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": ws.id,
                            "hidden": True
                        },
                        "fields": "hidden"
                    }
                }]
            })
        except Exception:
            pass
        
        return ws


def save_draft_survey(
    spreadsheet_id: str,
    idpel: str,
    nama: str,
    lokasi: str,
    pekerjaan: str,
    ulp: str,
    no_spk: str,
    vendor: str,
    barang_data: List[Dict[str, Any]]
) -> Dict[str, Any]:
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        all_values = ws.get_all_values()
        
        existing_row = None
        for idx, row in enumerate(all_values[1:], start=2):
            if len(row) > 0 and str(row[0]).strip() == str(idpel).strip():
                existing_row = idx
                break
        
        barang_json = json.dumps(barang_data, ensure_ascii=False)
        tanggal_save = now_jakarta().strftime("%d/%m/%Y %H:%M:%S")
        
        if existing_row:
            ws.update(f'A{existing_row}:J{existing_row}', [[
                idpel, nama, lokasi, pekerjaan, ulp, no_spk, vendor,
                barang_json, tanggal_save, "Updated"
            ]])
            
            return {
                "success": True,
                "message": f"Data pelanggan {nama} berhasil diperbarui",
                "is_new": False
            }
        else:
            ws.append_row([
                idpel, nama, lokasi, pekerjaan, ulp, no_spk, vendor,
                barang_json, tanggal_save, "Not Yet"
            ])
            
            return {
                "success": True,
                "message": f"Data pelanggan {nama} berhasil disimpan",
                "is_new": True
            }
    
    except Exception as e:
        return {
            "success": False,
            "message": f"Error saat menyimpan draft: {str(e)}",
            "is_new": False
        }


@st.cache_data(ttl=60, show_spinner=False)
def load_all_drafts(spreadsheet_id: str) -> pd.DataFrame:
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        data = ws.get_all_records()
        
        if not data:
            return pd.DataFrame()
        
        df = pd.DataFrame(data)
        
        def parse_barang(json_str):
            try:
                if json_str and str(json_str).strip():
                    return json.loads(json_str)
                return []
            except Exception:
                return []
        
        df['Barang_List'] = df['Data Barang (JSON)'].apply(parse_barang)
        df['Jumlah_Item'] = df['Barang_List'].apply(len)
        
        try:
            df['_sort_date'] = pd.to_datetime(
                df['Tanggal Save'],
                format="%d/%m/%Y %H:%M:%S",
                errors='coerce'
            )
            df = df.sort_values('_sort_date', ascending=False)
            df = df.drop(columns=['_sort_date'])
        except Exception:
            pass
        
        return df
    
    except Exception as e:
        st.error(f"Error loading drafts: {str(e)}")
        return pd.DataFrame()


def load_single_draft(spreadsheet_id: str, idpel: str) -> Dict[str, Any]:
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        all_values = ws.get_all_values()
        
        for row in all_values[1:]:
            if len(row) > 0 and str(row[0]).strip() == str(idpel).strip():
                try:
                    barang = json.loads(row[7]) if len(row) > 7 and row[7] else []
                except Exception:
                    barang = []
                
                return {
                    "found": True,
                    "data": {
                        "idpel": row[0] if len(row) > 0 else "",
                        "nama": row[1] if len(row) > 1 else "",
                        "lokasi": row[2] if len(row) > 2 else "",
                        "pekerjaan": row[3] if len(row) > 3 else "",
                        "ulp": row[4] if len(row) > 4 else "",
                        "no_spk": row[5] if len(row) > 5 else "",
                        "vendor": row[6] if len(row) > 6 else "",
                        "barang": barang,
                        "tanggal_save": row[8] if len(row) > 8 else "",
                        "status_survey": row[9] if len(row) > 9 else ""
                    }
                }
        
        return {"found": False, "data": None}
    
    except Exception as e:
        return {"found": False, "data": None, "error": str(e)}


def delete_draft_survey(spreadsheet_id: str, idpel: str) -> Dict[str, Any]:
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        all_values = ws.get_all_values()
        
        for idx, row in enumerate(all_values[1:], start=2):
            if len(row) > 0 and str(row[0]).strip() == str(idpel).strip():
                ws.delete_rows(idx)
                return {
                    "success": True,
                    "message": f"Draft untuk ID Pelanggan {idpel} berhasil dihapus"
                }
        
        return {
            "success": False,
            "message": f"Draft untuk ID Pelanggan {idpel} tidak ditemukan"
        }
    
    except Exception as e:
        return {
            "success": False,
            "message": f"Error saat menghapus draft: {str(e)}"
        }


def count_drafts(spreadsheet_id: str) -> int:
    """Count total number of saved drafts."""
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        data = ws.get_all_records()
        return len(data)
    except Exception:
        return 0