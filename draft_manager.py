import json
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
from auth import get_gspread_client
import gspread

# Configure timezone for Jakarta or fallback to UTC+7.
try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

# Retrieve draft sheet name from secrets or use default.
try:
    DRAFT_SHEET_NAME = str(st.secrets.get("DRAFT_SHEET_NAME", "_DRAFT"))
except Exception:
    DRAFT_SHEET_NAME = "_DRAFT"


def get_or_create_draft_sheet(spreadsheet_id: str) -> gspread.Worksheet:
    # Get or create the hidden draft worksheet.
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    
    try:
        ws = sh.worksheet(DRAFT_SHEET_NAME)
        
        # Verify header structure and update if necessary.
        header = ws.row_values(1)
        
        if len(header) < 11 or "Foto Survey (JSON)" not in header:
            headers = [
                "ID Pelanggan",
                "Nama",
                "Lokasi",
                "Pekerjaan",
                "ULP",
                "No SPK",
                "Vendor Pelaksana",
                "Data Barang (JSON)",
                "Foto Survey (JSON)",
                "Tanggal Save",
                "Status Tanggal Survey"
            ]
            ws.update('A1:K1', [headers])
            
            ws.format('A1:K1', {
                "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.6},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "horizontalAlignment": "CENTER"
            })
        
        return ws
    
    except gspread.WorksheetNotFound:
        # Create new hidden sheet if not found.
        ws = sh.add_worksheet(title=DRAFT_SHEET_NAME, rows=1000, cols=11)
        
        headers = [
            "ID Pelanggan",
            "Nama",
            "Lokasi",
            "Pekerjaan",
            "ULP",
            "No SPK",
            "Vendor Pelaksana",
            "Data Barang (JSON)",
            "Foto Survey (JSON)",
            "Tanggal Save",
            "Status Tanggal Survey"
        ]
        ws.update('A1:K1', [headers])
        
        ws.format('A1:K1', {
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
    barang_data: List[Dict[str, Any]],
    foto_survey_links: Optional[List[Dict[str, str]]] = None
) -> Dict[str, Any]:
    # Save or update survey data in the draft sheet.
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        all_values = ws.get_all_values()
        
        existing_row = None
        for idx, row in enumerate(all_values[1:], start=2):
            if len(row) > 0 and str(row[0]).strip() == str(idpel).strip():
                existing_row = idx
                break
        
        barang_json = json.dumps(barang_data, ensure_ascii=False)
        foto_json = json.dumps(foto_survey_links or [], ensure_ascii=False)
        tanggal_save = now_jakarta().strftime("%d/%m/%Y %H:%M:%S")
        
        if existing_row:
            ws.update(f'A{existing_row}:K{existing_row}', [[
                idpel, nama, lokasi, pekerjaan, ulp, no_spk, vendor,
                barang_json, foto_json, tanggal_save, "Updated"
            ]])
            
            return {
                "success": True,
                "message": f"Data pelanggan {nama} berhasil diperbarui",
                "is_new": False
            }
        else:
            ws.append_row([
                idpel, nama, lokasi, pekerjaan, ulp, no_spk, vendor,
                barang_json, foto_json, tanggal_save, "Not Yet"
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
    # Load all draft records into a DataFrame.
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
        
        if 'Data Barang (JSON)' in df.columns:
            df['Barang_List'] = df['Data Barang (JSON)'].apply(parse_barang)
            df['Jumlah_Item'] = df['Barang_List'].apply(len)
        else:
            df['Barang_List'] = [[] for _ in range(len(df))]
            df['Jumlah_Item'] = 0
        
        def parse_foto(json_str):
            try:
                if json_str and str(json_str).strip():
                    return json.loads(json_str)
                return []
            except Exception:
                return []
        
        if 'Foto Survey (JSON)' in df.columns:
            df['Foto_List'] = df['Foto Survey (JSON)'].apply(parse_foto)
            df['Jumlah_Foto'] = df['Foto_List'].apply(len)
        else:
            df['Foto_List'] = [[] for _ in range(len(df))]
            df['Jumlah_Foto'] = 0
        
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
    # Retrieve a specific draft by customer ID.
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        all_values = ws.get_all_values()
        
        for row in all_values[1:]:
            if len(row) > 0 and str(row[0]).strip() == str(idpel).strip():
                try:
                    barang = json.loads(row[7]) if len(row) > 7 and row[7] else []
                except Exception:
                    barang = []
                
                try:
                    foto_survey = json.loads(row[8]) if len(row) > 8 and row[8] else []
                except Exception:
                    foto_survey = []
                
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
                        "foto_survey": foto_survey,
                        "tanggal_save": row[9] if len(row) > 9 else "",
                        "status_survey": row[10] if len(row) > 10 else ""
                    }
                }
        
        return {"found": False, "data": None}
    
    except Exception as e:
        return {"found": False, "data": None, "error": str(e)}


def delete_draft_survey(spreadsheet_id: str, idpel: str) -> Dict[str, Any]:
    # Delete a draft record by customer ID.
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
    # Count the total number of saved drafts.
    try:
        ws = get_or_create_draft_sheet(spreadsheet_id)
        data = ws.get_all_records()
        return len(data)
    except Exception:
        return 0


def sync_to_sheet5(spreadsheet_id: str, target_sheet_name: str = "Rekap_Material") -> Dict[str, Any]:
    # Synchronize draft data to Rekap_Material starting from row 2, respecting existing headers.
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(spreadsheet_id)
        
        draft_ws = get_or_create_draft_sheet(spreadsheet_id)
        draft_records = draft_ws.get_all_records()
        
        if not draft_records:
            return {"success": True, "message": "Tidak ada data draft untuk disinkronisasi."}

        # Define item keys strictly matching Sheet5 column order (Cols D - N).
        target_items_order = [
            "Jasa Kegiatan Geser APP",
            "Jasa Kegiatan Geser Perubahan Situasi SR",
            "Service wedge clamp 2/4 x 6/10 mm",
            "Strainhook / ekor babi",
            "Imundex klem",
            "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover",
            "Paku Beton",
            "Pole Bracket 3-9\"",
            "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover",
            "Twisted Cable 2 x 10 mm² - Al"
        ]
        
        rows_to_write = []
        
        for idx, row in enumerate(draft_records, start=1):
            # Parse Barang (Flattening)
            try:
                barang_json = row.get("Data Barang (JSON)", "[]")
                barang_list = json.loads(str(barang_json)) if barang_json else []
            except:
                barang_list = []
                
            barang_dict = {item.get("Rincian"): item.get("Vol") for item in barang_list}
            
            # Construct row data: [No, Nama, Idpel, Item1, Item2, ..., Item10]
            current_row = [
                idx,                                    # Col A: No (Auto Number)
                str(row.get("Nama", "")),               # Col B: Nama
                "'" + str(row.get("ID Pelanggan", "")), # Col C: Idpel (Force String)
            ]
            
            # Map volumes to columns D - M (Total 10 items)
            for item_name in target_items_order:
                vol = barang_dict.get(item_name, 0)
                current_row.append(str(vol) if vol > 0 else "")
            
            rows_to_write.append(current_row)
            
        try:
            target_ws = sh.worksheet(target_sheet_name)
            
            # Clear old data starting from row 2 (preserve headers).
            target_ws.batch_clear(["A2:N1000"]) 
            
            # Write new data starting at A2.
            target_ws.update('A2', rows_to_write)
            
            return {"success": True, "message": f"Berhasil sinkronisasi {len(draft_records)} data ke {target_sheet_name} (Header aman)"}
            
        except gspread.WorksheetNotFound:
            return {"success": False, "message": f"Sheet '{target_sheet_name}' tidak ditemukan. Mohon buat sheet tersebut terlebih dahulu."}
        
    except Exception as e:
        return {"success": False, "message": f"Gagal sinkronisasi: {str(e)}"}