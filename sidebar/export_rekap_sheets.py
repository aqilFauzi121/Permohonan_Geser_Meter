# export_rekap_sheets.py - Simplified with Template Formulas
import re
from datetime import datetime, timedelta
from typing import Optional, List, Any
import pandas as pd
import streamlit as st
from auth import get_gspread_client

# Timezone helper
try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

# Retention: keep latest 40 tabs
KEEP_LATEST_TABS = 40
_RE_REKAP = re.compile(r"^REKAP\s+.+?\s*-\s*(\d{8}[_-]\d{4})_(Vendor|Pelanggan)$")

def _parse_dt_from_title(title: str) -> Optional[datetime]:
    m = _RE_REKAP.match(title)
    if not m:
        return None
    ts = m.group(1).replace("-", "_")
    try:
        return datetime.strptime(ts, "%Y%m%d_%H%M")
    except Exception:
        return None

def cleanup_old_rekap(sh, keep_latest: int = KEEP_LATEST_TABS) -> None:
    candidates: List[tuple[Optional[datetime], Any]] = []
    for ws in sh.worksheets():
        if ws.title.startswith("REKAP "):
            dt = _parse_dt_from_title(ws.title)
            candidates.append((dt, ws))
    
    if len(candidates) <= keep_latest:
        return
    
    candidates.sort(key=lambda x: (x[0] is not None, x[0]), reverse=True)
    
    for _, ws in candidates[keep_latest:]:
        try:
            sh.del_worksheet(ws)
        except Exception:
            pass

# Template titles
TEMPLATE_VENDOR_TITLE = "Template Vendor"
TEMPLATE_PELANGGAN_TITLE = "Template Pelanggan"

try:
    TEMPLATE_VENDOR_TITLE = str(st.secrets.get("TEMPLATE_VENDOR_TITLE", "Template Vendor"))
    TEMPLATE_PELANGGAN_TITLE = str(st.secrets.get("TEMPLATE_PELANGGAN_TITLE", "Template Pelanggan"))
except Exception:
    pass

_N_BARIS_ITEM = 13  # Row 14-26

# Item mapping to template rows
TEMPLATE_ORDER = [
    "Jasa Kegiatan Geser APP",
    "Jasa Kegiatan Geser Perubahan Situasi SR",
    "Service wedge clamp 2/4 x 6/10 mm",
    "Strainhook / ekor babi",
    "Imundex klem",
    "Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover",
    "Paku Beton",
    "Pole Bracket 3-9\"",
    "Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover",
    "Segel Plastik",
    "Twisted Cable 2 x 10 mm² - Al",
    "Asuransi",
    "Twisted Cable 2x10 mm² - Al",
]

def _normalize(s: str) -> str:
    s = str(s or "").lower()
    s = s.replace("–", "-").replace("—", "-").replace(""", '"').replace(""", '"').replace("'", "'")
    s = s.replace("mm2", "mm²").replace("mm^2", "mm²")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# Aliases for item name variations
ALIASES = {
    _normalize("Jasa Kegiatan"): _normalize("Jasa Kegiatan Geser APP"),
    _normalize("Jasa Kegiatan Geser APP"): _normalize("Jasa Kegiatan Geser APP"),
    _normalize("Jasa Kegiatan Perubahan Situasi SR"): _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"),
    _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"): _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"),
    _normalize("Service wedge clamp 2/4 x 6/10 mm"): _normalize("Service wedge clamp 2/4 x 6/10 mm"),
    _normalize("Service wedge clamp 2/4 x 6/10mm"): _normalize("Service wedge clamp 2/4 x 6/10 mm"),
    _normalize("Strainhook / ekor babi"): _normalize("Strainhook / ekor babi"),
    _normalize("Strainthook / ekor babi"): _normalize("Strainhook / ekor babi"),
    _normalize("Strainhook / Ekor babi"): _normalize("Strainhook / ekor babi"),
    _normalize("Imundex klem"): _normalize("Imundex klem"),
    _normalize("Imundex Klem"): _normalize("Imundex klem"),
    _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"),
    _normalize("Conn. press AL/AL 10-16 mm² + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"),
    _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"),
    _normalize("Conn. press AL/AL 50-70 mm² + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"),
    _normalize('Pole Bracket 3-9"'): _normalize('Pole Bracket 3-9"'),
    _normalize("Twisted Cable 2 x 10 mm² - Al"): _normalize("Twisted Cable 2 x 10 mm² - Al"),
    _normalize("Twisted Cable 2 x 10 mm² – Al"): _normalize("Twisted Cable 2 x 10 mm² - Al"),
    _normalize("Twisted Cable 2x10 mm² - Al"): _normalize("Twisted Cable 2x10 mm² - Al"),
    _normalize("Twisted Cable 2x10 mm² – Al"): _normalize("Twisted Cable 2x10 mm² - Al"),
}

_TEMPLATE_INDEX = {_normalize(n): i for i, n in enumerate(TEMPLATE_ORDER)}

def _find_template_row_index(item_name: str) -> Optional[int]:
    key = ALIASES.get(_normalize(item_name), _normalize(item_name))
    return _TEMPLATE_INDEX.get(key)

def _to_int(v, default=0) -> int:
    try:
        x = pd.to_numeric(v, errors="coerce")
        if pd.isna(x):
            return int(default)
        return int(x)
    except Exception:
        return int(default)

def _to_sheet_values(grid: List[List[Optional[Any]]]) -> List[List[Any]]:
    out: List[List[Any]] = []
    for row in grid:
        v = row[0] if row else None
        out.append(["" if v is None else v])
    return out

def update_tanggal_survey(spreadsheet_id: str, gid: str, idpel: str) -> dict:
    try:
        now = now_jakarta()
        gc = get_gspread_client()
        sh = gc.open_by_key(spreadsheet_id)
        
        target_ws = None
        for ws in sh.worksheets():
            if str(ws.id) == str(gid):
                target_ws = ws
                break
        
        if target_ws is None:
            return {"success": False, "message": "Worksheet dengan GID tidak ditemukan", "row": 0, "col": 0}
        
        header = target_ws.row_values(1)
        
        tanggal_survey_col = None
        for idx, col_name in enumerate(header):
            normalized = str(col_name).strip().lower()
            if "tanggal survey" in normalized or "tanggalsurvey" in normalized:
                tanggal_survey_col = idx + 1
                break
        
        if tanggal_survey_col is None:
            return {"success": False, "message": f"Kolom 'Tanggal Survey' tidak ditemukan", "row": 0, "col": 0}
        
        id_pelanggan_col = None
        for idx, col_name in enumerate(header):
            normalized = str(col_name).strip().lower()
            if "id pelanggan" in normalized or "idpelanggan" in normalized:
                id_pelanggan_col = idx + 1
                break
        
        if id_pelanggan_col is None:
            return {"success": False, "message": "Kolom 'ID Pelanggan' tidak ditemukan", "row": 0, "col": 0}
        
        id_column_values = target_ws.col_values(id_pelanggan_col)
        
        matched_row_index = None
        for i in reversed(range(1, len(id_column_values))):
            if str(id_column_values[i]).strip() == str(idpel).strip():
                matched_row_index = i + 1
                break
        
        if matched_row_index is None:
            return {"success": False, "message": f"ID Pelanggan {idpel} tidak ditemukan di sheet", "row": 0, "col": 0}
        
        timestamp_str = now.strftime("%d/%m/%Y %H:%M:%S")
        target_ws.update_cell(matched_row_index, tanggal_survey_col, timestamp_str)
        
        return {
            "success": True,
            "message": f"Berhasil update row {matched_row_index}, col {tanggal_survey_col} dengan waktu WIB: {timestamp_str}",
            "row": matched_row_index,
            "col": tanggal_survey_col
        }
    except Exception as e:
        return {"success": False, "message": f"Error: {str(e)}", "row": 0, "col": 0}

def export_rekap_to_sheet(
    spreadsheet_id: str,
    sheet_title: str,
    meta: dict,
    df_pilih: pd.DataFrame,
    template_title: str,
):
    """Export rekap with template formulas (only fill Identitas + Volume)"""
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)
    
    # Find template worksheet
    template_ws = None
    all_sheets = []
    for ws in sh.worksheets():
        all_sheets.append(ws.title)
        if ws.title == template_title:
            template_ws = ws
            break
    
    if template_ws is None:
        raise RuntimeError(
            f"Template '{template_title}' tidak ditemukan!\n"
            f"Available sheets: {', '.join(all_sheets)}"
        )
    
    template_id = template_ws.id
    
    # Duplicate template
    dup_result = sh.batch_update({
        "requests": [{
            "duplicateSheet": {
                "sourceSheetId": template_id,
                "insertSheetIndex": 0,
                "newSheetName": sheet_title,
            }
        }]
    })
    new_sheet_id = dup_result["replies"][0]["duplicateSheet"]["properties"]["sheetId"]
    
    # Identitas (C3:C8)
    identitas = [
        [meta.get("Pekerjaan", "-")],
        [meta.get("Nama", "-")],
        [meta.get("Lokasi", "-")],
        [meta.get("ULP", "-")],
        [meta.get("No SPK", "-")],
        [meta.get("Vendor", "-")],
    ]
    
    # Volume (C14:C26)
    if df_pilih is None or df_pilih.empty:
        df_norm = pd.DataFrame(columns=["Rincian", "Vol"])
    else:
        df_norm = df_pilih.copy()
    
    vol_values: List[List[Optional[int]]] = [[None] for _ in range(_N_BARIS_ITEM)]
    
    for _, row in df_norm.iterrows():
        nama_raw = str(row.get("Rincian", "")).strip()
        if not nama_raw:
            continue
        
        idx = _find_template_row_index(nama_raw)
        if idx is None or idx >= _N_BARIS_ITEM:
            continue
        
        qty = _to_int(row.get("Vol", 0))
        if qty > 0:
            vol_values[idx][0] = qty
    
    # Batch update: only 2 ranges
    payload = {
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": f"'{sheet_title}'!C3:C8", "values": identitas},
            {"range": f"'{sheet_title}'!C14:C26", "values": _to_sheet_values(vol_values)},
        ],
    }
    
    try:
        sh.values_batch_update(body=payload)
    except Exception:
        for item in payload["data"]:
            sh.values_update(
                item["range"],
                params={"valueInputOption": "USER_ENTERED"},
                body={"values": item["values"]},
            )
    
    return {
        "sheet_title": sheet_title,
        "new_sheet_id": new_sheet_id,
    }

def export_rekap_pair(
    spreadsheet_id: str,
    base_sheet_title_vendor: str,
    base_sheet_title_pelanggan: str,
    meta: dict,
    df_pilih: pd.DataFrame,
    idpel: Optional[str] = None,
    gid: Optional[str] = None,
):
    """Export Vendor + Pelanggan sheets with different templates"""
    info_vendor = export_rekap_to_sheet(
        spreadsheet_id=spreadsheet_id,
        sheet_title=base_sheet_title_vendor,
        meta=meta,
        df_pilih=df_pilih,
        template_title=TEMPLATE_VENDOR_TITLE,
    )
    
    info_pelanggan = export_rekap_to_sheet(
        spreadsheet_id=spreadsheet_id,
        sheet_title=base_sheet_title_pelanggan,
        meta=meta,
        df_pilih=df_pilih,
        template_title=TEMPLATE_PELANGGAN_TITLE,
    )
    
    sh = get_gspread_client().open_by_key(spreadsheet_id)
    cleanup_old_rekap(sh, keep_latest=KEEP_LATEST_TABS)
    
    survey_result = {"success": False, "message": "Parameter tidak lengkap"}
    if idpel is not None and gid is not None:
        survey_result = update_tanggal_survey(spreadsheet_id, gid, idpel)
    
    return {
        "vendor": info_vendor,
        "pelanggan": info_pelanggan,
        "survey_result": survey_result
    }