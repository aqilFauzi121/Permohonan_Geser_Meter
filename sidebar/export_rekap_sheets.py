# export_rekap_sheets.py
import re
from datetime import datetime
from typing import Optional, List, Any

import pandas as pd
import streamlit as st
from auth import get_gspread_client

# ==============================
#  TIMEZONE HELPER
# ==============================
try:
    from zoneinfo import ZoneInfo
    def now_jakarta():
        return datetime.now(tz=ZoneInfo("Asia/Jakarta"))
except Exception:
    from datetime import timedelta
    def now_jakarta():
        return datetime.utcnow() + timedelta(hours=7)

# ==============================
#  Konfigurasi Retention
# ==============================
KEEP_LATEST_TABS = 40

_RE_REKAP = re.compile(
    r"^REKAP\s+.+?\s*-\s*(\d{8}[_-]\d{4})_(Vendor|Pelanggan)$"
)

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

# ==============================
#  Konstanta Template - FINAL FIX
# ==============================
_TEMPLATE_CANDIDATES = ("Template", "Sheet1")
_N_BARIS_ITEM = 15
_RESERVED_TOP_ROWS = 2

TEMPLATE_ORDER = [
    "Jasa Kegiatan Geser APP",
    "Jasa Kegiatan Geser Perubahan Situasi SR",
    "Service wedge clamp 2/4 x 6/10mm",
    "Strainhook / Ekor babi",
    "Imundex Klem",
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
    s = s.replace("-", "-").replace("—", "-").replace(""", '"').replace(""", '"').replace("'", "'")
    s = s.replace("mm2", "mm²").replace("mm^2", "mm²")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

ALIASES = {
    _normalize("Jasa Kegiatan"): _normalize("Jasa Kegiatan Geser APP"),
    _normalize("Jasa Kegiatan Geser APP"): _normalize("Jasa Kegiatan Geser APP"),
    _normalize("Jasa Kegiatan Perubahan Situasi SR"): _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"),
    _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"): _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"),
    
    _normalize("Service wedge clamp 2/4 x 6/10 mm"): _normalize("Service wedge clamp 2/4 x 6/10mm"),
    
    _normalize("Strainhook / ekor babi"): _normalize("Strainhook / Ekor babi"),
    _normalize("Strainthook / ekor babi"): _normalize("Strainhook / Ekor babi"),
    
    _normalize("Imundex Klem"): _normalize("Imundex Klem"),
    
    _normalize("Cable support (50/80J/2009)"): _normalize("Cable support (508/U/2009)"),
    
    _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"),
    _normalize("Conn. press AL/AL 10-16 mm² + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"),
    _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"),
    _normalize("Conn. press AL/AL 50-70 mm² + Scoot + Cover"): _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"),
    
    _normalize('Pole Bracket 3-9"'): _normalize('Pole Bracket 3-9"'),

    _normalize("Twisted Cable 2 x 10 mm² - Al"): _normalize("Twisted Cable 2 x 10 mm² - Al"),
    _normalize("Twisted Cable 2x10 mm² - Al"): _normalize("Twisted Cable 2x10 mm² - Al"),
}
_TEMPLATE_INDEX = { _normalize(n): i for i, n in enumerate(TEMPLATE_ORDER) }

def _find_template_row_index(item_name: str) -> Optional[int]:
    key = ALIASES.get(_normalize(item_name), _normalize(item_name))
    return _TEMPLATE_INDEX.get(key)

def _to_int(v, default=0) -> int:
    try:
        x = pd.to_numeric(v, errors="coerce")
        if pd.isna(x): return int(default)
        return int(x)
    except Exception:
        return int(default)

def _find_template_worksheet(sh, preferred_title: str = "Template"):
    names = []
    if preferred_title: names.append(preferred_title)
    for n in _TEMPLATE_CANDIDATES:
        if n not in names: names.append(n)
    for name in names:
        try:
            return sh.worksheet(name)
        except Exception:
            continue
    raise RuntimeError("Template sheet tidak ditemukan. Buat tab 'Template' atau 'Sheet1'.")

# ==============================
#  Tabel Harga - FINAL FIX
# ==============================
DEFAULT_PRICE_VENDOR = {
    _normalize("Jasa Kegiatan Geser APP"): 93000,
    _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"): 79000,
    _normalize("Service wedge clamp 2/4 x 6/10 mm"): 3990,
    _normalize("Strainhook / Ekor babi"): 8000,
    _normalize("Imundex Klem"): 454,
    _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"): 11999,
    _normalize("Paku Beton"): 74,
    _normalize('Pole Bracket 3-9"'): 36823,
    _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"): 29400,
    _normalize("Segel Plastik"): 1754,
    _normalize("Twisted Cable 2 x 10 mm² - Al"): 4339,
    _normalize("Twisted Cable 2x10 mm² - Al"): 0,
    _normalize("Asuransi"): 0,
}

DEFAULT_PRICE_PELANGGAN = {
    _normalize("Jasa Kegiatan Geser APP"): 103230,
    _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"): 87690,
    _normalize("Service wedge clamp 2/4 x 6/10 mm"): 4428.9,
    _normalize("Strainhook / Ekor babi"): 8880,
    _normalize("Imundex Klem"): 503.94,
    _normalize("Conn. press AL/AL type 10-16 mm2 / 10-16 mm2 + Scoot + Cover"): 13318.89,
    _normalize("Paku Beton"): 82.14,
    _normalize('Pole Bracket 3-9"'): 40873.53,
    _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"): 32634,
    _normalize("Segel Plastik"): 1946.94,
    _normalize("Twisted Cable 2 x 10 mm² - Al"): 4816.29,
    _normalize("Twisted Cable 2x10 mm² - Al"): 0,
    _normalize("Asuransi"): 0,
}

def _load_price_from_secrets(key: str) -> dict:
    try:
        data = st.secrets.get(key, None)
        if not data: return {}
        out = {}
        for k, v in dict(data).items():
            out[_normalize(k)] = _to_int(v, 0)
        return out
    except Exception:
        return {}

def _resolve_prices():
    vendor_over = _load_price_from_secrets("PRICE_TABLE_VENDOR")
    pelanggan_over = _load_price_from_secrets("PRICE_TABLE_PELANGGAN")

    price_vendor = DEFAULT_PRICE_VENDOR.copy()
    price_vendor.update(vendor_over)

    price_pelanggan = DEFAULT_PRICE_PELANGGAN.copy()
    price_pelanggan.update(pelanggan_over)

    return {"vendor": price_vendor, "pelanggan": price_pelanggan}

# ==============================
#  Item Kategorisasi
# ==============================
PLN_ONLY_NAMES = {
    _normalize("Conn. press AL/AL type 10-16 mm2 / 50-70 mm2 + Scoot + Cover"),
    _normalize("Segel Plastik"),
    _normalize("Twisted Cable 2 x 10 mm² - Al"),
    _normalize("Twisted Cable 2x10 mm² - Al"),
    _normalize("Asuransi"),
}

JASA_ITEMS = {
    _normalize("Jasa Kegiatan Geser APP"),
    _normalize("Jasa Kegiatan Geser Perubahan Situasi SR"),
}

# ==============================
#  Util konversi grid
# ==============================
def _to_sheet_values(grid: List[List[Optional[int]]]) -> List[List[Any]]:
    out: List[List[Any]] = []
    for row in grid:
        v = row[0] if row else None
        out.append(["" if v is None else v])
    return out

# ==============================
#  Update Tanggal Survey - FIXED TIMEZONE
# ==============================
def update_tanggal_survey(spreadsheet_id: str, gid: str, idpel: str) -> dict:
    try:
        # Timezone helper di dalam function
        try:
            from zoneinfo import ZoneInfo
            now = datetime.now(tz=ZoneInfo("Asia/Jakarta"))
        except Exception:
            from datetime import timedelta
            now = datetime.utcnow() + timedelta(hours=7)
        
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
            return {"success": False, "message": f"Kolom 'Tanggal Survey' tidak ditemukan. Header: {header}", "row": 0, "col": 0}
        
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

# ==============================
#  Ekspor (satu sheet)
# ==============================
def export_rekap_to_sheet(
    spreadsheet_id: str,
    sheet_title: str,
    meta: dict,
    df_pilih: pd.DataFrame,
    template_title: str = "Template",
    price_profile: str = "vendor",
):
    price_profiles = _resolve_prices()
    price_table = price_profiles.get(price_profile, DEFAULT_PRICE_VENDOR)

    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_id)

    template_ws = _find_template_worksheet(sh, preferred_title=template_title)
    template_id = template_ws.id

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

    identitas = [
        [meta.get("Pekerjaan", "-")],
        [meta.get("Nama", "-")],
        [meta.get("Lokasi", "-")],
        [meta.get("ULP", "-")],
        [meta.get("No SPK", "-")],
        [meta.get("Vendor", "-")],
    ]

    if df_pilih is None or df_pilih.empty:
        df_norm = pd.DataFrame(columns=["Rincian", "SAT", "Vol"])
    else:
        df_norm = df_pilih.rename(columns={"Harga_Satuan_Material": "Harga Satuan Material"}).copy()

    vol_values:   List[List[Optional[int]]] = [[None] for _ in range(_N_BARIS_ITEM)]
    price_pln:    List[List[Optional[int]]] = [[None] for _ in range(_N_BARIS_ITEM)]
    price_tunai:  List[List[Optional[int]]] = [[None] for _ in range(_N_BARIS_ITEM)]
    price_jasa:   List[List[Optional[int]]] = [[None] for _ in range(_N_BARIS_ITEM)]

    subtotal = 0

    for _, row in df_norm.iterrows():
        nama_raw = str(row.get("Rincian", "")).strip()
        if not nama_raw:
            continue

        idx = _find_template_row_index(nama_raw)
        if idx is None:
            continue

        target_idx = _RESERVED_TOP_ROWS + idx
        if target_idx >= _N_BARIS_ITEM:
            continue

        qty = _to_int(row.get("Vol", 0))
        name_key = ALIASES.get(_normalize(nama_raw), _normalize(nama_raw))
        price = price_table.get(name_key, _to_int(row.get("Harga Satuan Material", 0)))

        if qty > 0:
            vol_values[target_idx][0] = qty

        is_jasa = name_key in JASA_ITEMS
        is_pln_only = name_key in PLN_ONLY_NAMES

        if price > 0:
            if is_jasa:
                price_jasa[target_idx][0] = price
            elif is_pln_only:
                price_pln[target_idx][0] = price
            else:
                price_tunai[target_idx][0] = price

        subtotal += qty * price

    ppn = int(round(subtotal * 0.11))
    total_biaya = int(subtotal + ppn)

    payload = {
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": f"'{sheet_title}'!C3:C8",   "values": identitas},
            {"range": f"'{sheet_title}'!C12:C26", "values": _to_sheet_values(vol_values)},
            {"range": f"'{sheet_title}'!D12:D26", "values": _to_sheet_values(price_pln)},
            {"range": f"'{sheet_title}'!E12:E26", "values": _to_sheet_values(price_tunai)},
            {"range": f"'{sheet_title}'!G12:G26", "values": _to_sheet_values(price_jasa)},  # ✅ TAMBAH INI!
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
        "subtotal": int(subtotal),
        "ppn": ppn,
        "total_after_ppn": total_biaya,
        "new_sheet_id": new_sheet_id,
    }

# ==============================
#  Ekspor (dua sheet + Retention + Update Tanggal Survey)
# ==============================
def export_rekap_pair(
    spreadsheet_id: str,
    base_sheet_title_vendor: str,
    base_sheet_title_pelanggan: str,
    meta: dict,
    df_pilih: pd.DataFrame,
    template_title: str = "Template",
    idpel: Optional[str] = None,
    gid: Optional[str] = None,
):
    info_vendor = export_rekap_to_sheet(
        spreadsheet_id=spreadsheet_id,
        sheet_title=base_sheet_title_vendor,
        meta=meta,
        df_pilih=df_pilih,
        template_title=template_title,
        price_profile="vendor",
    )
    info_pelanggan = export_rekap_to_sheet(
        spreadsheet_id=spreadsheet_id,
        sheet_title=base_sheet_title_pelanggan,
        meta=meta,
        df_pilih=df_pilih,
        template_title=template_title,
        price_profile="pelanggan",
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