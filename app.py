import streamlit as st
import pandas as pd
import os
import re
import math
import unicodedata
from io import BytesIO
from copy import copy

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import coordinate_to_tuple

APP_VERSION = "2026-02-02_dynamic_sections_safe_merged_cells"

# ----------------- CONFIG APP -----------------
st.set_page_config(page_title="FT Grands Comptes", page_icon="🚚", layout="wide")
st.title("Generateur de Fiches Techniques Grands Comptes")
st.caption("Version de test basée sur bdd_CG.xlsx")
st.sidebar.info(f"✅ Version: {APP_VERSION}")

st.sidebar.markdown("### Template Excel")
uploaded_template = st.sidebar.file_uploader(
    "Uploader le template (xlsx) à utiliser",
    type=["xlsx"]
)

TEMPLATE_FALLBACK = "FT_Grands_Comptes.xlsx"
IMG_ROOT = "images"

# ----------------- HELPERS -----------------
def _strip_accents(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = _strip_accents(s)
    s = s.lower()
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = s.replace("–", "-").replace("—", "-")
    s = " ".join(s.split())
    return s

def get_col(df: pd.DataFrame, wanted: str):
    if df is None or wanted is None:
        return None
    w = _norm(wanted)
    for c in df.columns:
        if _norm(c) == w:
            return c
    for c in df.columns:
        if w in _norm(c):
            return c
    return None

def clean_unique_list(series: pd.Series):
    if series is None:
        return []
    s = series.dropna().astype(str).map(lambda x: x.strip())
    s = s[(s != "") & (s.str.lower() != "nan")]
    return sorted(s.unique().tolist())

def extract_pf_key(code_pf: str):
    if not isinstance(code_pf, str) or code_pf.strip() == "":
        return ""
    return code_pf.split(" - ")[0].strip()

def choose_codes(prod_choice, opt_choice, veh_prod, veh_opt):
    prod_code = veh_prod
    opt_code = veh_opt
    if isinstance(prod_choice, str) and prod_choice not in (None, "", "Tous"):
        prod_code = prod_choice
    if isinstance(opt_choice, str) and opt_choice not in (None, "", "Tous"):
        opt_code = opt_choice
    return prod_code, opt_code

# ----------------- IMAGES -----------------
def resolve_image_path(cell_value, subdir):
    if not isinstance(cell_value, str) or not cell_value.strip():
        return None
    val = cell_value.strip()
    if val.lower().startswith(("http://", "https://")):
        return val
    val = val.replace("\\", "/")
    filename = os.path.basename(val)
    return os.path.join(IMG_ROOT, subdir, filename)

def show_image(path_or_url, caption):
    st.caption(caption)
    if not path_or_url:
        st.info("Pas d'image définie")
        return
    if isinstance(path_or_url, str) and path_or_url.lower().startswith(("http://", "https://")):
        st.image(path_or_url)
        return
    if os.path.exists(path_or_url):
        st.image(path_or_url)
    else:
        st.warning(f"Image introuvable : {path_or_url}")

# ----------------- LOAD DATA -----------------
@st.cache_data
def load_data():
    xls = pd.ExcelFile("bdd_CG.xlsx")
    vehicules = pd.read_excel(xls, "FS_referentiel_produits_std_Ver")
    cabines = pd.read_excel(xls, "CABINES")
    moteurs = pd.read_excel(xls, "MOTEURS")
    chassis = pd.read_excel(xls, "CHASSIS")
    caisses = pd.read_excel(xls, "CAISSES")
    frigo = pd.read_excel(xls, "FRIGO")
    hayons = pd.read_excel(xls, "HAYONS")
    return vehicules, cabines, moteurs, chassis, caisses, frigo, hayons

# ----------------- FILTERS -----------------
def filtre_select(df, col_wanted, label):
    col = get_col(df, col_wanted)
    if col is None:
        st.sidebar.write(f"(colonne '{col_wanted}' absente)")
        return df, None

    options = clean_unique_list(df[col])
    choix = st.sidebar.selectbox(label, ["Tous"] + options)

    if choix != "Tous":
        df = df[df[col].astype(str).str.strip() == str(choix).strip()]

    return df, choix

filtre_select_options = filtre_select

def format_vehicule(row):
    champs = []
    for c in ["code_pays", "Marque", "Modele", "Code_PF", "Standard_PF"]:
        if c in row and pd.notna(row[c]):
            champs.append(str(row[c]))
    return " – ".join(champs)

# ----------------- EXCEL HELPERS -----------------
def _find_title_row(ws, include_keywords, exclude_keywords=None, max_col_letter="L"):
    inc = [_norm(k) for k in include_keywords]
    exc = [_norm(k) for k in (exclude_keywords or [])]
    max_col = column_index_from_string(max_col_letter)

    for r in range(1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            tv = _norm(v)
            ok_inc = all(k in tv for k in inc)
            ok_exc = all(k not in tv for k in exc)
            if ok_inc and ok_exc:
                return r
    return None

def _region_rows(ws, start_row, next_row):
    if start_row is None:
        return []
    start = start_row + 1
    end = (next_row - 1) if (next_row is not None) else ws.max_row
    if end < start:
        return []
    return list(range(start, end + 1))

def _merged_range_including(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng
    return None

def safe_set(ws, addr, value):
    """Ecrit même si la cellule est fusionnée (écrit dans la top-left)."""
    r, c = coordinate_to_tuple(addr)
    rng = _merged_range_including(ws, r, c)
    if rng:
        ws.cell(rng.min_row, rng.min_col).value = value
    else:
        ws.cell(r, c).value = value

def _unmerge_overlapping_row(ws, row, c1, c2):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row == row and rng.max_row == row:
            if not (rng.max_col < c1 or rng.min_col > c2):
                try:
                    ws.unmerge_cells(str(rng))
                except Exception:
                    pass

def _ensure_merge_row(ws, row, c1, c2):
    if c2 <= c1:
        return
    _unmerge_overlapping_row(ws, row, c1, c2)
    try:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    except Exception:
        pass

def _write_merged(ws, row, col, value):
    rng = _merged_range_including(ws, row, col)
    r0, c0 = (rng.min_row, rng.min_col) if rng else (row, col)
    cell = ws.cell(r0, c0)
    if isinstance(cell, MergedCell):
        return
    cell.value = value
    cell.alignment = Alignment(wrap_text=True, vertical="top")

def _style_source_cell(ws, row, col):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        rng = _merged_range_including(ws, row, col)
        if rng:
            return ws.cell(rng.min_row, rng.min_col)
    return cell

def _snapshot_row_style(ws, src_row, max_col):
    snap = []
    for c in range(1, max_col + 1):
        src = _style_source_cell(ws, src_row, c)
        snap.append({
            "font": copy(src.font),
            "border": copy(src.border),
            "fill": copy(src.fill),
            "number_format": src.number_format,
            "protection": copy(src.protection),
            "alignment": copy(src.alignment),
        })
    height = ws.row_dimensions[src_row].height
    return snap, height

def _insert_rows_with_style(ws, insert_at_row, n_rows, style_src_row, max_col_letter="L"):
    if n_rows <= 0:
        return
    max_col = column_index_from_string(max_col_letter)

    # snapshot AVANT insertion
    snap, src_height = _snapshot_row_style(ws, style_src_row, max_col)

    ws.insert_rows(insert_at_row, amount=n_rows)

    for i in range(n_rows):
        r = insert_at_row + i
        ws.row_dimensions[r].height = src_height
        for c in range(1, max_col + 1):
            tgt = ws.cell(r, c)
            tgt.value = None
            stl = snap[c - 1]
            tgt.font = copy(stl["font"])
            tgt.border = copy(stl["border"])
            tgt.fill = copy(stl["fill"])
            tgt.number_format = stl["number_format"]
            tgt.protection = copy(stl["protection"])
            tgt.alignment = copy(stl["alignment"])

def _ensure_section_capacity(ws, title_row, next_title_row, rows_needed, max_col_letter="L", hard_cap_extra=300):
    """
    Insère des lignes JUSTE AVANT next_title_row pour agrandir la section.
    """
    if title_row is None or rows_needed is None or rows_needed <= 0:
        return 0

    rows = _region_rows(ws, title_row, next_title_row)
    current = len(rows)

    if rows_needed <= current:
        return 0

    extra = min(rows_needed - current, hard_cap_extra)
    insert_at = next_title_row if next_title_row is not None else (ws.max_row + 1)

    # ligne modèle : 1ère ligne de la zone, sinon title_row+1 si possible
    style_src_row = rows[0] if rows else (title_row + 1 if title_row + 1 <= ws.max_row else title_row)

    _insert_rows_with_style(ws, insert_at, extra, style_src_row, max_col_letter=max_col_letter)
    return extra

def fill_region(ws, rows, values, start_cols, end_cols_override):
    if not rows:
        return 0, 0
    if not values:
        values = []

    # force merges puis écrit
    for r in rows:
        for sc in start_cols:
            endc = end_cols_override.get(sc, sc)
            _ensure_merge_row(ws, r, sc, endc)

    capacity = len(rows) * len(start_cols)
    n = min(len(values), capacity)

    i = 0
    for r in rows:
        for sc in start_cols:
            if i >= n:
                return n, capacity
            _write_merged(ws, r, sc, values[i])
            i += 1
    return n, capacity

# ----------------- DATA BUILDERS -----------------
def build_values(row, code_col):
    if row is None:
        return []
    vals = []
    for col, val in row.items():
        if str(col).strip() == str(code_col).strip():
            continue
        if pd.isna(val) or str(val).strip() == "":
            continue
        name_lower = _norm(col)
        if ("produit" in name_lower and "option" in name_lower) or name_lower.startswith("zone libre"):
            continue
        if str(col).strip() == "_":
            continue
        vals.append(str(val).strip())
    return vals

def find_row(df, code, code_col_wanted, code_pf_fallback=None, prefer_po=None):
    if not isinstance(code, str) or code.strip() == "" or code == "Tous":
        code = ""

    code_col = get_col(df, code_col_wanted) or df.columns[0]

    po_col = None
    for c in df.columns:
        if "produit" in _norm(c) and "option" in _norm(c):
            po_col = c
            break

    if code:
        cand = df[df[code_col].astype(str).str.strip() == code.strip()]
        if not cand.empty:
            if prefer_po and po_col:
                cand_po = cand[cand[po_col].astype(str).str.strip().str.upper() == prefer_po.upper()]
                return cand_po.iloc[0] if not cand_po.empty else cand.iloc[0]
            return cand.iloc[0]

    if isinstance(code_pf_fallback, str) and code_pf_fallback.strip():
        key = extract_pf_key(code_pf_fallback)
        if key:
            cand = df[df[code_col].astype(str).str.contains(re.escape(key), na=False)]
            if not cand.empty:
                if prefer_po and po_col:
                    cand_po = cand[cand[po_col].astype(str).str.strip().str.upper() == prefer_po.upper()]
                    return cand_po.iloc[0] if not cand_po.empty else cand.iloc[0]
                return cand.iloc[0]

    return None

# ----------------- EXCEL GENERATION (DYNAMIC) -----------------
def genere_ft_excel_dynamic(
    veh,
    cab_prod_choice, cab_opt_choice,
    mot_prod_choice, mot_opt_choice,
    ch_prod_choice, ch_opt_choice,
    caisse_prod_choice, caisse_opt_choice,
    gf_prod_choice, gf_opt_choice,
    hay_prod_choice, hay_opt_choice,
    cabines, moteurs, chassis, caisses, frigo, hayons,
    template_bytes=None,
    template_path=TEMPLATE_FALLBACK,
    debug=False,
):
    # Load template
    if template_bytes:
        wb = load_workbook(BytesIO(template_bytes), read_only=False, data_only=False)
    else:
        if not os.path.exists(template_path):
            st.error(f"Template introuvable : {template_path}")
            return None, {}
        wb = load_workbook(template_path, read_only=False, data_only=False)

    ws = wb["date"] if "date" in wb.sheetnames else wb[wb.sheetnames[0]]
    ws.print_title_rows = None

    # ---------- build values ----------
    code_pf_ref = veh.get("Code_PF", "")

    cab_prod_code, cab_opt_code = choose_codes(cab_prod_choice, cab_opt_choice, veh.get("C_Cabine"), veh.get("C_Cabine-OPTIONS"))
    mot_prod_code, mot_opt_code = choose_codes(mot_prod_choice, mot_opt_choice, veh.get("M_moteur"), veh.get("M_moteur-OPTIONS"))
    ch_prod_code, ch_opt_code = choose_codes(ch_prod_choice, ch_opt_choice, veh.get("C_Chassis"), veh.get("C_Chassis-OPTIONS"))
    caisse_prod_code, caisse_opt_code = choose_codes(caisse_prod_choice, caisse_opt_choice, veh.get("C_Caisse"), veh.get("C_Caisse-OPTIONS"))
    gf_prod_code, gf_opt_code = choose_codes(gf_prod_choice, gf_opt_choice, veh.get("C_Groupe frigo"), veh.get("C_Groupe frigo-OPTIONS"))
    hay_prod_code, hay_opt_code = choose_codes(hay_prod_choice, hay_opt_choice, veh.get("C_Hayon elevateur"), veh.get("C_Hayon elevateur-OPTIONS"))

    cab_prod_row = find_row(cabines, cab_prod_code, "C_Cabine", code_pf_fallback=code_pf_ref, prefer_po="P")
    cab_opt_row  = find_row(cabines, cab_opt_code,  "C_Cabine", code_pf_fallback=code_pf_ref, prefer_po="O")

    mot_prod_row = find_row(moteurs, mot_prod_code, "M_moteur", code_pf_fallback=code_pf_ref, prefer_po="P")
    mot_opt_row  = find_row(moteurs, mot_opt_code,  "M_moteur", code_pf_fallback=code_pf_ref, prefer_po="O")

    ch_prod_row  = find_row(chassis, ch_prod_code, "CH_chassis", code_pf_fallback=code_pf_ref, prefer_po="P")
    ch_opt_row   = find_row(chassis, ch_opt_code,  "CH_chassis", code_pf_fallback=code_pf_ref, prefer_po="O")

    caisse_prod_row = find_row(caisses, caisse_prod_code, "CF_caisse", code_pf_fallback=code_pf_ref, prefer_po="P")
    caisse_opt_row  = find_row(caisses, caisse_opt_code,  "CF_caisse", code_pf_fallback=code_pf_ref, prefer_po="O")

    gf_prod_row = find_row(frigo, gf_prod_code, "GF_groupe frigo", code_pf_fallback=code_pf_ref, prefer_po="P")
    gf_opt_row  = find_row(frigo, gf_opt_code,  "GF_groupe frigo", code_pf_fallback=code_pf_ref, prefer_po="O")

    hay_prod_row = find_row(hayons, hay_prod_code, "HL_hayon elevateur", code_pf_fallback=code_pf_ref, prefer_po="P")
    hay_opt_row  = find_row(hayons, hay_opt_code,  "HL_hayon elevateur", code_pf_fallback=code_pf_ref, prefer_po="O")

    cab_codecol = get_col(cabines, "C_Cabine") or cabines.columns[0]
    mot_codecol = get_col(moteurs, "M_moteur") or moteurs.columns[0]
    ch_codecol  = get_col(chassis, "CH_chassis") or chassis.columns[0]
    caisse_codecol = get_col(caisses, "CF_caisse") or caisses.columns[0]
    gf_codecol = get_col(frigo, "GF_groupe frigo") or frigo.columns[0]
    hay_codecol = get_col(hayons, "HL_hayon elevateur") or hayons.columns[0]

    cab_vals = build_values(cab_prod_row, cab_codecol)
    cab_opt_vals = build_values(cab_opt_row, cab_codecol)
    mot_vals = build_values(mot_prod_row, mot_codecol)
    mot_opt_vals = build_values(mot_opt_row, mot_codecol)
    ch_vals = build_values(ch_prod_row, ch_codecol)
    ch_opt_vals = build_values(ch_opt_row, ch_codecol)

    caisse_vals = build_values(caisse_prod_row, caisse_codecol)
    caisse_opt_vals = build_values(caisse_opt_row, caisse_codecol)
    gf_vals = build_values(gf_prod_row, gf_codecol)
    gf_opt_vals = build_values(gf_opt_row, gf_codecol)
    hay_vals = build_values(hay_prod_row, hay_codecol)
    hay_opt_vals = build_values(hay_opt_row, hay_codecol)

    # ---------- locate sections (initial) ----------
    cab_row = _find_title_row(ws, ["cabine"], exclude_keywords=["options"])
    cab_opt_row_t = _find_title_row(ws, ["cabine", "options"])
    car_row = _find_title_row(ws, ["carrosserie"], exclude_keywords=["options"])
    car_opt_row_t = _find_title_row(ws, ["carrosserie", "options"])
    fr_row = _find_title_row(ws, ["groupe", "frigorifique"], exclude_keywords=["options"])
    fr_opt_row_t = _find_title_row(ws, ["groupe", "frigorifique", "options"])
    hy_row = _find_title_row(ws, ["hayon"], exclude_keywords=["options"])
    hy_opt_row_t = _find_title_row(ws, ["hayon", "options"])
    pub_row = _find_title_row(ws, ["publicite"])  # publicite / publicité

    # ---------- compute needed rows ----------
    cab_needed = max(len(cab_vals), len(mot_vals), len(ch_vals), 0)
    cab_opt_needed = max(len(cab_opt_vals), len(mot_opt_vals), len(ch_opt_vals), 0)

    car_needed = len(caisse_vals)
    car_opt_needed = len(caisse_opt_vals)

    fr_needed = int(math.ceil(len(gf_vals) / 2.0)) if gf_vals else 0
    fr_opt_needed = len(gf_opt_vals)

    hy_needed = int(math.ceil(len(hay_vals) / 2.0)) if hay_vals else 0
    hy_opt_needed = len(hay_opt_vals)

    # ---------- expand sections (BOTTOM -> TOP) ----------
    extras = {}
    extras["hy_opt"] = _ensure_section_capacity(ws, hy_opt_row_t, pub_row, hy_opt_needed)
    extras["hy"]     = _ensure_section_capacity(ws, hy_row,     hy_opt_row_t, hy_needed)
    extras["fr_opt"] = _ensure_section_capacity(ws, fr_opt_row_t, hy_row, fr_opt_needed)
    extras["fr"]     = _ensure_section_capacity(ws, fr_row,     fr_opt_row_t, fr_needed)
    extras["car_opt"]= _ensure_section_capacity(ws, car_opt_row_t, fr_row, car_opt_needed)
    extras["car"]    = _ensure_section_capacity(ws, car_row,    car_opt_row_t, car_needed)
    extras["cab_opt"]= _ensure_section_capacity(ws, cab_opt_row_t, car_row, cab_opt_needed)
    extras["cab"]    = _ensure_section_capacity(ws, cab_row,    cab_opt_row_t, cab_needed)

    # ---------- re-locate sections (after insertion) ----------
    cab_row = _find_title_row(ws, ["cabine"], exclude_keywords=["options"])
    cab_opt_row_t = _find_title_row(ws, ["cabine", "options"])
    car_row = _find_title_row(ws, ["carrosserie"], exclude_keywords=["options"])
    car_opt_row_t = _find_title_row(ws, ["carrosserie", "options"])
    fr_row = _find_title_row(ws, ["groupe", "frigorifique"], exclude_keywords=["options"])
    fr_opt_row_t = _find_title_row(ws, ["groupe", "frigorifique", "options"])
    hy_row = _find_title_row(ws, ["hayon"], exclude_keywords=["options"])
    hy_opt_row_t = _find_title_row(ws, ["hayon", "options"])
    pub_row = _find_title_row(ws, ["publicite"])

    cab_rows = _region_rows(ws, cab_row, cab_opt_row_t)
    cab_opt_rows = _region_rows(ws, cab_opt_row_t, car_row)

    car_rows = _region_rows(ws, car_row, car_opt_row_t)
    car_opt_rows = _region_rows(ws, car_opt_row_t, fr_row)

    fr_rows = _region_rows(ws, fr_row, fr_opt_row_t)
    fr_opt_rows = _region_rows(ws, fr_opt_row_t, hy_row)

    hy_rows = _region_rows(ws, hy_row, hy_opt_row_t)
    hy_opt_rows = _region_rows(ws, hy_opt_row_t, pub_row)

    # ---------- write headers (safe for merged cells) ----------
    header_map = {
        "code_pays": "C5",
        "Marque": "C6",
        "Modele": "C7",
        "Code_PF": "C8",
        "Standard_PF": "C9",
        "catalogue_1\n PF": "C10",
        "catalogue_2\nST": "C11",
        "catalogue_3\nZR": "C12",
        "catalogue_3\n LIBRE": "C12",
    }
    for k, addr in header_map.items():
        if k in veh.index and pd.notna(veh.get(k)):
            safe_set(ws, addr, veh.get(k))

    # ---------- fill sections (blocks fixed to avoid overwrite) ----------
    colB = column_index_from_string("B")
    colF = column_index_from_string("F")
    colH = column_index_from_string("H")

    endE = column_index_from_string("E")
    endG = column_index_from_string("G")
    endL = column_index_from_string("L")

    # CAB/MOT/CH details
    fill_region(ws, cab_rows, cab_vals, [colB], end_cols_override={colB: endE})
    fill_region(ws, cab_rows, mot_vals, [colF], end_cols_override={colF: endG})
    fill_region(ws, cab_rows, ch_vals,  [colH], end_cols_override={colH: endL})

    # CAB/MOT/CH options
    fill_region(ws, cab_opt_rows, cab_opt_vals, [colB], end_cols_override={colB: endE})
    fill_region(ws, cab_opt_rows, mot_opt_vals, [colF], end_cols_override={colF: endG})
    fill_region(ws, cab_opt_rows, ch_opt_vals,  [colH], end_cols_override={colH: endL})

    # CAISSE full width
    fill_region(ws, car_rows, caisse_vals, [colB], end_cols_override={colB: endL})
    fill_region(ws, car_opt_rows, caisse_opt_vals, [colB], end_cols_override={colB: endL})

    # FRIGO 2 colonnes : B:E puis F:L
    fill_region(ws, fr_rows, gf_vals, [colB, colF], end_cols_override={colB: endE, colF: endL})
    fill_region(ws, fr_opt_rows, gf_opt_vals, [colB], end_cols_override={colB: endL})

    # HAYON 2 colonnes : B:E puis F:L
    fill_region(ws, hy_rows, hay_vals, [colB, colF], end_cols_override={colB: endE, colF: endL})
    fill_region(ws, hy_opt_rows, hay_opt_vals, [colB], end_cols_override={colB: endL})

    # ---------- dimensions (safe merged cells) ----------
    safe_set(ws, "I5",  veh.get("W int\n utile \nsur plinthe"))
    safe_set(ws, "I6",  veh.get("L int \nutile \nsur plinthe"))
    safe_set(ws, "I7",  veh.get("H int"))
    safe_set(ws, "I8",  veh.get("H"))

    safe_set(ws, "K4",  veh.get("L"))
    safe_set(ws, "K5",  veh.get("Z"))
    safe_set(ws, "K6",  veh.get("Hc"))
    safe_set(ws, "K7",  veh.get("F"))
    safe_set(ws, "K8",  veh.get("X"))

    safe_set(ws, "I10", veh.get("PTAC"))
    safe_set(ws, "I11", veh.get("CU"))
    safe_set(ws, "I12", veh.get("Volume"))
    safe_set(ws, "I13", veh.get("palettes 800 x 1200 mm"))

    # ---------- images (AFTER insertion so anchors are correct) ----------
    img_veh_path = resolve_image_path(veh.get("Image Vehicule"), "Image Vehicule")
    img_client_path = resolve_image_path(veh.get("Image Client"), "Image Client")
    img_carbu_path = resolve_image_path(veh.get("Image Carburant"), "Image Carburant")

    logo_pf_path = os.path.join(IMG_ROOT, "logo_pf.png")
    if os.path.exists(logo_pf_path):
        xl_logo = XLImage(logo_pf_path)
        xl_logo.anchor = "B2"
        ws.add_image(xl_logo)

    if img_veh_path and isinstance(img_veh_path, str) and os.path.exists(img_veh_path):
        xl_img_veh = XLImage(img_veh_path)
        xl_img_veh.anchor = "B15"
        ws.add_image(xl_img_veh)

    if img_client_path and isinstance(img_client_path, str) and os.path.exists(img_client_path):
        xl_img_client = XLImage(img_client_path)
        xl_img_client.anchor = "H2"
        ws.add_image(xl_img_client)

    if img_carbu_path and isinstance(img_carbu_path, str) and os.path.exists(img_carbu_path):
        xl_img_carbu = XLImage(img_carbu_path)
        xl_img_carbu.anchor = "H15"
        ws.add_image(xl_img_carbu)

    if debug:
        # Optionnel : trace
        try:
            safe_set(ws, "A1", f"DEBUG inserted rows: {extras}")
        except Exception:
            pass

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, extras

# ----------------- APP -----------------
vehicules, cabines, moteurs, chassis, caisses, frigo, hayons = load_data()

st.sidebar.header("Filtres véhicule")
df_filtre = vehicules.copy()

df_filtre, code_pays = filtre_select(df_filtre, "code_pays", "Code pays")
df_filtre, marque = filtre_select(df_filtre, "Marque", "Marque")
df_filtre, modele = filtre_select(df_filtre, "Modele", "Modèle")
df_filtre, code_pf = filtre_select(df_filtre, "Code_PF", "Code PF")
df_filtre, std_pf = filtre_select(df_filtre, "Standard_PF", "Standard PF")

df_filtre, cab_prod_choice = filtre_select(df_filtre, "C_Cabine", "Cabine - code produit")
df_filtre, cab_opt_choice  = filtre_select_options(df_filtre, "C_Cabine-OPTIONS", "Cabine - code options")

df_filtre, mot_prod_choice = filtre_select(df_filtre, "M_moteur", "Moteur - code produit")
df_filtre, mot_opt_choice  = filtre_select_options(df_filtre, "M_moteur-OPTIONS", "Moteur - code options")

df_filtre, ch_prod_choice  = filtre_select(df_filtre, "C_Chassis", "Châssis - code produit")
df_filtre, ch_opt_choice   = filtre_select_options(df_filtre, "C_Chassis-OPTIONS", "Châssis - code options")

df_filtre, caisse_prod_choice = filtre_select(df_filtre, "C_Caisse", "Caisse - code produit")
df_filtre, caisse_opt_choice  = filtre_select_options(df_filtre, "C_Caisse-OPTIONS", "Caisse - code options")

df_filtre, gf_prod_choice  = filtre_select(df_filtre, "C_Groupe frigo", "Groupe frigo - code produit")
df_filtre, gf_opt_choice   = filtre_select_options(df_filtre, "C_Groupe frigo-OPTIONS", "Groupe frigo - code options")

df_filtre, hay_prod_choice = filtre_select(df_filtre, "C_Hayon elevateur", "Hayon - code produit")
df_filtre, hay_opt_choice  = filtre_select_options(df_filtre, "C_Hayon elevateur-OPTIONS", "Hayon - code options")

st.subheader("Résultats du filtrage")
st.write(f"{len(df_filtre)} combinaison(s) véhicule trouvée(s).")

if df_filtre.empty:
    st.warning("Aucun véhicule ne correspond aux filtres sélectionnés.")
    st.stop()

indices = list(df_filtre.index)
choix_idx = st.selectbox(
    "Sélectionne le véhicule pour générer la FT :",
    indices,
    format_func=lambda i: format_vehicule(df_filtre.loc[i]),
)
veh = df_filtre.loc[choix_idx]

st.markdown("---")
st.subheader("Synthèse véhicule")

cols_synthese = [
    "code_pays", "Marque", "Modele", "Code_PF", "Standard_PF",
    "catalogue_1\n PF", "catalogue_2\nST", "catalogue_3\nZR"
]
cols_existantes = [c for c in cols_synthese if c in veh.index]
st.table(veh[cols_existantes].to_frame(name="Valeur"))

img_veh_path = resolve_image_path(veh.get("Image Vehicule"), "Image Vehicule")
img_client_path = resolve_image_path(veh.get("Image Client"), "Image Client")
img_carbu_path = resolve_image_path(veh.get("Image Carburant"), "Image Carburant")

st.subheader("Images associées")
col1, col2, col3 = st.columns(3)
with col1:
    show_image(img_veh_path, "Image véhicule")
with col2:
    show_image(img_client_path, "Image client")
with col3:
    show_image(img_carbu_path, "Picto carburant")

st.markdown("---")
st.subheader("Génération de la fiche technique (dynamic)")

debug_mode = st.checkbox("Afficher debug (lignes insérées)", value=False)

if st.button("⚙️ Générer la FT (Excel)"):
    template_bytes = uploaded_template.getvalue() if uploaded_template else None

    ft_file, extras = genere_ft_excel_dynamic(
        veh,
        cab_prod_choice, cab_opt_choice,
        mot_prod_choice, mot_opt_choice,
        ch_prod_choice, ch_opt_choice,
        caisse_prod_choice, caisse_opt_choice,
        gf_prod_choice, gf_opt_choice,
        hay_prod_choice, hay_opt_choice,
        cabines, moteurs, chassis, caisses, frigo, hayons,
        template_bytes=template_bytes,
        template_path=TEMPLATE_FALLBACK,
        debug=debug_mode,
    )

    if ft_file is not None:
        import time
        codepf = str(veh.get("Code_PF", "")).strip() or "vehicule"
        filename = f"FT_{codepf}_{APP_VERSION}_{int(time.time())}.xlsx"

        st.success("✅ Fiche générée !")

        if debug_mode:
            st.info(f"Lignes insérées par section : {extras}")

        st.download_button(
            label="⬇️ Télécharger la fiche Excel",
            data=ft_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
