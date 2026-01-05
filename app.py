import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment
from openpyxl.cell.cell import MergedCell

APP_VERSION = "2026-01-05_simple_fill_FT_Grands_Comptes_only"

# ----------------- CONFIG APP -----------------
st.set_page_config(page_title="FT Grands Comptes", page_icon="🚚", layout="wide")
st.title("Generateur de Fiches Techniques Grands Comptes")
st.caption("Version de test basée sur bdd_CG.xlsx")
st.sidebar.info(f"✅ Version: {APP_VERSION}")

# ✅ Template upload (bypass repo caching / deployment issues)
st.sidebar.markdown("### Template Excel")
uploaded_template = st.sidebar.file_uploader(
    "Uploader le template (xlsx) à utiliser",
    type=["xlsx"]
)

TEMPLATE_FALLBACK = "FT_Grands_Comptes.xlsx"


IMG_ROOT = "images"


# ----------------- HELPERS -----------------
def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).lower()
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


def _infer_block_endcol(ws, rows, start_col, default_endcol):
    for r in rows:
        rng = _merged_range_including(ws, r, start_col)
        if rng and rng.min_row == rng.max_row == r and rng.min_col == start_col:
            return rng.max_col
    return default_endcol


def fill_region(ws, rows, values, start_cols, mode="auto"):
    if not rows:
        return 0, 0

    # defaults (ton template)
    default_full_end = column_index_from_string("L")
    default_left_end = column_index_from_string("E")
    default_right_end = column_index_from_string("L")

    ends = {}
    if mode == "full" or (mode == "auto" and len(start_cols) == 1):
        sc = start_cols[0]
        ends[sc] = _infer_block_endcol(ws, rows, sc, default_full_end)
    else:
        for sc in start_cols:
            if column_index_from_string("B") == sc:
                ends[sc] = _infer_block_endcol(ws, rows, sc, default_left_end)
            elif column_index_from_string("F") == sc:
                ends[sc] = _infer_block_endcol(ws, rows, sc, default_right_end)
            else:
                ends[sc] = _infer_block_endcol(ws, rows, sc, default_full_end)

    # force merges
    for r in rows:
        for sc in start_cols:
            _ensure_merge_row(ws, r, sc, ends.get(sc, sc))

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


# ----------------- EXCEL GENERATION (SIMPLE) -----------------
def genere_ft_excel_simple(
    veh,
    cab_prod_choice, cab_opt_choice,
    mot_prod_choice, mot_opt_choice,
    ch_prod_choice, ch_opt_choice,
    caisse_prod_choice, caisse_opt_choice,
    gf_prod_choice, gf_opt_choice,
    hay_prod_choice, hay_opt_choice,
    cabines, moteurs, chassis, caisses, frigo, hayons,
):
    TEMPLATE = "FT_Grands_Comptes.xlsx"
    if not os.path.exists(TEMPLATE):
        st.error(f"Template introuvable : {TEMPLATE}")
        return None

    wb = load_workbook(TEMPLATE, read_only=False, data_only=False)
    ws = wb["date"] if "date" in wb.sheetnames else wb[wb.sheetnames[0]]

    # Pas de répétition d'entête imprimée
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

    # ---------- header mapping ----------
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
    for k, cell in header_map.items():
        if k in veh.index and pd.notna(veh.get(k)):
            ws[cell] = veh.get(k)

    # ---------- images ----------
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

    # ---------- locate sections ----------
    cab_row = _find_title_row(ws, ["cabine"], exclude_keywords=["options"])
    cab_opt_row = _find_title_row(ws, ["cabine", "options"])
    car_row = _find_title_row(ws, ["carrosserie"], exclude_keywords=["options"])
    car_opt_row = _find_title_row(ws, ["carrosserie", "options"])
    fr_row = _find_title_row(ws, ["groupe", "frigorifique"], exclude_keywords=["options"])
    fr_opt_row = _find_title_row(ws, ["groupe", "frigorifique", "options"])
    hy_row = _find_title_row(ws, ["hayon"], exclude_keywords=["options"])
    hy_opt_row = _find_title_row(ws, ["hayon", "options"])
    pub_row = _find_title_row(ws, ["publicite"])

    cab_rows = _region_rows(ws, cab_row, cab_opt_row)
    cab_opt_rows = _region_rows(ws, cab_opt_row, car_row)

    car_rows = _region_rows(ws, car_row, car_opt_row)
    car_opt_rows = _region_rows(ws, car_opt_row, fr_row)

    fr_rows = _region_rows(ws, fr_row, fr_opt_row)
    fr_opt_rows = _region_rows(ws, fr_opt_row, hy_row)

    hy_rows = _region_rows(ws, hy_row, hy_opt_row)
    hy_opt_rows = _region_rows(ws, hy_opt_row, pub_row)

    colB = column_index_from_string("B")
    colF = column_index_from_string("F")
    colH = column_index_from_string("H")

    # CAB/MOT/CH details (same rows, different cols)
    fill_region(ws, cab_rows, cab_vals, [colB], mode="auto")
    fill_region(ws, cab_rows, mot_vals, [colF], mode="auto")
    fill_region(ws, cab_rows, ch_vals,  [colH], mode="auto")

    # CAB/MOT/CH options
    fill_region(ws, cab_opt_rows, cab_opt_vals, [colB], mode="auto")
    fill_region(ws, cab_opt_rows, mot_opt_vals, [colF], mode="auto")
    fill_region(ws, cab_opt_rows, ch_opt_vals,  [colH], mode="auto")

    # CAISSE full width B->L
    fill_region(ws, car_rows, caisse_vals, [colB], mode="full")
    fill_region(ws, car_opt_rows, caisse_opt_vals, [colB], mode="full")

    # FRIGO B then F
    fill_region(ws, fr_rows, gf_vals, [colB, colF], mode="two_col")
    fill_region(ws, fr_opt_rows, gf_opt_vals, [colB], mode="full")

    # HAYON B then F
    fill_region(ws, hy_rows, hay_vals, [colB, colF], mode="two_col")
    fill_region(ws, hy_opt_rows, hay_opt_vals, [colB], mode="full")

    # DIMENSIONS (si tes cellules sont différentes, adapte)
    ws["I5"]  = veh.get("W int\n utile \nsur plinthe")
    ws["I6"]  = veh.get("L int \nutile \nsur plinthe")
    ws["I7"]  = veh.get("H int")
    ws["I8"]  = veh.get("H")

    ws["K4"]  = veh.get("L")
    ws["K5"]  = veh.get("Z")
    ws["K6"]  = veh.get("Hc")
    ws["K7"]  = veh.get("F")
    ws["K8"]  = veh.get("X")

    ws["I10"] = veh.get("PTAC")
    ws["I11"] = veh.get("CU")
    ws["I12"] = veh.get("Volume")
    ws["I13"] = veh.get("palettes 800 x 1200 mm")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


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
st.subheader("Génération de la fiche technique (simple)")

if st.button("⚙️ Générer la FT (Excel)"):
    ft_file = genere_ft_excel_simple(
        veh,
        cab_prod_choice, cab_opt_choice,
        mot_prod_choice, mot_opt_choice,
        ch_prod_choice, ch_opt_choice,
        caisse_prod_choice, caisse_opt_choice,
        gf_prod_choice, gf_opt_choice,
        hay_prod_choice, hay_opt_choice,
        cabines, moteurs, chassis, caisses, frigo, hayons
    )

    if ft_file is not None:
        # ⚠️ nom unique pour être sûr d'ouvrir le bon fichier
        import time
        codepf = str(veh.get("Code_PF", "")).strip() or "vehicule"
        filename = f"FT_{codepf}_{APP_VERSION}_{int(time.time())}.xlsx"
        st.success("✅ Fiche générée !")
        st.download_button(
            label="⬇️ Télécharger la fiche Excel",
            data=ft_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
