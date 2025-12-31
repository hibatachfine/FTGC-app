import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment

APP_VERSION = "2025-12-31_no_extra_rows_v1"  # <- tu dois voir ça dans la sidebar


# ----------------- CONFIG APP -----------------

st.set_page_config(
    page_title="FT Grands Comptes",
    page_icon="🚚",
    layout="wide"
)

st.title("Generateur de Fiches Techniques Grands Comptes")
st.caption("Version de test basée sur bdd_CG.xlsx")
st.sidebar.info(f"✅ App version: {APP_VERSION}")

IMG_ROOT = "images"  # dossier racine des images dans le repo


# ----------------- HELPERS ROBUSTES -----------------

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


# ----------------- CHARGEMENT DATA -----------------

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


# ----------------- FILTRES -----------------

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


# ----------------- SYNTHÈSE COMPOSANTS (APP) -----------------

def affiche_composant(titre, code, df_ref, col_code_ref, code_pf_for_fallback=None, prefer_po=None):
    st.markdown("---")
    st.subheader(titre)

    if pd.isna(code) or str(code).strip() == "":
        st.info("Aucun code renseigné pour ce composant.")
        return

    st.write(f"Code composant : **{code}**")

    code_col = get_col(df_ref, col_code_ref) or df_ref.columns[0]
    code_str = str(code).strip()

    po_col = None
    for c in df_ref.columns:
        if "produit" in _norm(c) and "option" in _norm(c):
            po_col = c
            break

    comp = df_ref[df_ref[code_col].astype(str).str.strip() == code_str]

    if comp.empty and isinstance(code_pf_for_fallback, str) and code_pf_for_fallback.strip():
        key = extract_pf_key(code_pf_for_fallback)
        if key:
            cand = df_ref[df_ref[code_col].astype(str).str.contains(re.escape(key), na=False)]
            if prefer_po and po_col and not cand.empty:
                cand_po = cand[cand[po_col].astype(str).str.strip().str.upper() == prefer_po.upper()]
                comp = cand_po if not cand_po.empty else cand
            else:
                comp = cand

    if comp.empty:
        st.warning("Code non trouvé dans la base de référence.")
        return

    comp_row = comp.iloc[0].dropna()
    st.table(comp_row.to_frame(name="Valeur"))


# ----------------- GENERATION FT -----------------

def genere_ft_excel(
    veh,
    cab_prod_choice, cab_opt_choice,
    mot_prod_choice, mot_opt_choice,
    ch_prod_choice, ch_opt_choice,
    caisse_prod_choice, caisse_opt_choice,
    gf_prod_choice, gf_opt_choice,
    hay_prod_choice, hay_opt_choice,
):
    template_path = "FT_Grand_Compte.xlsx"
    if not os.path.exists(template_path):
        st.info("Le fichier modèle 'FT_Grand_Compte.xlsx' n'est pas présent dans le repo.")
        return None

    wb = load_workbook(template_path, read_only=False, data_only=False)
    ws = wb["date"] if "date" in wb.sheetnames else wb[wb.sheetnames[0]]

    # ---- helpers excel ----
    def cell_to_rc(cell_addr: str):
        col_letters = "".join(ch for ch in cell_addr if ch.isalpha())
        row_digits = "".join(ch for ch in cell_addr if ch.isdigit())
        return column_index_from_string(col_letters), int(row_digits)

    def merged_top_left(row, col):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return rng.min_row, rng.min_col
        return row, col

    def set_cell_value_merged_safe(row, col, value):
        r0, c0 = merged_top_left(row, col)
        cell = ws.cell(row=r0, column=c0)
        if isinstance(cell, MergedCell):
            return
        cell.value = value
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    def write_block_merged_safe(start_rc, values, n_rows):
        start_col, start_row = start_rc
        for i in range(n_rows):
            v = values[i] if i < len(values) else None
            set_cell_value_merged_safe(start_row + i, start_col, v)

    def insert_rows_and_shift(anchors_dict: dict, insert_at_row: int, n: int):
        if n <= 0:
            return anchors_dict
        ws.insert_rows(insert_at_row, n)
        new_anchors = {}
        for k, (c, r) in anchors_dict.items():
            new_anchors[k] = (c, r + n) if r >= insert_at_row else (c, r)
        return new_anchors

    # ---- data helpers ----
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

    def choose_codes(prod_choice, opt_choice, veh_prod, veh_opt):
        prod_code = veh_prod
        opt_code = veh_opt
        if isinstance(prod_choice, str) and prod_choice not in (None, "", "Tous"):
            prod_code = prod_choice
        if isinstance(opt_choice, str) and opt_choice not in (None, "", "Tous"):
            opt_code = opt_choice
        return prod_code, opt_code

    # ---- header ----
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

    # ---- images ----
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

    # ---- composants ----
    global cabines, moteurs, chassis, caisses, frigo, hayons
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

    anchors = {
        "CAB_START": cell_to_rc("B18"),
        "MOT_START": cell_to_rc("F18"),
        "CH_START":  cell_to_rc("H18"),
        "CAB_OPT": cell_to_rc("B38"),
        "MOT_OPT": cell_to_rc("F38"),
        "CH_OPT":  cell_to_rc("H38"),
        "CAISSE_START": cell_to_rc("B40"),
        "CAISSE_OPT":   cell_to_rc("B47"),
        "GF_START": cell_to_rc("B50"),
        "GF_OPT":   cell_to_rc("B58"),
        "HAY_START": cell_to_rc("B61"),
        "HAY_OPT":   cell_to_rc("B68"),
    }

    BASE = {
        "TOP_MAIN": 17,
        "TOP_OPT":  3,
        "CAISSE_MAIN": 5,
        "CAISSE_OPT":  2,
        "GF_MAIN": 6,
        "GF_OPT":  2,
        "HAY_MAIN": 5,
        "HAY_OPT":  3,
    }

    def ensure_space(start_anchor_key: str, base_rows: int, needed_rows: int):
        extra_rows = max(0, int(needed_rows) - int(base_rows))
    if extra_rows <= 0:
        return

    start_col, start_row = anchors[start_anchor_key]
    insert_at = start_row + int(base_rows)

    new_anchors = insert_rows_and_shift(anchors, insert_at, extra_rows)
    anchors.clear()
    anchors.update(new_anchors)


    top_needed = max(len(cab_vals), len(mot_vals), len(ch_vals), 1)
    ensure_space("CAB_START", BASE["TOP_MAIN"], top_needed)
    write_block_merged_safe(anchors["CAB_START"], cab_vals, top_needed)
    write_block_merged_safe(anchors["MOT_START"], mot_vals, top_needed)
    write_block_merged_safe(anchors["CH_START"],  ch_vals,  top_needed)

    top_opt_needed = max(len(cab_opt_vals), len(mot_opt_vals), len(ch_opt_vals), 1)
    ensure_space("CAB_OPT", BASE["TOP_OPT"], top_opt_needed)
    write_block_merged_safe(anchors["CAB_OPT"], cab_opt_vals, top_opt_needed)
    write_block_merged_safe(anchors["MOT_OPT"], mot_opt_vals, top_opt_needed)
    write_block_merged_safe(anchors["CH_OPT"],  ch_opt_vals,  top_opt_needed)

    caisse_needed = max(len(caisse_vals), 1)
    ensure_space("CAISSE_START", BASE["CAISSE_MAIN"], caisse_needed)
    write_block_merged_safe(anchors["CAISSE_START"], caisse_vals, caisse_needed)

    caisse_opt_needed = max(len(caisse_opt_vals), 1)
    ensure_space("CAISSE_OPT", BASE["CAISSE_OPT"], caisse_opt_needed)
    write_block_merged_safe(anchors["CAISSE_OPT"], caisse_opt_vals, caisse_opt_needed)

    gf_needed = max(len(gf_vals), 1)
    ensure_space("GF_START", BASE["GF_MAIN"], gf_needed)
    write_block_merged_safe(anchors["GF_START"], gf_vals, gf_needed)

    gf_opt_needed = max(len(gf_opt_vals), 1)
    ensure_space("GF_OPT", BASE["GF_OPT"], gf_opt_needed)
    write_block_merged_safe(anchors["GF_OPT"], gf_opt_vals, gf_opt_needed)

    hay_needed = max(len(hay_vals), 1)
    ensure_space("HAY_START", BASE["HAY_MAIN"], hay_needed)
    write_block_merged_safe(anchors["HAY_START"], hay_vals, hay_needed)

    hay_opt_needed = max(len(hay_opt_vals), 1)
    ensure_space("HAY_OPT", BASE["HAY_OPT"], hay_opt_needed)
    write_block_merged_safe(anchors["HAY_OPT"], hay_opt_vals, hay_opt_needed)

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

code_pf_ref = veh.get("Code_PF", "")

affiche_composant("Cabine", veh.get("C_Cabine"), cabines, "C_Cabine", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Châssis", veh.get("C_Chassis"), chassis, "CH_chassis", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Caisse", veh.get("C_Caisse"), caisses, "CF_caisse", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Moteur", veh.get("M_moteur"), moteurs, "M_moteur", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Groupe frigorifique", veh.get("C_Groupe frigo"), frigo, "GF_groupe frigo", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Hayon élévateur", veh.get("C_Hayon elevateur"), hayons, "HL_hayon elevateur", code_pf_for_fallback=code_pf_ref, prefer_po="P")

st.markdown("---")
st.subheader("Génération de la fiche technique")

ft_file = genere_ft_excel(
    veh,
    cab_prod_choice, cab_opt_choice,
    mot_prod_choice, mot_opt_choice,
    ch_prod_choice, ch_opt_choice,
    caisse_prod_choice, caisse_opt_choice,
    gf_prod_choice, gf_opt_choice,
    hay_prod_choice, hay_opt_choice,
)

if ft_file is not None:
    nom_fichier = f"FT_{veh.get('Code_PF', 'PF')}_{veh.get('Modele', 'MODELE')}.xlsx"
    st.download_button(
        label=" Télécharger la fiche technique remplie",
        data=ft_file,
        file_name=nom_fichier,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Ajoute le modèle 'FT_Grand_Compte.xlsx' dans le repo pour activer le téléchargement.")
