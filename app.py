import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell

# ----------------- CONFIG APP -----------------

st.set_page_config(
    page_title="FT Grands Comptes",
    page_icon="🚚",
    layout="wide"
)

st.title("Generateur de Fiches Techniques Grands Comptes")
st.caption("Version de test basée sur bdd_CG.xlsx")

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


# ----------------- FONCTIONS UTILES -----------------

def resolve_image_path(cell_value, subdir):
    if not isinstance(cell_value, str) or not cell_value.strip():
        return None

    val = cell_value.strip()

    if val.lower().startswith(("http://", "https://")):
        return val

    # Linux basename ne coupe pas sur "\" -> on convertit en "/"
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


def filtre_select_options(df, col_wanted, label):
    col = get_col(df, col_wanted)
    if col is None:
        st.sidebar.write(f"(colonne '{col_wanted}' absente)")
        return df, None

    options = clean_unique_list(df[col])
    choix = st.sidebar.selectbox(label, ["Tous"] + options)

    if choix != "Tous":
        df = df[df[col].astype(str).str.strip() == str(choix).strip()]

    return df, choix


def format_vehicule(row):
    champs = []
    for c in ["code_pays", "Marque", "Modele", "Code_PF", "Standard_PF"]:
        if c in row and pd.notna(row[c]):
            champs.append(str(row[c]))
    return " – ".join(champs)


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


# ----------------- GENERATION FT (avec décalage auto) -----------------

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
    ws = wb["date"]

    # ---- Réglages max lignes ----
    MAX = {
        "CAB": 20,
        "MOT": 8,
        "CH": 16,
        "CAISSE": 33,
        "GF": 10,
        "HAY": 11,
        "OPT": 8,  # options par bloc (tu peux monter à 12 si besoin)
    }

    def cell_to_rc(cell_addr: str):
        col_letters = "".join(ch for ch in cell_addr if ch.isalpha())
        row_digits = "".join(ch for ch in cell_addr if ch.isdigit())
        return column_index_from_string(col_letters), int(row_digits)

    def rc_to_cell(col_idx: int, row_idx: int):
        # convert col index -> letters
        letters = ""
        n = col_idx
        while n > 0:
            n, r = divmod(n - 1, 26)
            letters = chr(65 + r) + letters
        return f"{letters}{row_idx}"

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

    def write_block(start_cell, values, n_rows):
        """Ecrit values verticalement sur n_rows (n_rows >= len(values) en général)."""
        start_col, start_row = cell_to_rc(start_cell)
        for i in range(n_rows):
            cell = ws.cell(row=start_row + i, column=start_col)
            if isinstance(cell, MergedCell):
                continue
            cell.value = values[i] if i < len(values) else None

    def insert_rows_and_shift(anchors: dict, insert_at_row: int, n: int):
        """Insert n rows at insert_at_row and shift all anchors having row >= insert_at_row."""
        if n <= 0:
            return anchors
        ws.insert_rows(insert_at_row, n)
        new_anchors = {}
        for k, (c, r) in anchors.items():
            if r >= insert_at_row:
                new_anchors[k] = (c, r + n)
            else:
                new_anchors[k] = (c, r)
        return new_anchors

    def find_row(df, code, code_col_wanted, code_pf_fallback=None, prefer_po=None):
        if not isinstance(code, str) or code.strip() == "" or code == "Tous":
            code = ""

        code_col = get_col(df, code_col_wanted) or df.columns[0]

        po_col = None
        for c in df.columns:
            if "produit" in _norm(c) and "option" in _norm(c):
                po_col = c
                break

        # exact
        if code:
            cand = df[df[code_col].astype(str).str.strip() == code.strip()]
            if not cand.empty:
                if prefer_po and po_col:
                    cand_po = cand[cand[po_col].astype(str).str.strip().str.upper() == prefer_po.upper()]
                    return cand_po.iloc[0] if not cand_po.empty else cand.iloc[0]
                return cand.iloc[0]

        # fallback Code_PF
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

    # ---- HEADER ----
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

    # ---- IMAGES ----
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

    # ---- COMPOSANTS ----
    global cabines, moteurs, chassis, caisses, frigo, hayons
    code_pf_ref = veh.get("Code_PF", "")

    cab_prod_code, cab_opt_code = choose_codes(cab_prod_choice, cab_opt_choice, veh.get("C_Cabine"), veh.get("C_Cabine-OPTIONS"))
    mot_prod_code, mot_opt_code = choose_codes(mot_prod_choice, mot_opt_choice, veh.get("M_moteur"), veh.get("M_moteur-OPTIONS"))
    ch_prod_code, ch_opt_code = choose_codes(ch_prod_choice, ch_opt_choice, veh.get("C_Chassis"), veh.get("C_Chassis-OPTIONS"))
    caisse_prod_code, caisse_opt_code = choose_codes(caisse_prod_choice, caisse_opt_choice, veh.get("C_Caisse"), veh.get("C_Caisse-OPTIONS"))
    gf_prod_code, gf_opt_code = choose_codes(gf_prod_choice, gf_opt_choice, veh.get("C_Groupe frigo"), veh.get("C_Groupe frigo-OPTIONS"))
    hay_prod_code, hay_opt_code = choose_codes(hay_prod_choice, hay_opt_choice, veh.get("C_Hayon elevateur"), veh.get("C_Hayon elevateur-OPTIONS"))

    # Récup rows
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

    # Valeurs
    cab_codecol = get_col(cabines, "C_Cabine") or cabines.columns[0]
    mot_codecol = get_col(moteurs, "M_moteur") or moteurs.columns[0]
    ch_codecol  = get_col(chassis, "CH_chassis") or chassis.columns[0]
    caisse_codecol = get_col(caisses, "CF_caisse") or caisses.columns[0]
    gf_codecol = get_col(frigo, "GF_groupe frigo") or frigo.columns[0]
    hay_codecol = get_col(hayons, "HL_hayon elevateur") or hayons.columns[0]

    cab_vals = build_values(cab_prod_row, cab_codecol)
    mot_vals = build_values(mot_prod_row, mot_codecol)
    ch_vals  = build_values(ch_prod_row,  ch_codecol)

    cab_opt_vals = build_values(cab_opt_row, cab_codecol)
    mot_opt_vals = build_values(mot_opt_row, mot_codecol)
    ch_opt_vals  = build_values(ch_opt_row,  ch_codecol)

    caisse_vals = build_values(caisse_prod_row, caisse_codecol)
    caisse_opt_vals = build_values(caisse_opt_row, caisse_codecol)

    gf_vals = build_values(gf_prod_row, gf_codecol)
    gf_opt_vals = build_values(gf_opt_row, gf_codecol)

    hay_vals = build_values(hay_prod_row, hay_codecol)
    hay_opt_vals = build_values(hay_opt_row, hay_codecol)

    # ---- ANCHORS (positions initiales template) ----
    # IMPORTANT : on ne change pas tes cellules, on les décale si on insert des lignes
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

        # Dimensions
        "DIM_I5": cell_to_rc("I5"),
        "DIM_K4": cell_to_rc("K4"),
        "DIM_I10": cell_to_rc("I10"),
        "DIM_I13": cell_to_rc("I13"),
    }

    # ---- 1) Bande haute CAB/MOT/CH (détails) : on agrandit selon le max des 3 ----
    top_needed = max(len(cab_vals), len(mot_vals), len(ch_vals), MAX["CAB"], MAX["MOT"], MAX["CH"])
    # On garde les seuils composant, mais la "bande" doit couvrir le plus grand des 3
    # => on calcule une hauteur cible raisonnable :
    top_target = max(MAX["CAB"], MAX["MOT"], MAX["CH"], max(len(cab_vals), len(mot_vals), len(ch_vals)))
    # longueur actuelle prévue par template : 17 (ton ancienne valeur)
    TOP_DEFAULT = 17
    extra_top = max(0, top_target - TOP_DEFAULT)

    if extra_top > 0:
        # insert après la zone détail (row 18 + 17 = 35)
        insert_at = anchors["CAB_START"][1] + TOP_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_top)

    # Ecriture détails (on écrit top_target lignes pour aligner)
    cab_start = rc_to_cell(*anchors["CAB_START"])
    mot_start = rc_to_cell(*anchors["MOT_START"])
    ch_start  = rc_to_cell(*anchors["CH_START"])

    write_block(cab_start, cab_vals, top_target)
    write_block(mot_start, mot_vals, top_target)
    write_block(ch_start,  ch_vals,  top_target)

    # ---- 2) Bande options CAB/MOT/CH : on agrandit selon max options ----
    opt_target = max(MAX["OPT"], len(cab_opt_vals), len(mot_opt_vals), len(ch_opt_vals))
    OPT_DEFAULT = 3  # ton ancienne valeur
    extra_opt = max(0, opt_target - OPT_DEFAULT)

    if extra_opt > 0:
        insert_at = anchors["CAB_OPT"][1] + OPT_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_opt)

    cab_opt_cell = rc_to_cell(*anchors["CAB_OPT"])
    mot_opt_cell = rc_to_cell(*anchors["MOT_OPT"])
    ch_opt_cell  = rc_to_cell(*anchors["CH_OPT"])

    write_block(cab_opt_cell, cab_opt_vals, opt_target)
    write_block(mot_opt_cell, mot_opt_vals, opt_target)
    write_block(ch_opt_cell,  ch_opt_vals,  opt_target)

    # ---- 3) CAISSE (détails) ----
    caisse_target = max(MAX["CAISSE"], len(caisse_vals))
    CAISSE_DEFAULT = 5  # ton ancienne valeur
    extra_caisse = max(0, caisse_target - CAISSE_DEFAULT)

    if extra_caisse > 0:
        insert_at = anchors["CAISSE_START"][1] + CAISSE_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_caisse)

    caisse_cell = rc_to_cell(*anchors["CAISSE_START"])
    write_block(caisse_cell, caisse_vals, caisse_target)

    # ---- 4) CAISSE options ----
    caisse_opt_target = max(MAX["OPT"], len(caisse_opt_vals))
    CAISSE_OPT_DEFAULT = 2
    extra_caisse_opt = max(0, caisse_opt_target - CAISSE_OPT_DEFAULT)

    if extra_caisse_opt > 0:
        insert_at = anchors["CAISSE_OPT"][1] + CAISSE_OPT_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_caisse_opt)

    caisse_opt_cell = rc_to_cell(*anchors["CAISSE_OPT"])
    write_block(caisse_opt_cell, caisse_opt_vals, caisse_opt_target)

    # ---- 5) FRIGO ----
    gf_target = max(MAX["GF"], len(gf_vals))
    GF_DEFAULT = 6
    extra_gf = max(0, gf_target - GF_DEFAULT)

    if extra_gf > 0:
        insert_at = anchors["GF_START"][1] + GF_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_gf)

    gf_cell = rc_to_cell(*anchors["GF_START"])
    write_block(gf_cell, gf_vals, gf_target)

    # FRIGO options
    gf_opt_target = max(MAX["OPT"], len(gf_opt_vals))
    GF_OPT_DEFAULT = 2
    extra_gf_opt = max(0, gf_opt_target - GF_OPT_DEFAULT)

    if extra_gf_opt > 0:
        insert_at = anchors["GF_OPT"][1] + GF_OPT_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_gf_opt)

    gf_opt_cell = rc_to_cell(*anchors["GF_OPT"])
    write_block(gf_opt_cell, gf_opt_vals, gf_opt_target)

    # ---- 6) HAYON ----
    hay_target = max(MAX["HAY"], len(hay_vals))
    HAY_DEFAULT = 5
    extra_hay = max(0, hay_target - HAY_DEFAULT)

    if extra_hay > 0:
        insert_at = anchors["HAY_START"][1] + HAY_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_hay)

    hay_cell = rc_to_cell(*anchors["HAY_START"])
    write_block(hay_cell, hay_vals, hay_target)

    # HAYON options
    hay_opt_target = max(MAX["OPT"], len(hay_opt_vals))
    HAY_OPT_DEFAULT = 3
    extra_hay_opt = max(0, hay_opt_target - HAY_OPT_DEFAULT)

    if extra_hay_opt > 0:
        insert_at = anchors["HAY_OPT"][1] + HAY_OPT_DEFAULT
        anchors = insert_rows_and_shift(anchors, insert_at, extra_hay_opt)

    hay_opt_cell = rc_to_cell(*anchors["HAY_OPT"])
    write_block(hay_opt_cell, hay_opt_vals, hay_opt_target)

    # ---- DIMENSIONS (inchangé, mais si tu veux les déplacer aussi : on peut) ----
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


# ----------------- CHARGEMENT DES DONNÉES -----------------

vehicules, cabines, moteurs, chassis, caisses, frigo, hayons = load_data()


# ----------------- SIDEBAR : FILTRES -----------------

st.sidebar.header("Filtres véhicule")

df_filtre = vehicules.copy()

df_filtre, code_pays = filtre_select(df_filtre, "code_pays", "Code pays")
df_filtre, marque = filtre_select(df_filtre, "Marque", "Marque")
df_filtre, modele = filtre_select(df_filtre, "Modele", "Modèle")
df_filtre, code_pf = filtre_select(df_filtre, "Code_PF", "Code PF")
df_filtre, std_pf = filtre_select(df_filtre, "Standard_PF", "Standard PF")

# composants (produit + options)
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

# ----------------- CHOIX D'UNE LIGNE -----------------

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

# ----------------- APERÇU IMAGES -----------------

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

# ----------------- DETAIL COMPOSANTS (tu peux supprimer si tu veux) -----------------

code_pf_ref = veh.get("Code_PF", "")

affiche_composant("Cabine", veh.get("C_Cabine"), cabines, "C_Cabine", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Châssis", veh.get("C_Chassis"), chassis, "CH_chassis", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Caisse", veh.get("C_Caisse"), caisses, "CF_caisse", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Moteur", veh.get("M_moteur"), moteurs, "M_moteur", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Groupe frigorifique", veh.get("C_Groupe frigo"), frigo, "GF_groupe frigo", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Hayon élévateur", veh.get("C_Hayon elevateur"), hayons, "HL_hayon elevateur", code_pf_for_fallback=code_pf_ref, prefer_po="P")

# ----------------- BOUTON FT -----------------

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
