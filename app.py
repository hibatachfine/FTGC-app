import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell

# ----------------- CONFIG APP -----------------

st.set_page_config(
    page_title="FT Grand Compte",
    page_icon="🚚",
    layout="wide"
)

st.title("Generateur de Fiches Techniques Grand Compte")
st.caption("Version de test basée sur bdd_CG.xlsx")

IMG_ROOT = "images"  # dossier racine des images dans le repo


# ----------------- HELPERS ROBUSTES (COLONNES / VALEURS) -----------------

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).lower()
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = s.replace("–", "-").replace("—", "-")
    s = " ".join(s.split())
    return s

def get_col(df: pd.DataFrame, wanted: str):
    """Retrouve une colonne même si variations (casse / espaces / retours ligne)."""
    if df is None or wanted is None:
        return None
    w = _norm(wanted)

    # match exact normalisé
    for c in df.columns:
        if _norm(c) == w:
            return c

    # fallback: contient
    for c in df.columns:
        if w in _norm(c):
            return c

    return None

def clean_unique_list(series: pd.Series):
    """Nettoie les valeurs affichées dans un selectbox."""
    if series is None:
        return []
    s = series.dropna().astype(str).map(lambda x: x.strip())
    s = s[(s != "") & (s.str.lower() != "nan")]
    return sorted(s.unique().tolist())


# ----------------- FONCTIONS UTILES -----------------

def resolve_image_path(cell_value, subdir):
    """
    Transforme la valeur Excel (chemin complet, nom de fichier ou URL)
    en chemin exploitable.
    """
    if not isinstance(cell_value, str) or not cell_value.strip():
        return None

    val = cell_value.strip()

    if val.lower().startswith(("http://", "https://")):
        return val

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
    """Selectbox 'Tous + valeurs uniques' dans la sidebar et renvoie le DF filtré."""
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
    """
    Même logique que filtre_select mais pour les options.
    Construit la liste depuis df (déjà filtré) => options cohérentes avec le véhicule/PF.
    """
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


def affiche_composant(titre, code, df_ref, col_code_ref):
    st.markdown("---")
    st.subheader(titre)

    if pd.isna(code) or str(code).strip() == "":
        st.info("Aucun code renseigné pour ce composant.")
        return

    st.write(f"Code composant : **{code}**")

    # colonne code : on essaye col_code_ref, sinon 1ère colonne (souvent la colonne code)
    col_ref = get_col(df_ref, col_code_ref)
    if col_ref is None:
        col_ref = df_ref.columns[0]

    comp = df_ref[df_ref[col_ref].astype(str).str.strip() == str(code).strip()]
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
        st.info("Le fichier modèle 'FT_Grand_Compte.xlsx' n'est pas présent dans le repo. "
                "Ajoute-le à la racine pour activer la génération automatique.")
        return None

    wb = load_workbook(template_path, read_only=False, data_only=False)
    ws = wb["date"]  # feuille FT

    def build_values(row, code_col):
        """Liste de valeurs sans titres/colonnes techniques."""
        if row is None:
            return []
        vals = []
        for col, val in row.items():
            if str(col).strip() == str(code_col).strip():
                continue
            if pd.isna(val) or str(val).strip() == "":
                continue
            name_lower = str(col).lower()
            if ("produit" in name_lower and "option" in name_lower) or name_lower.startswith("zone libre"):
                continue
            if str(col).strip() == "_":
                continue
            vals.append(str(val).strip())
        return vals

    def write_block(start_cell, values, max_rows):
        """Ecrit values verticalement dans une colonne à partir de start_cell."""
        col_letters = "".join(ch for ch in start_cell if ch.isalpha())
        row_digits = "".join(ch for ch in start_cell if ch.isdigit())
        start_col = column_index_from_string(col_letters)
        start_row = int(row_digits)

        for i in range(max_rows):
            cell = ws.cell(row=start_row + i, column=start_col)
            if isinstance(cell, MergedCell):
                continue
            cell.value = values[i] if i < len(values) else None

    def find_row(df, code, code_col_wanted):
        """
        Cherche le code dans df.
        Si code_col_wanted n'existe pas, on prend df.columns[0] (colonne code du sheet).
        """
        if not isinstance(code, str) or code.strip() == "" or code == "Tous":
            return None

        code_col = get_col(df, code_col_wanted)
        if code_col is None:
            code_col = df.columns[0]

        cand = df[df[code_col].astype(str).str.strip() == code.strip()]
        if cand.empty:
            return None
        return cand.iloc[0]

    def choose_codes(prod_choice, opt_choice, veh_prod, veh_opt):
        """Override par filtres si pas 'Tous'."""
        prod_code = veh_prod
        opt_code = veh_opt

        if isinstance(prod_choice, str) and prod_choice not in (None, "", "Tous"):
            prod_code = prod_choice
        if isinstance(opt_choice, str) and opt_choice not in (None, "", "Tous"):
            opt_code = opt_choice

        return prod_code, opt_code

    # ---- Header (robuste sur veh) ----
    header_map = [
        ("code_pays", "C5"),
        ("Marque", "C6"),
        ("Modele", "C7"),
        ("Code_PF", "C8"),
        ("Standard_PF", "C9"),
        ("catalogue_1\n PF", "C10"),
        ("catalogue_2\nST", "C11"),
        ("catalogue_3\nZR", "C12"),
        ("catalogue_3\n LIBRE", "C12"),  # au cas où
    ]
    for col_name, cell in header_map:
        if col_name in veh.index and pd.notna(veh.get(col_name)):
            ws[cell] = veh.get(col_name)

    # ---- Images ----
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

    # ---- Composants ----
    global cabines, moteurs, chassis, caisses, frigo, hayons

    cab_prod_code, cab_opt_code = choose_codes(
        cab_prod_choice, cab_opt_choice,
        veh.get("C_Cabine"), veh.get("C_Cabine-OPTIONS")
    )
    mot_prod_code, mot_opt_code = choose_codes(
        mot_prod_choice, mot_opt_choice,
        veh.get("M_moteur"), veh.get("M_moteur-OPTIONS")
    )
    ch_prod_code, ch_opt_code = choose_codes(
        ch_prod_choice, ch_opt_choice,
        veh.get("C_Chassis"), veh.get("C_Chassis-OPTIONS")
    )
    caisse_prod_code, caisse_opt_code = choose_codes(
        caisse_prod_choice, caisse_opt_choice,
        veh.get("C_Caisse"), veh.get("C_Caisse-OPTIONS")
    )
    gf_prod_code, gf_opt_code = choose_codes(
        gf_prod_choice, gf_opt_choice,
        veh.get("C_Groupe frigo"), veh.get("C_Groupe frigo-OPTIONS")
    )
    hay_prod_code, hay_opt_code = choose_codes(
        hay_prod_choice, hay_opt_choice,
        veh.get("C_Hayon elevateur"), veh.get("C_Hayon elevateur-OPTIONS")
    )

    # CABINE
    cab_prod_row = find_row(cabines, cab_prod_code, "C_Cabine")
    cab_opt_row = find_row(cabines, cab_opt_code, "C_Cabine")
    write_block("B18", build_values(cab_prod_row, (get_col(cabines, "C_Cabine") or cabines.columns[0])), 17)
    write_block("B38", build_values(cab_opt_row, (get_col(cabines, "C_Cabine") or cabines.columns[0])), 3)

    # MOTEUR
    mot_prod_row = find_row(moteurs, mot_prod_code, "M_moteur")
    mot_opt_row = find_row(moteurs, mot_opt_code, "M_moteur")
    write_block("F18", build_values(mot_prod_row, (get_col(moteurs, "M_moteur") or moteurs.columns[0])), 17)
    write_block("F38", build_values(mot_opt_row, (get_col(moteurs, "M_moteur") or moteurs.columns[0])), 3)

    # CHASSIS
    ch_prod_row = find_row(chassis, ch_prod_code, "c_chassis")
    ch_opt_row = find_row(chassis, ch_opt_code, "c_chassis")
    write_block("H18", build_values(ch_prod_row, chassis.columns[0]), 17)
    write_block("H38", build_values(ch_opt_row, chassis.columns[0]), 3)

    # CAISSE / CARROSSERIE
    caisse_prod_row = find_row(caisses, caisse_prod_code, "c_caisse")
    caisse_opt_row = find_row(caisses, caisse_opt_code, "c_caisse")
    write_block("B40", build_values(caisse_prod_row, caisses.columns[0]), 5)
    write_block("B47", build_values(caisse_opt_row, caisses.columns[0]), 2)

    # GROUPE FRIGO
    gf_prod_row = find_row(frigo, gf_prod_code, "c_groupe frigo")
    gf_opt_row = find_row(frigo, gf_opt_code, "c_groupe frigo")
    write_block("B50", build_values(gf_prod_row, frigo.columns[0]), 6)
    write_block("B58", build_values(gf_opt_row, frigo.columns[0]), 2)

    # HAYON
    hay_prod_row = find_row(hayons, hay_prod_code, "c_hayon elevateur")
    hay_opt_row = find_row(hayons, hay_opt_code, "c_hayon elevateur")
    write_block("B61", build_values(hay_prod_row, hayons.columns[0]), 5)
    write_block("B68", build_values(hay_opt_row, hayons.columns[0]), 3)

    # ---- DIMENSIONS (colonnes exactes de ta BDD) ----
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

# Filtres principaux
df_filtre, code_pays = filtre_select(df_filtre, "code_pays", "Code pays")
df_filtre, marque = filtre_select(df_filtre, "Marque", "Marque")
df_filtre, modele = filtre_select(df_filtre, "Modele", "Modèle")
df_filtre, code_pf = filtre_select(df_filtre, "Code_PF", "Code PF")
df_filtre, std_pf = filtre_select(df_filtre, "Standard_PF", "Standard PF")

# Filtres composants : produit puis options (les options viennent du DF filtré)
df_filtre, cab_prod_choice = filtre_select(df_filtre, "C_Cabine", "Cabine - code produit")
df_filtre, cab_opt_choice = filtre_select_options(df_filtre, "C_Cabine-OPTIONS", "Cabine - code options")

df_filtre, mot_prod_choice = filtre_select(df_filtre, "M_moteur", "Moteur - code produit")
df_filtre, mot_opt_choice = filtre_select_options(df_filtre, "M_moteur-OPTIONS", "Moteur - code options")

df_filtre, ch_prod_choice = filtre_select(df_filtre, "C_Chassis", "Châssis - code produit")
df_filtre, ch_opt_choice = filtre_select_options(df_filtre, "C_Chassis-OPTIONS", "Châssis - code options")

df_filtre, caisse_prod_choice = filtre_select(df_filtre, "C_Caisse", "Caisse - code produit")
df_filtre, caisse_opt_choice = filtre_select_options(df_filtre, "C_Caisse-OPTIONS", "Caisse - code options")

df_filtre, gf_prod_choice = filtre_select(df_filtre, "C_Groupe frigo", "Groupe frigo - code produit")
df_filtre, gf_opt_choice = filtre_select_options(df_filtre, "C_Groupe frigo-OPTIONS", "Groupe frigo - code options")

df_filtre, hay_prod_choice = filtre_select(df_filtre, "C_Hayon elevateur", "Hayon - code produit")
df_filtre, hay_opt_choice = filtre_select_options(df_filtre, "C_Hayon elevateur-OPTIONS", "Hayon - code options")

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
    "code_pays",
    "Marque",
    "Modele",
    "Code_PF",
    "Standard_PF",
    "catalogue_1\n PF",
    "catalogue_2\nST",
    "catalogue_3\nZR",
    "catalogue_3\n LIBRE",
]
cols_existantes = [c for c in cols_synthese if c in veh.index]
st.table(veh[cols_existantes].to_frame(name="Valeur"))

# ----------------- APERÇU IMAGES -----------------

img_veh_val = veh.get("Image Vehicule")
img_client_val = veh.get("Image Client")
img_carbu_val = veh.get("Image Carburant")

img_veh_path = resolve_image_path(img_veh_val, "Image Vehicule")
img_client_path = resolve_image_path(img_client_val, "Image Client")
img_carbu_path = resolve_image_path(img_carbu_val, "Image Carburant")

st.subheader("Images associées")
col1, col2, col3 = st.columns(3)

with col1:
    show_image(img_veh_path, "Image véhicule")

with col2:
    show_image(img_client_path, "Image client")

with col3:
    show_image(img_carbu_path, "Picto carburant")

# ----------------- DETAIL COMPOSANTS (inchangé) -----------------

affiche_composant("Cabine", veh.get("C_Cabine"), cabines, "C_Cabine")
affiche_composant("Châssis", veh.get("C_Chassis"), chassis, "c_chassis")
affiche_composant("Caisse", veh.get("C_Caisse"), caisses, "c_caisse")
affiche_composant("Moteur", veh.get("M_moteur"), moteurs, "M_moteur")
affiche_composant("Groupe frigorifique", veh.get("C_Groupe frigo"), frigo, "c_groupe frigo")
affiche_composant("Hayon élévateur", veh.get("C_Hayon elevateur"), hayons, "c_hayon elevateur")

# ----------------- BOUTON DE TÉLÉCHARGEMENT FT -----------------

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
