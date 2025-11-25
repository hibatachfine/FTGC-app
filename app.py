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

IMG_ROOT = "images"  # dossier racine des images (Image Vehicule / Image Client / Image Carburant)


# ----------------- FONCTIONS UTILES -----------------

def resolve_image_path(cell_value, subdir):
    """Transforme la valeur Excel (chemin, nom de fichier ou URL) en chemin exploitable."""
    if not isinstance(cell_value, str) or not cell_value.strip():
        return None

    val = cell_value.strip()

    # URL
    if val.lower().startswith(("http://", "https://")):
        return val

    # Chemin local / nom de fichier
    filename = os.path.basename(val)
    return os.path.join(IMG_ROOT, subdir, filename)


def show_image(path_or_url, caption):
    """Affiche une image si possible, sinon un message explicite."""
    st.caption(caption)

    if not path_or_url:
        st.info("Pas d'image définie")
        return

    # URL
    if isinstance(path_or_url, str) and path_or_url.lower().startswith(("http://", "https://")):
        st.image(path_or_url)
        return

    # Fichier local
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


def filtre_select(df, col, label):
    """Selectbox dans la sidebar sur une colonne (avec 'Tous')."""
    if col not in df.columns:
        # On ne casse rien si la colonne d'options n'existe pas
        st.sidebar.write(f"(colonne '{col}' absente)")
        return df, None

    vals = df[col].dropna()
    vals = [v for v in vals.unique().tolist() if str(v).strip() != ""]
    vals = sorted(vals)

    options_display = ["Tous"] + vals
    choix = st.sidebar.selectbox(label, options_display)

    if choix != "Tous":
        df = df[df[col] == choix]

    return df, choix


def format_vehicule(row):
    champs = []
    for c in ["code_pays", "Marque", "Modele", "Code_PF", "Standard_PF"]:
        if c in row and pd.notna(row[c]):
            champs.append(str(row[c]))
    return " – ".join(champs)


# ----------------- GENERATION DE LA FT -----------------

def genere_ft_excel(veh):
    """
    Génère une fiche technique Excel à partir du modèle 'FT_Grand_Compte.xlsx'
    et de la ligne véhicule sélectionnée.

    Pour chaque composant (Cabine, Moteur, Châssis, Caisse, Groupe frigo, Hayon) :
      - récupère la ligne PRODUIT (P) dans l'onglet de référence (code sans -OPTIONS)
      - récupère la ligne OPTION (O) (code ...-OPTIONS)
    Dans la FT on écrit uniquement les VALEURS (2 places, Vitres électriques…).
    """

    template_path = "FT_Grand_Compte.xlsx"

    if not os.path.exists(template_path):
        st.info(
            "Le fichier modèle 'FT_Grand_Compte.xlsx' n'est pas présent dans le repo. "
            "Ajoute-le à la racine pour activer la génération automatique."
        )
        return None

    wb = load_workbook(template_path, read_only=False, data_only=False)
    ws = wb["date"]  # nom de la feuille dans ton modèle

    # ---- helpers internes ----

    def get_prodopt_col(df):
        """Retourne la colonne 'Produit (P) / Option (O)' si elle existe."""
        for c in df.columns:
            name = str(c).lower()
            if "produit (p" in name and "option" in name:
                return c
        return None

    def build_lines_from_row(row_series, code_col):
        """
        Transforme une ligne de BDD composant en liste de valeurs (sans libellé).
        Ignore :
          - la colonne code (C_Cabine, M_moteur, c_chassis, ...)
          - la colonne Produit (P) / Option (O)
          - les 'zone libre'
          - la colonne '_' éventuelle
        """
        if row_series is None:
            return []

        lines = []
        for col, val in row_series.items():
            if pd.isna(val) or str(val).strip() == "":
                continue
            if col == code_col:
                continue
            name_lower = str(col).strip().lower()
            if name_lower.startswith("zone libre"):
                continue
            if "produit (p" in name_lower and "option" in name_lower:
                continue
            if col == "_":
                continue
            lines.append(str(val))  # uniquement la valeur
        return lines

    def fill_lines(ws_local, start_cell, lines, max_rows):
        """
        Ecrit les valeurs dans une seule colonne, à partir de start_cell,
        sur max_rows lignes max, en ignorant les cellules fusionnées.
        """
        if not lines:
            return

        col_letters = "".join(ch for ch in start_cell if ch.isalpha())
        row_digits = "".join(ch for ch in start_cell if ch.isdigit())
        start_col = column_index_from_string(col_letters)
        start_row = int(row_digits)

        line_idx = 0
        for i in range(max_rows):
            cell = ws_local.cell(row=start_row + i, column=start_col)

            # merged cell → on ne touche pas
            if isinstance(cell, MergedCell):
                continue

            if line_idx < len(lines):
                cell.value = lines[line_idx]
                line_idx += 1
            else:
                cell.value = None  # nettoyage

    def find_component_row_exact(df_ref, code_col, code, prod_or_opt=None):
        """
        Cherche STRICTEMENT le code dans la colonne code_col, et optionnellement P/O.
        Si plusieurs lignes -> on prend la première.
        """
        if not isinstance(code, str) or not code or code != code:
            return None

        df = df_ref
        prodopt_col = get_prodopt_col(df)

        mask = df[code_col] == code
        if prod_or_opt and prodopt_col is not None:
            mask &= df[prodopt_col] == prod_or_opt

        cand = df[mask]
        if cand.empty:
            return None
        return cand.iloc[0]

    # ----------- 1) EN-TÊTE VÉHICULE -----------

    mapping = {
        "code_pays": "C5",
        "Marque": "C6",
        "Modele": "C7",
        "Code_PF": "C8",
        "Standard_PF": "C9",
        "catalogue_1\n PF": "C10",
        "catalogue_2\nST": "C11",
        "catalogue_3\n LIBRE": "C12",
    }

    for col_bdd, cell_excel in mapping.items():
        if col_bdd in veh.index:
            ws[cell_excel] = veh[col_bdd]

    # ----------- 2) IMAGES -----------

    img_veh_val = veh.get("Image Vehicule")
    img_client_val = veh.get("Image Client")
    img_carbu_val = veh.get("Image Carburant")

    img_veh_path = resolve_image_path(img_veh_val, "Image Vehicule")
    img_client_path = resolve_image_path(img_client_val, "Image Client")
    img_carbu_path = resolve_image_path(img_carbu_val, "Image Carburant")

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

    # ----------- 3) DÉTAILS COMPOSANTS + OPTIONS -----------

    global cabines, moteurs, chassis, caisses, frigo, hayons

    # CABINE
    cab_code = veh.get("C_Cabine")
    cab_opt_code = veh.get("C_Cabine-OPTIONS")

    cab_row = find_component_row_exact(cabines, "C_Cabine", cab_code, prod_or_opt="P")
    cab_opt_row = find_component_row_exact(cabines, "C_Cabine", cab_opt_code, prod_or_opt="O")

    fill_lines(ws, "B18", build_lines_from_row(cab_row, "C_Cabine"), max_rows=17)
    fill_lines(ws, "B38", build_lines_from_row(cab_opt_row, "C_Cabine"), max_rows=3)

    # MOTEUR
    mot_code = veh.get("M_moteur")
    mot_opt_code = veh.get("M_moteur-OPTIONS")

    mot_row = find_component_row_exact(moteurs, "M_moteur", mot_code, prod_or_opt="P")
    mot_opt_row = find_component_row_exact(moteurs, "M_moteur", mot_opt_code, prod_or_opt="O")

    fill_lines(ws, "F18", build_lines_from_row(mot_row, "M_moteur"), max_rows=17)
    fill_lines(ws, "F38", build_lines_from_row(mot_opt_row, "M_moteur"), max_rows=3)

    # CHASSIS
    ch_code = veh.get("C_Chassis")
    ch_opt_code = veh.get("C_Chassis-OPTIONS")

    ch_row = find_component_row_exact(chassis, "c_chassis", ch_code, prod_or_opt="P")
    ch_opt_row = find_component_row_exact(chassis, "c_chassis", ch_opt_code, prod_or_opt="O")

    fill_lines(ws, "H18", build_lines_from_row(ch_row, "c_chassis"), max_rows=17)
    fill_lines(ws, "H38", build_lines_from_row(ch_opt_row, "c_chassis"), max_rows=3)

    # CAISSE (CARROSSERIE)
    caisse_code = veh.get("C_Caisse")
    caisse_opt_code = veh.get("C_Caisse-OPTIONS")

    caisse_row = find_component_row_exact(caisses, "c_caisse", caisse_code, prod_or_opt="P")
    caisse_opt_row = find_component_row_exact(caisses, "c_caisse", caisse_opt_code, prod_or_opt="O")

    fill_lines(ws, "B40", build_lines_from_row(caisse_row, "c_caisse"), max_rows=5)
    fill_lines(ws, "B47", build_lines_from_row(caisse_opt_row, "c_caisse"), max_rows=2)

    # GROUPE FRIGORIFIQUE
    gf_code = veh.get("C_Groupe frigo")
    gf_opt_code = veh.get("C_Groupe frigo-OPTIONS")

    gf_row = find_component_row_exact(frigo, "c_groupe frigo", gf_code, prod_or_opt="P")
    gf_opt_row = find_component_row_exact(frigo, "c_groupe frigo", gf_opt_code, prod_or_opt="O")

    fill_lines(ws, "B50", build_lines_from_row(gf_row, "c_groupe frigo"), max_rows=6)
    fill_lines(ws, "B58", build_lines_from_row(gf_opt_row, "c_groupe frigo"), max_rows=2)

    # HAYON
    hay_code = veh.get("C_Hayon elevateur")
    hay_opt_code = veh.get("C_Hayon elevateur-OPTIONS")

    hay_row = find_component_row_exact(hayons, "c_hayon elevateur", hay_code, prod_or_opt="P")
    hay_opt_row = find_component_row_exact(hayons, "c_hayon elevateur", hay_opt_code, prod_or_opt="O")

    fill_lines(ws, "B61", build_lines_from_row(hay_row, "c_hayon elevateur"), max_rows=5)
    fill_lines(ws, "B68", build_lines_from_row(hay_opt_row, "c_hayon elevateur"), max_rows=3)

    # ----------- 4) SAUVEGARDE -----------

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

# Filtres codes PRODUIT
df_filtre, cab_code = filtre_select(df_filtre, "C_Cabine", "Cabine")
df_filtre, ch_code = filtre_select(df_filtre, "C_Chassis", "Châssis")
df_filtre, caisse_code = filtre_select(df_filtre, "C_Caisse", "Caisse")
df_filtre, mot_code = filtre_select(df_filtre, "M_moteur", "Moteur")
df_filtre, gf_code = filtre_select(df_filtre, "C_Groupe frigo", "Groupe frigorifique")
df_filtre, hay_code = filtre_select(df_filtre, "C_Hayon elevateur", "Hayon élévateur")

# Filtres codes OPTIONS (…-OPTIONS)
df_filtre, cab_opt_code = filtre_select(df_filtre, "C_Cabine-OPTIONS", "Cabine - options")
df_filtre, ch_opt_code = filtre_select(df_filtre, "C_Chassis-OPTIONS", "Châssis - options")
df_filtre, caisse_opt_code = filtre_select(df_filtre, "C_Caisse-OPTIONS", "Caisse - options")
df_filtre, mot_opt_code = filtre_select(df_filtre, "M_moteur-OPTIONS", "Moteur - options")
df_filtre, gf_opt_code = filtre_select(df_filtre, "C_Groupe frigo-OPTIONS", "Groupe frigo - options")
df_filtre, hay_opt_code = filtre_select(df_filtre, "C_Hayon elevateur-OPTIONS", "Hayon - options")

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

# ----------------- BOUTON DE TÉLÉCHARGEMENT FT -----------------

st.markdown("---")
st.subheader("Génération de la fiche technique")

ft_file = genere_ft_excel(veh)

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
