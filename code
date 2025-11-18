import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# ----------------- CONFIG APP -----------------

st.set_page_config(
    page_title="FT Grand Compte",
    page_icon="🚚",
    layout="wide"
)

st.title("Générateur de Fiches Techniques Grand Compte")
st.caption("Version de test basée sur bdd_CG.xlsx")

IMG_ROOT = "images"  # dossier racine des images dans le repo


# ----------------- FONCTIONS UTILES -----------------

def resolve_image_path(cell_value, subdir):
    """
    Prend la valeur de la cellule Excel (chemin complet, nom de fichier ou URL)
    et renvoie un chemin utilisable par Streamlit / openpyxl.
    - subdir = 'vehicules', 'clients', 'carburant'
    """
    if not isinstance(cell_value, str) or not cell_value.strip():
        return None

    val = cell_value.strip()

    # Cas URL -> on renvoie tel quel
    if val.lower().startswith("http://") or val.lower().startswith("https://"):
        return val

    # Cas chemin Windows / nom simple -> on garde juste le fichier
    filename = os.path.basename(val)
    return os.path.join(IMG_ROOT, subdir, filename)


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
    """Créé un selectbox 'Tous + valeurs uniques' et renvoie le DF filtré."""
    if col not in df.columns:
        st.error(f"Colonne manquante dans la base : {col}")
        return df, None

    options = sorted(
        [v for v in df[col].dropna().unique().tolist() if str(v).strip() != ""]
    )
    options_display = ["Tous"] + options
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


def affiche_composant(titre, code, df_ref, col_code_ref):
    st.markdown("---")
    st.subheader(titre)

    if pd.isna(code) or str(code).strip() == "":
        st.info("Aucun code renseigné pour ce composant.")
        return

    st.write(f"Code composant : **{code}**")

    comp = df_ref[df_ref[col_code_ref] == code]

    if comp.empty:
        st.warning("Code non trouvé dans la base de référence.")
        return

    comp_row = comp.iloc[0].dropna()
    st.table(comp_row.to_frame(name="Valeur"))


def genere_ft_excel(veh):
    """
    Génère une fiche technique Excel à partir du modèle 'FT_Grand_Compte.xlsx'
    et de la ligne véhicule sélectionnée.
    A ADAPTER : les cellules Excel utilisées pour chaque champ.
    """

    # 1. Charger le modèle
    template_path = "FT_Grand_Compte.xlsx"   # <- adapte le nom si différent
    if not os.path.exists(template_path):
        st.error("Le fichier modèle 'FT_Grand_Compte.xlsx' est introuvable dans le repo.")
        return None

    wb = load_workbook(template_path)
    ws = wb.active  # ou wb['NomDeLaFeuille'] si besoin

    # 2. Remplir quelques champs texte (ADAPTER LES CELLULES)
    #  --> Mets ici les correspondances exactes entre champs de la BDD
    #      et les cellules de ta fiche technique.
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

    # 3. Images (véhicule / client / carburant)
    img_veh_val = veh.get("Image Vehicule")
    img_client_val = veh.get("Image Client")
    img_carbu_val = veh.get("Image Carburant")

    img_veh_path = resolve_image_path(img_veh_val, "vehicules")
    img_client_path = resolve_image_path(img_client_val, "clients")
    img_carbu_path = resolve_image_path(img_carbu_val, "carburant")

    # Logo PF fixe (si tu veux en mettre un)
    logo_pf_path = os.path.join(IMG_ROOT, "logo_pf.png")
    if os.path.exists(logo_pf_path):
        xl_img = XLImage(logo_pf_path)
        xl_img.anchor = "B2"   # cellule d’ancrage à adapter
        ws.add_image(xl_img)

    if img_veh_path and os.path.exists(img_veh_path):
        xl_img_veh = XLImage(img_veh_path)
        xl_img_veh.anchor = "B15"  # à adapter à la mise en page de ta FT
        ws.add_image(xl_img_veh)

    if img_client_path and os.path.exists(img_client_path):
        xl_img_client = XLImage(img_client_path)
        xl_img_client.anchor = "H2"  # à adapter
        ws.add_image(xl_img_client)

    if img_carbu_path and os.path.exists(img_carbu_path):
        xl_img_carbu = XLImage(img_carbu_path)
        xl_img_carbu.anchor = "H15"  # à adapter
        ws.add_image(xl_img_carbu)

    # 4. Sauvegarde dans un buffer pour téléchargement
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

df_filtre, cab_code = filtre_select(df_filtre, "C_Cabine", "Cabine")
df_filtre, ch_code = filtre_select(df_filtre, "C_Chassis", "Châssis")
df_filtre, caisse_code = filtre_select(df_filtre, "C_Caisse", "Caisse")
df_filtre, mot_code = filtre_select(df_filtre, "M_moteur", "Moteur")
df_filtre, gf_code = filtre_select(df_filtre, "C_Groupe frigo", "Groupe frigorifique")
df_filtre, hay_code = filtre_select(df_filtre, "C_Hayon elevateur", "Hayon élévateur")

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

# ----------------- APERCU IMAGES -----------------

img_veh_val = veh.get("Image Vehicule")
img_client_val = veh.get("Image Client")
img_carbu_val = veh.get("Image Carburant")

img_veh_path = resolve_image_path(img_veh_val, "vehicules")
img_client_path = resolve_image_path(img_client_val, "clients")
img_carbu_path = resolve_image_path(img_carbu_val, "carburant")

st.subheader("Images associées")
col1, col2, col3 = st.columns(3)

with col1:
    st.caption("Image véhicule")
    if img_veh_path:
        st.image(img_veh_path)
    else:
        st.info("Pas d'image véhicule définie")

with col2:
    st.caption("Image client")
    if img_client_path:
        st.image(img_client_path)
    else:
        st.info("Pas d'image client définie")

with col3:
    st.caption("Picto carburant")
    if img_carbu_path:
        st.image(img_carbu_path)
    else:
        st.info("Pas de picto carburant défini")


# ----------------- DETAIL COMPOSANTS -----------------

affiche_composant("Cabine", veh.get("C_Cabine"), cabines, "C_Cabine")
affiche_composant("Châssis", veh.get("C_Chassis"), chassis, "c_chassis")
affiche_composant("Caisse", veh.get("C_Caisse"), caisses, "c_caisse")
affiche_composant("Moteur", veh.get("M_moteur"), moteurs, "M_moteur")
affiche_composant("Groupe frigorifique", veh.get("C_Groupe frigo"), frigo, "c_groupe frigo")
affiche_composant("Hayon élévateur", veh.get("C_Hayon elevateur"), hayons, "c_hayon elevateur")


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
