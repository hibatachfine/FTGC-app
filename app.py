import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from copy import copy

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell

APP_VERSION = "2026-01-05_fix_insert_rows_preserve_merges_titles_fullwidth"

# ----------------- CONFIG APP -----------------
st.set_page_config(page_title="FT Grands Comptes", page_icon="🚚", layout="wide")
st.title("Generateur de Fiches Techniques Grands Comptes")
st.caption("Version de test basée sur bdd_CG.xlsx")
st.sidebar.info(f"✅ Version: {APP_VERSION}")

IMG_ROOT = "images"  # dossier racine des images dans le repo


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


# ----------------- SYNTH COMPONENT -----------------
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


# ----------------- EXCEL GENERATION -----------------
def genere_ft_excel(
    veh,
    cab_prod_choice, cab_opt_choice,
    mot_prod_choice, mot_opt_choice,
    ch_prod_choice, ch_opt_choice,
    caisse_prod_choice, caisse_opt_choice,
    gf_prod_choice, gf_opt_choice,
    hay_prod_choice, hay_opt_choice,
    cabines, moteurs, chassis, caisses, frigo, hayons
):
    template_path = "FT_Grand_Compte.xlsx"
    if not os.path.exists(template_path):
        st.error("Le fichier modèle 'FT_Grand_Compte.xlsx' n'est pas présent dans le repo.")
        return None

    wb = load_workbook(template_path, read_only=False, data_only=False)
    if "date" in wb.sheetnames:
        ws = wb["date"]
    elif "data" in wb.sheetnames:
        ws = wb["data"]
    else:
        ws = wb[wb.sheetnames[0]]

    # ---------- Excel helpers ----------
    def cell_to_rc(cell_addr: str):
        col_letters = "".join(ch for ch in cell_addr if ch.isalpha())
        row_digits = "".join(ch for ch in cell_addr if ch.isdigit())
        return column_index_from_string(col_letters), int(row_digits)

    def merged_range_containing(row, col):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return rng
        return None

    def merged_top_left(row, col):
        rng = merged_range_containing(row, col)
        if rng:
            return rng.min_row, rng.min_col
        return row, col

    def col_width(col_idx: int) -> float:
        letter = get_column_letter(col_idx)
        w = ws.column_dimensions[letter].width
        return float(w) if w is not None else 8.43

    def snap_to_wide_area(row: int, col: int, search_right: int = 18):
        """
        Si on pointe une zone étroite, on se décale vers la zone la plus large à droite
        (utile si merges ont été déplacés ou si ancre tombe sur une colonne d'indent).
        """
        best_left = col
        best_score = -1e18
        start_w = col_width(col)

        for c in range(col, col + search_right + 1):
            rng = merged_range_containing(row, c)
            if rng:
                width_cols = rng.max_col - rng.min_col + 1
                left = rng.min_col
            else:
                width_cols = 1
                left = c

            dist = abs(c - col)
            w = col_width(c)
            score = (width_cols * 120.0) + (w * 2.0) - (dist * 10.0)

            if start_w < 4.5 and c == col:
                score -= 500

            if score > best_score:
                best_score = score
                best_left = left

        return row, best_left

    def set_cell_value_merged_safe(row, col, value):
        row, col = snap_to_wide_area(row, col)
        r0, c0 = merged_top_left(row, col)
        cell = ws.cell(row=r0, column=c0)
        if isinstance(cell, MergedCell):
            return

        cell.value = value

        # garder style du template, on force juste wrap_text=True
        try:
            al = copy(cell.alignment)
            al.wrap_text = True
            cell.alignment = al
        except Exception:
            pass

    def write_block_merged_safe(start_rc, values, n_rows):
        start_col, start_row = start_rc
        for i in range(n_rows):
            v = values[i] if i < len(values) else None
            set_cell_value_merged_safe(start_row + i, start_col, v)

    def _collect_merges():
        out = []
        for rng in list(ws.merged_cells.ranges):
            out.append((rng.min_row, rng.max_row, rng.min_col, rng.max_col))
        return out

    def _unmerge_all(merges):
        for (r1, r2, c1, c2) in merges:
            try:
                ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
            except Exception:
                pass

    def _remerge_shifted(merges, insert_at_row, n):
        """
        Recrée TOUS les merges du template en les décalant/étendant correctement :
        - merges entièrement sous l'insertion -> shift
        - merges qui traversent l'insertion -> expand (max_row += n)
        """
        for (r1, r2, c1, c2) in merges:
            nr1, nr2 = r1, r2
            if r2 < insert_at_row:
                pass
            elif r1 >= insert_at_row:
                nr1 += n
                nr2 += n
            else:
                nr2 += n  # merge traverse l'endroit où on insère -> on étend

            try:
                ws.merge_cells(start_row=nr1, start_column=c1, end_row=nr2, end_column=c2)
            except Exception:
                pass

    def insert_rows_and_shift_keep_layout(
        anchors_dict: dict,
        insert_at_row: int,
        n: int,
        template_row: int,
        max_col_letter: str = "F",   # pages 2/3 chez toi = jusqu'à F (sinon mets "L")
    ):
        """
        ✅ Insert rows SANS casser la mise en page :
        - on sauvegarde toutes les fusions, on les enlève
        - on insère les lignes
        - on recrée les fusions en les shiftant/étendant
        - on clone styles/hauteur depuis template_row sur les lignes ajoutées
        """
        if n <= 0:
            return anchors_dict

        max_col = column_index_from_string(max_col_letter)

        # 0) sauver + retirer merges (sinon insert_rows les casse chez openpyxl)
        merges_before = _collect_merges()
        _unmerge_all(merges_before)

        # 1) insert rows
        ws.insert_rows(insert_at_row, n)

        # 2) remettre les merges correctement shiftés/étendus
        _remerge_shifted(merges_before, insert_at_row, n)

        # 3) shift page breaks
        try:
            for br in ws.row_breaks.brk:
                if br.id >= insert_at_row:
                    br.id += n
        except Exception:
            pass

        # 4) clone row height
        base_height = ws.row_dimensions[template_row].height

        # 5) clone styles (cell-by-cell) sur nouvelles lignes
        for i in range(n):
            new_r = insert_at_row + i

            if base_height is not None:
                ws.row_dimensions[new_r].height = base_height

            for c in range(1, max_col + 1):
                src = ws.cell(row=template_row, column=c)
                dst = ws.cell(row=new_r, column=c)

                # si src est une MergedCell (pas top-left), pas fiable
                if isinstance(src, MergedCell):
                    continue

                if src.has_style:
                    dst._style = copy(src._style)
                    dst.font = copy(src.font)
                    dst.fill = copy(src.fill)
                    dst.border = copy(src.border)
                    dst.number_format = src.number_format
                    dst.protection = copy(src.protection)
                    dst.alignment = copy(src.alignment)

        # 6) update anchors
        new_anchors = {}
        for k, (c, r) in anchors_dict.items():
            new_anchors[k] = (c, r + n) if r >= insert_at_row else (c, r)
        return new_anchors

    # ---------- Data helpers ----------
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
                    return ca
