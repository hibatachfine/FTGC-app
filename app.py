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

APP_VERSION = "2026-01-05_force_fullwidth_merges_pages_2_3"

st.set_page_config(page_title="FT Grands Comptes", page_icon="🚚", layout="wide")
st.title("Generateur de Fiches Techniques Grands Comptes")
st.caption("Version de test basée sur bdd_CG.xlsx")
st.sidebar.info(f"✅ Version: {APP_VERSION}")

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
    ws = wb["date"] if "date" in wb.sheetnames else wb[wb.sheetnames[0]]

    # ✅ Réglage largeur pages 2/3 : c’est la zone blanche de texte.
    # Sur ta capture c’est B → F. Si ta zone va plus loin, mets "L".
    P23_TEXT_START_COL = column_index_from_string("B")
    P23_TEXT_END_COL = column_index_from_string("F")

    # ---------- excel helpers ----------
    def cell_to_rc(cell_addr: str):
        col_letters = "".join(ch for ch in cell_addr if ch.isalpha())
        row_digits = "".join(ch for ch in cell_addr if ch.isdigit())
        return column_index_from_string(col_letters), int(row_digits)

    def _collect_merges():
        return [(rng.min_row, rng.max_row, rng.min_col, rng.max_col) for rng in list(ws.merged_cells.ranges)]

    def _unmerge_all(merges):
        for (r1, r2, c1, c2) in merges:
            try:
                ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
            except Exception:
                pass

    def _remerge_shifted(merges, insert_at_row, n):
        # shift / expand merges properly when inserting rows
        for (r1, r2, c1, c2) in merges:
            nr1, nr2 = r1, r2
            if r2 < insert_at_row:
                pass
            elif r1 >= insert_at_row:
                nr1 += n
                nr2 += n
            else:
                nr2 += n  # merge crosses insertion -> expand
            try:
                ws.merge_cells(start_row=nr1, start_column=c1, end_row=nr2, end_column=c2)
            except Exception:
                pass

    def insert_rows_preserve_merges(insert_at_row: int, n: int):
        if n <= 0:
            return
        merges_before = _collect_merges()
        _unmerge_all(merges_before)
        ws.insert_rows(insert_at_row, n)
        _remerge_shifted(merges_before, insert_at_row, n)

        # shift page breaks too
        try:
            for br in ws.row_breaks.brk:
                if br.id >= insert_at_row:
                    br.id += n
        except Exception:
            pass

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
        try:
            al = copy(cell.alignment)
            al.wrap_text = True
            cell.alignment = al
        except Exception:
            pass

    def write_block(start_rc, values, n_rows):
        start_col, start_row = start_rc
        for i in range(n_rows):
            v = values[i] if i < len(values) else None
            set_cell_value_merged_safe(start_row + i, start_col, v)

    def copy_row_styles(template_row: int, target_row: int, max_col_letter: str):
        max_col = column_index_from_string(max_col_letter)
        if ws.row_dimensions[template_row].height is not None:
            ws.row_dimensions[target_row].height = ws.row_dimensions[template_row].height
        for c in range(1, max_col + 1):
            src = ws.cell(row=template_row, column=c)
            dst = ws.cell(row=target_row, column=c)
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

    def force_fullwidth_merges(start_row: int, n_rows: int, start_col: int, end_col: int, template_row: int):
        """
        ✅ Force des merges horizontaux start_col→end_col pour chaque ligne.
        Ça donne une vraie largeur de wrap sur toute la zone blanche.
        """
        # Unmerge merges qui touchent la zone (mais qui ne commencent pas en A)
        merges_to_remove = []
        for rng in list(ws.merged_cells.ranges):
            if rng.min_col >= start_col and not (rng.max_col < start_col or rng.min_col > end_col):
                if not (rng.max_row < start_row or rng.min_row > start_row + n_rows - 1):
                    merges_to_remove.append(rng)

        for rng in merges_to_remove:
            try:
                ws.unmerge_cells(str(rng))
            except Exception:
                pass

        # Merge par ligne + recopier le style du template_row (col start_col)
        src_cell = ws.cell(row=template_row, column=start_col)
        for i in range(n_rows):
            r = start_row + i
            try:
                ws.merge_cells(start_row=r, start_column=start_col, end_row=r, end_column=end_col)
            except Exception:
                pass

            dst_cell = ws.cell(row=r, column=start_col)
            if not isinstance(src_cell, MergedCell) and src_cell.has_style:
                dst_cell._style = copy(src_cell._style)
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.number_format = src_cell.number_format
                dst_cell.protection = copy(src_cell.protection)
                dst_cell.alignment = copy(src_cell.alignment)

            # wrap on
            try:
                al = copy(dst_cell.alignment)
                al.wrap_text = True
                dst_cell.alignment = al
            except Exception:
                pass

    # ---------- data helpers ----------
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

    if img_veh_path and os.path.exists(img_veh_path):
        xl_img_veh = XLImage(img_veh_path)
        xl_img_veh.anchor = "B15"
        ws.add_image(xl_img_veh)

    if img_client_path and os.path.exists(img_client_path):
        xl_img_client = XLImage(img_client_path)
        xl_img_client.anchor = "H2"
        ws.add_image(xl_img_client)

    if img_carbu_path and os.path.exists(img_carbu_path):
        xl_img_carbu = XLImage(img_carbu_path)
        xl_img_carbu.anchor = "H15"
        ws.add_image(xl_img_carbu)

    # ---- COMPOSANTS ----
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
        "CAB_OPT":   cell_to_rc("B38"),
        "MOT_OPT":   cell_to_rc("F38"),
        "CH_OPT":    cell_to_rc("H38"),

        "CAISSE_START": cell_to_rc("B40"),
        "CAISSE_OPT":   cell_to_rc("B47"),
        "GF_START":     cell_to_rc("B50"),
        "GF_OPT":       cell_to_rc("B58"),
        "HAY_START":    cell_to_rc("B61"),
        "HAY_OPT":      cell_to_rc("B68"),
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

    def ensure_space(anchor_key: str, base_rows: int, needed_rows: int, template_row: int, max_col_letter: str):
        extra = max(0, int(needed_rows) - int(base_rows))
        if extra <= 0:
            return
        start_col, start_row = anchors[anchor_key]
        insert_at = start_row + int(base_rows)
        insert_rows_preserve_merges(insert_at, extra)

        # copier style sur les nouvelles lignes
        for i in range(extra):
            copy_row_styles(template_row=template_row, target_row=insert_at + i, max_col_letter=max_col_letter)

        # shift anchors after insertion
        new_anchors = {}
        for k, (c, r) in anchors.items():
            new_anchors[k] = (c, r + extra) if r >= insert_at else (c, r)
        anchors.clear()
        anchors.update(new_anchors)

    # ---- WRITE blocks TOP (page 1) ----
    top_needed = max(len(cab_vals), len(mot_vals), len(ch_vals), 1)
    ensure_space("CAB_START", BASE["TOP_MAIN"], top_needed, template_row=anchors["CAB_START"][1], max_col_letter="L")
    write_block(anchors["CAB_START"], cab_vals, top_needed)
    write_block(anchors["MOT_START"], mot_vals, top_needed)
    write_block(anchors["CH_START"],  ch_vals,  top_needed)

    top_opt_needed = max(len(cab_opt_vals), len(mot_opt_vals), len(ch_opt_vals), 1)
    ensure_space("CAB_OPT", BASE["TOP_OPT"], top_opt_needed, template_row=anchors["CAB_OPT"][1], max_col_letter="L")
    write_block(anchors["CAB_OPT"], cab_opt_vals, top_opt_needed)
    write_block(anchors["MOT_OPT"], mot_opt_vals, top_opt_needed)
    write_block(anchors["CH_OPT"],  ch_opt_vals,  top_opt_needed)

    # ---- PAGES 2/3 : force full-width merges B→F ----
    def write_p23(anchor_key: str, base_rows: int, values: list):
        needed = max(len(values), 1)
        start_col, start_row = anchors[anchor_key]

        # 1) add rows if needed (preserve merges / titles)
        ensure_space(anchor_key, base_rows, needed, template_row=start_row, max_col_letter="F")

        # 2) force merges B→F for every line we will write
        force_fullwidth_merges(
            start_row=anchors[anchor_key][1],
            n_rows=needed,
            start_col=P23_TEXT_START_COL,
            end_col=P23_TEXT_END_COL,
            template_row=anchors[anchor_key][1]
        )

        # 3) write line by line into column B (top-left of merge)
        write_block((P23_TEXT_START_COL, anchors[anchor_key][1]), values, needed)

    write_p23("CAISSE_START", BASE["CAISSE_MAIN"], caisse_vals)
    write_p23("CAISSE_OPT",   BASE["CAISSE_OPT"],  caisse_opt_vals)
    write_p23("GF_START",     BASE["GF_MAIN"],     gf_vals)
    write_p23("GF_OPT",       BASE["GF_OPT"],      gf_opt_vals)
    write_p23("HAY_START",    BASE["HAY_MAIN"],    hay_vals)
    write_p23("HAY_OPT",      BASE["HAY_OPT"],     hay_opt_vals)

    # ---- DIMENSIONS ----
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

cab_prod_code, cab_opt_code = choose_codes(cab_prod_choice, cab_opt_choice, veh.get("C_Cabine"), veh.get("C_Cabine-OPTIONS"))
mot_prod_code, mot_opt_code = choose_codes(mot_prod_choice, mot_opt_choice, veh.get("M_moteur"), veh.get("M_moteur-OPTIONS"))
ch_prod_code, ch_opt_code = choose_codes(ch_prod_choice, ch_opt_choice, veh.get("C_Chassis"), veh.get("C_Chassis-OPTIONS"))
caisse_prod_code, caisse_opt_code = choose_codes(caisse_prod_choice, caisse_opt_choice, veh.get("C_Caisse"), veh.get("C_Caisse-OPTIONS"))
gf_prod_code, gf_opt_code = choose_codes(gf_prod_choice, gf_opt_choice, veh.get("C_Groupe frigo"), veh.get("C_Groupe frigo-OPTIONS"))
hay_prod_code, hay_opt_code = choose_codes(hay_prod_choice, hay_opt_choice, veh.get("C_Hayon elevateur"), veh.get("C_Hayon elevateur-OPTIONS"))

affiche_composant("Cabine (Produit)", cab_prod_code, cabines, "C_Cabine", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Cabine (Options)", cab_opt_code, cabines, "C_Cabine", code_pf_for_fallback=code_pf_ref, prefer_po="O")

affiche_composant("Moteur (Produit)", mot_prod_code, moteurs, "M_moteur", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Moteur (Options)", mot_opt_code, moteurs, "M_moteur", code_pf_for_fallback=code_pf_ref, prefer_po="O")

affiche_composant("Châssis (Produit)", ch_prod_code, chassis, "CH_chassis", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Châssis (Options)", ch_opt_code, chassis, "CH_chassis", code_pf_for_fallback=code_pf_ref, prefer_po="O")

affiche_composant("Caisse (Produit)", caisse_prod_code, caisses, "CF_caisse", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Caisse (Options)", caisse_opt_code, caisses, "CF_caisse", code_pf_for_fallback=code_pf_ref, prefer_po="O")

affiche_composant("Groupe frigo (Produit)", gf_prod_code, frigo, "GF_groupe frigo", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Groupe frigo (Options)", gf_opt_code, frigo, "GF_groupe frigo", code_pf_for_fallback=code_pf_ref, prefer_po="O")

affiche_composant("Hayon (Produit)", hay_prod_code, hayons, "HL_hayon elevateur", code_pf_for_fallback=code_pf_ref, prefer_po="P")
affiche_composant("Hayon (Options)", hay_opt_code, hayons, "HL_hayon elevateur", code_pf_for_fallback=code_pf_ref, prefer_po="O")

st.markdown("---")
st.subheader("Génération de la fiche technique")

if st.button("⚙️ Générer la FT (Excel)"):
    ft_file = genere_ft_excel(
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
        filename = f"FT_{str(veh.get('Code_PF','')).strip() or 'vehicule'}.xlsx"
        st.success("✅ Fiche générée !")
        st.download_button(
            label="⬇️ Télécharger la fiche Excel",
            data=ft_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
