import re
import unicodedata
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

try:
    import openpyxl
    from openpyxl import Workbook
except Exception:
    openpyxl = None
    Workbook = None


# =========================
# CONFIG (bdd_CG-JL.xlsx)
# =========================
BDD_FALLBACK = "bdd_CG-JL.xlsx"

# Colonnes "code produit" par feuille POUR bdd_CG-JL.xlsx
CODE_COL = {
    "CABINES": "c_cabine",
    "MOTEURS": "m_moteur",
    "CHASSIS": "ch_chassis",            # IMPORTANT : pas "c_chassis"
    "CAISSES": "cf_caisse",
    "FRIGO": "gf_groupe_frigo",
    "HAYONS": "hl_hayon_elevateur",
}

SHEETS = list(CODE_COL.keys())


# =========================
# HELPERS
# =========================
_BAD_TOKENS = {"na", "-", "_", ""}

def _pick_header_label(raw: object) -> str:
    """
    Dans ta bdd, les headers sont multi-lignes.
    On prend la première ligne utile (≠ NA / - / _), sinon la 1ère non vide.
    """
    s = "" if raw is None else str(raw)
    lines = [ln.strip() for ln in s.splitlines() if ln.strip()]
    for ln in lines:
        if ln.lower() not in _BAD_TOKENS:
            return ln
    return lines[0] if lines else ""

def _strip_accents(s: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFD", s)
        if unicodedata.category(ch) != "Mn"
    )

def _clean_colname(raw: object) -> str:
    label = _pick_header_label(raw)
    label = _strip_accents(str(label)).lower().strip()
    # supprime ponctuation, normalise espaces, puis espace -> underscore
    label = re.sub(r"[^\w\s-]", "", label)
    label = re.sub(r"\s+", " ", label).strip().replace(" ", "_")
    return label

def normalize_df_cols(df: pd.DataFrame) -> pd.DataFrame:
    """
    - normalise les colonnes (multi-lignes -> 1 label)
    - drop la colonne concat (header énorme qui recopie tout)
    - garantit unicité (suffixes _2, _3...)
    """
    df = df.copy()

    keep_raw = []
    new_cols = []
    seen = {}

    for raw in df.columns:
        raw_s = "" if raw is None else str(raw)

        # drop "concat" (très long header avec énormément de \n)
        if raw_s.count("\n") > 20 and len(raw_s) > 200:
            continue

        col = _clean_colname(raw)
        if not col:
            col = "col"

        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 1

        keep_raw.append(raw)
        new_cols.append(col)

    df = df[keep_raw]
    df.columns = new_cols
    return df

def _norm_code(x: object) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s.lower() in {"nan", "none"}:
        return ""
    return s

def find_row(df: pd.DataFrame, code: object, code_col: str):
    """
    Retourne la 1ère ligne matching (Series) ou None.
    """
    target = _norm_code(code)
    if not target:
        return None

    if code_col not in df.columns:
        raise KeyError(f"Colonne '{code_col}' introuvable. Colonnes dispo: {list(df.columns)}")

    col = df[code_col].map(_norm_code)
    cand = df[col == target]
    if cand.empty:
        return None
    return cand.iloc[0]

def list_codes(df: pd.DataFrame, code_col: str) -> list[str]:
    if code_col not in df.columns:
        return []
    vals = df[code_col].map(_norm_code)
    vals = [v for v in vals.tolist() if v]
    return sorted(set(vals))

@st.cache_data(show_spinner=False)
def load_bdd(bdd_source) -> dict[str, pd.DataFrame]:
    out = {}
    for sheet in SHEETS:
        df = pd.read_excel(bdd_source, sheet_name=sheet)
        df = normalize_df_cols(df)
        out[sheet] = df
    return out


# =========================
# EXPORT EXCEL (simple)
# =========================
def export_selection_to_excel(selection: dict) -> bytes:
    """
    Exporte un Excel simple (summary + 1 feuille par composant).
    """
    if Workbook is None:
        raise RuntimeError("openpyxl n'est pas dispo. Ajoute 'openpyxl' dans requirements.txt")

    wb = Workbook()
    ws = wb.active
    ws.title = "SUMMARY"

    ws.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])
    ws.append(["Component", "Code", "Found"])
    for comp, data in selection.items():
        ws.append([comp, data.get("code", ""), "YES" if data.get("row") is not None else "NO"])

    # feuilles détaillées
    for comp, data in selection.items():
        ws2 = wb.create_sheet(title=comp[:31])
        ws2.append(["Code", data.get("code", "")])
        ws2.append([])

        row = data.get("row")
        if row is None:
            ws2.append(["NOT FOUND"])
            continue

        # row est une Series
        ws2.append(["Field", "Value"])
        for k, v in row.to_dict().items():
            ws2.append([k, "" if pd.isna(v) else str(v)])

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="FTGC", layout="wide")
st.title("FTGC")

with st.sidebar:
    st.header("Base de données")
    uploaded_bdd = st.file_uploader("Upload bdd_CG-JL.xlsx", type=["xlsx"])
    use_debug = st.checkbox("Debug colonnes", value=False)

bdd_source = uploaded_bdd if uploaded_bdd is not None else BDD_FALLBACK

try:
    dfs = load_bdd(bdd_source)
except Exception as e:
    st.error("Impossible de charger la base de données.")
    st.exception(e)
    st.stop()

# Debug colonnes
if use_debug:
    st.subheader("Colonnes détectées (après normalisation)")
    for sh in SHEETS:
        st.write(f"**{sh}** → code_col attendu: `{CODE_COL[sh]}`")
        st.write(list(dfs[sh].columns))

st.divider()
st.subheader("Sélection composants")

cols = st.columns(3)
selection = {}

for i, sh in enumerate(SHEETS):
    df = dfs[sh]
    code_col = CODE_COL[sh]
    codes = list_codes(df, code_col)

    with cols[i % 3]:
        st.markdown(f"### {sh}")
        if not codes:
            st.warning(f"Aucune liste de codes (colonne `{code_col}` introuvable).")
            code = st.text_input(f"Code {sh}", key=f"code_{sh}")
        else:
            code = st.selectbox(f"Code {sh}", options=[""] + codes, key=f"code_{sh}")

        row = None
        err = None
        if code:
            try:
                row = find_row(df, code, code_col)
            except Exception as e:
                err = e

        if err is not None:
            st.error("Erreur recherche code")
            st.exception(err)
        elif code and row is None:
            st.warning("Code introuvable dans la feuille.")
        elif row is not None:
            # Affiche quelques champs
            preview = row.to_dict()
            preview = {k: ("" if pd.isna(v) else v) for k, v in preview.items()}
            st.dataframe(pd.DataFrame([preview]), use_container_width=True)

        selection[sh] = {"code": code, "row": row}

st.divider()
st.subheader("Export")

if openpyxl is None:
    st.warning("⚠️ openpyxl manquant. Ajoute `openpyxl` dans requirements.txt pour exporter un Excel.")

btn_col1, btn_col2 = st.columns([1, 2])
with btn_col1:
    if st.button("Générer Excel"):
        try:
            out = export_selection_to_excel(selection)
            st.session_state["last_xlsx"] = out
            st.success("Excel généré ✅")
        except Exception as e:
            st.error("Échec génération Excel")
            st.exception(e)

with btn_col2:
    if "last_xlsx" in st.session_state:
        st.download_button(
            "Télécharger le fichier Excel",
            data=st.session_state["last_xlsx"],
            file_name="FTGC_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
