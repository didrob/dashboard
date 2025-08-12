import io
import numpy as np
import pandas as pd
import streamlit as st

# ---------- App setup ----------
st.set_page_config(page_title="ILMS Ledger Viewer", layout="wide")
st.title("ILMS Ledger Viewer")
st.caption("Last opp Excel-filen (ILMS data TNG July 2025.xlsx) og filtrer/pivoter ledger-data.")

# ---------- Helpers ----------
@st.cache_data
def load_excel(uploaded_file):
    """
    Leser hele Excel-arbeidsboken til et dict {sheet_name: DataFrame}
    Bruker openpyxl eksplisitt (fungerer i Streamlit Cloud).
    """
    file_bytes = uploaded_file.getvalue()
    # sheet_name=None -> alle ark
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl")
    return sheets

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """Rydder vanlige problemer: tomme rader/kolonner, #N/A, feil header, typer."""
    df = df.dropna(axis=0, how='all').dropna(axis=1, how='all').copy()
    if df.empty:
        return df

    # Hvis mange "Unnamed" kolonner, løft en rad til header
    col_names = [str(c) for c in df.columns]
    if len(col_names) and sum(c.startswith("Unnamed") for c in col_names) / len(col_names) >= 0.5:
        nonnull_counts = df.notna().sum(axis=1)
        header_idx = int(nonnull_counts.idxmax())
        new_header = df.loc[header_idx].astype(str).tolist()
        df = df.loc[header_idx + 1:].copy()
        df.columns = new_header

    # Bytt ut NA-markører
    df = df.replace({"#N/A": np.nan, "N/A": np.nan, "NA": np.nan, "": np.nan})

    # Prøv å oppdage dato/tall automatisk i objekt-kolonner
    for c in df.columns:
        s = df[c]
        if s.dtype == "O":
            dt = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
            if dt.notna().mean() > 0.7:
                df[c] = dt.dt.date
            else:
                df[c] = pd.to_numeric(s, errors="ignore")
    return df

def unify_column_names(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=lambda x: str(x).strip())

def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce")

# ---------- UI: File upload ----------
uploaded = st.file_uploader("Dra hit Excel-filen din (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("👆 Last opp Excel-filen for å komme i gang.")
    st.stop()

with st.spinner("Leser Excel …"):
    sheets = load_excel(uploaded)

# ---------- Velg relevante ark og vask ----------
wanted = ["Ledger entry", "ILMS Ton", "ILMS Water", "ILMS Mooring", "Mapping"]
present = [s for s in wanted if s in sheets]

if not present:
    st.error("Fant ingen av arkene: " + ", ".join(wanted))
    st.stop()

cleaned = {name: unify_column_names(clean_df(sheets[name])) for name in present}

with st.expander("Forhåndsvis ark (før filtrering)"):
    for s in present:
        st.markdown(f"**{s}**  \n{cleaned[s].shape[0]} rader, {cleaned[s].shape[1]} kolonner")
        st.dataframe(cleaned[s].head(10), use_container_width=True)

# ---------- Velg 'base' (Ledger entry hvis mulig) ----------
base_name = "Ledger entry" if "Ledger entry" in cleaned else present[0]
base = cleaned[base_name].copy()

# Normaliser kolonnenavn som ofte varierer
rename_map = {
    "Project": "Project Name",
    "Project name": "Project Name",
    "Item": "Item Name",
    "Unit": "Unit of measure",
    "UoM": "Unit of measure",
}
base = base.rename(columns=rename_map)

# Sørg for at de viktigste kolonnene finnes (skaper tomme ved behov)
expected_cols = [
    "Date", "Customer", "Project Name", "Source", "Item Name",
    "Quantity", "Unit of measure", "Time Group", "Location"
]
for col in expected_cols:
    if col not in base.columns:
        base[col] = np.nan

# ---------- Sidebar filtere ----------
st.sidebar.header("Filter")
min_date = pd.to_datetime(base["Date"], errors="coerce").min()
max_date = pd.to_datetime(base["Date"], errors="coerce").max()
if pd.notna(min_date) and pd.notna(max_date):
    date_range = st.sidebar.date_input("Dato (fra–til)", value=(min_date, max_date))
else:
    date_range = None

def multiselect_for(col):
    vals = sorted([v for v in base[col].dropna().unique().tolist() if str(v).strip() != ""])
    return st.sidebar.multiselect(col, vals, default=[])

sel_customer = multiselect_for("Customer")
sel_project  = multiselect_for("Project Name")
sel_item     = multiselect_for("Item Name")
sel_source   = multiselect_for("Source")
sel_timegrp  = multiselect_for("Time Group")
sel_location = multiselect_for("Location")

# ---------- Filtrer data ----------
df = base.copy()

if date_range and isinstance(date_range, (list, tuple)) and len(date_range) == 2:
    d1, d2 = date_range
    dcol = pd.to_datetime(df["Date"], errors="coerce")
    df = df[(dcol >= pd.to_datetime(d1)) & (dco_]()
