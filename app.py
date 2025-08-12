
import io
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="ILMS Ledger Viewer", layout="wide")

st.title("ILMS Ledger Viewer")
st.caption("Last opp Excel-filen (ILMS data TNG July 2025.xlsx) og filtrer/pivoter ledger-data.")

@st.cache_data
def load_excel(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    return {name: xls.parse(name) for name in xls.sheet_names}

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(axis=0, how='all').dropna(axis=1, how='all').copy()
    if df.empty:
        return df
    col_names = [str(c) for c in df.columns]
    if len(col_names) and sum(c.startswith("Unnamed") for c in col_names)/len(col_names) >= 0.5:
        nonnull_counts = df.notna().sum(axis=1)
        header_idx = int(nonnull_counts.idxmax())
        new_header = df.loc[header_idx].astype(str).tolist()
        df = df.loc[header_idx+1:].copy()
        df.columns = new_header
    df = df.replace({"#N/A": np.nan, "N/A": np.nan, "NA": np.nan, "": np.nan})
    for c in df.columns:
        s = df[c]
        if s.dtype == "O":
            dt = pd.to_datetime(s, errors="coerce", dayfirst=True, infer_datetime_format=True)
            if dt.notna().mean() > 0.7:
                df[c] = dt.dt.date
            else:
                df[c] = pd.to_numeric(s, errors="ignore")
    return df

def unify_column_names(df):
    df = df.rename(columns=lambda x: str(x).strip())
    return df

def safe_numeric(s):
    return pd.to_numeric(s, errors="coerce")

uploaded = st.file_uploader("Dra hit Excel-filen din (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("👆 Last opp Excel-filen for å komme i gang.")
    st.stop()

with st.spinner("Leser Excel ..."):
    sheets = load_excel(uploaded)

wanted = ["Ledger entry", "ILMS Ton", "ILMS Water", "ILMS Mooring", "Mapping"]
present = [s for s in wanted if s in sheets]

if not present:
    st.error("Fant ingen av arkene: " + ", ".join(wanted))
    st.stop()

cleaned = {s: unify_column_names(clean_df(sheets[s])) for s in present}

with st.expander("Forhåndsvis ark (før filtrering)"):
    for s in present:
        st.markdown(f"**{s}**  \n{cleaned[s].shape[0]} rader, {cleaned[s].shape[1]} kolonner")
        st.dataframe(cleaned[s].head(10))

base_name = "Ledger entry" if "Ledger entry" in cleaned else present[0]
base = cleaned[base_name].copy()

rename_map = {
    "Project": "Project Name",
    "Project name": "Project Name",
    "Item": "Item Name",
    "Unit": "Unit of measure",
    "UoM": "Unit of measure",
}
base = base.rename(columns=rename_map)

for col in ["Date","Customer","Project Name","Source","Item Name","Quantity","Time Group","Location","Unit of measure"]:
    if col not in base.columns:
        base[col] = np.nan

st.sidebar.header("Filter")
min_date = pd.to_datetime(base["Date"], errors="coerce").min()
max_date = pd.to_datetime(base["Date"], errors="coerce").max()
date_range = st.sidebar.date_input("Dato (fra–til)", value=(min_date, max_date)) if pd.notna(min_date) and pd.notna(max_date) else None

def pick(col):
    vals = sorted([v for v in base[col].dropna().unique().tolist() if v != "" ])
    return st.sidebar.multiselect(col, vals)

sel_customer = pick("Customer")
sel_project  = pick("Project Name")
sel_item     = pick("Item Name")
sel_source   = pick("Source")
sel_timegrp  = pick("Time Group")
sel_location = pick("Location")

df = base.copy()
if date_range and isinstance(date_range, (list, tuple)) and len(date_range)==2:
    d1, d2 = date_range
    dcol = pd.to_datetime(df["Date"], errors="coerce")
    df = df[(dcol >= pd.to_datetime(d1)) & (dcol <= pd.to_datetime(d2))]
for col, sel in [
    ("Customer", sel_customer),
    ("Project Name", sel_project),
    ("Item Name", sel_item),
    ("Source", sel_source),
    ("Time Group", sel_timegrp),
    ("Location", sel_location),
]:
    if sel:
        df = df[df[col].isin(sel)]

df["Month"] = pd.to_datetime(df["Date"], errors="coerce").dt.to_period("M").astype(str)
df["Quantity_num"] = safe_numeric(df["Quantity"])

st.subheader("Ledger-tabell")
st.dataframe(df[["Date","Customer","Project Name","Source","Item Name","Quantity","Unit of measure","Time Group","Location"]].reset_index(drop=True))

st.subheader("Pivot: Sum Quantity per måned × Time Group")
pivot = pd.pivot_table(
    df, index="Month", columns="Time Group", values="Quantity_num", aggfunc="sum", fill_value=0
).sort_index()
st.dataframe(pivot)

st.subheader("Trend: Quantity per måned")
trend = df.groupby("Month", as_index=False)["Quantity_num"].sum().sort_values("Month")
st.line_chart(trend.set_index("Month"))

def to_excel_bytes(dataframe_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, d in dataframe_dict.items():
            d.to_excel(writer, sheet_name=name[:31], index=False)
    buf.seek(0)
    return buf

col1, col2 = st.columns(2)
with col1:
    st.download_button("Last ned filtrert tabell (CSV)", data=df.to_csv(index=False).encode("utf-8"), file_name="ledger_filtered.csv", mime="text/csv")
with col2:
    xbuf = to_excel_bytes({"ledger_filtered": df, "pivot": pivot.reset_index()})
    st.download_button("Last ned Excel (tabell + pivot)", data=xbuf, file_name="ledger_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.success("Ferdig! Bruk filtrene i venstre side for å se ulike utsnitt.")
