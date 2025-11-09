import io
from typing import List, Tuple

import pandas as pd
import streamlit as st

# -------------------------
# Config (fixed per your rules)
# -------------------------
FINAL_COLS = [
    "Row No.",
    "Store Code",
    "Store Name",
    "Date",
    "Gross Sales",
    "Net Sales",
    "Discounted",
    "Item Void",
    "Void Value",
    "Item Refund",
    "Refund Value",
    "Terminal",
    "Unnamed",
    "Quantity",
    "Transaction",
    "Average Transaction Value",
]

NUMERIC_COLS = {
    "Gross Sales",
    "Net Sales",
    "Discounted",
    "Item Void",
    "Void Value",
    "Item Refund",
    "Refund Value",
    "Quantity",
    "Transaction",
    "Average Transaction Value",
}
NON_NUMERIC_COLS = set(FINAL_COLS) - NUMERIC_COLS  # (1,2,3,4,12,13)

HEADER_ROWS = [3, 4]  # Excel is 1-indexed rows 4 & 5 => pandas header=[3,4]

# -------------------------
# Helpers
# -------------------------
def _flatten_columns(cols) -> List[str]:
    """
    Flattens a MultiIndex header into a single row of strings.
    We will immediately overwrite names with FINAL_COLS by position,
    but flattening helps robust reading.
    """
    flat = []
    for c in cols:
        if isinstance(c, tuple):
            parts = [str(x).strip() for x in c if x is not None and str(x).strip() != ""]
            flat.append(" | ".join(parts) if parts else "")
        else:
            s = "" if c is None else str(c).strip()
            flat.append(s)
    return flat

def _coerce_numeric(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace(",", "", regex=False), errors="coerce"
            )
    return df

def _parse_date(series: pd.Series) -> pd.Series:
    # dayfirst=True per your Vietnam context
    return pd.to_datetime(series, errors="coerce", dayfirst=True)

def _rename_by_position(df: pd.DataFrame, final_cols: List[str]) -> pd.DataFrame:
    # Keep first 16 columns; pad if fewer; drop extras if more
    n = len(final_cols)
    df = df.iloc[:, :n].copy()
    df.columns = final_cols
    return df

def _store_name_last4(series: pd.Series) -> pd.Series:
    return series.astype(str).str[-4:]

def _filter_gross_sales_numeric_positive(df: pd.DataFrame) -> pd.DataFrame:
    # Enforce numeric
    df["Gross Sales"] = pd.to_numeric(
        df["Gross Sales"].astype(str).str.replace(",", "", regex=False),
        errors="coerce",
    )
    return df[df["Gross Sales"].notna() & (df["Gross Sales"] > 0)]

# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Template → Raw (Fixed Rules)", page_icon="📊", layout="wide")
st.title("📊 Template → Raw Data Normalizer (Fixed Rules)")

uploaded = st.file_uploader("Upload Excel (.xlsx or .xls)", type=["xlsx", "xls"])
if not uploaded:
    st.info("Upload an Excel file to begin.")
    st.stop()

# Choose sheet
try:
    # NOTE: do NOT use a variable named 'xl' to avoid the earlier NameError confusion
    xls = pd.ExcelFile(uploaded)
    sheets = xls.sheet_names
except Exception as e:
    st.error(f"Failed to open Excel: {e}")
    st.stop()

sheet = st.selectbox("Select sheet", sheets, index=0)

# Read with two header rows (rows 4 & 5)
try:
    df_raw = pd.read_excel(
        xls, sheet_name=sheet, header=HEADER_ROWS, dtype=object
    )
except Exception as e:
    st.error(f"Failed to read sheet with header rows {HEADER_ROWS}: {e}")
    st.stop()

# Flatten header (for robustness), then immediately rename by position
df_raw.columns = _flatten_columns(df_raw.columns)
df = _rename_by_position(df_raw, FINAL_COLS)

# Trim/clean strings in all columns
for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})

# Drop fully empty rows
df = df.dropna(how="all").reset_index(drop=True)

# Parse Date (Col 4)
df["Date"] = _parse_date(df["Date"])

# Column 3: keep last 4 characters as Store Name
df["Store Name"] = _store_name_last4(df["Store Name"])

# Coerce numeric for required columns (all except 1,2,3,4,12,13)
df = _coerce_numeric(df, [c for c in FINAL_COLS if c in NUMERIC_COLS])

# Filter: keep only rows with Gross Sales numeric and > 0
df = _filter_gross_sales_numeric_positive(df)

# Final clean once more
df = df.dropna(how="all").reset_index(drop=True)

st.subheader("Output preview")
st.dataframe(df.head(200))

# Downloads
csv_bytes = df.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download CSV", data=csv_bytes, file_name="normalized.csv", mime="text/csv")

xlsx_buf = io.BytesIO()
with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="normalized")
st.download_button("⬇️ Download Excel", data=xlsx_buf.getvalue(),
                   file_name="normalized.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
