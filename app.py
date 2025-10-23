import io
from typing import List, Tuple

import pandas as pd
import streamlit as st


# ---------- Page setup ----------
st.set_page_config(page_title="CSV ↔ Excel Comparator", page_icon="🧪", layout="wide")
st.title("CSV ↔ Excel Comparator")
st.caption("Upload one CSV and one Excel file. The app checks if **all rows** in the CSV equal **all rows** in the Excel (including duplicates).")

with st.expander("How it works", expanded=False):
    st.markdown(
        """
        - Column headers are expected to be the same in both files (order can differ).
        - The comparison is **order-independent** and **counts duplicates** correctly.
        - You can optionally trim whitespace in text fields.
        - Results include detailed diffs:
            - **Only in CSV** (and how many times a row appears more than in Excel)
            - **Only in Excel** (and how many times a row appears more than in CSV)
        """
    )

# ---------- Sidebar options ----------
st.sidebar.header("Options")
trim_ws = st.sidebar.checkbox("Trim whitespace in text cells", value=True)
lowercase_colnames = st.sidebar.checkbox("Lowercase column names", value=False)
show_preview_rows = st.sidebar.number_input("Preview rows to display", min_value=5, max_value=200, value=20, step=5)

st.sidebar.divider()
st.sidebar.caption("Reading options (advanced)")
csv_sep = st.sidebar.text_input("CSV delimiter (leave empty for auto)", value="")
excel_sheet_name: str | None = None

# ---------- File uploaders ----------
left, right = st.columns(2, vertical_alignment="top")
with left:
    csv_file = st.file_uploader("Upload CSV", type=["csv"], accept_multiple_files=False)
with right:
    xls_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"], accept_multiple_files=False)

def _normalize_colnames(cols: List[str]) -> List[str]:
    cols2 = [c.strip() for c in cols]
    if lowercase_colnames:
        cols2 = [c.lower() for c in cols2]
    return cols2

def _trim_ws_df(df: pd.DataFrame) -> pd.DataFrame:
    if not trim_ws:
        return df
    df = df.copy()
    # Only trim strings to avoid impacting numerics/dates
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]):
            df[c] = df[c].str.strip()
    return df

def _read_csv(file) -> pd.DataFrame:
    # Try to sniff delimiter if not provided
    if csv_sep.strip():
        return pd.read_csv(file, dtype_backend="pyarrow", sep=csv_sep)
    # Auto: try pandas default, if that fails try semicolon
    try:
        return pd.read_csv(file, dtype_backend="pyarrow")
    except Exception:
        file.seek(0)
        return pd.read_csv(file, dtype_backend="pyarrow", sep=";")

def _read_excel(file, sheet: str | int | None) -> Tuple[pd.DataFrame, List[str]]:
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    chosen = sheet if sheet in sheets else sheets[0]
    df = pd.read_excel(xls, sheet_name=chosen, dtype_backend="pyarrow")
    return df, sheets

def _align_and_normalize(csv_df: pd.DataFrame, xls_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, List[str], List[str], List[str]]:
    # Normalize column names
    csv_df = csv_df.copy()
    xls_df = xls_df.copy()
    csv_df.columns = _normalize_colnames(list(csv_df.columns))
    xls_df.columns = _normalize_colnames(list(xls_df.columns))

    # Trim whitespace in text columns if option is set
    csv_df = _trim_ws_df(csv_df)
    xls_df = _trim_ws_df(xls_df)

    csv_cols = list(csv_df.columns)
    xls_cols = list(xls_df.columns)

    csv_set = set(csv_cols)
    xls_set = set(xls_cols)

    only_in_csv = sorted(list(csv_set - xls_set))
    only_in_xls = sorted(list(xls_set - csv_set))

    common_cols = sorted(list(csv_set & xls_set))

    # Restrict to common columns for row comparison
    csv_df = csv_df[common_cols]
    xls_df = xls_df[common_cols]

    return csv_df, xls_df, common_cols, only_in_csv, only_in_xls

def _multiset_counts(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """
    Returns a DataFrame keyed by all columns with a 'count' column
    representing how many times each row appears.
    We keep original dtypes, but fill NA consistently so groupby works.
    """
    # For grouping purposes, Pandas treats NaNs as separate values per row unless filled.
    g = (
        df[cols]
        .astype(object)  # robust grouping across mixed dtypes
        .where(pd.notna(df[cols]), None)  # make all missing values a shared sentinel
        .groupby(cols, dropna=False, as_index=False, sort=False)
        .size()
        .rename(columns={"size": "count"})
    )
    return g

def _compare(csv_df: pd.DataFrame, xls_df: pd.DataFrame, cols: List[str]) -> Tuple[bool, pd.DataFrame, pd.DataFrame]:
    """
    Compare the two datasets as multisets of rows.
    Returns:
      - equal (bool)
      - only_in_csv (DataFrame): rows where CSV has a higher count than Excel, with a 'delta' column
      - only_in_xls (DataFrame): rows where Excel has a higher count than CSV, with a 'delta' column
    """
    csv_counts = _multiset_counts(csv_df, cols).rename(columns={"count": "count_csv"})
    xls_counts = _multiset_counts(xls_df, cols).rename(columns={"count": "count_excel"})

    merged = csv_counts.merge(xls_counts, on=cols, how="outer")
    merged["count_csv"] = merged["count_csv"].fillna(0).astype(int)
    merged["count_excel"] = merged["count_excel"].fillna(0).astype(int)
    merged["delta"] = merged["count_csv"] - merged["count_excel"]

    only_in_csv = merged[merged["delta"] > 0].copy()
    only_in_xls = merged[merged["delta"] < 0].copy()
    # Normalize deltas to positive numbers for readability
    only_in_csv["delta"] = only_in_csv["delta"].abs()
    only_in_xls["delta"] = only_in_xls["delta"].abs()

    equal = only_in_csv.empty and only_in_xls.empty
    return equal, only_in_csv, only_in_xls

# ---------- Main flow ----------
if csv_file and xls_file:
    # Excel sheet picker (after sniffing sheet names)
    # We need a fresh handle for reading multiple times
    xls_bytes = io.BytesIO(xls_file.getvalue())
    df_xls_preview, sheet_names = _read_excel(xls_bytes, None)
    excel_sheet_name = st.selectbox("Choose Excel sheet", options=sheet_names, index=0, help="Select which sheet from the Excel file to compare.")
    xls_bytes.seek(0)

    # Read full data
    try:
        df_csv = _read_csv(csv_file)
    except Exception as e:
        st.error(f"Failed to read CSV: {e}")
        st.stop()

    try:
        df_xls, _ = _read_excel(xls_bytes, excel_sheet_name)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        st.stop()

    # Align columns and normalize
    df_csv, df_xls, common_cols, only_in_csv_cols, only_in_xls_cols = _align_and_normalize(df_csv, df_xls)

    # Headers check
    if only_in_csv_cols or only_in_xls_cols:
        col1, col2 = st.columns(2)
        if only_in_csv_cols:
            with col1:
                st.warning("Columns only in CSV")
                st.code(", ".join(only_in_csv_cols))
        if only_in_xls_cols:
            with col2:
                st.warning("Columns only in Excel")
                st.code(", ".join(only_in_xls_cols))
        st.info("Only common columns will be compared.")

    st.subheader("Column set used for comparison")
    st.code(", ".join(common_cols) if common_cols else "(no common columns)")
    if not common_cols:
        st.error("No common columns to compare. Please upload files with matching headers.")
        st.stop()

    # Quick previews
    prev_left, prev_right = st.columns(2)
    with prev_left:
        st.markdown("**CSV preview**")
        st.dataframe(df_csv.head(show_preview_rows), use_container_width=True)
    with prev_right:
        st.markdown("**Excel preview**")
        st.dataframe(df_xls.head(show_preview_rows), use_container_width=True)

    # Compare
    equal, only_in_csv_rows, only_in_xls_rows = _compare(df_csv, df_xls, common_cols)

    st.divider()
    if equal:
        st.success("✅ The CSV and Excel contain exactly the same rows (including duplicate counts) over the common columns.")
    else:
        st.error("❌ The datasets differ.")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Rows only in CSV (or appearing more times in CSV)**")
            if only_in_csv_rows.empty:
                st.write("None 🎉")
            else:
                st.dataframe(only_in_csv_rows.head(show_preview_rows), use_container_width=True)
                csv_buf = io.StringIO()
                only_in_csv_rows.to_csv(csv_buf, index=False)
                st.download_button("Download full CSV-only rows", csv_buf.getvalue(), file_name="only_in_csv.csv", mime="text/csv")

        with c2:
            st.markdown("**Rows only in Excel (or appearing more times in Excel)**")
            if only_in_xls_rows.empty:
                st.write("None 🎉")
            else:
                st.dataframe(only_in_xls_rows.head(show_preview_rows), use_container_width=True)
                xls_buf = io.StringIO()
                only_in_xls_rows.to_csv(xls_buf, index=False)
                st.download_button("Download full Excel-only rows", xls_buf.getvalue(), file_name="only_in_excel.csv", mime="text/csv")

else:
    st.info("Upload a CSV and an Excel file to begin.")
