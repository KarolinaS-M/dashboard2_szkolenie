import io
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px


# ----------------------------
# App configuration
# ----------------------------
st.set_page_config(
    page_title="Excel Dashboard (Streamlit)",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Excel Dashboard")
st.caption("Upload an Excel file, explore sheets, filter data, and visualize columns.")


# ----------------------------
# Helpers
# ----------------------------
@dataclass
class LoadResult:
    sheets: Dict[str, pd.DataFrame]
    warnings: List[str]


def _safe_read_excel(file_bytes: bytes) -> LoadResult:
    """
    Read an Excel file from bytes and return all sheets as DataFrames.
    Uses openpyxl under the hood (via pandas).
    """
    warnings: List[str] = []
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets: Dict[str, pd.DataFrame] = {}

    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            # Normalize column names a bit (optional)
            df.columns = [str(c).strip() for c in df.columns]
            sheets[name] = df
        except Exception as e:
            warnings.append(f"Could not read sheet '{name}': {e}")

    if not sheets:
        warnings.append("No readable sheets were found in the Excel file.")

    return LoadResult(sheets=sheets, warnings=warnings)


@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes) -> LoadResult:
    # Caching by file content keeps the app responsive on repeated interactions.
    return _safe_read_excel(file_bytes)


def dataframe_overview(df: pd.DataFrame) -> pd.DataFrame:
    """
    Create a compact overview table with dtypes, missingness, and basic stats.
    """
    overview = pd.DataFrame({
        "dtype": df.dtypes.astype(str),
        "missing_count": df.isna().sum(),
        "missing_pct": (df.isna().mean() * 100).round(2),
        "n_unique": df.nunique(dropna=True),
    })

    # Add numeric summaries where applicable
    num_cols = df.select_dtypes(include="number").columns
    if len(num_cols) > 0:
        overview.loc[num_cols, "min"] = df[num_cols].min()
        overview.loc[num_cols, "max"] = df[num_cols].max()
        overview.loc[num_cols, "mean"] = df[num_cols].mean()

    return overview.reset_index().rename(columns={"index": "column"})


def choose_filter_columns(df: pd.DataFrame) -> Tuple[List[str], List[str], List[str]]:
    """
    Classify columns into numeric, datetime-like, and categorical-ish.
    """
    numeric_cols = df.select_dtypes(include="number").columns.tolist()

    # Try to detect datetime-like columns (already datetime or parseable)
    datetime_cols: List[str] = []
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            datetime_cols.append(col)
        else:
            # Try parsing a sample safely
            if df[col].dtype == object:
                sample = df[col].dropna().astype(str).head(50)
                if len(sample) > 0:
                    parsed = pd.to_datetime(sample, errors="coerce", utc=False)
                    if parsed.notna().mean() > 0.8:
                        datetime_cols.append(col)

    # Categorical-ish: object/bool/category + small-ish cardinality
    cat_candidates = df.select_dtypes(include=["object", "bool", "category"]).columns.tolist()
    categorical_cols = []
    for col in cat_candidates:
        nun = df[col].nunique(dropna=True)
        if nun <= 200:
            categorical_cols.append(col)

    # Remove datetime columns from categorical list if overlap
    categorical_cols = [c for c in categorical_cols if c not in datetime_cols]

    return numeric_cols, datetime_cols, categorical_cols


def apply_filters(
    df: pd.DataFrame,
    numeric_cols: List[str],
    categorical_cols: List[str],
    datetime_cols: List[str],
) -> pd.DataFrame:
    """
    Apply sidebar filters to the DataFrame.
    """
    out = df.copy()

    st.sidebar.subheader("Filters")

    # Categorical filters
    if categorical_cols:
        with st.sidebar.expander("Categorical", expanded=True):
            for col in categorical_cols:
                values = out[col].dropna().unique().tolist()
                if len(values) == 0:
                    continue
                # Limit options shown if very large
                values_sorted = sorted(values, key=lambda x: str(x))[:500]
                selected = st.multiselect(f"{col}", options=values_sorted, default=None)
                if selected:
                    out = out[out[col].isin(selected)]

    # Numeric range filters
    if numeric_cols:
        with st.sidebar.expander("Numeric", expanded=False):
            for col in numeric_cols:
                series = out[col].dropna()
                if series.empty:
                    continue
                min_v = float(series.min())
                max_v = float(series.max())
                if min_v == max_v:
                    continue
                chosen = st.slider(
                    f"{col} range",
                    min_value=min_v,
                    max_value=max_v,
                    value=(min_v, max_v),
                )
                out = out[out[col].between(chosen[0], chosen[1], inclusive="both")]

    # Datetime range filters (best-effort)
    if datetime_cols:
        with st.sidebar.expander("Date/Time", expanded=False):
            for col in datetime_cols:
                col_series = out[col]
                if not pd.api.types.is_datetime64_any_dtype(col_series):
                    # attempt parse for filtering
                    parsed = pd.to_datetime(col_series, errors="coerce", utc=False)
                else:
                    parsed = col_series

                parsed_nonnull = parsed.dropna()
                if parsed_nonnull.empty:
                    continue

                dmin = parsed_nonnull.min().to_pydatetime()
                dmax = parsed_nonnull.max().to_pydatetime()

                start, end = st.date_input(
                    f"{col} date range",
                    value=(dmin.date(), dmax.date()),
                )
                # Re-apply mask
                mask = parsed.between(pd.to_datetime(start), pd.to_datetime(end) + pd.Timedelta(days=1), inclusive="left")
                out = out[mask.fillna(False)]

    return out


def plot_section(df: pd.DataFrame) -> None:
    """
    Simple charting section for numeric columns.
    """
    st.subheader("Visualizations")

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if not numeric_cols:
        st.info("No numeric columns found to plot.")
        return

    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        x_col = st.selectbox("X axis", options=["(index)"] + df.columns.tolist(), index=0)
    with c2:
        y_col = st.selectbox("Y axis (numeric)", options=numeric_cols, index=0)
    with c3:
        chart_type = st.selectbox("Chart type", options=["Line", "Scatter", "Histogram", "Box"], index=1)

    plot_df = df.copy()
    if x_col == "(index)":
        plot_df = plot_df.reset_index(drop=False).rename(columns={"index": "index"})
        x_col_use = "index"
    else:
        x_col_use = x_col

    try:
        if chart_type == "Line":
            fig = px.line(plot_df, x=x_col_use, y=y_col)
        elif chart_type == "Scatter":
            fig = px.scatter(plot_df, x=x_col_use, y=y_col)
        elif chart_type == "Histogram":
            fig = px.histogram(plot_df, x=y_col)
        else:
            fig = px.box(plot_df, x=x_col_use if x_col != "(index)" else None, y=y_col)

        st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Plot error: {e}")


# ----------------------------
# UI: Upload
# ----------------------------
with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader("Upload an Excel file (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Please upload an Excel file to begin.")
    st.stop()

file_bytes = uploaded.getvalue()
result = load_excel_cached(file_bytes)

for w in result.warnings:
    st.warning(w)

if not result.sheets:
    st.stop()

# ----------------------------
# UI: Sheet selection
# ----------------------------
sheet_names = list(result.sheets.keys())
selected_sheet = st.selectbox("Select sheet", options=sheet_names, index=0)
df_raw = result.sheets[selected_sheet]

st.subheader("Dataset snapshot")
cA, cB, cC, cD = st.columns(4)
cA.metric("Rows", f"{len(df_raw):,}")
cB.metric("Columns", f"{df_raw.shape[1]:,}")
cC.metric("Missing cells", f"{int(df_raw.isna().sum().sum()):,}")
cD.metric("Memory (approx.)", f"{df_raw.memory_usage(deep=True).sum() / (1024**2):.2f} MB")

# ----------------------------
# Filters + main table
# ----------------------------
numeric_cols, datetime_cols, categorical_cols = choose_filter_columns(df_raw)
df = apply_filters(df_raw, numeric_cols, categorical_cols, datetime_cols)

tab1, tab2, tab3 = st.tabs(["Data", "Overview", "Quality"])

with tab1:
    st.write("Filtered data (preview):")
    st.dataframe(df, use_container_width=True, height=420)

    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download filtered CSV",
        data=csv,
        file_name=f"{selected_sheet}_filtered.csv",
        mime="text/csv",
    )

with tab2:
    st.write("Column overview:")
    st.dataframe(dataframe_overview(df_raw), use_container_width=True, height=420)

with tab3:
    st.write("Missingness by column:")
    missing = df_raw.isna().mean().sort_values(ascending=False).reset_index()
    missing.columns = ["column", "missing_rate"]
    st.dataframe(missing, use_container_width=True, height=420)

    # Simple bar chart for top missing columns
    top_missing = missing.head(30).copy()
    top_missing["missing_rate_pct"] = (top_missing["missing_rate"] * 100).round(2)
    try:
        fig = px.bar(top_missing, x="column", y="missing_rate_pct")
        st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"Plot error: {e}")

st.divider()
plot_section(df)