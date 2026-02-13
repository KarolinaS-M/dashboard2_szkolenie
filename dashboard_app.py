import io
from typing import Tuple, List

import pandas as pd
import streamlit as st
import plotly.express as px


# -----------------------------
# Page configuration
# -----------------------------
st.set_page_config(
    page_title="Excel Dashboard (Streamlit)",
    page_icon="📊",
    layout="wide",
)


# -----------------------------
# Helpers
# -----------------------------
def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize column names to strings (keeps original names, but ensures consistent type).
    """
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> Tuple[pd.ExcelFile, List[str]]:
    """
    Load Excel file and return the ExcelFile object + sheet names.
    Cached by file content (bytes).
    """
    bio = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(bio)
    return xls, xls.sheet_names


@st.cache_data(show_spinner=False)
def read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    """
    Read a chosen sheet into a DataFrame.
    """
    bio = io.BytesIO(file_bytes)
    df = pd.read_excel(bio, sheet_name=sheet_name)
    df = _normalize_columns(df)
    return df


def guess_likert_columns(df: pd.DataFrame) -> List[str]:
    """
    Heuristic: Likert columns often contain "(1–5)" or "(1-5)" in name and are numeric.
    """
    candidates = []
    for c in df.columns:
        name = c.lower()
        if any(x in name for x in ["(1–5)", "(1-5)", "1–5", "1-5"]) and pd.api.types.is_numeric_dtype(df[c]):
            candidates.append(c)
    return candidates


def safe_numeric(df: pd.DataFrame, col: str) -> pd.Series:
    """
    Convert to numeric if possible, otherwise return NaNs for non-convertible rows.
    """
    return pd.to_numeric(df[col], errors="coerce")


def add_sidebar_filters(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build sidebar filters based on detected columns.
    Works well for typical survey datasets (gender, age, tenure, yes/no, etc.).
    """
    filtered = df.copy()

    st.sidebar.header("Filtry")

    # Common columns in PL datasets
    col_gender = "Płeć" if "Płeć" in filtered.columns else None
    col_age = "Wiek" if "Wiek" in filtered.columns else None
    col_tenure = "Staż pracy (lata)" if "Staż pracy (lata)" in filtered.columns else None

    # 1) Categorical multi-select filters (limited to low-cardinality columns)
    cat_cols = [c for c in filtered.columns if filtered[c].dtype == "object"]
    # Limit to columns with not-too-many unique values
    cat_cols = [c for c in cat_cols if filtered[c].nunique(dropna=True) <= 25]

    # Put the most common first if present
    preferred_order = []
    for c in [col_gender, "Czy szkolenie spełniło oczekiwania (tak/nie)"]:
        if c and c in cat_cols:
            preferred_order.append(c)
    for c in cat_cols:
        if c not in preferred_order:
            preferred_order.append(c)

    with st.sidebar.expander("Filtry kategoryczne", expanded=True):
        for c in preferred_order:
            options = sorted([x for x in filtered[c].dropna().unique().tolist()])
            if not options:
                continue
            default = options  # show all by default
            sel = st.multiselect(f"{c}", options=options, default=default)
            if sel and len(sel) < len(options):
                filtered = filtered[filtered[c].isin(sel)]

    # 2) Numeric range filters for age/tenure if available
    with st.sidebar.expander("Filtry liczbowe", expanded=True):
        for label, col in [("Wiek", col_age), ("Staż pracy (lata)", col_tenure)]:
            if col and col in filtered.columns:
                s = safe_numeric(filtered, col)
                if s.notna().any():
                    mn = int(s.min())
                    mx = int(s.max())
                    a, b = st.slider(label, min_value=mn, max_value=mx, value=(mn, mx))
                    filtered = filtered[s.between(a, b)]

    return filtered


def kpi_card(title: str, value, help_text: str | None = None):
    st.metric(label=title, value=value, help=help_text)


# -----------------------------
# UI
# -----------------------------
st.title("📊 Dashboard z pliku Excel (Streamlit)")
st.write(
    "Wgraj plik Excel, a aplikacja zbuduje dashboard: podsumowania, rozkłady ocen, porównania grup oraz korelacje."
)

uploaded = st.file_uploader(
    "Wgraj plik Excel (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False,
)

if uploaded is None:
    st.info("Aby rozpocząć, wgraj plik Excel (.xlsx) powyżej.")
    st.stop()

file_bytes = uploaded.getvalue()

try:
    xls, sheets = load_excel(file_bytes)
except Exception as e:
    st.error("Nie udało się otworzyć pliku Excel. Sprawdź, czy to poprawny plik .xlsx.")
    st.exception(e)
    st.stop()

st.sidebar.header("Arkusz")
sheet = st.sidebar.selectbox("Wybierz arkusz", options=sheets, index=0)

try:
    df = read_sheet(file_bytes, sheet)
except Exception as e:
    st.error("Nie udało się wczytać wybranego arkusza.")
    st.exception(e)
    st.stop()

if df.empty:
    st.warning("Wybrany arkusz jest pusty.")
    st.stop()

# Basic cleanup: keep a copy of original and a trimmed working version
df_raw = df.copy()
df = df.dropna(axis=1, how="all")  # drop fully empty columns

# Sidebar filters
df_f = add_sidebar_filters(df)

# Layout
left, right = st.columns([1.2, 1])

with left:
    st.subheader("Podgląd danych")
    st.caption(f"Arkusz: **{sheet}** | Wiersze (po filtrach): **{len(df_f)}** / {len(df)}")
    st.dataframe(df_f, use_container_width=True, hide_index=True)

    st.download_button(
        "Pobierz dane po filtrach (CSV)",
        data=df_f.to_csv(index=False).encode("utf-8"),
        file_name="filtered_data.csv",
        mime="text/csv",
        use_container_width=True,
    )

with right:
    st.subheader("Szybkie wskaźniki (KPI)")

    likert_cols = guess_likert_columns(df_f)
    numeric_cols = [c for c in df_f.columns if pd.api.types.is_numeric_dtype(df_f[c])]
    # Prefer Likert if detected, else use numeric columns
    rating_cols = likert_cols if likert_cols else numeric_cols

    if len(df_f) == 0:
        st.warning("Brak danych po filtrach.")
    else:
        if rating_cols:
            # Show up to 4 KPIs
            for c in rating_cols[:4]:
                s = safe_numeric(df_f, c)
                avg = float(s.mean()) if s.notna().any() else None
                med = float(s.median()) if s.notna().any() else None
                if avg is None:
                    kpi_card(c, "—")
                else:
                    kpi_card(c, f"{avg:.2f}", help_text=f"Mediana: {med:.2f}")
        else:
            kpi_card("Liczba obserwacji", f"{len(df_f)}")

    st.divider()

    st.subheader("Struktura odpowiedzi")
    # Show a pie for a likely column if present
    candidate_cat = None
    for c in ["Płeć", "Czy szkolenie spełniło oczekiwania (tak/nie)"]:
        if c in df_f.columns and df_f[c].dtype == "object":
            candidate_cat = c
            break

    if candidate_cat:
        vc = df_f[candidate_cat].value_counts(dropna=False).reset_index()
        vc.columns = [candidate_cat, "Liczba"]
        fig = px.pie(vc, names=candidate_cat, values="Liczba", title=candidate_cat)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption("Brak oczywistej kolumny kategorycznej do szybkiego wykresu kołowego.")


st.divider()
st.header("Analizy")

tab1, tab2, tab3 = st.tabs(["Rozkłady ocen", "Porównania grup", "Korelacje"])

with tab1:
    st.subheader("Rozkłady ocen (histogramy)")
    if not rating_cols:
        st.info("Nie wykryto kolumn liczbowych do analizy rozkładów.")
    else:
        col = st.selectbox("Wybierz zmienną liczbową", options=rating_cols, index=0)
        s = safe_numeric(df_f, col).dropna()
        if s.empty:
            st.warning("Brak danych (po filtrach) dla tej zmiennej.")
        else:
            fig = px.histogram(s, nbins=min(20, max(5, int(s.nunique()))), title=f"Histogram: {col}")
            st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("Porównania grup (boxplot)")
    # Choose y as numeric and x as categorical
    cat_cols = [c for c in df_f.columns if df_f[c].dtype == "object" and df_f[c].nunique(dropna=True) <= 25]
    if not rating_cols or not cat_cols:
        st.info("Aby zrobić porównania grup, potrzebujesz co najmniej jednej kolumny liczbowej i jednej kategorycznej.")
    else:
        y = st.selectbox("Zmienna liczbowa (Y)", options=rating_cols, index=0)
        x = st.selectbox("Zmienna grupująca (X)", options=cat_cols, index=0)

        tmp = df_f[[x, y]].copy()
        tmp[y] = safe_numeric(tmp, y)
        tmp = tmp.dropna(subset=[x, y])

        if tmp.empty:
            st.warning("Brak danych do wykresu po zastosowaniu filtrów.")
        else:
            fig = px.box(tmp, x=x, y=y, points="all", title=f"{y} wg {x}")
            st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("Korelacje (heatmapa)")
    # Correlation among numeric columns
    num_cols = [c for c in df_f.columns if pd.api.types.is_numeric_dtype(df_f[c])]
    if len(num_cols) < 2:
        st.info("Za mało kolumn liczbowych, aby policzyć korelacje.")
    else:
        tmp = df_f[num_cols].copy()
        corr_method = st.radio("Metoda", options=["pearson", "spearman"], horizontal=True, index=0)
        corr = tmp.corr(method=corr_method)

        fig = px.imshow(
            corr,
            text_auto=True,
            aspect="auto",
            title=f"Macierz korelacji ({corr_method})",
        )
        st.plotly_chart(fig, use_container_width=True)


st.divider()
st.caption(
    "Wskazówka: jeśli Twoje dane mają inne nazwy kolumn niż typowe ankietowe (np. Wiek/Płeć), "
    "dashboard i tak zadziała, tylko filtry „specjalne” pojawią się dla wykrytych kolumn."
)