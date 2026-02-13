import io
from dataclasses import dataclass
from typing import Optional, Tuple, List

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt


# -----------------------------
# App config
# -----------------------------
st.set_page_config(
    page_title="Dashboard ankiety szkoleniowej",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Dashboard ankiety szkoleniowej (Excel → Streamlit)")
st.caption(
    "Wgraj plik Excel (.xlsx). Aplikacja wczyta dane, pozwoli filtrować odpowiedzi i pokaże podsumowania oraz wykresy."
)


# -----------------------------
# Helpers
# -----------------------------
def _normalize_column_name(col: str) -> str:
    """Normalize column names to help matching even if there are minor differences."""
    if col is None:
        return ""
    col = str(col).strip().lower()
    col = col.replace("\n", " ").replace("\r", " ")
    col = " ".join(col.split())
    return col


def _safe_to_numeric(series: pd.Series) -> pd.Series:
    """Convert series to numeric where possible, keeping NaN for non-convertible values."""
    return pd.to_numeric(series, errors="coerce")


def _safe_to_bool_from_pl(series: pd.Series) -> pd.Series:
    """Convert Polish yes/no strings to boolean: 'tak' -> True, 'nie' -> False."""
    s = series.astype(str).str.strip().str.lower()
    s = s.replace({"tak": True, "nie": False, "true": True, "false": False, "1": True, "0": False})
    # For other values keep NaN
    return s.where(s.isin([True, False]), np.nan)


def _download_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def _plot_hist(series: pd.Series, title: str, bins: int = 10):
    fig, ax = plt.subplots()
    data = series.dropna()
    ax.hist(data, bins=bins)
    ax.set_title(title)
    ax.set_xlabel(series.name if series.name else "")
    ax.set_ylabel("Liczba odpowiedzi")
    st.pyplot(fig, clear_figure=True)


def _plot_bar_counts(series: pd.Series, title: str, top_n: int = 12):
    fig, ax = plt.subplots()
    vc = series.dropna().astype(str).str.strip()
    vc = vc[vc != ""].value_counts().head(top_n)
    ax.bar(vc.index, vc.values)
    ax.set_title(title)
    ax.set_xlabel("")
    ax.set_ylabel("Liczba wskazań")
    ax.tick_params(axis="x", labelrotation=45)
    st.pyplot(fig, clear_figure=True)


def _plot_box_by_category(values: pd.Series, category: pd.Series, title: str):
    dfp = pd.DataFrame({"value": values, "cat": category}).dropna()
    if dfp.empty:
        st.info("Brak danych do wykresu pudełkowego po filtrach.")
        return
    cats = dfp["cat"].astype(str)
    groups = [dfp.loc[cats == c, "value"].values for c in sorted(cats.unique())]

    fig, ax = plt.subplots()
    ax.boxplot(groups, tick_labels=sorted(cats.unique()))
    ax.set_title(title)
    ax.set_xlabel(category.name if category.name else "Kategoria")
    ax.set_ylabel(values.name if values.name else "Wartość")
    st.pyplot(fig, clear_figure=True)


def _plot_scatter(x: pd.Series, y: pd.Series, title: str):
    dfp = pd.DataFrame({"x": x, "y": y}).dropna()
    if dfp.empty:
        st.info("Brak danych do wykresu rozrzutu po filtrach.")
        return
    fig, ax = plt.subplots()
    ax.scatter(dfp["x"].values, dfp["y"].values)
    ax.set_title(title)
    ax.set_xlabel(x.name if x.name else "X")
    ax.set_ylabel(y.name if y.name else "Y")
    st.pyplot(fig, clear_figure=True)


@dataclass
class ColumnMap:
    id_col: Optional[str]
    gender_col: Optional[str]
    age_col: Optional[str]
    tenure_col: Optional[str]
    org_rating_col: Optional[str]
    trainer_rating_col: Optional[str]
    expectations_col: Optional[str]
    best_element_col: Optional[str]
    improve_col: Optional[str]
    overall_sat_col: Optional[str]


def _auto_map_columns(columns: List[str]) -> ColumnMap:
    """Try to auto-detect expected columns based on normalized names."""
    norm = {c: _normalize_column_name(c) for c in columns}

    def find_contains(*needles: str) -> Optional[str]:
        for c, n in norm.items():
            if all(needle in n for needle in needles):
                return c
        return None

    return ColumnMap(
        id_col=find_contains("id"),
        gender_col=find_contains("płeć") or find_contains("plec"),
        age_col=find_contains("wiek"),
        tenure_col=find_contains("staż") or find_contains("staz") or find_contains("staż pracy") or find_contains("staz pracy"),
        org_rating_col=find_contains("ocena organizacji"),
        trainer_rating_col=find_contains("ocena prowadzącego") or find_contains("ocena prowadzacego"),
        expectations_col=find_contains("spełniło oczekiwania") or find_contains("spelnilo oczekiwania") or find_contains("oczekiwania"),
        best_element_col=find_contains("najbardziej wartościowy") or find_contains("najbardziej wartosciowy"),
        improve_col=find_contains("co można poprawić") or find_contains("co mozna poprawic") or find_contains("poprawić") or find_contains("poprawic"),
        overall_sat_col=find_contains("ogólna satysfakcja") or find_contains("ogolna satysfakcja") or find_contains("satysfakcja"),
    )


@st.cache_data(show_spinner=False)
def load_excel(uploaded_bytes: bytes) -> Tuple[pd.ExcelFile, List[str]]:
    """Load Excel bytes and return ExcelFile + sheet names."""
    bio = io.BytesIO(uploaded_bytes)
    xls = pd.ExcelFile(bio)
    return xls, xls.sheet_names


@st.cache_data(show_spinner=False)
def read_sheet(uploaded_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    """Read a specific sheet into a DataFrame."""
    bio = io.BytesIO(uploaded_bytes)
    df = pd.read_excel(bio, sheet_name=sheet_name)
    # Ensure columns are strings
    df.columns = [str(c) for c in df.columns]
    return df


def _apply_filters(
    df: pd.DataFrame,
    cmap: ColumnMap,
    gender_sel: Optional[List[str]],
    age_range: Optional[Tuple[float, float]],
    tenure_range: Optional[Tuple[float, float]],
    exp_sel: Optional[List[str]],
) -> pd.DataFrame:
    out = df.copy()

    # Gender filter
    if cmap.gender_col and gender_sel:
        out = out[out[cmap.gender_col].astype(str).isin(gender_sel)]

    # Age filter
    if cmap.age_col and age_range is not None:
        age_num = _safe_to_numeric(out[cmap.age_col])
        out = out[(age_num >= age_range[0]) & (age_num <= age_range[1])]

    # Tenure filter
    if cmap.tenure_col and tenure_range is not None:
        ten_num = _safe_to_numeric(out[cmap.tenure_col])
        out = out[(ten_num >= tenure_range[0]) & (ten_num <= tenure_range[1])]

    # Expectations filter (string-based to keep robust)
    if cmap.expectations_col and exp_sel:
        out = out[out[cmap.expectations_col].astype(str).str.strip().str.lower().isin([s.lower() for s in exp_sel])]

    return out


def _kpi_block(df: pd.DataFrame, cmap: ColumnMap):
    col1, col2, col3, col4 = st.columns(4)

    n = len(df)
    col1.metric("Liczba odpowiedzi (po filtrach)", f"{n}")

    # Overall satisfaction
    if cmap.overall_sat_col:
        sat = _safe_to_numeric(df[cmap.overall_sat_col])
        col2.metric("Średnia satysfakcja", f"{sat.mean():.2f}" if sat.notna().any() else "–")
    else:
        col2.metric("Średnia satysfakcja", "–")

    # Expectations met %
    if cmap.expectations_col:
        exp_bool = _safe_to_bool_from_pl(df[cmap.expectations_col])
        if exp_bool.notna().any():
            pct = 100.0 * exp_bool.mean()
            col3.metric("Spełnione oczekiwania", f"{pct:.1f}%")
        else:
            col3.metric("Spełnione oczekiwania", "–")
    else:
        col3.metric("Spełnione oczekiwania", "–")

    # Trainer rating
    if cmap.trainer_rating_col:
        tr = _safe_to_numeric(df[cmap.trainer_rating_col])
        col4.metric("Średnia ocena prowadzącego", f"{tr.mean():.2f}" if tr.notna().any() else "–")
    else:
        col4.metric("Średnia ocena prowadzącego", "–")


# -----------------------------
# Sidebar: upload + settings
# -----------------------------
with st.sidebar:
    st.header("⚙️ Wejście danych")
    uploaded = st.file_uploader("Wgraj plik Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

    use_demo = st.toggle("Użyj danych demo (opcjonalnie)", value=False, help="Pozwala uruchomić aplikację bez wgrywania pliku.")

    st.divider()
    st.header("🔎 Filtry")

if use_demo and uploaded is None:
    # Built-in demo dataset aligned with your file structure
    demo = pd.DataFrame(
        {
            "ID": range(1, 11),
            "Płeć": ["Kobieta", "Mężczyzna"] * 5,
            "Wiek": [24, 39, 31, 47, 28, 42, 36, 29, 51, 33],
            "Staż pracy (lata)": [1, 7, 5, 8, 3, 16, 10, 4, 20, 6],
            "Ocena organizacji szkolenia (1–5)": [3, 3, 4, 3, 4, 3, 5, 4, 3, 4],
            "Ocena prowadzącego (1–5)": [4, 4, 5, 4, 5, 4, 5, 4, 4, 5],
            "Czy szkolenie spełniło oczekiwania (tak/nie)": ["tak", "tak", "tak", "nie", "tak", "tak", "tak", "tak", "nie", "tak"],
            "Najbardziej wartościowy element": ["Materiały szkoleniowe", "Dyskusje w grupach", "Przykłady z życia", "Dyskusje w grupach", "Materiały szkoleniowe",
                                               "Materiały szkoleniowe", "Przykłady z życia", "Dyskusje w grupach", "Przykłady z życia", "Materiały szkoleniowe"],
            "Co można poprawić": ["Więcej przykładów praktycznych", "Dłuższe przerwy", "Więcej interakcji", "Więcej interakcji", "Dłuższe przerwy",
                                  "Więcej przykładów praktycznych", "Więcej interakcji", "Dłuższe przerwy", "Więcej interakcji", "Więcej przykładów praktycznych"],
            "Ogólna satysfakcja (1–5)": [5, 4, 4, 4, 5, 4, 5, 4, 3, 5],
        }
    )
    df_raw = demo.copy()
    sheet_name = "DEMO"
else:
    if uploaded is None:
        st.info("Wgraj plik Excel w panelu po lewej stronie lub włącz dane demo.")
        st.stop()

    uploaded_bytes = uploaded.getvalue()
    xls, sheets = load_excel(uploaded_bytes)

    with st.sidebar:
        sheet_name = st.selectbox(
            "Arkusz do wczytania",
            options=sheets,
            index=sheets.index("Dane surowe") if "Dane surowe" in sheets else 0,
        )

    df_raw = read_sheet(uploaded_bytes, sheet_name=sheet_name)

# Auto-map columns
cmap_auto = _auto_map_columns(list(df_raw.columns))

# Allow manual mapping (robustness)
with st.sidebar:
    st.subheader("Mapowanie kolumn (auto + ręcznie)")
    cols = ["(brak)"] + list(df_raw.columns)

    def pick(label: str, current: Optional[str]) -> Optional[str]:
        idx = cols.index(current) if current in cols else 0
        chosen = st.selectbox(label, options=cols, index=idx)
        return None if chosen == "(brak)" else chosen

    cmap = ColumnMap(
        id_col=pick("ID", cmap_auto.id_col),
        gender_col=pick("Płeć", cmap_auto.gender_col),
        age_col=pick("Wiek", cmap_auto.age_col),
        tenure_col=pick("Staż pracy", cmap_auto.tenure_col),
        org_rating_col=pick("Ocena organizacji", cmap_auto.org_rating_col),
        trainer_rating_col=pick("Ocena prowadzącego", cmap_auto.trainer_rating_col),
        expectations_col=pick("Spełnione oczekiwania", cmap_auto.expectations_col),
        best_element_col=pick("Najbardziej wartościowy element", cmap_auto.best_element_col),
        improve_col=pick("Co można poprawić", cmap_auto.improve_col),
        overall_sat_col=pick("Ogólna satysfakcja", cmap_auto.overall_sat_col),
    )

# Filters UI
with st.sidebar:
    # Gender
    gender_sel = None
    if cmap.gender_col:
        genders = sorted(df_raw[cmap.gender_col].dropna().astype(str).unique().tolist())
        if genders:
            gender_sel = st.multiselect("Płeć", options=genders, default=genders)

    # Age range
    age_range = None
    if cmap.age_col:
        age_num = _safe_to_numeric(df_raw[cmap.age_col]).dropna()
        if not age_num.empty:
            amin, amax = float(age_num.min()), float(age_num.max())
            age_range = st.slider("Wiek", min_value=amin, max_value=amax, value=(amin, amax))

    # Tenure range
    tenure_range = None
    if cmap.tenure_col:
        ten_num = _safe_to_numeric(df_raw[cmap.tenure_col]).dropna()
        if not ten_num.empty:
            tmin, tmax = float(ten_num.min()), float(ten_num.max())
            tenure_range = st.slider("Staż pracy (lata)", min_value=tmin, max_value=tmax, value=(tmin, tmax))

    # Expectations
    exp_sel = None
    if cmap.expectations_col:
        exps = sorted(df_raw[cmap.expectations_col].dropna().astype(str).str.strip().str.lower().unique().tolist())
        if exps:
            exp_sel = st.multiselect("Czy spełniło oczekiwania", options=exps, default=exps)

# Apply filters
df = _apply_filters(df_raw, cmap, gender_sel, age_range, tenure_range, exp_sel)

# -----------------------------
# Main layout
# -----------------------------
st.subheader(f"Źródło danych: {sheet_name}")
_kpi_block(df, cmap)

tab1, tab2, tab3 = st.tabs(["📈 Wykresy", "🧾 Tabela i eksport", "🧪 Jakość danych"])

with tab1:
    left, right = st.columns(2)

    with left:
        if cmap.overall_sat_col:
            s = _safe_to_numeric(df[cmap.overall_sat_col])
            s.name = cmap.overall_sat_col
            _plot_hist(s, "Rozkład ogólnej satysfakcji", bins=5)
        else:
            st.info("Nie wskazano kolumny ogólnej satysfakcji.")

        if cmap.best_element_col:
            _plot_bar_counts(df[cmap.best_element_col], "Najbardziej wartościowy element (top)", top_n=12)
        else:
            st.info("Nie wskazano kolumny „Najbardziej wartościowy element”.")

    with right:
        if cmap.gender_col and cmap.overall_sat_col:
            s = _safe_to_numeric(df[cmap.overall_sat_col])
            s.name = cmap.overall_sat_col
            _plot_box_by_category(s, df[cmap.gender_col], "Satysfakcja a płeć")
        else:
            st.info("Aby pokazać wykres pudełkowy, potrzebne są kolumny: płeć i ogólna satysfakcja.")

        if cmap.age_col and cmap.overall_sat_col:
            x = _safe_to_numeric(df[cmap.age_col])
            x.name = cmap.age_col
            y = _safe_to_numeric(df[cmap.overall_sat_col])
            y.name = cmap.overall_sat_col
            _plot_scatter(x, y, "Wiek a satysfakcja (scatter)")
        else:
            st.info("Aby pokazać scatter, potrzebne są kolumny: wiek i ogólna satysfakcja.")

with tab2:
    st.write("Dane po filtrach:")
    st.dataframe(df, use_container_width=True)

    c1, c2 = st.columns([1, 1])
    with c1:
        st.download_button(
            label="⬇️ Pobierz CSV (dane po filtrach)",
            data=_download_csv_bytes(df),
            file_name="ankieta_po_filtrach.csv",
            mime="text/csv",
        )
    with c2:
        st.download_button(
            label="⬇️ Pobierz CSV (dane surowe)",
            data=_download_csv_bytes(df_raw),
            file_name="ankieta_surowe.csv",
            mime="text/csv",
        )

with tab3:
    st.write("Podstawowa diagnostyka braków i typów danych.")
    qc = pd.DataFrame(
        {
            "kolumna": df_raw.columns,
            "typ": [str(df_raw[c].dtype) for c in df_raw.columns],
            "braki (count)": [int(df_raw[c].isna().sum()) for c in df_raw.columns],
            "braki (%)": [float(df_raw[c].isna().mean() * 100.0) for c in df_raw.columns],
            "unikalne (count)": [int(df_raw[c].nunique(dropna=True)) for c in df_raw.columns],
        }
    )
    st.dataframe(qc, use_container_width=True)

    st.markdown("#### Podgląd mapowania kolumn")
    st.json(cmap.__dict__, expanded=False)

st.caption("Uwaga: wykresy są generowane w Matplotlib; aplikacja nie wysyła danych nigdzie poza Twoją sesję.")