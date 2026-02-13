import io
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Dashboard ankiety satysfakcji (Excel → Streamlit)",
    page_icon="📊",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def get_sheet_names(uploaded_bytes: bytes) -> list[str]:
    bio = io.BytesIO(uploaded_bytes)
    xls = pd.ExcelFile(bio)
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def load_sheet(uploaded_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    bio = io.BytesIO(uploaded_bytes)
    df = pd.read_excel(bio, sheet_name=sheet_name)
    # Normalizacja nazw kolumn (bez utraty polskich znaków)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _first_existing_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    return None


def _to_numeric_series(s: pd.Series) -> pd.Series:
    if s is None:
        return s
    # Obsługa wartości typu "1–5" w nagłówku nie ma znaczenia, ale w danych mogą być stringi
    return pd.to_numeric(s.astype(str).str.replace(",", ".", regex=False), errors="coerce")


def _kpi_card(label: str, value: str, help_text: Optional[str] = None) -> None:
    st.metric(label=label, value=value, help=help_text)


def _safe_value_counts(df: pd.DataFrame, col: str) -> pd.DataFrame:
    vc = (
        df[col]
        .astype(str)
        .replace({"nan": np.nan, "None": np.nan})
        .dropna()
        .str.strip()
        .replace({"": np.nan})
        .dropna()
        .value_counts()
    )
    return vc.rename_axis("Odpowiedź").reset_index(name="Liczba")


def main() -> None:
    st.title("📊 Dashboard ankiety satysfakcji (Streamlit + Excel)")
    st.caption(
        "Wgraj plik Excel. Aplikacja wczyta arkusze, pozwoli filtrować dane i pokaże analizy (oceny + odpowiedzi otwarte)."
    )

    uploaded = st.file_uploader(
        "Wgraj plik Excel (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        help="Plik nie jest zapisywany na serwerze, jest przetwarzany w pamięci sesji.",
    )

    if not uploaded:
        st.info("Wgraj plik Excel, aby rozpocząć.")
        return

    uploaded_bytes = uploaded.getvalue()

    try:
        sheets = get_sheet_names(uploaded_bytes)
    except Exception as e:
        st.error("Nie udało się odczytać arkuszy z pliku Excel.")
        st.exception(e)
        return

    with st.sidebar:
        st.header("Ustawienia")
        sheet = st.selectbox("Arkusz", options=sheets, index=0)

    try:
        df_raw = load_sheet(uploaded_bytes, sheet)
    except Exception as e:
        st.error("Nie udało się wczytać danych z wybranego arkusza.")
        st.exception(e)
        return

    if df_raw.empty:
        st.warning("Wybrany arkusz jest pusty.")
        return

    # --- Mapowanie kolumn (dopasowane do Twojego pliku; działa też, gdy ktoś zmieni nagłówki na podobne) ---
    col_id = _first_existing_col(df_raw, ["ID", "Id", "id"])
    col_gender = _first_existing_col(df_raw, ["Płeć", "Plec", "Gender", "PŁEĆ"])
    col_age = _first_existing_col(df_raw, ["Wiek", "Age"])
    col_tenure = _first_existing_col(df_raw, ["Staż pracy (lata)", "Staż pracy", "Staz pracy (lata)", "Tenure"])
    col_org = _first_existing_col(df_raw, ["Ocena organizacji szkolenia (1–5)", "Ocena organizacji szkolenia"])
    col_trainer = _first_existing_col(df_raw, ["Ocena prowadzącego (1–5)", "Ocena prowadzącego"])
    col_expect = _first_existing_col(
        df_raw,
        ["Czy szkolenie spełniło oczekiwania (tak/nie)", "Czy szkolenie spełniło oczekiwania", "Spełniło oczekiwania"],
    )
    col_best = _first_existing_col(df_raw, ["Najbardziej wartościowy element", "Najbardziej wartosciowy element"])
    col_improve = _first_existing_col(df_raw, ["Co można poprawić", "Co mozna poprawic", "Co można poprawic"])
    col_overall = _first_existing_col(df_raw, ["Ogólna satysfakcja (1–5)", "Ogolna satysfakcja (1–5)", "Ogólna satysfakcja"])

    # --- Przygotowanie danych ---
    df = df_raw.copy()

    # Numeryczne pola
    if col_age:
        df[col_age] = _to_numeric_series(df[col_age])
    if col_tenure:
        df[col_tenure] = _to_numeric_series(df[col_tenure])
    for c in [col_org, col_trainer, col_overall]:
        if c:
            df[c] = _to_numeric_series(df[c])

    # Normalizacja tak/nie
    if col_expect:
        df[col_expect] = (
            df[col_expect]
            .astype(str)
            .str.strip()
            .str.lower()
            .replace({"tak": "tak", "nie": "nie", "yes": "tak", "no": "nie", "true": "tak", "false": "nie"})
        )

    # --- Filtry ---
    with st.sidebar:
        st.subheader("Filtry")

        df_f = df.copy()

        if col_gender:
            genders = sorted(df_f[col_gender].dropna().astype(str).unique().tolist())
            sel_genders = st.multiselect("Płeć", options=genders, default=genders)
            if sel_genders:
                df_f = df_f[df_f[col_gender].astype(str).isin(sel_genders)]

        if col_expect:
            expects = ["tak", "nie"]
            present = sorted(set(df_f[col_expect].dropna().astype(str).unique().tolist()) & set(expects))
            if present:
                sel_expect = st.multiselect("Spełnienie oczekiwań", options=present, default=present)
                if sel_expect:
                    df_f = df_f[df_f[col_expect].astype(str).isin(sel_expect)]

        if col_age and df_f[col_age].notna().any():
            amin = int(np.nanmin(df_f[col_age].values))
            amax = int(np.nanmax(df_f[col_age].values))
            age_range = st.slider("Wiek", min_value=amin, max_value=amax, value=(amin, amax))
            df_f = df_f[df_f[col_age].between(age_range[0], age_range[1], inclusive="both")]

        if col_tenure and df_f[col_tenure].notna().any():
            tmin = float(np.nanmin(df_f[col_tenure].values))
            tmax = float(np.nanmax(df_f[col_tenure].values))
            tenure_range = st.slider("Staż pracy (lata)", min_value=float(tmin), max_value=float(tmax), value=(float(tmin), float(tmax)))
            df_f = df_f[df_f[col_tenure].between(tenure_range[0], tenure_range[1], inclusive="both")]

        st.divider()
        st.caption(f"Wiersze po filtrach: **{len(df_f)}** / {len(df)}")

    # --- KPI ---
    kpi_cols = st.columns(4)

    with kpi_cols[0]:
        _kpi_card("Liczba odpowiedzi", str(len(df_f)))

    def mean_fmt(series: Optional[pd.Series]) -> str:
        if series is None or series.dropna().empty:
            return "—"
        return f"{series.mean():.2f}"

    with kpi_cols[1]:
        _kpi_card("Śr. satysfakcja", mean_fmt(df_f[col_overall]) if col_overall else "—")

    with kpi_cols[2]:
        _kpi_card("Śr. ocena prowadzącego", mean_fmt(df_f[col_trainer]) if col_trainer else "—")

    with kpi_cols[3]:
        if col_expect and len(df_f) > 0:
            share_yes = (df_f[col_expect].astype(str) == "tak").mean() * 100
            _kpi_card("Oczekiwania spełnione", f"{share_yes:.1f}%")
        else:
            _kpi_card("Oczekiwania spełnione", "—")

    # --- Zakładki ---
    tab_overview, tab_ratings, tab_segments, tab_text, tab_data = st.tabs(
        ["Przegląd", "Oceny", "Segmenty", "Odpowiedzi otwarte", "Dane"]
    )

    with tab_overview:
        left, right = st.columns([1.1, 0.9], gap="large")

        with left:
            st.subheader("Struktura próby")
            chart_rows = []

            if col_gender:
                g = _safe_value_counts(df_f, col_gender)
                if not g.empty:
                    fig = px.pie(g, names="Odpowiedź", values="Liczba", title="Płeć")
                    st.plotly_chart(fig, use_container_width=True)

            if col_expect:
                e = _safe_value_counts(df_f, col_expect)
                if not e.empty:
                    fig = px.bar(e, x="Odpowiedź", y="Liczba", title="Spełnienie oczekiwań", text="Liczba")
                    fig.update_layout(xaxis_title="", yaxis_title="Liczba")
                    st.plotly_chart(fig, use_container_width=True)

        with right:
            st.subheader("Wiek i staż")
            if col_age and df_f[col_age].notna().any():
                fig = px.histogram(df_f, x=col_age, nbins=10, title="Rozkład wieku")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Brak kolumny wieku lub brak danych liczbowych.")

            if col_tenure and df_f[col_tenure].notna().any():
                fig = px.histogram(df_f, x=col_tenure, nbins=10, title="Rozkład stażu pracy (lata)")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Brak kolumny stażu lub brak danych liczbowych.")

    with tab_ratings:
        st.subheader("Oceny (1–5)")

        rating_cols = [c for c in [col_org, col_trainer, col_overall] if c]
        if not rating_cols:
            st.warning("Nie wykryto kolumn z ocenami (np. satysfakcja, prowadzący, organizacja).")
        else:
            c1, c2 = st.columns([1, 1], gap="large")

            with c1:
                pick = st.selectbox("Wybierz zmienną do rozkładu", options=rating_cols)
                if df_f[pick].notna().any():
                    fig = px.histogram(df_f, x=pick, nbins=5, title=f"Rozkład: {pick}")
                    fig.update_layout(xaxis_title="", yaxis_title="Liczba")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Brak danych do wykresu.")

            with c2:
                st.markdown("**Porównanie średnich**")
                melted = df_f[rating_cols].melt(var_name="Zmienna", value_name="Ocena").dropna()
                if not melted.empty:
                    means = melted.groupby("Zmienna", as_index=False)["Ocena"].mean().sort_values("Ocena", ascending=False)
                    fig = px.bar(means, x="Zmienna", y="Ocena", title="Średnia ocena wg zmiennej", text=means["Ocena"].round(2))
                    fig.update_layout(xaxis_title="", yaxis_title="Średnia")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Brak danych do porównania.")

            st.divider()
            st.markdown("**Korelacje (jeśli dostępne)**")
            num_cols = [c for c in [col_age, col_tenure, col_org, col_trainer, col_overall] if c]
            num_df = df_f[num_cols].copy()
            for c in num_cols:
                num_df[c] = _to_numeric_series(num_df[c])

            corr = num_df.corr(numeric_only=True)
            if corr.shape[0] >= 2:
                fig = px.imshow(corr, text_auto=True, title="Macierz korelacji")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Za mało zmiennych liczbowych, aby policzyć korelacje.")

    with tab_segments:
        st.subheader("Segmenty i porównania")

        # Grupowanie po zmiennej kategorycznej (płeć / oczekiwania)
        group_candidates = [c for c in [col_gender, col_expect] if c]
        rating_cols = [c for c in [col_org, col_trainer, col_overall] if c]

        if not group_candidates or not rating_cols:
            st.info("Aby zrobić segmentację, potrzebne są: (1) kolumna kategoryczna oraz (2) kolumny z ocenami.")
        else:
            group_col = st.selectbox("Grupuj według", options=group_candidates)

            tmp = df_f[[group_col] + rating_cols].copy()
            for c in rating_cols:
                tmp[c] = _to_numeric_series(tmp[c])

            tmp[group_col] = tmp[group_col].astype(str).str.strip()
            tmp = tmp.dropna(subset=[group_col])

            grouped = tmp.groupby(group_col, as_index=False)[rating_cols].mean(numeric_only=True)

            st.dataframe(grouped, use_container_width=True)

            melted = grouped.melt(id_vars=[group_col], var_name="Zmienna", value_name="Średnia")
            fig = px.bar(
                melted,
                x=group_col,
                y="Średnia",
                color="Zmienna",
                barmode="group",
                title="Średnie oceny w grupach",
                text=melted["Średnia"].round(2),
            )
            fig.update_layout(xaxis_title="", yaxis_title="Średnia")
            st.plotly_chart(fig, use_container_width=True)

    with tab_text:
        st.subheader("Odpowiedzi otwarte")
        st.caption("Analiza częstotliwości najczęściej pojawiających się odpowiedzi (bez NLP).")

        left, right = st.columns(2, gap="large")

        with left:
            st.markdown("**Najbardziej wartościowy element**")
            if col_best:
                vc = _safe_value_counts(df_f, col_best)
                if vc.empty:
                    st.info("Brak danych tekstowych.")
                else:
                    fig = px.bar(vc.head(20), x="Liczba", y="Odpowiedź", orientation="h", title="Top odpowiedzi")
                    st.plotly_chart(fig, use_container_width=True)
                    st.dataframe(vc, use_container_width=True, height=320)
            else:
                st.info("Nie wykryto kolumny z 'Najbardziej wartościowy element'.")

        with right:
            st.markdown("**Co można poprawić**")
            if col_improve:
                vc = _safe_value_counts(df_f, col_improve)
                if vc.empty:
                    st.info("Brak danych tekstowych.")
                else:
                    fig = px.bar(vc.head(20), x="Liczba", y="Odpowiedź", orientation="h", title="Top sugestie")
                    st.plotly_chart(fig, use_container_width=True)
                    st.dataframe(vc, use_container_width=True, height=320)
            else:
                st.info("Nie wykryto kolumny z 'Co można poprawić'.")

    with tab_data:
        st.subheader("Podgląd danych + eksport")

        show_cols = st.multiselect(
            "Wybierz kolumny do podglądu",
            options=df_f.columns.tolist(),
            default=df_f.columns.tolist(),
        )

        st.dataframe(df_f[show_cols], use_container_width=True, height=420)

        csv = df_f.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Pobierz dane po filtrach (CSV)",
            data=csv,
            file_name="dane_po_filtrach.csv",
            mime="text/csv",
        )

        st.caption("Uwaga: eksport jest w UTF-8 (polskie znaki zachowane).")


if __name__ == "__main__":
    main()