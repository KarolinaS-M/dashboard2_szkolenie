import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


APP_TITLE = "Dashboard satysfakcji ze szkolenia"
DEFAULT_XLSX_PATH = Path("data/ankieta_satysfakcja_szkolenie.xlsx")
SHEET_NAME = "Dane surowe"


@st.cache_data(show_spinner=False)
def load_data_from_path(xlsx_path: Path) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME)
    return df


@st.cache_data(show_spinner=False)
def load_data_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SHEET_NAME)
    return df


def normalize_yes_no(s: pd.Series) -> pd.Series:
    # Normalizacja do {"tak","nie"} (bezpiecznie dla różnych wariantów zapisu)
    s2 = (
        s.astype(str)
        .str.strip()
        .str.lower()
        .replace({"yes": "tak", "no": "nie", "true": "tak", "false": "nie"})
    )
    # jeśli pojawi się coś innego, zostawiamy oryginał po obróbce
    return s2


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    # Ujednolicenia i bezpieczne typy
    df = df.copy()

    # Standardowe nazwy kolumn z Twojego pliku
    col_map = {
        "ID": "ID",
        "Płeć": "Płeć",
        "Wiek": "Wiek",
        "Staż pracy (lata)": "Staż",
        "Ocena organizacji szkolenia (1–5)": "Ocena_organizacji",
        "Ocena prowadzącego (1–5)": "Ocena_prowadzącego",
        "Czy szkolenie spełniło oczekiwania (tak/nie)": "Oczekiwania",
        "Najbardziej wartościowy element": "Wartościowy_element",
        "Co można poprawić": "Do_poprawy",
        "Ogólna satysfakcja (1–5)": "Satysfakcja",
    }

    missing = [c for c in col_map.keys() if c not in df.columns]
    if missing:
        raise ValueError(
            "Brakuje oczekiwanych kolumn w danych: "
            + ", ".join(missing)
            + ". Sprawdź arkusz i nagłówki."
        )

    df = df.rename(columns=col_map)

    for c in ["Wiek", "Staż", "Ocena_organizacji", "Ocena_prowadzącego", "Satysfakcja"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Oczekiwania"] = normalize_yes_no(df["Oczekiwania"])
    df["Płeć"] = df["Płeć"].astype(str).str.strip()
    df["Wartościowy_element"] = df["Wartościowy_element"].astype(str).str.strip()
    df["Do_poprawy"] = df["Do_poprawy"].astype(str).str.strip()

    # Dodatkowe zmienne do dashboardu
    df["Grupa_wieku"] = pd.cut(
        df["Wiek"],
        bins=[0, 24, 34, 44, 54, 64, np.inf],
        labels=["≤24", "25–34", "35–44", "45–54", "55–64", "65+"],
        right=True,
        include_lowest=True,
    )

    df["Grupa_stażu"] = pd.cut(
        df["Staż"],
        bins=[-np.inf, 1, 3, 5, 10, 20, np.inf],
        labels=["≤1", "2–3", "4–5", "6–10", "11–20", "21+"],
        right=True,
    )

    return df


def kpi_block(df: pd.DataFrame) -> None:
    n = len(df)
    if n == 0:
        st.warning("Brak obserwacji po zastosowaniu filtrów.")
        return

    avg_sat = float(df["Satysfakcja"].mean())
    avg_org = float(df["Ocena_organizacji"].mean())
    avg_trainer = float(df["Ocena_prowadzącego"].mean())
    share_yes = float((df["Oczekiwania"] == "tak").mean() * 100.0)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Liczba odpowiedzi", f"{n}")
    c2.metric("Średnia satysfakcja (1–5)", f"{avg_sat:.2f}")
    c3.metric("Śr. ocena organizacji (1–5)", f"{avg_org:.2f}")
    c4.metric("Śr. ocena prowadzącego (1–5)", f"{avg_trainer:.2f}")

    st.caption(f"Odsetek odpowiedzi „tak” (spełnienie oczekiwań): **{share_yes:.1f}%**.")


def bar_counts(df: pd.DataFrame, column: str, title: str) -> None:
    tmp = (
        df[column]
        .fillna("Brak danych")
        .astype(str)
        .value_counts(dropna=False)
        .rename_axis(column)
        .reset_index(name="Liczba")
    )
    fig = px.bar(tmp, x=column, y="Liczba", title=title)
    fig.update_layout(xaxis_title="", yaxis_title="Liczba", margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


def ratings_distribution(df: pd.DataFrame, column: str, title: str) -> None:
    tmp = (
        df[column]
        .dropna()
        .astype(int)
        .value_counts()
        .reindex([1, 2, 3, 4, 5], fill_value=0)
        .rename_axis("Ocena")
        .reset_index(name="Liczba")
    )
    fig = px.bar(tmp, x="Ocena", y="Liczba", title=title)
    fig.update_layout(xaxis_title="Ocena", yaxis_title="Liczba", margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


def scatter_relation(df: pd.DataFrame) -> None:
    tmp = df.dropna(subset=["Satysfakcja", "Ocena_organizacji", "Ocena_prowadzącego"])
    if len(tmp) < 2:
        st.info("Za mało danych do wykresu zależności po filtrach.")
        return

    fig = px.scatter(
        tmp,
        x="Ocena_organizacji",
        y="Satysfakcja",
        color="Oczekiwania",
        size="Ocena_prowadzącego",
        hover_data=["Płeć", "Wiek", "Staż", "Wartościowy_element"],
        title="Satysfakcja a ocena organizacji (rozmiar: ocena prowadzącego)",
    )
    fig.update_layout(
        xaxis_title="Ocena organizacji (1–5)",
        yaxis_title="Ogólna satysfakcja (1–5)",
        margin=dict(l=10, r=10, t=50, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)


def correlation_heatmap(df: pd.DataFrame) -> None:
    cols = ["Wiek", "Staż", "Ocena_organizacji", "Ocena_prowadzącego", "Satysfakcja"]
    tmp = df[cols].copy()
    if tmp.dropna().shape[0] < 2:
        st.info("Za mało kompletnych obserwacji do korelacji po filtrach.")
        return

    corr = tmp.corr(numeric_only=True)
    fig = px.imshow(
        corr.round(3),
        text_auto=True,
        title="Korelacje (Pearson) pomiędzy zmiennymi liczbowymi",
        aspect="auto",
    )
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.write(
        "Dashboard działa na danych z pliku Excel (arkusz **Dane surowe**). "
        "Możesz korzystać z pliku w repozytorium lub wgrać własny plik w sidebarze."
    )

    with st.sidebar:
        st.header("Źródło danych")
        uploaded = st.file_uploader("Wgraj plik Excel (.xlsx)", type=["xlsx"])

        st.divider()
        st.header("Filtry")

    # Wczytanie danych
    try:
        if uploaded is not None:
            raw = load_data_from_bytes(uploaded.getvalue())
            data_source_note = "Dane wczytane z pliku wgranego w aplikacji."
        else:
            if not DEFAULT_XLSX_PATH.exists():
                st.error(
                    "Nie znaleziono domyślnego pliku danych w repozytorium: "
                    f"`{DEFAULT_XLSX_PATH}`. Wgraj plik w sidebarze lub dodaj go do repo."
                )
                st.stop()
            raw = load_data_from_path(DEFAULT_XLSX_PATH)
            data_source_note = f"Dane wczytane z repozytorium: `{DEFAULT_XLSX_PATH}`."

        df = prepare(raw)

    except Exception as e:
        st.error("Nie udało się wczytać lub przygotować danych.")
        st.exception(e)
        st.stop()

    st.caption(data_source_note)

    # Filtry w sidebarze (po przygotowaniu df)
    with st.sidebar:
        # Płeć
        płcie = sorted([p for p in df["Płeć"].dropna().unique().tolist() if p != "nan"])
        selected_płeć = st.multiselect("Płeć", options=płcie, default=płcie)

        # Oczekiwania
        ocz_opts = sorted(df["Oczekiwania"].dropna().unique().tolist())
        selected_ocz = st.multiselect("Spełnienie oczekiwań", options=ocz_opts, default=ocz_opts)

        # Wiek i staż
        min_w, max_w = int(np.nanmin(df["Wiek"])), int(np.nanmax(df["Wiek"]))
        min_s, max_s = int(np.nanmin(df["Staż"])), int(np.nanmax(df["Staż"]))

        age_range = st.slider("Wiek", min_value=min_w, max_value=max_w, value=(min_w, max_w))
        tenure_range = st.slider("Staż pracy (lata)", min_value=min_s, max_value=max_s, value=(min_s, max_s))

        # Wartościowy element
        elem_opts = sorted(df["Wartościowy_element"].dropna().unique().tolist())
        selected_elem = st.multiselect("Najbardziej wartościowy element", options=elem_opts, default=elem_opts)

        show_raw = st.checkbox("Pokaż tabelę danych (po filtrach)", value=False)

    # Zastosowanie filtrów
    f = df.copy()
    f = f[f["Płeć"].isin(selected_płeć)]
    f = f[f["Oczekiwania"].isin(selected_ocz)]
    f = f[(f["Wiek"] >= age_range[0]) & (f["Wiek"] <= age_range[1])]
    f = f[(f["Staż"] >= tenure_range[0]) & (f["Staż"] <= tenure_range[1])]
    f = f[f["Wartościowy_element"].isin(selected_elem)]

    # KPI
    kpi_block(f)
    st.divider()

    # Układ wykresów
    colA, colB = st.columns(2)
    with colA:
        ratings_distribution(f, "Satysfakcja", "Rozkład ogólnej satysfakcji (1–5)")
        bar_counts(f, "Wartościowy_element", "Najbardziej wartościowy element (liczności)")

    with colB:
        ratings_distribution(f, "Ocena_organizacji", "Rozkład oceny organizacji (1–5)")
        ratings_distribution(f, "Ocena_prowadzącego", "Rozkład oceny prowadzącego (1–5)")

    st.divider()

    colC, colD = st.columns(2)
    with colC:
        bar_counts(f, "Do_poprawy", "Co można poprawić (liczności)")
    with colD:
        bar_counts(f, "Grupa_wieku", "Struktura wieku (przedziały)")

    st.divider()

    colE, colF = st.columns(2)
    with colE:
        scatter_relation(f)
    with colF:
        correlation_heatmap(f)

    if show_raw:
        st.divider()
        st.subheader("Dane po filtrach")
        st.dataframe(
            f.sort_values("ID"),
            use_container_width=True,
            hide_index=True,
        )


if __name__ == "__main__":
    main()