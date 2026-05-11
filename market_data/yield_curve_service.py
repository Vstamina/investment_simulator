from __future__ import annotations

import pandas as pd


def build_simple_focus_curve(focus_df: pd.DataFrame) -> pd.DataFrame:
    """
    Monta uma curva simples de juros usando a mediana Focus da Selic por ano.

    Esta curva é uma aproximação consultiva.
    Não substitui curva DI futuro, B3 ou ETTJ da ANBIMA.
    """

    if focus_df is None or focus_df.empty:
        return pd.DataFrame()

    if "Indicador" not in focus_df.columns:
        return pd.DataFrame()

    selic_df = focus_df[
        focus_df["Indicador"].astype(str).str.lower() == "selic"
    ].copy()

    if selic_df.empty:
        return pd.DataFrame()

    required_columns = [
        "DataReferencia",
        "Mediana",
    ]

    for column in required_columns:
        if column not in selic_df.columns:
            return pd.DataFrame()

    curve_df = selic_df[required_columns].copy()
    curve_df = curve_df.dropna(subset=["DataReferencia", "Mediana"])
    curve_df = curve_df.drop_duplicates(subset=["DataReferencia"])
    curve_df = curve_df.sort_values("DataReferencia")

    curve_df = curve_df.rename(
        columns={
            "DataReferencia": "Vértice",
            "Mediana": "Taxa Selic Esperada (%)",
        }
    )

    curve_df["Fonte"] = "Focus/Bacen"
    curve_df["Tipo"] = "Curva simplificada por expectativa"

    return curve_df


def classify_curve_shape(curve_df: pd.DataFrame) -> str:
    """
    Classifica a inclinação da curva de forma simples.
    """

    if curve_df is None or curve_df.empty:
        return "sem dados suficientes"

    if len(curve_df) < 2:
        return "curva curta, sem comparação suficiente"

    try:
        first_rate = float(curve_df.iloc[0]["Taxa Selic Esperada (%)"])
        last_rate = float(curve_df.iloc[-1]["Taxa Selic Esperada (%)"])
    except Exception:
        return "sem dados suficientes"

    difference = last_rate - first_rate

    if difference > 0.50:
        return "curva ascendente"

    if difference < -0.50:
        return "curva descendente"

    return "curva relativamente estável"