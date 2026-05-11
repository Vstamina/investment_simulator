from __future__ import annotations

import requests
import pandas as pd


FOCUS_BASE_URL = (
    "https://olinda.bcb.gov.br/olinda/servico/"
    "Expectativas/versao/v1/odata"
)


def fetch_focus_annual_expectations(
    indicator: str,
    top: int = 10,
) -> pd.DataFrame:
    """
    Busca expectativas anuais do Focus pela API OData do Banco Central.
    """

    url = (
        f"{FOCUS_BASE_URL}/ExpectativasMercadoAnuais"
        f"?$top={top}"
        f"&$filter=Indicador eq '{indicator}'"
        f"&$orderby=Data desc"
        f"&$format=json"
    )

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()

        data = response.json().get("value", [])

        if not data:
            return pd.DataFrame()

        df = pd.DataFrame(data)

        selected_columns = [
            column for column in [
                "Indicador",
                "Data",
                "DataReferencia",
                "Media",
                "Mediana",
                "DesvioPadrao",
                "Minimo",
                "Maximo",
                "numeroRespondentes",
            ]
            if column in df.columns
        ]

        return df[selected_columns]

    except Exception:
        return pd.DataFrame()


def get_focus_snapshot() -> pd.DataFrame:
    """
    Consolida expectativas principais do Focus.
    """

    indicators = [
        "Selic",
        "IPCA",
    ]

    frames = []

    for indicator in indicators:
        df = fetch_focus_annual_expectations(
            indicator=indicator,
            top=10,
        )

        if not df.empty:
            frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)
