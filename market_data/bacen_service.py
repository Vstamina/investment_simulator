from __future__ import annotations

from dataclasses import dataclass
import requests
import pandas as pd


BCB_SGS_BASE_URL = "https://api.bcb.gov.br/dados/serie/bcdata.sgs"


@dataclass
class BacenIndicator:
    name: str
    code: int
    value: float | None
    date: str | None
    unit: str
    source: str = "Banco Central do Brasil - SGS"


def _parse_bacen_value(value: str | float | int | None) -> float | None:
    if value is None:
        return None

    try:
        return float(str(value).replace(",", "."))
    except ValueError:
        return None


def fetch_latest_sgs_value(
    series_code: int,
    indicator_name: str,
    unit: str,
) -> BacenIndicator:
    """
    Busca o último valor disponível de uma série SGS do Banco Central.
    Se a API falhar, retorna value=None sem derrubar o app.
    """

    url = f"{BCB_SGS_BASE_URL}.{series_code}/dados/ultimos/1?formato=json"

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()

        if not data:
            return BacenIndicator(
                name=indicator_name,
                code=series_code,
                value=None,
                date=None,
                unit=unit,
            )

        last_item = data[-1]

        return BacenIndicator(
            name=indicator_name,
            code=series_code,
            value=_parse_bacen_value(last_item.get("valor")),
            date=last_item.get("data"),
            unit=unit,
        )

    except Exception:
        return BacenIndicator(
            name=indicator_name,
            code=series_code,
            value=None,
            date=None,
            unit=unit,
        )


def get_bacen_snapshot() -> pd.DataFrame:
    """
    Retorna um quadro inicial de indicadores do Bacen.

    Séries usadas nesta primeira versão:
    - 432: Meta Selic definida pelo Copom, em % ao ano.
    - 11: Selic diária, em % ao dia.
    """

    indicators = [
        fetch_latest_sgs_value(
            series_code=432,
            indicator_name="Meta Selic",
            unit="% a.a.",
        ),
        fetch_latest_sgs_value(
            series_code=11,
            indicator_name="Selic diária",
            unit="% a.d.",
        ),
    ]

    rows = []

    for item in indicators:
        rows.append(
            {
                "Indicador": item.name,
                "Código SGS": item.code,
                "Valor": item.value,
                "Unidade": item.unit,
                "Data": item.date,
                "Fonte": item.source,
            }
        )

    return pd.DataFrame(rows)


def get_latest_selic_meta(default_value: float = 10.75) -> float:
    """
    Retorna a Selic meta mais recente.
    Se a API falhar, devolve o valor padrão informado.
    """

    indicator = fetch_latest_sgs_value(
        series_code=432,
        indicator_name="Meta Selic",
        unit="% a.a.",
    )

    if indicator.value is None:
        return default_value

    return indicator.value