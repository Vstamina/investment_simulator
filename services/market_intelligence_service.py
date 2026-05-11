import datetime as dt
import json
import urllib.parse
import urllib.request

import pandas as pd


BACEN_SELIC_META_URL = (
    "https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/"
    "dados/ultimos/1?formato=json"
)

FOCUS_BASE_URL = (
    "https://olinda.bcb.gov.br/olinda/servico/Expectativas/"
    "versao/v1/odata/ExpectativaMercadoAnuais"
)


def _fetch_json(url):
    try:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "application/json",
            },
        )

        with urllib.request.urlopen(request, timeout=25) as response:
            raw_data = response.read().decode("utf-8")

        return json.loads(raw_data)

    except Exception:
        return None


def _safe_float(value):
    try:
        if value is None:
            return None

        if isinstance(value, str):
            value = value.replace(",", ".")

        return float(value)

    except Exception:
        return None


def _current_year():
    return dt.datetime.now().year


def _format_percent(value):
    try:
        return f"{float(value):.2f}%".replace(".", ",")
    except Exception:
        return "não disponível"


def _build_focus_indicator_url(indicator):
    filter_text = f"Indicador eq '{indicator}'"
    encoded_filter = urllib.parse.quote(filter_text, safe="")

    return (
        f"{FOCUS_BASE_URL}?"
        f"$top=100&"
        f"$filter={encoded_filter}&"
        f"$orderby=Data desc&"
        f"$format=json"
    )


# =========================================================
# BACEN — SELIC ATUAL
# =========================================================

def fetch_current_selic_from_bacen():
    data = _fetch_json(BACEN_SELIC_META_URL)

    if not data:
        return None

    try:
        last_item = data[-1]
        value = last_item.get("valor")

        return _safe_float(value)

    except Exception:
        return None


def build_bacen_dataframe(selic_atual):
    rows = []

    if selic_atual is not None:
        rows.append(
            {
                "Indicador": "Selic Meta Atual",
                "Fonte": "Bacen SGS 432",
                "Valor": selic_atual,
                "Unidade": "% a.a.",
                "Data de referência": dt.datetime.now().strftime("%d/%m/%Y"),
            }
        )

    return pd.DataFrame(rows)


# =========================================================
# FOCUS — EXPECTATIVAS
# =========================================================

def fetch_focus_indicator(indicator):
    url = _build_focus_indicator_url(indicator)
    data = _fetch_json(url)

    if not data:
        return pd.DataFrame()

    try:
        values = data.get("value", [])

        if not values:
            return pd.DataFrame()

        df = pd.DataFrame(values)

        required_columns = [
            "Indicador",
            "Data",
            "DataReferencia",
            "Mediana",
        ]

        for column in required_columns:
            if column not in df.columns:
                return pd.DataFrame()

        df["Mediana"] = df["Mediana"].apply(_safe_float)

        df["Ano Referência"] = (
            df["DataReferencia"]
            .astype(str)
            .str.extract(r"(\d{4})")[0]
        )

        df["Ano Referência"] = pd.to_numeric(
            df["Ano Referência"],
            errors="coerce",
        )

        year_now = _current_year()

        df = df[
            df["Ano Referência"].between(
                year_now,
                year_now + 4,
            )
        ].copy()

        if df.empty:
            return pd.DataFrame()

        df["Data"] = pd.to_datetime(
            df["Data"],
            errors="coerce",
        )

        df = df.sort_values(
            by=[
                "Ano Referência",
                "Data",
            ],
            ascending=[
                True,
                False,
            ],
        )

        latest_df = (
            df.dropna(subset=["Ano Referência"])
            .drop_duplicates(
                subset=[
                    "Indicador",
                    "Ano Referência",
                ],
                keep="first",
            )
            .sort_values(
                by=[
                    "Indicador",
                    "Ano Referência",
                ]
            )
        )

        latest_df = latest_df[
            [
                "Indicador",
                "Ano Referência",
                "Mediana",
                "Data",
                "DataReferencia",
            ]
        ].copy()

        latest_df = latest_df.rename(
            columns={
                "Mediana": "Expectativa Focus (%)",
                "Data": "Data da consulta Focus",
                "DataReferencia": "Referência Focus",
            }
        )

        latest_df["Fonte"] = "Bacen Focus"

        return latest_df.reset_index(drop=True)

    except Exception:
        return pd.DataFrame()


def build_focus_fallback():
    year_now = _current_year()
    today = pd.Timestamp.today()

    fallback_rows = [
        {
            "Indicador": "Selic",
            "Ano Referência": year_now,
            "Expectativa Focus (%)": 13.00,
            "Data da consulta Focus": today,
            "Referência Focus": str(year_now),
            "Fonte": "Fallback técnico",
        },
        {
            "Indicador": "Selic",
            "Ano Referência": year_now + 1,
            "Expectativa Focus (%)": 11.25,
            "Data da consulta Focus": today,
            "Referência Focus": str(year_now + 1),
            "Fonte": "Fallback técnico",
        },
        {
            "Indicador": "Selic",
            "Ano Referência": year_now + 2,
            "Expectativa Focus (%)": 10.00,
            "Data da consulta Focus": today,
            "Referência Focus": str(year_now + 2),
            "Fonte": "Fallback técnico",
        },
        {
            "Indicador": "Selic",
            "Ano Referência": year_now + 3,
            "Expectativa Focus (%)": 10.00,
            "Data da consulta Focus": today,
            "Referência Focus": str(year_now + 3),
            "Fonte": "Fallback técnico",
        },
        {
            "Indicador": "Selic",
            "Ano Referência": year_now + 4,
            "Expectativa Focus (%)": 10.00,
            "Data da consulta Focus": today,
            "Referência Focus": str(year_now + 4),
            "Fonte": "Fallback técnico",
        },
        {
            "Indicador": "IPCA",
            "Ano Referência": year_now,
            "Expectativa Focus (%)": 4.50,
            "Data da consulta Focus": today,
            "Referência Focus": str(year_now),
            "Fonte": "Fallback técnico",
        },
    ]

    return pd.DataFrame(fallback_rows)


def fetch_focus_expectations():
    indicators = [
        "Selic",
        "IPCA",
        "Câmbio",
        "PIB Total",
    ]

    frames = []

    for indicator in indicators:
        indicator_df = fetch_focus_indicator(indicator)

        if indicator_df is not None and not indicator_df.empty:
            frames.append(indicator_df)

    if frames:
        return pd.concat(
            frames,
            ignore_index=True,
        )

    return build_focus_fallback()


# =========================================================
# CURVA SIMPLIFICADA DE JUROS
# =========================================================

def build_simplified_interest_curve(focus_df):
    rows = []

    if focus_df is not None and not focus_df.empty:
        selic_focus_df = focus_df[
            focus_df["Indicador"] == "Selic"
        ].copy()

        if not selic_focus_df.empty:
            selic_focus_df = selic_focus_df.sort_values(
                by="Ano Referência"
            )

            for _, row in selic_focus_df.iterrows():
                year = row.get("Ano Referência")
                expected_selic = _safe_float(
                    row.get("Expectativa Focus (%)")
                )

                if expected_selic is None:
                    continue

                try:
                    year_label = int(year)
                except Exception:
                    year_label = str(year)

                rows.append(
                    {
                        "Vértice": str(year_label),
                        "Taxa Selic Esperada (%)": expected_selic,
                        "Fonte": "Focus/Bacen",
                        "Tipo": "Curva simplificada por expectativa",
                    }
                )

    curve_df = pd.DataFrame(rows)

    if curve_df.empty:
        return pd.DataFrame(
            columns=[
                "Vértice",
                "Taxa Selic Esperada (%)",
                "Fonte",
                "Tipo",
            ]
        )

    curve_df = curve_df.drop_duplicates(
        subset=[
            "Vértice",
        ],
        keep="first",
    )

    return curve_df.reset_index(drop=True)


def classify_curve(curve_df, selic_atual=None):
    if curve_df is None or curve_df.empty:
        return "curva não disponível"

    try:
        rates = (
            curve_df["Taxa Selic Esperada (%)"]
            .astype(float)
            .tolist()
        )

        if len(rates) < 2:
            return "curva sem vértices suficientes"

        first_rate = rates[0]
        last_rate = rates[-1]

        if selic_atual is not None:
            reference_rate = float(selic_atual)
        else:
            reference_rate = first_rate

        difference = last_rate - reference_rate
        neutral_threshold = 0.10

        if difference > neutral_threshold:
            return "curva ascendente"

        if difference < -neutral_threshold:
            return "curva descendente"

        return "curva praticamente estável"

    except Exception:
        return "curva não classificada"


def build_curve_movement_reading(curve_df, selic_atual):
    if curve_df is None or curve_df.empty:
        return None, None, None

    try:
        chart_df = curve_df.copy()

        chart_df["Taxa Selic Esperada (%)"] = (
            chart_df["Taxa Selic Esperada (%)"].astype(float)
        )

        if len(chart_df) < 2:
            return (
                "curva sem vértices suficientes",
                0.0,
                "A curva possui apenas um vértice. Para uma leitura prospectiva, "
                "é necessário carregar expectativas futuras de mercado.",
            )

        if selic_atual is None:
            selic_atual = chart_df["Taxa Selic Esperada (%)"].iloc[0]

        taxa_final = chart_df["Taxa Selic Esperada (%)"].iloc[-1]
        spread_final = taxa_final - float(selic_atual)

        limite_neutro = 0.10

        if spread_final > limite_neutro:
            movimento_curva = "curva abrindo"
            leitura_movimento = (
                "A curva está abrindo em relação à Selic atual. "
                "Isso indica que as expectativas de mercado apontam para juros futuros "
                "acima da taxa corrente, o que pode favorecer uma conversa consultiva "
                "sobre proteção de taxa, prazo, prêmio de risco e alternativas prefixadas "
                "ou híbridas, sempre conforme perfil, liquidez e horizonte do cliente."
            )

        elif spread_final < -limite_neutro:
            movimento_curva = "curva fechando"
            leitura_movimento = (
                "A curva está fechando em relação à Selic atual. "
                "Isso indica que as expectativas de mercado apontam para juros futuros "
                "abaixo da taxa corrente. Nesse cenário, produtos pós-fixados seguem "
                "relevantes para liquidez e flexibilidade, mas cresce a importância "
                "de avaliar risco de reinvestimento e alternativas que possam travar "
                "taxas ou proteger poder de compra."
            )

        else:
            movimento_curva = "curva praticamente estável"
            leitura_movimento = (
                "A curva está praticamente estável em relação à Selic atual. "
                "Nesse cenário, a análise consultiva deve priorizar liquidez, prazo, "
                "tributação, previsibilidade, risco de crédito e aderência ao objetivo "
                "financeiro do cliente."
            )

        return movimento_curva, spread_final, leitura_movimento

    except Exception:
        return None, None, None


# =========================================================
# LEITURA FORESIGHT
# =========================================================

def build_market_reading(
    selic_atual,
    focus_df,
    curve_df,
    curve_shape,
    movimento_curva,
    spread_final,
    leitura_movimento,
):
    if curve_df is None or curve_df.empty:
        return (
            "Os dados públicos do Banco Central foram consultados, mas a curva "
            "simplificada de juros não pôde ser construída nesta execução."
        )

    selic_atual_text = _format_percent(selic_atual)

    spread_text = "não disponível"

    if spread_final is not None:
        try:
            spread_text = (
                f"{str(round(float(spread_final), 2)).replace('.', ',')} "
                "ponto percentual"
            )
        except Exception:
            pass

    ipca_text = ""

    if focus_df is not None and not focus_df.empty:
        ipca_df = focus_df[
            focus_df["Indicador"] == "IPCA"
        ].copy()

        if not ipca_df.empty:
            try:
                first_ipca = ipca_df.sort_values(
                    by="Ano Referência"
                ).iloc[0]

                ipca_text = (
                    f"\n\nA expectativa de IPCA também deve ser acompanhada, "
                    f"pois a inflação esperada influencia o espaço para cortes ou "
                    f"manutenção da Selic. Para {int(first_ipca['Ano Referência'])}, "
                    f"a referência utilizada está em "
                    f"{_format_percent(first_ipca['Expectativa Focus (%)'])}."
                )
            except Exception:
                ipca_text = ""

    reading = f"""
Os dados públicos do Banco Central foram carregados para apoiar a leitura macroeconômica da simulação.

A Selic atual utilizada como linha de referência da curva está em {selic_atual_text}. As expectativas de mercado foram utilizadas para os vértices futuros da curva, permitindo observar a direção projetada dos juros ao longo do tempo.

A curva simplificada apresenta configuração de {curve_shape}. Em relação à Selic atual, a leitura indica {movimento_curva}, com spread do último vértice de {spread_text} frente à taxa corrente.

{leitura_movimento}{ipca_text}

Essa leitura ajuda a contextualizar o comportamento relativo entre produtos pós-fixados, prefixados e indexados à inflação. Em ambiente de juros projetados em queda, produtos pós-fixados tendem a preservar flexibilidade e liquidez, mas podem carregar risco de reinvestimento. Em ambiente de juros projetados em alta ou de abertura da curva, a análise deve considerar prêmio de prazo, risco fiscal, inflação esperada, qualidade do emissor e necessidade de liquidez.
"""

    return reading.strip()


# =========================================================
# FUNÇÃO PRINCIPAL CHAMADA PELO APP
# =========================================================

def generate_market_intelligence():
    selic_atual = fetch_current_selic_from_bacen()

    if selic_atual is None:
        selic_atual = 14.50

    bacen_df = build_bacen_dataframe(
        selic_atual
    )

    focus_df = fetch_focus_expectations()

    curve_df = build_simplified_interest_curve(
        focus_df=focus_df
    )

    curve_shape = classify_curve(
        curve_df=curve_df,
        selic_atual=selic_atual,
    )

    movimento_curva, spread_final, leitura_movimento = (
        build_curve_movement_reading(
            curve_df=curve_df,
            selic_atual=selic_atual,
        )
    )

    if movimento_curva is None:
        movimento_curva = "movimento não classificado"

    if leitura_movimento is None:
        leitura_movimento = (
            "Não foi possível gerar uma leitura completa da curva nesta execução."
        )

    reading = build_market_reading(
        selic_atual=selic_atual,
        focus_df=focus_df,
        curve_df=curve_df,
        curve_shape=curve_shape,
        movimento_curva=movimento_curva,
        spread_final=spread_final,
        leitura_movimento=leitura_movimento,
    )

    return {
        "selic_atual": selic_atual,
        "bacen_df": bacen_df,
        "focus_df": focus_df,
        "curve_df": curve_df,
        "curve_shape": curve_shape,
        "reading": reading,
        "movimento_curva": movimento_curva,
        "spread_final": spread_final,
        "leitura_movimento": leitura_movimento,
    }