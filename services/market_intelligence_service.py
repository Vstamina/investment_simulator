from __future__ import annotations

import pandas as pd

from market_data.bacen_service import get_bacen_snapshot
from market_data.focus_service import get_focus_snapshot
from market_data.yield_curve_service import (
    build_simple_focus_curve,
    classify_curve_shape,
)
from market_data.foresight_service import generate_foresight_reading


def generate_market_intelligence() -> dict:
    """
    Consolida dados públicos de mercado e gera uma leitura consultiva
    baseada em Bacen, Focus, curva simplificada de juros e foresight.

    A função é resiliente: se alguma fonte falhar, o app continua rodando.
    """

    try:
        bacen_df = get_bacen_snapshot()
    except Exception:
        bacen_df = pd.DataFrame()

    try:
        focus_df = get_focus_snapshot()
    except Exception:
        focus_df = pd.DataFrame()

    try:
        curve_df = build_simple_focus_curve(focus_df)
        curve_shape = classify_curve_shape(curve_df)
    except Exception:
        curve_df = pd.DataFrame()
        curve_shape = "sem dados suficientes"

    try:
        reading = generate_foresight_reading(
            bacen_df=bacen_df,
            focus_df=focus_df,
            curve_df=curve_df,
            curve_shape=curve_shape,
        )
    except Exception:
        reading = (
            "Não foi possível gerar a leitura de inteligência de mercado nesta execução. "
            "A simulação principal permanece disponível com base nas premissas informadas."
        )

    return {
        "bacen_df": bacen_df,
        "focus_df": focus_df,
        "curve_df": curve_df,
        "curve_shape": curve_shape,
        "reading": reading,
    }