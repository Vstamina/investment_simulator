from __future__ import annotations

import pandas as pd


def generate_foresight_reading(
    bacen_df: pd.DataFrame,
    focus_df: pd.DataFrame,
    curve_df: pd.DataFrame,
    curve_shape: str,
) -> str:
    """
    Gera leitura consultiva de mercado com linguagem segura.
    Não emite recomendação individualizada.
    """

    if bacen_df is None or bacen_df.empty:
        bacen_text = (
            "Os dados do Bacen não foram carregados neste momento. "
            "A leitura deve ser feita com base nas premissas manuais da simulação."
        )
    else:
        bacen_text = (
            "Os dados públicos do Banco Central foram carregados para apoiar "
            "a leitura macroeconômica da simulação."
        )

    if focus_df is None or focus_df.empty:
        focus_text = (
            "As expectativas Focus não foram carregadas neste momento. "
            "A análise prospectiva fica limitada aos parâmetros informados manualmente."
        )
    else:
        focus_text = (
            "As expectativas de mercado do Focus foram utilizadas como referência "
            "para observar possíveis direções de juros e inflação."
        )

    if curve_df is None or curve_df.empty:
        curve_text = (
            "Não foi possível montar uma curva simplificada de juros nesta execução."
        )
    else:
        curve_text = (
            f"A curva simplificada construída a partir das expectativas de Selic "
            f"apresenta configuração de {curve_shape}. Essa leitura ajuda a contextualizar "
            "o comportamento relativo entre pós-fixados, prefixados e indexados à inflação."
        )

    if curve_shape == "curva descendente":
        strategic_text = (
            "Em um ambiente de curva descendente, produtos pós-fixados tendem a preservar "
            "flexibilidade, enquanto alternativas prefixadas podem ganhar relevância na conversa "
            "consultiva, desde que sejam avaliados prazo, liquidez e perfil do cliente."
        )

    elif curve_shape == "curva ascendente":
        strategic_text = (
            "Em um ambiente de curva ascendente, a leitura sugere cautela com travamento "
            "prematuro de taxas, mantendo atenção à liquidez, ao reinvestimento e à sensibilidade "
            "dos produtos aos movimentos futuros de juros."
        )

    elif curve_shape == "curva relativamente estável":
        strategic_text = (
            "Em um ambiente de curva relativamente estável, a comparação entre produtos passa "
            "a depender mais de liquidez, tributação, prazo e percentual do indexador."
        )

    else:
        strategic_text = (
            "Sem curva suficiente, a leitura deve priorizar a comparação objetiva entre rendimento "
            "líquido, tributação, liquidez e prazo da simulação."
        )

    questions = (
        "Perguntas consultivas sugeridas: qual é a necessidade de liquidez do cliente? "
        "O horizonte de investimento é compatível com o prazo simulado? "
        "Há preferência por previsibilidade, flexibilidade ou proteção contra inflação? "
        "O cliente aceita risco de oportunidade em caso de mudança relevante na trajetória dos juros?"
    )

    return "\n\n".join(
        [
            bacen_text,
            focus_text,
            curve_text,
            strategic_text,
            questions,
        ]
    )