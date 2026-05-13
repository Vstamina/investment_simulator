import pandas as pd
import streamlit as st
import plotly.graph_objects as go

def aliquota_ir_regressiva(dias: int) -> float:
    if dias <= 180:
        return 0.225
    if dias <= 360:
        return 0.20
    if dias <= 720:
        return 0.175
    return 0.15


def taxa_anual_para_mensal(taxa_anual_pct: float) -> float:
    taxa_anual = taxa_anual_pct / 100
    return (1 + taxa_anual) ** (1 / 12) - 1


def simular_investimento_pela_curva(
    valor_inicial: float,
    curva_df: pd.DataFrame,
    produto_tributavel: bool = True
) -> tuple[pd.DataFrame, dict]:
    """
    Simula aplicação de aporte único usando a curva anual informada.
    Espera curva_df com colunas:
    - ano
    - taxa
    """

    if curva_df.empty:
        return pd.DataFrame(), {}

    curva = curva_df.copy()
    curva = curva.sort_values("ano")

    saldo = valor_inicial
    linhas = []
    mes_global = 0

    for _, row in curva.iterrows():
        ano = int(row["ano"])
        taxa_anual_pct = float(row["taxa"])
        taxa_mensal = taxa_anual_para_mensal(taxa_anual_pct)

        saldo_inicio_ano = saldo

        for mes in range(1, 13):
            mes_global += 1
            saldo = saldo * (1 + taxa_mensal)

            rendimento_bruto = saldo - valor_inicial
            dias_estimados = mes_global * 30

            if produto_tributavel:
                aliquota_ir = aliquota_ir_regressiva(dias_estimados)
                ir_estimado = max(rendimento_bruto, 0) * aliquota_ir
            else:
                aliquota_ir = 0.0
                ir_estimado = 0.0

            saldo_liquido = saldo - ir_estimado

            linhas.append(
                {
                    "ano": ano,
                    "mes": mes,
                    "mes_global": mes_global,
                    "taxa_anual_pct": taxa_anual_pct,
                    "taxa_mensal_pct": taxa_mensal * 100,
                    "saldo_bruto": saldo,
                    "rendimento_bruto": rendimento_bruto,
                    "aliquota_ir": aliquota_ir,
                    "ir_estimado": ir_estimado,
                    "saldo_liquido": saldo_liquido,
                }
            )

        saldo_fim_ano = saldo

    resultado_df = pd.DataFrame(linhas)

    if resultado_df.empty:
        return resultado_df, {}

    valor_bruto_final = resultado_df["saldo_bruto"].iloc[-1]
    valor_liquido_final = resultado_df["saldo_liquido"].iloc[-1]
    ir_total = resultado_df["ir_estimado"].iloc[-1]

    meses = len(resultado_df)
    anos = meses / 12

    retorno_bruto = valor_bruto_final / valor_inicial - 1
    retorno_liquido = valor_liquido_final / valor_inicial - 1

    retorno_mensal_liquido = (1 + retorno_liquido) ** (1 / meses) - 1
    retorno_anual_liquido = (1 + retorno_liquido) ** (1 / anos) - 1

    metricas = {
        "valor_inicial": valor_inicial,
        "valor_bruto_final": valor_bruto_final,
        "valor_liquido_final": valor_liquido_final,
        "ir_total": ir_total,
        "retorno_bruto": retorno_bruto,
        "retorno_liquido": retorno_liquido,
        "retorno_mensal_liquido": retorno_mensal_liquido,
        "retorno_anual_liquido": retorno_anual_liquido,
        "meses": meses,
        "anos": anos,
    }

    return resultado_df, metricas


def grafico_investimento_pela_curva(resultado_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=resultado_df["mes_global"],
            y=resultado_df["saldo_bruto"],
            mode="lines",
            name="Valor bruto",
            line=dict(width=3),
        )
    )

    fig.add_trace(
        go.Scatter(
            x=resultado_df["mes_global"],
            y=resultado_df["saldo_liquido"],
            mode="lines",
            name="Valor líquido",
            line=dict(width=3),
        )
    )

    fig.update_layout(
        title="Projeção do investimento pela curva",
        height=480,
        hovermode="x unified",
        xaxis_title="Mês",
        yaxis_title="Valor projetado",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
    )

    return fig


def formatar_moeda(valor: float) -> str:
    return f"R$ {valor:,.2f}"


def formatar_pct(valor: float) -> str:
    return f"{valor * 100:.2f}%"


def render_curve_investment_module(curva_df: pd.DataFrame):
    st.markdown("### Simular investimento pela curva")

    st.info(
        "Este módulo projeta um aporte único com base nos vértices da curva de juros. "
        "A taxa anual de cada ano é convertida para taxa mensal equivalente, permitindo "
        "estimar valor bruto, imposto de renda e valor líquido ao longo do tempo."
    )

    col1, col2 = st.columns(2)

    with col1:
        valor_inicial = st.number_input(
            "Valor inicial investido",
            min_value=0.0,
            value=100000.0,
            step=1000.0,
            key="curva_valor_inicial"
        )

    with col2:
        produto_tributavel = st.selectbox(
            "Classificação fiscal",
            ["Tributável", "Isento"],
            key="curva_classificacao_fiscal"
        ) == "Tributável"

    resultado_df, metricas = simular_investimento_pela_curva(
        valor_inicial=valor_inicial,
        curva_df=curva_df,
        produto_tributavel=produto_tributavel
    )

    if resultado_df.empty:
        st.warning("Não foi possível simular o investimento pela curva.")
        return

    col_a, col_b, col_c, col_d = st.columns(4)

    col_a.metric(
        "Valor bruto final",
        formatar_moeda(metricas["valor_bruto_final"])
    )

    col_b.metric(
        "Valor líquido final",
        formatar_moeda(metricas["valor_liquido_final"])
    )

    col_c.metric(
        "IR estimado",
        formatar_moeda(metricas["ir_total"])
    )

    col_d.metric(
        "Retorno líquido acumulado",
        formatar_pct(metricas["retorno_liquido"])
    )

    col_e, col_f, col_g = st.columns(3)

    col_e.metric(
        "Retorno bruto acumulado",
        formatar_pct(metricas["retorno_bruto"])
    )

    col_f.metric(
        "Retorno líquido ao mês",
        formatar_pct(metricas["retorno_mensal_liquido"])
    )

    col_g.metric(
        "Retorno líquido ao ano",
        formatar_pct(metricas["retorno_anual_liquido"])
    )

    st.plotly_chart(
        grafico_investimento_pela_curva(resultado_df),
        use_container_width=True
    )

    with st.expander("Ver tabela mensal da simulação pela curva"):
        tabela = resultado_df.copy()
        tabela["taxa_anual_pct"] = tabela["taxa_anual_pct"].map(lambda x: f"{x:.2f}%")
        tabela["taxa_mensal_pct"] = tabela["taxa_mensal_pct"].map(lambda x: f"{x:.4f}%")
        tabela["saldo_bruto"] = tabela["saldo_bruto"].map(formatar_moeda)
        tabela["rendimento_bruto"] = tabela["rendimento_bruto"].map(formatar_moeda)
        tabela["ir_estimado"] = tabela["ir_estimado"].map(formatar_moeda)
        tabela["saldo_liquido"] = tabela["saldo_liquido"].map(formatar_moeda)
        tabela["aliquota_ir"] = tabela["aliquota_ir"].map(lambda x: f"{x * 100:.1f}%")

        st.dataframe(
            tabela[
                [
                    "ano",
                    "mes",
                    "taxa_anual_pct",
                    "taxa_mensal_pct",
                    "saldo_bruto",
                    "rendimento_bruto",
                    "aliquota_ir",
                    "ir_estimado",
                    "saldo_liquido",
                ]
            ],
            use_container_width=True,
            hide_index=True
        )

def render_curve_investment_module(curva_df: pd.DataFrame):
    st.markdown("### Simular investimento pela curva")

    st.info(
        "Este módulo projeta um aporte único com base nos vértices da curva de juros. "
        "A taxa anual de cada ano é convertida para taxa mensal equivalente, permitindo "
        "estimar valor bruto, imposto de renda e valor líquido ao longo do tempo."
    )

    col1, col2 = st.columns(2)

    with col1:
        valor_inicial = st.number_input(
            "Valor inicial investido",
            min_value=0.0,
            value=100000.0,
            step=1000.0,
            key="curva_valor_inicial"
        )

    with col2:
        produto_tributavel = st.selectbox(
            "Classificação fiscal",
            ["Tributável", "Isento"],
            key="curva_classificacao_fiscal"
        ) == "Tributável"

    resultado_df, metricas = simular_investimento_pela_curva(
        valor_inicial=valor_inicial,
        curva_df=curva_df,
        produto_tributavel=produto_tributavel
    )

    if resultado_df.empty:
        st.warning("Não foi possível simular o investimento pela curva.")
        return

    col_a, col_b, col_c, col_d = st.columns(4)

    col_a.metric(
        "Valor bruto final",
        formatar_moeda(metricas["valor_bruto_final"])
    )

    col_b.metric(
        "Valor líquido final",
        formatar_moeda(metricas["valor_liquido_final"])
    )

    col_c.metric(
        "IR estimado",
        formatar_moeda(metricas["ir_total"])
    )

    col_d.metric(
        "Retorno líquido acumulado",
        formatar_pct(metricas["retorno_liquido"])
    )

    col_e, col_f, col_g = st.columns(3)

    col_e.metric(
        "Retorno bruto acumulado",
        formatar_pct(metricas["retorno_bruto"])
    )

    col_f.metric(
        "Retorno líquido ao mês",
        formatar_pct(metricas["retorno_mensal_liquido"])
    )

    col_g.metric(
        "Retorno líquido ao ano",
        formatar_pct(metricas["retorno_anual_liquido"])
    )

    st.plotly_chart(
        grafico_investimento_pela_curva(resultado_df),
        use_container_width=True
    )

    with st.expander("Ver tabela mensal da simulação pela curva"):
        tabela = resultado_df.copy()

        tabela["taxa_anual_pct"] = tabela["taxa_anual_pct"].map(lambda x: f"{x:.2f}%")
        tabela["taxa_mensal_pct"] = tabela["taxa_mensal_pct"].map(lambda x: f"{x:.4f}%")
        tabela["saldo_bruto"] = tabela["saldo_bruto"].map(formatar_moeda)
        tabela["rendimento_bruto"] = tabela["rendimento_bruto"].map(formatar_moeda)
        tabela["ir_estimado"] = tabela["ir_estimado"].map(formatar_moeda)
        tabela["saldo_liquido"] = tabela["saldo_liquido"].map(formatar_moeda)
        tabela["aliquota_ir"] = tabela["aliquota_ir"].map(lambda x: f"{x * 100:.1f}%")

        st.dataframe(
            tabela[
                [
                    "ano",
                    "mes",
                    "taxa_anual_pct",
                    "taxa_mensal_pct",
                    "saldo_bruto",
                    "rendimento_bruto",
                    "aliquota_ir",
                    "ir_estimado",
                    "saldo_liquido",
                ]
            ],
            use_container_width=True,
            hide_index=True
        )