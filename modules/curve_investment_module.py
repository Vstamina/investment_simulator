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


def formatar_moeda(valor: float) -> str:
    return f"R$ {valor:,.2f}"


def formatar_pct(valor: float) -> str:
    return f"{valor * 100:.2f}%"


def simular_investimento_pela_curva(
    valor_inicial: float,
    curva_df: pd.DataFrame,
    produto_tributavel: bool = True,
    prazo_meses: int | None = None,
    percentual_cdi: float = 100.0,
    taxa_custo_anual: float = 0.0,
) -> tuple[pd.DataFrame, dict]:

    if curva_df is None or curva_df.empty:
        return pd.DataFrame(), {}

    curva = curva_df.copy()
    curva = curva.sort_values("ano")

    saldo = valor_inicial
    linhas = []
    mes_global = 0

    for _, row in curva.iterrows():
        ano = int(row["ano"])

        taxa_anual_pct = float(row["taxa"]) * (percentual_cdi / 100)
        taxa_anual_pct = taxa_anual_pct - taxa_custo_anual

        taxa_mensal = taxa_anual_para_mensal(taxa_anual_pct)

        for mes in range(1, 13):
            if prazo_meses is not None and mes_global >= prazo_meses:
                break

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

        if prazo_meses is not None and mes_global >= prazo_meses:
            break

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
            x=1,
        ),
    )

    return fig


def simular_produtos_pela_curva(
    valor_inicial: float,
    curva_df: pd.DataFrame,
    prazo_meses: int,
    cdb_percentual: float,
    lci_lca_percentual: float,
    tesouro_percentual: float,
    tesouro_custo_anual: float,
    fundo_percentual: float,
    fundo_taxa_anual: float,
) -> pd.DataFrame:

    produtos = [
        {
            "Produto": "CDB / LC",
            "Percentual": cdb_percentual,
            "Tributável": True,
            "Custo anual": 0.0,
        },
        {
            "Produto": "LCI / LCA",
            "Percentual": lci_lca_percentual,
            "Tributável": False,
            "Custo anual": 0.0,
        },
        {
            "Produto": "Tesouro Selic",
            "Percentual": tesouro_percentual,
            "Tributável": True,
            "Custo anual": tesouro_custo_anual,
        },
        {
            "Produto": "Fundo DI",
            "Percentual": fundo_percentual,
            "Tributável": True,
            "Custo anual": fundo_taxa_anual,
        },
    ]

    linhas = []

    for produto in produtos:
        _, metricas = simular_investimento_pela_curva(
            valor_inicial=valor_inicial,
            curva_df=curva_df,
            produto_tributavel=produto["Tributável"],
            prazo_meses=prazo_meses,
            percentual_cdi=produto["Percentual"],
            taxa_custo_anual=produto["Custo anual"],
        )

        if metricas:
            linhas.append(
                {
                    "Produto": produto["Produto"],
                    "% da curva/CDI": produto["Percentual"],
                    "Tributação": "Tributável" if produto["Tributável"] else "Isento",
                    "Custo anual": produto["Custo anual"],
                    "Valor bruto final": metricas["valor_bruto_final"],
                    "IR estimado": metricas["ir_total"],
                    "Valor líquido final": metricas["valor_liquido_final"],
                    "Retorno líquido": metricas["retorno_liquido"],
                    "Retorno líquido ao ano": metricas["retorno_anual_liquido"],
                }
            )

    return pd.DataFrame(linhas)


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
            key="curva_valor_inicial",
        )

    with col2:
        prazo_meses = st.slider(
            "Prazo da simulação",
            min_value=12,
            max_value=60,
            value=60,
            step=12,
            format="%d meses",
            key="curva_prazo_meses",
        )

    col3, col4 = st.columns(2)

    with col3:
        produto_tributavel = st.selectbox(
            "Classificação fiscal da simulação individual",
            ["Tributável", "Isento"],
            key="curva_classificacao_fiscal",
        ) == "Tributável"

    with col4:
        percentual_cdi = st.number_input(
            "Rentabilidade estimada da simulação individual (% da curva/CDI)",
            min_value=0.0,
            max_value=250.0,
            value=100.0,
            step=1.0,
            key="curva_percentual_cdi",
        )

    resultado_df, metricas = simular_investimento_pela_curva(
        valor_inicial=valor_inicial,
        curva_df=curva_df,
        produto_tributavel=produto_tributavel,
        prazo_meses=prazo_meses,
        percentual_cdi=percentual_cdi,
    )

    if resultado_df.empty:
        st.warning("Não foi possível simular o investimento pela curva.")
        return

    st.markdown("#### Resultado da simulação individual")

    col_a, col_b, col_c, col_d = st.columns(4)

    col_a.metric(
        "Valor bruto final",
        formatar_moeda(metricas["valor_bruto_final"]),
    )

    col_b.metric(
        "Valor líquido final",
        formatar_moeda(metricas["valor_liquido_final"]),
    )

    col_c.metric(
        "IR estimado",
        formatar_moeda(metricas["ir_total"]),
    )

    col_d.metric(
        "Retorno líquido acumulado",
        formatar_pct(metricas["retorno_liquido"]),
    )

    col_e, col_f, col_g = st.columns(3)

    col_e.metric(
        "Retorno bruto acumulado",
        formatar_pct(metricas["retorno_bruto"]),
    )

    col_f.metric(
        "Retorno líquido ao mês",
        formatar_pct(metricas["retorno_mensal_liquido"]),
    )

    col_g.metric(
        "Retorno líquido ao ano",
        formatar_pct(metricas["retorno_anual_liquido"]),
    )

    st.plotly_chart(
        grafico_investimento_pela_curva(resultado_df),
        width="stretch",
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
            width="stretch",
            hide_index=True,
        )

    st.markdown("#### Comparativo de produtos pela curva")

    st.caption(
        "Comparação estimada entre produtos usando a mesma curva de juros, "
        "com diferenças de tributação, percentual da curva/CDI e custos anuais."
    )

    colp1, colp2 = st.columns(2)

    with colp1:
        cdb_percentual = st.number_input(
            "CDB / LC (% da curva/CDI)",
            min_value=0.0,
            max_value=250.0,
            value=105.0,
            step=1.0,
            key="curva_cdb_percentual",
        )

        lci_lca_percentual = st.number_input(
            "LCI / LCA (% da curva/CDI)",
            min_value=0.0,
            max_value=250.0,
            value=93.0,
            step=1.0,
            key="curva_lci_lca_percentual",
        )

    with colp2:
        tesouro_percentual = st.number_input(
            "Tesouro Selic (% da curva/CDI)",
            min_value=0.0,
            max_value=250.0,
            value=100.0,
            step=1.0,
            key="curva_tesouro_percentual",
        )

        tesouro_custo_anual = st.number_input(
            "Custo anual Tesouro Selic (%)",
            min_value=0.0,
            max_value=5.0,
            value=0.20,
            step=0.05,
            key="curva_tesouro_custo_anual",
        )

    colp3, colp4 = st.columns(2)

    with colp3:
        fundo_percentual = st.number_input(
            "Fundo DI (% da curva/CDI)",
            min_value=0.0,
            max_value=250.0,
            value=100.0,
            step=1.0,
            key="curva_fundo_percentual",
        )

    with colp4:
        fundo_taxa_anual = st.number_input(
            "Taxa de administração Fundo DI (%)",
            min_value=0.0,
            max_value=5.0,
            value=0.50,
            step=0.05,
            key="curva_fundo_taxa_anual",
        )

    comparativo_df = simular_produtos_pela_curva(
        valor_inicial=valor_inicial,
        curva_df=curva_df,
        prazo_meses=prazo_meses,
        cdb_percentual=cdb_percentual,
        lci_lca_percentual=lci_lca_percentual,
        tesouro_percentual=tesouro_percentual,
        tesouro_custo_anual=tesouro_custo_anual,
        fundo_percentual=fundo_percentual,
        fundo_taxa_anual=fundo_taxa_anual,
    )

    if not comparativo_df.empty:
        comparativo_df = comparativo_df.sort_values(
            "Valor líquido final",
            ascending=False,
        )

        vencedor = comparativo_df.iloc[0]["Produto"]

        st.success(
            f"Na projeção pela curva, o produto com maior valor líquido final foi: {vencedor}."
        )

        tabela_comparativa = comparativo_df.copy()
        tabela_comparativa["% da curva/CDI"] = tabela_comparativa["% da curva/CDI"].map(
            lambda x: f"{x:.2f}%"
        )
        tabela_comparativa["Custo anual"] = tabela_comparativa["Custo anual"].map(
            lambda x: f"{x:.2f}%"
        )
        tabela_comparativa["Valor bruto final"] = tabela_comparativa[
            "Valor bruto final"
        ].map(formatar_moeda)
        tabela_comparativa["IR estimado"] = tabela_comparativa["IR estimado"].map(
            formatar_moeda
        )
        tabela_comparativa["Valor líquido final"] = tabela_comparativa[
            "Valor líquido final"
        ].map(formatar_moeda)
        tabela_comparativa["Retorno líquido"] = tabela_comparativa[
            "Retorno líquido"
        ].map(formatar_pct)
        tabela_comparativa["Retorno líquido ao ano"] = tabela_comparativa[
            "Retorno líquido ao ano"
        ].map(formatar_pct)

        st.dataframe(
            tabela_comparativa,
            width="stretch",
            hide_index=True,
        )