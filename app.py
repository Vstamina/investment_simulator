from datetime import date
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from services.simulation_service import (
    run_cdi_simulation,
    run_cdi_cashflow_simulation,
)
from services.interpretation_service import generate_consultive_analysis
from reports.word_report_generator import generate_word_report
from services.market_intelligence_service import generate_market_intelligence


st.set_page_config(
    page_title="Simulador de Investimentos",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# FUNÇÕES DE FORMATAÇÃO
# =========================================================

def format_currency(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_percent(value: float) -> str:
    return f"{value:.2f}%".replace(".", ",")


def format_compact_currency(value: float) -> str:
    """
    Formata valores grandes para leitura executiva.
    Exemplo:
    500000 -> R$ 500 mil
    1000000 -> R$ 1,00 mi
    100000000 -> R$ 100,00 mi
    """
    if value >= 1_000_000_000:
        return f"R$ {value / 1_000_000_000:.2f} bi".replace(".", ",")

    if value >= 1_000_000:
        return f"R$ {value / 1_000_000:.2f} mi".replace(".", ",")

    if value >= 1_000:
        return f"R$ {value / 1_000:.0f} mil".replace(".", ",")

    return format_currency(value)


def build_cashflows_from_editor(cashflow_df: pd.DataFrame) -> list[dict]:
    cashflows = []

    if cashflow_df.empty:
        return cashflows

    for _, row in cashflow_df.iterrows():
        if pd.isna(row["Data"]) or pd.isna(row["Valor"]):
            continue

        value = float(row["Valor"])

        if value <= 0:
            continue

        cashflows.append(
            {
                "data": row["Data"],
                "tipo": row["Tipo"],
                "valor": value,
                "descricao": row.get("Descrição", ""),
            }
        )

    return cashflows


def metric_card(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card-custom">
            <div class="metric-label-custom">{label}</div>
            <div class="metric-value-custom">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>
        .main {
            background-color: #f5f8fc;
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }

        .title-box {
            background: linear-gradient(135deg, #061b3a 0%, #083d77 55%, #0a84ff 100%);
            padding: 28px;
            border-radius: 18px;
            margin-bottom: 24px;
            color: white;
        }

        .title-box h1 {
            margin: 0;
            font-size: 32px;
            font-weight: 800;
            letter-spacing: -0.5px;
        }

        .title-box p {
            margin-top: 8px;
            font-size: 15px;
            opacity: 0.92;
        }

        .section-title {
            color: #062b5f;
            font-size: 21px;
            font-weight: 750;
            margin-top: 22px;
            margin-bottom: 10px;
        }

        .support-text {
            color: #516173;
            font-size: 14px;
            margin-bottom: 12px;
        }

        .metric-card-custom {
            background: rgba(255,255,255,0.02);
            border: 1px solid rgba(255,255,255,0.08);
            padding: 12px 14px;
            border-radius: 12px;
            min-height: 92px;
            overflow: hidden;
        }

        .metric-label-custom {
            font-size: 0.86rem;
            color: #cfd8e3;
            margin-bottom: 8px;
            font-weight: 600;
        }

        .metric-value-custom {
            font-size: 1.75rem;
            line-height: 1.12;
            color: white;
            font-weight: 600;
            word-break: break-word;
            overflow-wrap: anywhere;
        }

        div[data-testid="stMetricValue"] {
            font-size: 1.75rem !important;
            line-height: 1.1 !important;
            white-space: normal !important;
            overflow-wrap: anywhere !important;
        }

        div[data-testid="stMetricLabel"] {
            font-size: 0.90rem !important;
        }

        div[data-testid="stMetric"] {
            background-color: rgba(255,255,255,0.02);
            padding: 10px 12px;
            border-radius: 10px;
        }

        .warning-box {
            background-color: #fff7e6;
            border-left: 5px solid #f5a623;
            padding: 16px;
            border-radius: 10px;
            color: #4a3b16;
            margin-top: 18px;
            font-size: 14px;
        }

        .info-box {
            background-color: #eef6ff;
            border-left: 5px solid #0a84ff;
            padding: 14px;
            border-radius: 10px;
            color: #0b2447;
            font-size: 14px;
            margin-bottom: 12px;
        }

        /* Ajuste visual das tabelas */
        div[data-testid="stDataFrame"] {
            font-size: 0.86rem;
        }

        /* Sidebar */
        section[data-testid="stSidebar"] {
            background-color: #20222b;
        }

        section[data-testid="stSidebar"] label {
            font-size: 0.86rem !important;
            font-weight: 600 !important;
        }

        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {
            color: white;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# CABEÇALHO
# =========================================================

st.markdown(
    """
    <div class="title-box">
        <h1>Simulador Estratégico de Investimentos</h1>
        <p>MVP CDI | Comparação de renda fixa, calendário de aportes, valor líquido, tributação e leitura consultiva</p>
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# SIDEBAR — INPUTS
# =========================================================

with st.sidebar:
    st.header("Dados da Simulação")

    client_name = st.text_input(
        "Nome do cliente",
        value="Cliente Exemplo"
    )

    advisor_name = st.text_input(
        "Assessor / Banker",
        value="Nome do Assessor"
    )

    st.divider()

    initial_amount = st.number_input(
        "Valor inicial investido",
        min_value=0.0,
        value=500000.0,
        step=10000.0,
        format="%.2f"
    )

    simulation_mode = st.radio(
        "Modo de movimentação",
        [
            "Sem aportes adicionais",
            "Aporte mensal fixo",
            "Aportes e resgates por calendário",
        ],
        index=2
    )

    monthly_contribution = 0.0
    months = 24

    if simulation_mode == "Aporte mensal fixo":
        monthly_contribution = st.number_input(
            "Aporte mensal",
            min_value=0.0,
            value=0.0,
            step=1000.0,
            format="%.2f"
        )

        months = st.number_input(
            "Prazo da simulação em meses",
            min_value=1,
            max_value=360,
            value=24,
            step=1
        )

    elif simulation_mode == "Sem aportes adicionais":
        months = st.number_input(
            "Prazo da simulação em meses",
            min_value=1,
            max_value=360,
            value=24,
            step=1
        )

    st.divider()

    st.subheader("Período")

    start_date = st.date_input(
        "Data inicial",
        value=date(2026, 1, 1)
    )

    end_date = st.date_input(
        "Data final",
        value=date(2026, 12, 31)
    )

    st.divider()

    st.subheader("Premissas Econômicas")

    annual_cdi_rate = st.number_input(
        "CDI anual estimado (%)",
        min_value=0.0,
        value=10.65,
        step=0.05,
        format="%.2f"
    )

    selic_rate = st.number_input(
        "Selic meta estimada (%)",
        min_value=0.0,
        value=10.75,
        step=0.05,
        format="%.2f"
    )

    tr_rate = st.number_input(
        "TR anual estimada (%)",
        min_value=0.0,
        value=0.0,
        step=0.05,
        format="%.2f"
    )

    st.divider()

    st.subheader("Inteligência de Mercado")

    use_market_intelligence = st.checkbox(
        "Ativar leitura Bacen/Focus/Foresight",
        value=False
    )

    st.divider()

    st.subheader("Parâmetros por Produto")

    cdb_percentage = st.number_input(
        "CDB / LC (% do CDI)",
        min_value=0.0,
        value=105.0,
        step=1.0,
        format="%.2f"
    )

    lci_lca_percentage = st.number_input(
        "LCI / LCA (% do CDI)",
        min_value=0.0,
        value=92.0,
        step=1.0,
        format="%.2f"
    )

    treasury_percentage = st.number_input(
        "Tesouro Selic (% do CDI aproximado)",
        min_value=0.0,
        value=100.0,
        step=1.0,
        format="%.2f"
    )

    treasury_annual_fee = st.number_input(
        "Taxa/custo anual Tesouro Selic (%)",
        min_value=0.0,
        value=0.20,
        step=0.05,
        format="%.2f"
    )

    fund_percentage = st.number_input(
        "Fundo DI (% do CDI)",
        min_value=0.0,
        value=100.0,
        step=1.0,
        format="%.2f"
    )

    fund_annual_fee = st.number_input(
        "Taxa de administração Fundo DI (%)",
        min_value=0.0,
        value=0.50,
        step=0.05,
        format="%.2f"
    )


# =========================================================
# CALENDÁRIO DE MOVIMENTAÇÕES
# =========================================================

cashflows = []
cashflow_editor_df = pd.DataFrame()

if simulation_mode == "Aportes e resgates por calendário":
    st.markdown(
        '<div class="section-title">Calendário de Movimentações</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="info-box">
            Informe aportes e resgates programados. O sistema recalcula a projeção considerando as datas escolhidas.
        </div>
        """,
        unsafe_allow_html=True
    )

    default_cashflows = pd.DataFrame(
        [
            {
                "Data": date(2026, 3, 15),
                "Tipo": "Aporte",
                "Valor": 80000.0,
                "Descrição": "Aporte programado",
            },
            {
                "Data": date(2026, 6, 10),
                "Tipo": "Aporte",
                "Valor": 120000.0,
                "Descrição": "Entrada prevista",
            },
            {
                "Data": date(2026, 9, 20),
                "Tipo": "Resgate",
                "Valor": 50000.0,
                "Descrição": "Necessidade de liquidez",
            },
        ]
    )

    cashflow_editor_df = st.data_editor(
        default_cashflows,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Data": st.column_config.DateColumn(
                "Data",
                format="DD/MM/YYYY",
            ),
            "Tipo": st.column_config.SelectboxColumn(
                "Tipo",
                options=["Aporte", "Resgate"],
                required=True,
            ),
            "Valor": st.column_config.NumberColumn(
                "Valor",
                min_value=0.0,
                step=1000.0,
                format="R$ %.2f",
            ),
            "Descrição": st.column_config.TextColumn(
                "Descrição"
            ),
        },
    )

    cashflows = build_cashflows_from_editor(cashflow_editor_df)


# =========================================================
# VALIDAÇÃO DE PERÍODO
# =========================================================

if end_date <= start_date:
    st.error("A data final precisa ser posterior à data inicial.")
    st.stop()


# =========================================================
# EXECUÇÃO DA SIMULAÇÃO
# =========================================================

if simulation_mode == "Aportes e resgates por calendário":
    comparison_df, daily_df, monthly_df = run_cdi_cashflow_simulation(
        initial_amount=initial_amount,
        start_date=start_date,
        end_date=end_date,
        annual_cdi_rate=annual_cdi_rate,
        selic_rate=selic_rate,
        tr_rate=tr_rate,
        cdb_percentage=cdb_percentage,
        lci_lca_percentage=lci_lca_percentage,
        treasury_percentage=treasury_percentage,
        fund_percentage=fund_percentage,
        fund_annual_fee=fund_annual_fee,
        treasury_annual_fee=treasury_annual_fee,
        cashflows=cashflows,
    )

    evolution_df = daily_df.copy()
    evolution_x = "Data"

else:
    if simulation_mode == "Sem aportes adicionais":
        monthly_contribution = 0.0

    comparison_df, evolution_df = run_cdi_simulation(
        initial_amount=initial_amount,
        monthly_contribution=monthly_contribution,
        months=int(months),
        annual_cdi_rate=annual_cdi_rate,
        selic_rate=selic_rate,
        tr_rate=tr_rate,
        cdb_percentage=cdb_percentage,
        lci_lca_percentage=lci_lca_percentage,
        treasury_percentage=treasury_percentage,
        fund_percentage=fund_percentage,
        fund_annual_fee=fund_annual_fee,
        treasury_annual_fee=treasury_annual_fee,
    )

    daily_df = pd.DataFrame()
    monthly_df = pd.DataFrame()
    evolution_x = "Mês"


# =========================================================
# CARDS DE RESUMO
# =========================================================

best_product = comparison_df.iloc[0]

col1, col2, col3, col4 = st.columns(4)

with col1:
    metric_card(
        "Cliente",
        client_name
    )

with col2:
    if simulation_mode == "Aportes e resgates por calendário":
        period_text = f"{start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}"
        metric_card(
            "Período",
            period_text
        )
    else:
        metric_card(
            "Prazo",
            f"{months} meses"
        )

with col3:
    metric_card(
        "Melhor alternativa",
        str(best_product["Produto"])
    )

with col4:
    metric_card(
        "Valor líquido projetado",
        format_currency(best_product["Valor Líquido"])
    )


# =========================================================
# TABELA COMPARATIVA
# =========================================================

st.markdown(
    '<div class="section-title">Tabela Comparativa — Módulo CDI</div>',
    unsafe_allow_html=True
)

# Copia a base calculada para criar a versão formatada da tabela
display_df = comparison_df.copy()

# Colunas financeiras que serão exibidas em reais
currency_columns = [
    "Valor Inicial",
    "Valor Investido",
    "Total Aportado",
    "Total Resgatado",
    "Valor Bruto",
    "Rendimento Bruto",
    "IR",
    "Valor Líquido",
    "Rendimento Líquido",
]

# Colunas percentuais que serão exibidas com %
percent_columns = [
    "% CDI",
    "Taxa Efetiva a.a. (%)",
    "Alíquota IR (%)",
    "Rentab. Líq. Período (%)",
    "Rentab. Líq. Mês (%)",
    "Rentab. Líq. Ano (%)",
]

# Formata valores monetários
for col in currency_columns:
    if col in display_df.columns:
        display_df[col] = display_df[col].apply(format_currency)

# Formata percentuais
for col in percent_columns:
    if col in display_df.columns:
        display_df[col] = display_df[col].apply(format_percent)

# Tabela executiva: versão mais limpa para visualização principal
executive_columns = [
    "Produto",
    "Taxa Efetiva a.a. (%)",
    "Valor Inicial",
    "Valor Investido",
    "Total Aportado",
    "Total Resgatado",
    "Valor Líquido",
    "Rendimento Líquido",
    "Rentab. Líq. Período (%)",
    "Rentab. Líq. Ano (%)",
    "Tributável",
]

available_executive_columns = [
    column for column in executive_columns if column in display_df.columns
]

executive_df = display_df[available_executive_columns]

st.dataframe(
    executive_df,
    use_container_width=True,
    hide_index=True
)

# Tabela completa: fica disponível para conferência técnica
with st.expander("Ver tabela técnica completa"):
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )
    

# =========================================================
# EVOLUÇÃO PATRIMONIAL
# =========================================================

st.markdown(
    '<div class="section-title">Evolução Patrimonial Simulada</div>',
    unsafe_allow_html=True
)

fig = px.line(
    evolution_df,
    x=evolution_x,
    y="Saldo Bruto",
    color="Produto",
    markers=False,
    title="Evolução patrimonial por alternativa"
)

max_evolution_value = evolution_df["Saldo Bruto"].max()
evolution_y_max = max_evolution_value * 1.08 if max_evolution_value > 0 else 1

fig.update_layout(
    height=560,
    xaxis_title="Data" if evolution_x == "Data" else "Mês",
    yaxis_title="Saldo Bruto",
    legend_title="Produto",
    hovermode="x unified",
    margin=dict(t=60, r=20, l=20, b=40),
    font=dict(size=11),
)

fig.update_yaxes(
    range=[0, evolution_y_max],
    tickprefix="R$ ",
    tickformat="~s",
    tickfont=dict(size=11),
    title_font=dict(size=12),
    automargin=True
)

fig.update_xaxes(
    tickfont=dict(size=11),
    title_font=dict(size=12),
    automargin=True
)

st.plotly_chart(
    fig,
    use_container_width=True
)


# =========================================================
# RESUMO MENSAL DO CALENDÁRIO
# =========================================================

if simulation_mode == "Aportes e resgates por calendário":
    st.markdown(
        '<div class="section-title">Resumo Mensal do Calendário</div>',
        unsafe_allow_html=True
    )

    monthly_display_df = monthly_df.copy()

    monthly_currency_columns = [
        "Aportes",
        "Resgates",
        "Rendimento Bruto no Mês",
        "Saldo Bruto Final",
    ]

    for col in monthly_currency_columns:
        if col in monthly_display_df.columns:
            monthly_display_df[col] = monthly_display_df[col].apply(format_currency)

    st.dataframe(
        monthly_display_df,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# GRÁFICOS COM ESCALA DINÂMICA PARA VALORES PRIVATE
# =========================================================

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    max_liquid_value = comparison_df["Valor Líquido"].max()
    y_axis_max = max_liquid_value * 1.15 if max_liquid_value > 0 else 1

    ranking_fig = px.bar(
        comparison_df,
        x="Produto",
        y="Valor Líquido",
        title="Ranking por Valor Líquido",
        text="Valor Líquido"
    )

    ranking_fig.update_traces(
        texttemplate="%{customdata}",
        textposition="outside",
        textfont_size=10,
        cliponaxis=False,
        customdata=comparison_df["Valor Líquido"].apply(format_compact_currency)
    )

    ranking_fig.update_layout(
        height=540,
        xaxis_title="Produto",
        yaxis_title="Valor Líquido",
        margin=dict(t=70, r=20, l=20, b=50),
        showlegend=False,
        uniformtext_minsize=9,
        uniformtext_mode="hide",
        font=dict(size=11),
    )

    ranking_fig.update_xaxes(
        tickfont=dict(size=10),
        title_font=dict(size=11),
        automargin=True
    )

    ranking_fig.update_yaxes(
        range=[0, y_axis_max],
        tickprefix="R$ ",
        tickformat="~s",
        tickfont=dict(size=10),
        title_font=dict(size=11),
        automargin=True
    )

    st.plotly_chart(
        ranking_fig,
        use_container_width=True
    )

with col_chart2:
    profit_fig = px.pie(
        comparison_df,
        names="Produto",
        values="Rendimento Líquido",
        title="Distribuição do Rendimento Líquido"
    )

    profit_fig.update_traces(
        textinfo="percent",
        textfont_size=10,
        hovertemplate=(
            "<b>%{label}</b><br>"
            "Rendimento: R$ %{value:,.2f}<br>"
            "Participação: %{percent}"
            "<extra></extra>"
        )
    )

    profit_fig.update_layout(
        height=540,
        margin=dict(t=70, r=20, l=20, b=20),
        legend=dict(
            font=dict(size=10),
            orientation="v"
        ),
        font=dict(size=11)
    )

    st.plotly_chart(
        profit_fig,
        use_container_width=True
    )


# =========================================================
# LEITURA CONSULTIVA
# =========================================================

st.markdown(
    '<div class="section-title">Leitura Consultiva</div>',
    unsafe_allow_html=True
)

analysis = generate_consultive_analysis(comparison_df)
st.markdown(
    '<div class="section-title">Relatório Consultivo</div>',
    unsafe_allow_html=True
)

st.markdown(analysis)



# =========================================================
# INTELIGÊNCIA DE MERCADO E FORESIGHT
# =========================================================

if use_market_intelligence:
    st.markdown(
        '<div class="section-title">Inteligência de Mercado e Foresight</div>',
        unsafe_allow_html=True
    )

    st.info(
        "Módulo experimental ativado. Esta camada busca dados públicos do Bacen/Focus "
        "e gera uma leitura consultiva inicial, sem alterar os cálculos da simulação CDI."
    )

    update_market_data = st.button(
        "Atualizar inteligência de mercado",
        type="primary"
    )

    if update_market_data:
        try:
            with st.spinner("Buscando dados públicos de mercado..."):
                market_intelligence = generate_market_intelligence()

                st.session_state["market_intelligence"] = market_intelligence
                
                market_intelligence = generate_market_intelligence()

            bacen_df = market_intelligence.get("bacen_df")
            focus_df = market_intelligence.get("focus_df")
            curve_df = market_intelligence.get("curve_df")
            curve_shape = market_intelligence.get("curve_shape")
            market_reading = market_intelligence.get("reading")

            st.success("Inteligência de mercado carregada com sucesso.")

            # =========================================================
            # DADOS BACEN
            # =========================================================

            st.markdown("#### Dados Bacen")

            if bacen_df is None or bacen_df.empty:
                st.warning("Dados Bacen não disponíveis nesta execução.")
            else:
                st.dataframe(
                    bacen_df,
                    width="stretch",
                    hide_index=True
                )

            # =========================================================
            # EXPECTATIVAS FOCUS
            # =========================================================

            st.markdown("#### Expectativas Focus")

            if focus_df is None or focus_df.empty:
                st.warning("Expectativas Focus não disponíveis nesta execução.")
            else:
                st.dataframe(
                    focus_df,
                    width="stretch",
                    hide_index=True
                )

            # =========================================================
            # CURVA SIMPLIFICADA DE JUROS
            # =========================================================

            st.markdown("#### Curva Simplificada de Juros")

            movimento_curva = None
            spread_final = 0
            leitura_movimento = ""

            if curve_df is None or curve_df.empty:
                st.warning("Curva simplificada não disponível nesta execução.")
            else:
                st.caption(
                    f"Classificação da curva: {curve_shape}"
                )

                curve_chart_df = curve_df.copy()

                curve_chart_df["Vértice"] = (
                    curve_chart_df["Vértice"].astype(str)
                )

                curve_chart_df["Taxa Selic Esperada (%)"] = (
                    curve_chart_df["Taxa Selic Esperada (%)"].astype(float)
                )

                selic_atual_referencia = curve_chart_df[
                    "Taxa Selic Esperada (%)"
                ].iloc[0]

                curve_chart_df["Spread vs Selic Atual (p.p.)"] = (
                    curve_chart_df["Taxa Selic Esperada (%)"]
                    - selic_atual_referencia
                )

                curve_chart_df["Rótulo"] = curve_chart_df[
                    "Taxa Selic Esperada (%)"
                ].apply(
                    lambda value: f"{value:.2f}%"
                )

                spread_final = curve_chart_df[
                    "Spread vs Selic Atual (p.p.)"
                ].iloc[-1]

                limite_neutro = 0.10

                if spread_final > limite_neutro:
                    movimento_curva = "abertura da curva"
                    leitura_movimento = (
                        "A curva está abrindo em relação à Selic atual. "
                        "Isso indica que as expectativas de mercado apontam para juros futuros "
                        "acima da taxa corrente, o que pode favorecer uma conversa consultiva "
                        "sobre proteção de taxa, prazo e alternativas prefixadas, sempre conforme "
                        "o perfil e a necessidade de liquidez do cliente."
                    )
                elif spread_final < -limite_neutro:
                    movimento_curva = "fechamento da curva"
                    leitura_movimento = (
                        "A curva está fechando em relação à Selic atual. "
                        "Isso indica que as expectativas de mercado apontam para juros futuros "
                        "abaixo da taxa corrente, o que reforça a importância de avaliar o risco "
                        "de reinvestimento, o horizonte da aplicação e o momento adequado para "
                        "travar taxas em produtos prefixados ou híbridos."
                    )
                else:
                    movimento_curva = "curva praticamente estável"
                    leitura_movimento = (
                        "A curva está praticamente estável em relação à Selic atual. "
                        "Nesse cenário, a leitura consultiva deve priorizar liquidez, prazo, "
                        "tributação, previsibilidade e aderência ao objetivo financeiro do cliente."
                    )

                curve_fig = go.Figure()

                curve_fig.add_trace(
                    go.Scatter(
                        x=curve_chart_df["Vértice"],
                        y=curve_chart_df["Taxa Selic Esperada (%)"],
                        mode="lines+markers+text",
                        text=curve_chart_df["Rótulo"],
                        textposition="top center",
                        name="Selic esperada",
                        line=dict(
                            width=4,
                            shape="spline"
                        ),
                        marker=dict(
                            size=12,
                            symbol="circle"
                        ),
                        fill="tozeroy",
                        hovertemplate=(
                            "<b>Vértice:</b> %{x}<br>"
                            "<b>Taxa Selic esperada:</b> %{y:.2f}%"
                            "<extra></extra>"
                        )
                    )
                )

                curve_fig.add_trace(
                    go.Scatter(
                        x=curve_chart_df["Vértice"],
                        y=[selic_atual_referencia] * len(curve_chart_df),
                        mode="lines",
                        name="Selic atual",
                        line=dict(
                            width=2,
                            dash="dash"
                        ),
                        hovertemplate=(
                            "<b>Referência:</b> Selic atual<br>"
                            "<b>Taxa:</b> %{y:.2f}%"
                            "<extra></extra>"
                        )
                    )
                )

                curve_fig.update_layout(
                    title={
                        "text": (
                            f"Curva Simplificada de Juros — "
                            f"{str(curve_shape).title()} | "
                            f"{movimento_curva.title()}"
                        ),
                        "x": 0.03,
                        "xanchor": "left"
                    },
                    height=500,
                    template="plotly_white",
                    xaxis_title="Horizonte da expectativa",
                    yaxis_title="Taxa Selic esperada (%)",
                    showlegend=True,
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    ),
                    margin=dict(
                        t=100,
                        r=30,
                        l=40,
                        b=60
                    ),
                    font=dict(
                        size=12
                    )
                )

                curve_fig.update_yaxes(
                    ticksuffix="%",
                    showgrid=True,
                    zeroline=False
                )

                curve_fig.update_xaxes(
                    showgrid=False
                )

                st.plotly_chart(
                    curve_fig,
                    width="stretch"
                )

                curve_values = curve_chart_df[
                    "Taxa Selic Esperada (%)"
                ].tolist()

                curve_labels = curve_chart_df[
                    "Vértice"
                ].tolist()

                if len(curve_values) >= 3:
                    col_curve_1, col_curve_2, col_curve_3, col_curve_4 = (
                        st.columns(4)
                    )

                    col_curve_1.metric(
                        "Selic atual",
                        f"{selic_atual_referencia:.2f}%"
                    )

                    col_curve_2.metric(
                        curve_labels[1],
                        f"{curve_values[1]:.2f}%",
                        delta=(
                            f"{curve_values[1] - selic_atual_referencia:.2f} p.p."
                        )
                    )

                    col_curve_3.metric(
                        curve_labels[2],
                        f"{curve_values[2]:.2f}%",
                        delta=(
                            f"{curve_values[2] - selic_atual_referencia:.2f} p.p."
                        )
                    )

                    col_curve_4.metric(
                        "Movimento",
                        movimento_curva.title(),
                        delta=f"{spread_final:.2f} p.p."
                    )

                st.markdown("##### Leitura da curva em relação à Selic atual")

                st.markdown(
                    f"""
A Selic atual foi utilizada como referência da curva. 
O último vértice da curva apresenta diferença de **{spread_final:.2f} ponto percentual** 
em relação à taxa corrente, caracterizando **{movimento_curva}**.

{leitura_movimento}
"""
                )

                with st.expander("Ver tabela técnica da curva"):
                    st.dataframe(
                        curve_chart_df,
                        width="stretch",
                        hide_index=True
                    )

            # =========================================================
            # LEITURA FORESIGHT
            # =========================================================

            st.markdown("#### Leitura Foresight")

            if market_reading:
                st.markdown(market_reading)

                if movimento_curva:
                    st.markdown(
                        f"""
**Leitura complementar da curva:** em relação à Selic atual, a estrutura observada indica **{movimento_curva}**. 
O spread do último vértice frente à taxa corrente é de **{spread_final:.2f} ponto percentual**. 
Essa informação ajuda a qualificar a conversa sobre pós-fixados, prefixados e produtos híbridos, especialmente na avaliação de prazo, liquidez, previsibilidade e risco de reinvestimento.
"""
                    )
            else:
                st.warning("Leitura foresight não gerada nesta execução.")

        except Exception as error:
            st.error("Erro ao carregar a inteligência de mercado.")
            st.exception(error)


# =========================================================
# RELATÓRIO CONSULTIVO
# =========================================================

st.markdown(
    '<div class="section-title">Relatório Consultivo</div>',
    unsafe_allow_html=True
)

if simulation_mode == "Aportes e resgates por calendário":
    report_cashflow_df = cashflow_editor_df.copy()
    report_monthly_df = monthly_df.copy()
else:
    report_cashflow_df = pd.DataFrame()
    report_monthly_df = pd.DataFrame()

# =========================================================
# OPÇÕES DO RELATÓRIO
# =========================================================

st.markdown("### Opções do relatório")

with st.expander("Escolher o que incluir no relatório", expanded=False):
    incluir_visao_geral = st.checkbox(
        "Visão geral da simulação",
        value=True
    )

    incluir_premissas = st.checkbox(
        "Premissas utilizadas",
        value=True
    )

    incluir_comparativo = st.checkbox(
        "Comparativo dos produtos",
        value=True
    )

    incluir_calendario = st.checkbox(
        "Calendário de movimentações",
        value=True
    )

    incluir_resumo_mensal = st.checkbox(
        "Resumo mensal",
        value=False
    )

    incluir_leitura_consultiva = st.checkbox(
        "Leitura consultiva",
        value=True
    )

    incluir_inteligencia_mercado = st.checkbox(
        "Inteligência de mercado Bacen/Focus",
        value=True
    )

    incluir_curva_juros = st.checkbox(
        "Curva simplificada de juros",
        value=True
    )

    incluir_leitura_foresight = st.checkbox(
        "Leitura Foresight da curva",
        value=True
    )

    incluir_aviso_tecnico = st.checkbox(
        "Aviso técnico",
        value=True
    )

report_options = {
    "visao_geral": incluir_visao_geral,
    "premissas": incluir_premissas,
    "comparativo": incluir_comparativo,
    "calendario": incluir_calendario,
    "resumo_mensal": incluir_resumo_mensal,
    "leitura_consultiva": incluir_leitura_consultiva,
    "inteligencia_mercado": incluir_inteligencia_mercado,
    "curva_juros": incluir_curva_juros,
    "leitura_foresight": incluir_leitura_foresight,
    "aviso_tecnico": incluir_aviso_tecnico,
}

word_file = generate_word_report(
    client_name=client_name,
    advisor_name=advisor_name,
    simulation_mode=simulation_mode,
    start_date=start_date,
    end_date=end_date,
    months=int(months),
    initial_amount=initial_amount,
    annual_cdi_rate=annual_cdi_rate,
    selic_rate=selic_rate,
    tr_rate=tr_rate,
    cdb_percentage=cdb_percentage,
    lci_lca_percentage=lci_lca_percentage,
    treasury_percentage=treasury_percentage,
    treasury_annual_fee=treasury_annual_fee,
    fund_percentage=fund_percentage,
    fund_annual_fee=fund_annual_fee,
    comparison_df=comparison_df,
    cashflow_df=report_cashflow_df,
    monthly_df=report_monthly_df,
    consultive_analysis=analysis,
    market_intelligence=st.session_state.get("market_intelligence"),
    report_options=report_options,
)

file_name = f"relatorio_simulacao_{client_name.replace(' ', '_').lower()}.docx"

st.download_button(
    label="Baixar relatório Word",
    data=word_file,
    file_name=file_name,
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)



# =========================================================
# AVISO TÉCNICO
# =========================================================

st.markdown(
    """
    <div class="warning-box">
        <strong>Aviso:</strong> os valores apresentados são estimativas para fins de simulação.
        Não constituem promessa de rentabilidade, oferta, recomendação individualizada ou garantia de resultado.
        A análise final deve considerar perfil do cliente, adequação, liquidez, risco, tributação e condições vigentes.
    </div>
    """,
    unsafe_allow_html=True
)