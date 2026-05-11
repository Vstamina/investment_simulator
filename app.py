from datetime import date
import streamlit as st
import pandas as pd
import plotly.express as px

from services.simulation_service import (
    run_cdi_simulation,
    run_cdi_cashflow_simulation,
)
from services.interpretation_service import generate_consultive_analysis
from reports.word_report_generator import generate_word_report


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

display_df = comparison_df.copy()

currency_columns = [
    "Valor Inicial",
    "Total Aportado",
    "Total Resgatado",
    "Valor Bruto",
    "Rendimento Bruto",
    "IR",
    "Valor Líquido",
    "Rendimento Líquido",
]

percent_columns = [
    "% CDI",
    "Taxa Efetiva a.a. (%)",
    "Alíquota IR (%)",
    "Rentab. Líq. Período (%)",
    "Rentab. Líq. Mês (%)",
    "Rentab. Líq. Ano (%)",
]

for col in currency_columns:
    if col in display_df.columns:
        display_df[col] = display_df[col].apply(format_currency)

for col in percent_columns:
    if col in display_df.columns:
        display_df[col] = display_df[col].apply(format_percent)

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