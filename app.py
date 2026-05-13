from datetime import date
from modules.ibovespa_cdi_module import render_ibovespa_cdi_module
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go



# =========================================================
# FORMATADORES
# =========================================================

def format_currency(value):
    """
    Formata número em moeda brasileira.
    Exemplo:
    500000 -> R$ 500.000,00
    """
    try:
        value = float(value)
        return (
            f"R$ {value:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except Exception:
        return value


def format_percent(value):
    """
    Formata percentuais em padrão brasileiro.

    Aceita dois padrões:
    - 9.80  -> 9,80%
    - 0.098 -> 9,80%

    Isso evita o erro de exibir 0,10% quando o correto é 9,80%.
    """
    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return value

    if numeric_value == 0:
        percent_value = 0.0
    elif abs(numeric_value) < 1:
        percent_value = numeric_value * 100
    else:
        percent_value = numeric_value

    return f"{percent_value:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


from services.simulation_service import (
    run_cdi_simulation,
    run_cdi_cashflow_simulation,
)
from services.interpretation_service import generate_consultive_analysis
from reports.word_report_generator import generate_word_report
from services.market_intelligence_service import generate_market_intelligence
from services.fund_tax_service import FundTaxService


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

fund_type = st.selectbox(
    "Classificação fiscal do Fundo DI",
    [
        FundTaxService.LONG_TERM,
        FundTaxService.SHORT_TERM,
    ],
    index=0,
    help=(
        "Fundos de longo prazo usam come-cotas de 15%. "
        "Fundos de curto prazo usam come-cotas de 20%."
    )
)

apply_come_cotas = st.checkbox(
    "Aplicar come-cotas no Fundo DI",
    value=True,
    help=(
        "O come-cotas é a antecipação semestral do IR, "
        "normalmente em maio e novembro."
    )
)


# =========================================================
# CENÁRIO TRIBUTÁRIO COMPLEMENTAR: DIVIDENDOS
# =========================================================

include_dividend_scenario = st.checkbox(
    "Incluir cenário tributário com dividendos",
    value=st.session_state.get("input_include_dividend_scenario", False),
    key="input_include_dividend_scenario",
    help=(
        "Inclui uma camada complementar para clientes com recebimento "
        "relevante de lucros e dividendos. A simulação é estimativa "
        "e não substitui avaliação contábil ou fiscal."
    )
)

monthly_dividends = 0.0
same_payer_dividends = True
months_with_dividends = 12
taxable_monthly_dividends = 0.0
estimated_monthly_dividend_ir = 0.0
estimated_annual_dividend_ir = 0.0
annual_total_income = 0.0
minimum_tax_rate = 0.0
minimum_tax_due = 0.0
integrated_tax_scenario = {}

if include_dividend_scenario:
    st.markdown("#### Cenário tributário com dividendos")

    monthly_dividends = st.number_input(
        "Dividendos mensais estimados",
        min_value=0.0,
        value=float(st.session_state.get("input_monthly_dividends", 0.0)),
        step=1000.0,
        format="%.2f",
        key="input_monthly_dividends",
        help=(
            "Informe o valor mensal estimado de lucros e dividendos "
            "recebidos pelo cliente."
        )
    )

    same_payer_dividends = st.checkbox(
        "Dividendos pagos por uma mesma pessoa jurídica",
        value=st.session_state.get("input_same_payer_dividends", True),
        key="input_same_payer_dividends",
        help=(
            "A camada considera a hipótese de pagamentos concentrados "
            "em uma mesma fonte pagadora."
        )
    )

    months_with_dividends = st.number_input(
        "Quantidade de meses com dividendos no ano",
        min_value=0,
        max_value=12,
        value=int(st.session_state.get("input_months_with_dividends", 12)),
        step=1,
        key="input_months_with_dividends"
    )

    annual_total_income = st.number_input(
        "Renda anual total estimada do cliente",
        min_value=0.0,
        value=float(st.session_state.get("input_annual_total_income", 0.0)),
        step=10000.0,
        format="%.2f",
        key="input_annual_total_income",
        help=(
            "Informe a renda anual total estimada do cliente para simular "
            "a tributação mínima. Esse valor deve considerar a vida fiscal "
            "anual do cliente, conforme orientação contábil."
        )
    )

    DIVIDEND_MONTHLY_REFERENCE_LIMIT = 50000.0
    DIVIDEND_IR_RATE = 0.10

    if same_payer_dividends:
        if monthly_dividends > DIVIDEND_MONTHLY_REFERENCE_LIMIT:
            taxable_monthly_dividends = monthly_dividends
        else:
            taxable_monthly_dividends = 0.0

        estimated_monthly_dividend_ir = (
            taxable_monthly_dividends * DIVIDEND_IR_RATE
        )

        estimated_annual_dividend_ir = (
            estimated_monthly_dividend_ir * months_with_dividends
        )
    else:
        taxable_monthly_dividends = 0.0
        estimated_monthly_dividend_ir = 0.0
        estimated_annual_dividend_ir = 0.0

    if annual_total_income <= 600000:
        minimum_tax_rate = 0.0
    elif annual_total_income < 1200000:
        minimum_tax_rate = (
            (annual_total_income - 600000) / 600000
        ) * 0.10
    else:
        minimum_tax_rate = 0.10

    minimum_tax_due = annual_total_income * minimum_tax_rate

    st.info(
        "Esta é uma simulação complementar e simplificada. Ela não altera "
        "automaticamente a rentabilidade dos produtos, mas ajuda a contextualizar "
        "a decisão dentro da situação tributária global do cliente."
    )

    col_div_1, col_div_2, col_div_3, col_div_4 = st.columns(4)

    col_div_1.metric(
        "Base mensal dividendos",
        format_currency(taxable_monthly_dividends)
    )

    col_div_2.metric(
        "IRRF anual dividendos",
        format_currency(estimated_annual_dividend_ir)
    )

    col_div_3.metric(
        "Alíquota mínima estimada",
        f"{minimum_tax_rate * 100:.2f}%"
    )

    col_div_4.metric(
        "IR mínimo anual estimado",
        format_currency(minimum_tax_due)
    )


# =========================================================
# CALENDÁRIO DE MOVIMENTAÇÕES
# =========================================================

cashflows = []

if simulation_mode == "Aportes e resgates por calendário":
    st.markdown(
        '<div class="section-title">Calendário de Movimentações</div>',
        unsafe_allow_html=True
    )

    st.info(
        "Informe aportes e resgates programados. Para resgates, selecione de qual produto o valor será retirado."
    )

    default_cashflow_df = pd.DataFrame(
        [
            {
                "Data": start_date,
                "Produto": "Todos os produtos",
                "Tipo": "Aporte",
                "Valor": 0.0,
                "Descrição": "Aporte programado",
            }
        ]
    )

    cashflow_df = st.data_editor(
        default_cashflow_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Data": st.column_config.DateColumn(
                "Data",
                format="DD/MM/YYYY",
                required=True,
            ),
            "Produto": st.column_config.SelectboxColumn(
                "Produto",
                options=[
                    "Todos os produtos",
                    "LCI / LCA",
                    "CDB / LC",
                    "Tesouro Selic",
                    "Fundo DI",
                    "Poupança",
                ],
                required=True,
            ),
            "Tipo": st.column_config.SelectboxColumn(
                "Tipo",
                options=[
                    "Aporte",
                    "Resgate",
                ],
                required=True,
            ),
            "Valor": st.column_config.NumberColumn(
                "Valor",
                min_value=0.0,
                step=1000.0,
                format="R$ %.2f",
                required=True,
            ),
            "Descrição": st.column_config.TextColumn(
                "Descrição"
            ),
        },
        key="cashflow_editor",
    )

    valid_cashflow_df = cashflow_df[
        cashflow_df["Valor"].fillna(0) > 0
    ].copy()

    cashflows = valid_cashflow_df.to_dict("records")


# =========================================================
# VALIDAÇÃO DE PERÍODO
# =========================================================

if end_date <= start_date:
    st.error("A data final precisa ser posterior à data inicial.")
    st.stop()


# =========================================================
# PARÂMETROS POR PRODUTO
# =========================================================

st.markdown("### Parâmetros por Produto")

col_produto_1, col_produto_2 = st.columns(2)

with col_produto_1:
    cdb_percentage = st.number_input(
        "CDB / LC (% do CDI)",
        min_value=0.0,
        max_value=250.0,
        value=105.0,
        step=1.0,
        format="%.2f",
        key="cdb_percentage_input"
    )

    lci_lca_percentage = st.number_input(
        "LCI / LCA (% do CDI)",
        min_value=0.0,
        max_value=250.0,
        value=95.0,
        step=1.0,
        format="%.2f",
        key="lci_lca_percentage_input"
    )

    treasury_percentage = st.number_input(
        "Tesouro Selic (% do CDI aproximado)",
        min_value=0.0,
        max_value=250.0,
        value=100.0,
        step=1.0,
        format="%.2f",
        key="treasury_percentage_input"
    )

with col_produto_2:
    treasury_annual_fee = st.number_input(
        "Taxa/custo anual Tesouro Selic (%)",
        min_value=0.0,
        max_value=10.0,
        value=0.20,
        step=0.05,
        format="%.2f",
        key="treasury_annual_fee_input"
    )

    fund_percentage = st.number_input(
        "Fundo DI (% do CDI)",
        min_value=0.0,
        max_value=250.0,
        value=100.0,
        step=1.0,
        format="%.2f",
        key="fund_percentage_input"
    )

    fund_annual_fee = st.number_input(
        "Taxa de administração Fundo DI (%)",
        min_value=0.0,
        max_value=10.0,
        value=0.50,
        step=0.05,
        format="%.2f",
        key="fund_annual_fee_input"
    )

# =========================================================
# EXECUÇÃO DA SIMULAÇÃO
# =========================================================

daily_df = pd.DataFrame()
monthly_df = pd.DataFrame()
evolution_df = pd.DataFrame()
evolution_x = "Mês"

if simulation_mode == "Aportes e resgates por calendário":

    try:
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

    except Exception as error:
        st.warning(
            "O modo de aportes e resgates por calendário ainda está em ajuste. "
            "Para manter o simulador aberto, esta execução usará temporariamente "
            "a simulação padrão pelo período informado, sem aplicar os movimentos "
            "do calendário ao cálculo."
        )

        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)

        days_between_dates = max((end - start).days, 1)
        fallback_months = max(round(days_between_dates / 30), 1)

        comparison_df, evolution_df = run_cdi_simulation(
            initial_amount=initial_amount,
            monthly_contribution=0.0,
            months=int(fallback_months),
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

        daily_df = evolution_df.copy()
        monthly_df = pd.DataFrame()
        evolution_x = "Mês"

        with st.expander("Detalhes técnicos do fallback"):
            st.exception(error)

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
# AJUSTE DO FUNDO DI COM COME-COTAS
# =========================================================

fund_tax_service = FundTaxService()

annual_cdi_decimal = (
    annual_cdi_rate / 100
    if annual_cdi_rate > 1
    else annual_cdi_rate
)

fund_cdi_decimal = fund_percentage / 100

fund_admin_fee_decimal = fund_annual_fee / 100

if simulation_mode == "Aportes e resgates por calendário":
    fund_months = max(
        1,
        (end_date.year - start_date.year) * 12
        + (end_date.month - start_date.month)
    )
else:
    fund_months = int(months)

fund_result = fund_tax_service.simulate_fund_di(
    initial_amount=initial_amount,
    annual_cdi_rate=annual_cdi_decimal,
    fund_cdi_percentage=fund_cdi_decimal,
    admin_fee_rate=fund_admin_fee_decimal,
    months=fund_months,
    fund_type=fund_type,
    start_year=start_date.year,
    start_month=start_date.month,
    apply_come_cotas=apply_come_cotas,
)

# =========================================================
# SUBSTITUIÇÃO DA LINHA DO FUNDO DI NO COMPARATIVO
# =========================================================

if not comparison_df.empty and "Produto" in comparison_df.columns:
    fund_mask = comparison_df["Produto"].astype(str).str.contains(
        "Fundo DI",
        case=False,
        na=False
    )

    if fund_mask.any():

        if "Valor Bruto Final" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "Valor Bruto Final"
            ] = fund_result.gross_final_amount

        if "Valor Líquido Final" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "Valor Líquido Final"
            ] = fund_result.net_final_amount

        if "Lucro Líquido" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "Lucro Líquido"
            ] = fund_result.net_profit

        if "Rentabilidade Líquida (%)" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "Rentabilidade Líquida (%)"
            ] = fund_result.net_return_percentage

        if "IR Total" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "IR Total"
            ] = fund_result.total_tax

        if "Come-cotas" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "Come-cotas"
            ] = fund_result.come_cotas_tax

        if "IR no Resgate" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "IR no Resgate"
            ] = fund_result.redemption_tax

        if "Taxa de Administração" in comparison_df.columns:
            comparison_df.loc[
                fund_mask,
                "Taxa de Administração"
            ] = fund_result.admin_fee_impact


# =========================================================
# CENÁRIO ANUAL: TRIBUTAÇÃO MÍNIMA, DEDUÇÕES E ALOCAÇÃO
# =========================================================

if include_dividend_scenario:
    st.markdown("#### Cenário anual: tributação mínima, deduções e alocação")

    st.caption(
        "Este quadro compara o resultado líquido isolado dos produtos com uma "
        "leitura anual da vida tributária do cliente. A lógica não é somar "
        "impostos como vantagem, mas estimar se impostos já pagos ou retidos "
        "podem reduzir eventual saldo adicional de tributação mínima, quando "
        "aplicável ao caso concreto."
    )

    cdb_ir = 0.0
    lci_lca_ir = 0.0
    treasury_ir = 0.0
    fund_ir = 0.0

    cdb_net = 0.0
    lci_lca_net = 0.0
    treasury_net = 0.0
    fund_net = 0.0

    capital_base = 0.0

    if (
        comparison_df is not None
        and not comparison_df.empty
        and "Produto" in comparison_df.columns
    ):
        product_map = {}

        for _, row in comparison_df.iterrows():
            product_name = str(row["Produto"]).strip()
            product_map[product_name] = row

        if "CDB / LC" in product_map:
            cdb_ir = float(product_map["CDB / LC"].get("IR", 0.0))
            cdb_net = float(
                product_map["CDB / LC"].get("Valor Líquido", 0.0)
            )

        if "LCI / LCA" in product_map:
            lci_lca_ir = float(product_map["LCI / LCA"].get("IR", 0.0))
            lci_lca_net = float(
                product_map["LCI / LCA"].get("Valor Líquido", 0.0)
            )

        if "Tesouro Selic" in product_map:
            treasury_ir = float(
                product_map["Tesouro Selic"].get("IR", 0.0)
            )
            treasury_net = float(
                product_map["Tesouro Selic"].get("Valor Líquido", 0.0)
            )

        if "Fundo DI" in product_map:
            fund_ir = float(product_map["Fundo DI"].get("IR", 0.0))
            fund_net = float(
                product_map["Fundo DI"].get("Valor Líquido", 0.0)
            )

        if "Total Aportado" in comparison_df.columns:
            capital_base = float(comparison_df["Total Aportado"].max())

    if capital_base <= 0:
        capital_base = float(initial_amount)

    annual_dividend_ir = estimated_annual_dividend_ir

    products_for_analysis = [
        {
            "Produto": "LCI / LCA",
            "Valor líquido da aplicação": lci_lca_net,
            "IR do produto": lci_lca_ir,
            "IRRF dividendos": annual_dividend_ir,
        },
        {
            "Produto": "CDB / LC",
            "Valor líquido da aplicação": cdb_net,
            "IR do produto": cdb_ir,
            "IRRF dividendos": annual_dividend_ir,
        },
        {
            "Produto": "Tesouro Selic",
            "Valor líquido da aplicação": treasury_net,
            "IR do produto": treasury_ir,
            "IRRF dividendos": annual_dividend_ir,
        },
        {
            "Produto": "Fundo DI",
            "Valor líquido da aplicação": fund_net,
            "IR do produto": fund_ir,
            "IRRF dividendos": annual_dividend_ir,
        },
    ]

    scenario_rows = []

    for item in products_for_analysis:
        tax_already_paid_or_withheld = (
            item["IR do produto"] + item["IRRF dividendos"]
        )

        estimated_additional_tax = max(
            0.0,
            minimum_tax_due - tax_already_paid_or_withheld
        )

        lci_tax_already_paid = lci_lca_ir + annual_dividend_ir

        lci_additional_tax = max(
            0.0,
            minimum_tax_due - lci_tax_already_paid
        )

        fiscal_effect_vs_lci = 0.0

        if item["Produto"] != "LCI / LCA":
            fiscal_effect_vs_lci = (
                lci_additional_tax - estimated_additional_tax
            )

        comparable_net_value = (
            item["Valor líquido da aplicação"] + fiscal_effect_vs_lci
        )

        isolated_net_return_rate = 0.0
        comparable_net_return_rate = 0.0

        if capital_base > 0:
            isolated_net_return_rate = (
                (item["Valor líquido da aplicação"] / capital_base) - 1
            ) * 100

            comparable_net_return_rate = (
                (comparable_net_value / capital_base) - 1
            ) * 100

        if item["Produto"] == "LCI / LCA":
            decision_reading = (
                "Referência isenta. Vence quando a análise considera apenas "
                "o valor líquido da aplicação."
            )
        else:
            if fiscal_effect_vs_lci > 0:
                decision_reading = (
                    "Produto tributado com efeito fiscal potencial. Deve ser "
                    "comparado com a LCI/LCA pelo valor comparável no cenário anual."
                )
            else:
                decision_reading = (
                    "Produto tributado sem ganho fiscal relevante neste cenário. "
                    "A decisão deve priorizar rentabilidade líquida, liquidez e prazo."
                )

        scenario_rows.append(
            {
                "Produto": item["Produto"],
                "Valor líquido da aplicação": item[
                    "Valor líquido da aplicação"
                ],
                "Rentab. líquida isolada (%)": isolated_net_return_rate,
                "IR do produto": item["IR do produto"],
                "IRRF dividendos": item["IRRF dividendos"],
                "Saldo adicional estimado": estimated_additional_tax,
                "Efeito fiscal potencial vs LCI/LCA": fiscal_effect_vs_lci,
                "Valor comparável no cenário": comparable_net_value,
                "Rentab. comparável no cenário (%)": comparable_net_return_rate,
                "Decisão consultiva": decision_reading,
            }
        )

    annual_tax_scenario_df = pd.DataFrame(scenario_rows)

    display_annual_tax_scenario_df = annual_tax_scenario_df.copy()

    money_columns = [
        "Valor líquido da aplicação",
        "IR do produto",
        "IRRF dividendos",
        "Saldo adicional estimado",
        "Efeito fiscal potencial vs LCI/LCA",
        "Valor comparável no cenário",
    ]

    for column in money_columns:
        display_annual_tax_scenario_df[column] = (
            display_annual_tax_scenario_df[column].apply(format_currency)
        )

    percent_columns = [
        "Rentab. líquida isolada (%)",
        "Rentab. comparável no cenário (%)",
    ]

    for column in percent_columns:
        display_annual_tax_scenario_df[column] = (
            display_annual_tax_scenario_df[column].apply(
                lambda value: f"{value:.2f}%"
            )
        )

    st.dataframe(
        display_annual_tax_scenario_df,
        width="stretch",
        hide_index=True
    )

    isolated_winner_row = annual_tax_scenario_df.sort_values(
        by="Valor líquido da aplicação",
        ascending=False
    ).iloc[0]

    adjusted_winner_row = annual_tax_scenario_df.sort_values(
        by="Valor comparável no cenário",
        ascending=False
    ).iloc[0]

    cdb_row = annual_tax_scenario_df[
        annual_tax_scenario_df["Produto"] == "CDB / LC"
    ].iloc[0]

    lci_row = annual_tax_scenario_df[
        annual_tax_scenario_df["Produto"] == "LCI / LCA"
    ].iloc[0]

    net_difference_lci_vs_cdb = (
        lci_row["Valor líquido da aplicação"]
        - cdb_row["Valor líquido da aplicação"]
    )

    fiscal_effect_cdb_vs_lci = cdb_row[
        "Efeito fiscal potencial vs LCI/LCA"
    ]

    st.markdown("##### Leitura consultiva do cenário")

    st.success(
        f"No resultado líquido isolado da aplicação, o melhor produto é "
        f"{isolated_winner_row['Produto']}, com valor líquido estimado de "
        f"{format_currency(isolated_winner_row['Valor líquido da aplicação'])}."
    )

    if minimum_tax_due > 0:
        st.info(
            f"Na leitura fiscal anual, a renda anual informada gera IR mínimo "
            f"estimado de {format_currency(minimum_tax_due)}, com alíquota "
            f"mínima estimada de {minimum_tax_rate * 100:.2f}%. O CDB/LC "
            f"gera {format_currency(cdb_ir)} de IR próprio no produto. "
            f"Frente à LCI/LCA, isso representa efeito fiscal potencial de "
            f"{format_currency(fiscal_effect_cdb_vs_lci)}."
        )

        if fiscal_effect_cdb_vs_lci > net_difference_lci_vs_cdb:
            st.success(
                f"A diferença líquida da LCI/LCA sobre o CDB/LC é de "
                f"{format_currency(net_difference_lci_vs_cdb)}. Como o "
                f"efeito fiscal potencial do CDB/LC supera essa diferença, "
                f"o CDB/LC merece análise prioritária no cenário fiscal anual."
            )
        else:
            st.warning(
                f"A diferença líquida da LCI/LCA sobre o CDB/LC é de "
                f"{format_currency(net_difference_lci_vs_cdb)}. Como o "
                f"efeito fiscal potencial do CDB/LC não supera essa diferença, "
                f"a LCI/LCA permanece mais forte neste cenário."
            )
    else:
        st.info(
            "Como a renda anual informada não gera IR mínimo estimado, o "
            "cenário fiscal anual não altera a decisão. A comparação deve "
            "priorizar valor líquido da aplicação, liquidez, prazo e risco."
        )

    st.warning(
        "Atenção: esta simulação é estimativa. Ela não representa compensação "
        "automática mensal entre produtos financeiros e dividendos. A "
        "elegibilidade das deduções e o cálculo definitivo dependem da "
        "composição da renda anual do cliente e devem ser validados por "
        "contador ou especialista tributário."
    )

    integrated_tax_scenario = {
        "include_dividend_scenario": include_dividend_scenario,
        "monthly_dividends": monthly_dividends,
        "same_payer_dividends": same_payer_dividends,
        "months_with_dividends": months_with_dividends,
        "taxable_monthly_dividends": taxable_monthly_dividends,
        "estimated_monthly_dividend_ir": estimated_monthly_dividend_ir,
        "estimated_annual_dividend_ir": estimated_annual_dividend_ir,
        "annual_total_income": annual_total_income,
        "minimum_tax_rate": minimum_tax_rate,
        "minimum_tax_due": minimum_tax_due,
        "isolated_winner_product": isolated_winner_row["Produto"],
        "isolated_winner_net_value": float(
            isolated_winner_row["Valor líquido da aplicação"]
        ),
        "adjusted_winner_product": adjusted_winner_row["Produto"],
        "adjusted_winner_value": float(
            adjusted_winner_row["Valor comparável no cenário"]
        ),
        "net_difference_lci_vs_cdb": float(net_difference_lci_vs_cdb),
        "fiscal_effect_cdb_vs_lci": float(fiscal_effect_cdb_vs_lci),
        "capital_base": float(capital_base),
        "annual_tax_scenario_records": annual_tax_scenario_df.to_dict("records"),
    }
    

# =========================================================
# DETALHAMENTO DO FUNDO DI
# =========================================================

with st.expander("Detalhamento tributário do Fundo DI"):
    st.write("Classificação fiscal:", fund_result.fund_type)
    st.write("Aplicar come-cotas:", "Sim" if apply_come_cotas else "Não")

    st.write(
        "Alíquota de come-cotas:",
        "15%" if fund_result.fund_type == FundTaxService.LONG_TERM else "20%"
    )

    st.write(
        "Alíquota final de IR no resgate:",
        f"{fund_tax_service.get_final_ir_rate(fund_months * 30, fund_result.fund_type) * 100:.1f}%"
    )

    st.write("Valor bruto final:", round(fund_result.gross_final_amount, 2))
    st.write("Valor líquido final:", round(fund_result.net_final_amount, 2))
    st.write("Lucro líquido:", round(fund_result.net_profit, 2))
    st.write(
        "Impacto da taxa de administração:",
        round(fund_result.admin_fee_impact, 2)
    )
    st.write("Come-cotas:", round(fund_result.come_cotas_tax, 2))
    st.write("IR no resgate:", round(fund_result.redemption_tax, 2))
    st.write("IR total:", round(fund_result.total_tax, 2))
    st.write(
        "Rentabilidade líquida %:",
        round(fund_result.net_return_percentage, 2)
    )


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

# =========================================================
# NORMALIZAÇÃO DAS RENTABILIDADES
# =========================================================
# O motor pode devolver rentabilidade em decimal:
# 0.0978 = 9,78%
# Este bloco converte para escala percentual apenas quando necessário.

return_columns_to_normalize = [
    "Rentab. Líq. Período (%)",
    "Rentab. Líq. Mês (%)",
    "Rentab. Líq. Ano (%)",
]

for col in return_columns_to_normalize:
    if col in display_df.columns:
        display_df[col] = display_df[col].apply(
            lambda x: x * 100 if abs(float(x)) <= 1 else x
        )

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
                    width="stretch",
                    key="curve_simplified_interest_rate_chart"
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

        except Exception as e:
            st.warning(
                f"Não foi possível carregar a inteligência de mercado: {e}"
            )


# =========================================================
# RELATÓRIO CONSULTIVO
# =========================================================

st.markdown(
    '<div class="section-title">Relatório Consultivo</div>',
    unsafe_allow_html=True
)

if simulation_mode == "Aportes e resgates por calendário":
    report_cashflow_df = cashflow_df.copy()
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
    include_dividend_scenario=include_dividend_scenario,
    monthly_dividends=monthly_dividends,
    same_payer_dividends=same_payer_dividends,
    months_with_dividends=months_with_dividends,
    taxable_monthly_dividends=taxable_monthly_dividends,
    estimated_monthly_dividend_ir=estimated_monthly_dividend_ir,
    estimated_annual_dividend_ir=estimated_annual_dividend_ir,
    integrated_tax_scenario=integrated_tax_scenario,
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
# RADAR IBOVESPA x CDI
# =========================================================

st.divider()

with st.expander("Radar Ibovespa x CDI", expanded=False):

    st.info(
        "Módulo complementar de inteligência de mercado. "
        "Compara o desempenho histórico do Ibovespa com o CDI, identifica ciclos relevantes "
        "e gera uma leitura consultiva para apoio à conversa com o cliente."
    )

    ativar_ibovespa_cdi = st.checkbox(
        "Ativar análise Ibovespa x CDI",
        value=False,
        key="ativar_ibovespa_cdi_relatorio_final"
    )

    if ativar_ibovespa_cdi:
        render_ibovespa_cdi_module()


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