# =========================================================
# CDI CALCULATOR
# Motor de cálculo com juros compostos
# =========================================================

from calculations.tax_calculator import calculate_tax


# =========================================================
# CONVERSÕES DE TAXA
# =========================================================

def annual_to_monthly_rate(annual_rate_percent: float) -> float:
    """
    Converte taxa anual efetiva em taxa mensal equivalente composta.

    Exemplo:
    14,40% a.a. vira aproximadamente 1,13% a.m.
    """
    annual_rate = annual_rate_percent / 100
    return (1 + annual_rate) ** (1 / 12) - 1


def monthly_to_annual_rate(monthly_rate_decimal: float) -> float:
    """
    Converte taxa mensal decimal em taxa anual percentual composta.

    Exemplo:
    0,0113 ao mês vira aproximadamente 14,40% ao ano.
    """
    return ((1 + monthly_rate_decimal) ** 12 - 1)


def period_decimal_to_monthly_percent(
    period_return_decimal: float,
    months: int | None = None,
    days: int | None = None,
) -> float:
    """
    Converte rentabilidade decimal do período em rentabilidade mensal percentual.

    Entrada:
    0.144 = 14,40%

    Saída:
    1.13 = 1,13% ao mês, aproximadamente.
    """

    if period_return_decimal <= -1:
        return -100.0

    if months is not None and months > 0:
        monthly_return = (1 + period_return_decimal) ** (1 / months) - 1
        return monthly_return * 100

    if days is not None and days > 0:
        monthly_return = (1 + period_return_decimal) ** (30 / days) - 1
        return monthly_return * 100

    return period_return_decimal


def period_decimal_to_annual_percent(
    period_return_decimal: float,
    months: int | None = None,
    days: int | None = None,
    base_days: int = 360,
) -> float:
    """
    Converte rentabilidade decimal do período em rentabilidade anual percentual equivalente.

    Entrada:
    0.144 = 14,40%

    Saída:
    Se o período for de 12 meses, retorna 14,40%.
    Se o período for menor ou maior, anualiza de forma composta.
    """

    if period_return_decimal <= -1:
        return -100.0

    if months is not None and months > 0:
        annual_return = (1 + period_return_decimal) ** (12 / months) - 1
        return annual_return * 100

    if days is not None and days > 0:
        annual_return = (1 + period_return_decimal) ** (base_days / days) - 1
        return annual_return * 100

    return period_return_decimal * 100


# =========================================================
# TAXAS DOS PRODUTOS
# =========================================================

def get_effective_cdi_annual_rate(
    annual_cdi_rate: float,
    cdi_percentage: float,
    annual_fee: float = 0.0,
) -> float:
    """
    Calcula a taxa anual efetiva do produto como percentual do CDI.

    Exemplo:
    CDI 14,40% e produto 100% CDI = 14,40% a.a.
    CDI 14,40% e produto 105% CDI = 15,12% a.a.

    A taxa/custo anual é descontada por fator composto.
    """

    gross_annual_rate = annual_cdi_rate * (cdi_percentage / 100)

    gross_factor = 1 + (gross_annual_rate / 100)
    fee_factor = 1 + (annual_fee / 100)

    effective_factor = gross_factor / fee_factor

    return (effective_factor - 1) * 100


def get_savings_annual_rate(
    selic_rate: float,
    tr_rate: float = 0.0,
) -> float:
    """
    Calcula a taxa anual efetiva da poupança.

    Regra simplificada:
    - Selic acima de 8,5%: 0,5% ao mês + TR
    - Selic até 8,5%: 70% da Selic + TR

    Observação:
    Para Selic acima de 8,5%, 0,5% ao mês equivale a cerca de 6,17% ao ano.
    """

    if selic_rate > 8.5:
        monthly_basic_rate = 0.005
        monthly_tr_rate = annual_to_monthly_rate(tr_rate)
        monthly_rate = monthly_basic_rate + monthly_tr_rate
        return monthly_to_annual_rate(monthly_rate)

    return (selic_rate * 0.70) + tr_rate


# =========================================================
# SIMULAÇÃO CDI SEM CALENDÁRIO DE APORTES
# =========================================================

def simulate_cdi_product(
    product_name: str,
    initial_amount: float,
    monthly_contribution: float,
    months: int,
    annual_cdi_rate: float,
    cdi_percentage: float,
    taxable: bool,
    annual_fee: float = 0.0,
) -> dict:
    """
    Simula produto pós-fixado CDI com juros compostos mensais.

    Premissas:
    - Taxa anual efetiva convertida para taxa mensal composta.
    - Aporte mensal considerado ao fim de cada mês.
    - IR calculado sobre o rendimento bruto.
    """

    effective_annual_rate = get_effective_cdi_annual_rate(
        annual_cdi_rate=annual_cdi_rate,
        cdi_percentage=cdi_percentage,
        annual_fee=annual_fee,
    )

    monthly_rate = annual_to_monthly_rate(effective_annual_rate)

    balance = initial_amount
    invested_amount = initial_amount
    total_contributions = initial_amount
    total_withdrawals = 0.0

    evolution = []

    for month in range(1, months + 1):
        balance = balance * (1 + monthly_rate)

        if monthly_contribution > 0:
            balance += monthly_contribution
            invested_amount += monthly_contribution
            total_contributions += monthly_contribution

        evolution.append(
            {
                "Mês": month,
                "Produto": product_name,
                "Saldo Bruto": balance,
            }
        )

    gross_value = balance
    gross_profit = gross_value - invested_amount

    days = months * 30

    ir_value, ir_rate = calculate_tax(
        gross_profit=gross_profit,
        days=days,
        taxable=taxable,
    )

    net_value = gross_value - ir_value
    net_profit = net_value - invested_amount

    net_period_return_decimal = (
        net_profit / invested_amount if invested_amount > 0 else 0.0
    )

    return {
        "Produto": product_name,
        "% CDI": cdi_percentage,
        "Taxa Efetiva a.a.": effective_annual_rate,
        "Valor Inicial": initial_amount,
        "Valor Investido": invested_amount,
        "Total Aportado": total_contributions,
        "Total Resgatado": total_withdrawals,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_profit,
        "IR": ir_value,
        "Alíquota IR": ir_rate,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_profit,
        "Rentabilidade Líquida no Período (%)": net_period_return_decimal,
        "Rentabilidade Líquida ao Mês (%)": period_decimal_to_monthly_percent(
            period_return_decimal=net_period_return_decimal,
            months=months,
        ),
        "Rentabilidade Líquida ao Ano (%)": period_decimal_to_annual_percent(
            period_return_decimal=net_period_return_decimal,
            months=months,
        ),
        "Tributável": "Sim" if taxable else "Não",
        "Evolução Mensal": evolution,
    }


# =========================================================
# SIMULAÇÃO POUPANÇA SEM CALENDÁRIO DE APORTES
# =========================================================

def simulate_savings(
    initial_amount: float,
    monthly_contribution: float,
    months: int,
    selic_rate: float,
    tr_rate: float,
) -> dict:
    """
    Simula poupança com capitalização composta mensal.

    Premissas:
    - Poupança isenta de IR.
    - Aporte mensal considerado ao fim de cada mês.
    """

    effective_annual_rate = get_savings_annual_rate(
        selic_rate=selic_rate,
        tr_rate=tr_rate,
    )

    monthly_rate = annual_to_monthly_rate(effective_annual_rate)

    balance = initial_amount
    invested_amount = initial_amount
    total_contributions = initial_amount
    total_withdrawals = 0.0

    evolution = []

    for month in range(1, months + 1):
        balance = balance * (1 + monthly_rate)

        if monthly_contribution > 0:
            balance += monthly_contribution
            invested_amount += monthly_contribution
            total_contributions += monthly_contribution

        evolution.append(
            {
                "Mês": month,
                "Produto": "Poupança",
                "Saldo Bruto": balance,
            }
        )

    gross_value = balance
    gross_profit = gross_value - invested_amount

    net_value = gross_value
    net_profit = gross_profit

    net_period_return_decimal = (
        net_profit / invested_amount if invested_amount > 0 else 0.0
    )

    return {
        "Produto": "Poupança",
        "% CDI": 0.0,
        "Taxa Efetiva a.a.": effective_annual_rate,
        "Valor Inicial": initial_amount,
        "Valor Investido": invested_amount,
        "Total Aportado": total_contributions,
        "Total Resgatado": total_withdrawals,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_profit,
        "IR": 0.0,
        "Alíquota IR": 0.0,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_profit,
        "Rentabilidade Líquida no Período (%)": net_period_return_decimal,
        "Rentabilidade Líquida ao Mês (%)": period_decimal_to_monthly_percent(
            period_return_decimal=net_period_return_decimal,
            months=months,
        ),
        "Rentabilidade Líquida ao Ano (%)": period_decimal_to_annual_percent(
            period_return_decimal=net_period_return_decimal,
            months=months,
        ),
        "Tributável": "Não",
        "Evolução Mensal": evolution,
    }


# =========================================================
# FUNÇÕES DE COMPATIBILIDADE PARA CASHFLOW
# =========================================================

def get_effective_annual_rate(
    annual_cdi_rate: float,
    cdi_percentage: float,
    annual_fee: float = 0.0,
) -> float:
    """
    Compatibilidade com o módulo de cashflow.

    Calcula a taxa anual efetiva do produto CDI.
    """
    return get_effective_cdi_annual_rate(
        annual_cdi_rate=annual_cdi_rate,
        cdi_percentage=cdi_percentage,
        annual_fee=annual_fee,
    )


def annual_to_business_daily_rate(
    annual_rate_percent: float,
    base_days: int = 252,
) -> float:
    """
    Converte taxa anual efetiva em taxa diária composta para dias úteis.

    Exemplo:
    14,40% a.a. em base 252.
    """
    annual_rate_decimal = annual_rate_percent / 100

    return (1 + annual_rate_decimal) ** (1 / base_days) - 1


def get_savings_monthly_rate(
    selic_rate: float,
    tr_rate: float = 0.0,
) -> float:
    """
    Retorna a taxa mensal da poupança em decimal.

    Se Selic > 8,5%:
    0,5% ao mês + TR mensal equivalente.

    Se Selic <= 8,5%:
    70% da Selic anual convertida para taxa mensal composta.
    """
    if selic_rate > 8.5:
        monthly_basic_rate = 0.005
        monthly_tr_rate = annual_to_monthly_rate(tr_rate)

        return monthly_basic_rate + monthly_tr_rate

    annual_savings_rate = (selic_rate * 0.70) + tr_rate

    return annual_to_monthly_rate(annual_savings_rate)