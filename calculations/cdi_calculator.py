from calculations.tax_calculator import calculate_tax, get_ir_rate


def annual_to_monthly_rate(annual_rate: float) -> float:
    """
    Converte taxa anual percentual em taxa mensal decimal.
    Exemplo: 10.65% ao ano -> 0.00847 ao mês aproximadamente.
    """
    return (1 + annual_rate / 100) ** (1 / 12) - 1


def simulate_cdi_product(
    product_name: str,
    initial_amount: float,
    monthly_contribution: float,
    months: int,
    annual_cdi_rate: float,
    cdi_percentage: float,
    taxable: bool = True,
    annual_fee: float = 0.0,
) -> dict:
    """
    Simula produto indexado ao CDI.
    """

    days = months * 30

    effective_annual_rate = (annual_cdi_rate * (cdi_percentage / 100)) - annual_fee

    if effective_annual_rate < 0:
        effective_annual_rate = 0

    monthly_rate = annual_to_monthly_rate(effective_annual_rate)

    balance = initial_amount
    total_contributions = initial_amount

    monthly_evolution = []

    for month in range(1, months + 1):
        balance += monthly_contribution
        total_contributions += monthly_contribution

        balance = balance * (1 + monthly_rate)

        monthly_evolution.append(
            {
                "Mês": month,
                "Produto": product_name,
                "Saldo Bruto": balance,
            }
        )

    gross_value = balance
    gross_profit = gross_value - total_contributions
    tax_value = calculate_tax(gross_profit, days, taxable)
    net_value = gross_value - tax_value
    net_profit = net_value - total_contributions

    net_return_period = (net_profit / total_contributions) * 100 if total_contributions > 0 else 0
    net_return_monthly = ((1 + net_return_period / 100) ** (1 / months) - 1) * 100 if months > 0 else 0
    net_return_annual = ((1 + net_return_monthly / 100) ** 12 - 1) * 100

    return {
        "Produto": product_name,
        "% CDI": cdi_percentage,
        "Taxa Efetiva a.a.": effective_annual_rate,
        "Valor Investido": total_contributions,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_profit,
        "IR": tax_value,
        "Alíquota IR": get_ir_rate(days) * 100 if taxable else 0,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_profit,
        "Rentabilidade Líquida no Período (%)": net_return_period,
        "Rentabilidade Líquida ao Mês (%)": net_return_monthly,
        "Rentabilidade Líquida ao Ano (%)": net_return_annual,
        "Tributável": "Sim" if taxable else "Não",
        "Evolução Mensal": monthly_evolution,
    }


def simulate_savings(
    initial_amount: float,
    monthly_contribution: float,
    months: int,
    selic_rate: float,
    tr_rate: float,
) -> dict:
    """
    Simulação simplificada da poupança.
    Regra aproximada:
    - Selic acima de 8,5% a.a.: 0,5% ao mês + TR
    - Selic até 8,5% a.a.: 70% da Selic + TR
    """

    if selic_rate > 8.5:
        annual_rate = 6.17 + tr_rate
    else:
        annual_rate = (selic_rate * 0.70) + tr_rate

    monthly_rate = annual_to_monthly_rate(annual_rate)

    balance = initial_amount
    total_contributions = initial_amount
    monthly_evolution = []

    for month in range(1, months + 1):
        balance += monthly_contribution
        total_contributions += monthly_contribution
        balance = balance * (1 + monthly_rate)

        monthly_evolution.append(
            {
                "Mês": month,
                "Produto": "Poupança",
                "Saldo Bruto": balance,
            }
        )

    gross_value = balance
    gross_profit = gross_value - total_contributions
    net_value = gross_value
    net_profit = gross_profit

    net_return_period = (net_profit / total_contributions) * 100 if total_contributions > 0 else 0
    net_return_monthly = ((1 + net_return_period / 100) ** (1 / months) - 1) * 100 if months > 0 else 0
    net_return_annual = ((1 + net_return_monthly / 100) ** 12 - 1) * 100

    return {
        "Produto": "Poupança",
        "% CDI": 0,
        "Taxa Efetiva a.a.": annual_rate,
        "Valor Investido": total_contributions,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_profit,
        "IR": 0,
        "Alíquota IR": 0,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_profit,
        "Rentabilidade Líquida no Período (%)": net_return_period,
        "Rentabilidade Líquida ao Mês (%)": net_return_monthly,
        "Rentabilidade Líquida ao Ano (%)": net_return_annual,
        "Tributável": "Não",
        "Evolução Mensal": monthly_evolution,
    }