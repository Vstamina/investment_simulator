from datetime import date
import pandas as pd

from calculations.tax_calculator import calculate_tax, get_ir_rate


def annual_to_daily_rate(annual_rate: float, business_days: int = 252) -> float:
    """
    Converte taxa anual percentual em taxa diária decimal.
    Usa 252 dias úteis como aproximação de mercado.
    """
    return (1 + annual_rate / 100) ** (1 / business_days) - 1


def simulate_product_with_cashflows(
    product_name: str,
    initial_amount: float,
    start_date: date,
    end_date: date,
    annual_cdi_rate: float,
    cdi_percentage: float,
    taxable: bool = True,
    annual_fee: float = 0.0,
    cashflows: list[dict] | None = None,
) -> dict:
    """
    Simula um produto CDI considerando aportes e resgates em datas específicas.
    """

    if cashflows is None:
        cashflows = []

    if end_date <= start_date:
        raise ValueError("A data final deve ser posterior à data inicial.")

    effective_annual_rate = (annual_cdi_rate * (cdi_percentage / 100)) - annual_fee

    if effective_annual_rate < 0:
        effective_annual_rate = 0

    daily_rate = annual_to_daily_rate(effective_annual_rate)

    balance = initial_amount
    total_contributions = initial_amount
    total_withdrawals = 0.0

    timeline_rows = []

    all_dates = pd.date_range(start=start_date, end=end_date, freq="D")

    cashflow_df = pd.DataFrame(cashflows)

    if not cashflow_df.empty:
        cashflow_df["data"] = pd.to_datetime(cashflow_df["data"]).dt.date

    for current_timestamp in all_dates:
        current_date = current_timestamp.date()

        day_contribution = 0.0
        day_withdrawal = 0.0

        if not cashflow_df.empty:
            day_events = cashflow_df[cashflow_df["data"] == current_date]

            for _, event in day_events.iterrows():
                value = float(event["valor"])
                event_type = event["tipo"]

                if event_type == "Aporte":
                    balance += value
                    total_contributions += value
                    day_contribution += value

                elif event_type == "Resgate":
                    withdrawal = min(value, balance)
                    balance -= withdrawal
                    total_withdrawals += withdrawal
                    day_withdrawal += withdrawal

        opening_balance_after_cashflow = balance

        daily_profit = balance * daily_rate
        balance += daily_profit

        timeline_rows.append(
            {
                "Data": current_date,
                "Produto": product_name,
                "Aportes": day_contribution,
                "Resgates": day_withdrawal,
                "Rendimento Bruto Diário": daily_profit,
                "Saldo Bruto": balance,
                "Saldo Após Movimentação": opening_balance_after_cashflow,
            }
        )

    gross_value = balance
    gross_profit = gross_value + total_withdrawals - total_contributions

    total_days = (end_date - start_date).days
    tax_value = calculate_tax(gross_profit, total_days, taxable)

    net_value = gross_value - tax_value
    net_profit = net_value + total_withdrawals - total_contributions

    net_return_period = (
        (net_profit / total_contributions) * 100
        if total_contributions > 0
        else 0
    )

    total_months = max(total_days / 30, 1)

    net_return_monthly = (
        ((1 + net_return_period / 100) ** (1 / total_months) - 1) * 100
        if net_return_period > -100
        else 0
    )

    net_return_annual = ((1 + net_return_monthly / 100) ** 12 - 1) * 100

    timeline_df = pd.DataFrame(timeline_rows)

    monthly_summary = (
        timeline_df.assign(Mês=timeline_df["Data"].apply(lambda x: x.strftime("%Y-%m")))
        .groupby(["Produto", "Mês"], as_index=False)
        .agg(
            {
                "Aportes": "sum",
                "Resgates": "sum",
                "Rendimento Bruto Diário": "sum",
                "Saldo Bruto": "last",
            }
        )
        .rename(
            columns={
                "Rendimento Bruto Diário": "Rendimento Bruto no Mês",
                "Saldo Bruto": "Saldo Bruto Final",
            }
        )
    )

    return {
        "Produto": product_name,
        "% CDI": cdi_percentage,
        "Taxa Efetiva a.a.": effective_annual_rate,
        "Valor Inicial": initial_amount,
        "Total Aportado": total_contributions,
        "Total Resgatado": total_withdrawals,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_profit,
        "IR": tax_value,
        "Alíquota IR": get_ir_rate(total_days) * 100 if taxable else 0,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_profit,
        "Rentabilidade Líquida no Período (%)": net_return_period,
        "Rentabilidade Líquida ao Mês (%)": net_return_monthly,
        "Rentabilidade Líquida ao Ano (%)": net_return_annual,
        "Tributável": "Sim" if taxable else "Não",
        "Evolução Diária": timeline_rows,
        "Resumo Mensal": monthly_summary.to_dict("records"),
    }


def simulate_savings_with_cashflows(
    initial_amount: float,
    start_date: date,
    end_date: date,
    selic_rate: float,
    tr_rate: float,
    cashflows: list[dict] | None = None,
) -> dict:
    """
    Simulação simplificada de poupança com movimentações por data.
    """

    if selic_rate > 8.5:
        annual_rate = 6.17 + tr_rate
    else:
        annual_rate = (selic_rate * 0.70) + tr_rate

    return simulate_product_with_cashflows(
        product_name="Poupança",
        initial_amount=initial_amount,
        start_date=start_date,
        end_date=end_date,
        annual_cdi_rate=annual_rate,
        cdi_percentage=100,
        taxable=False,
        annual_fee=0.0,
        cashflows=cashflows,
    )