from datetime import timedelta

import pandas as pd
from dateutil.relativedelta import relativedelta

from calculations.cdi_calculator import (
    annual_to_business_daily_rate,
    get_effective_annual_rate,
    get_savings_monthly_rate,
    get_savings_annual_rate,
)
from calculations.tax_calculator import calculate_tax
from market_data.business_calendar import (
    count_business_days,
    count_calendar_days,
    get_brazil_national_holidays,
    is_business_day,
)


def normalize_cashflows(cashflows: list[dict]) -> list[dict]:
    """
    Normaliza calendário de aportes e resgates.
    """
    normalized = []

    for item in cashflows:
        cashflow_date = item.get("Data")
        amount = float(item.get("Valor", 0.0) or 0.0)
        movement_type = str(item.get("Tipo", "Aporte")).strip()

        if amount <= 0:
            continue

        if cashflow_date is None:
            continue

        cashflow_date = pd.to_datetime(cashflow_date).date()

        if movement_type.lower() == "resgate":
            movement_type = "Resgate"
        else:
            movement_type = "Aporte"

        normalized.append(
            {
                "Data": cashflow_date,
                "Valor": amount,
                "Tipo": movement_type,
            }
        )

    return normalized


def get_cashflow_amount_for_date(
    cashflows: list[dict],
    current_date,
    movement_type: str
) -> float:
    """
    Soma aportes ou resgates de uma data específica.
    """
    total = 0.0

    for item in cashflows:
        if item["Data"] == current_date and item["Tipo"] == movement_type:
            total += float(item["Valor"])

    return total


def simulate_product_with_cashflows(
    product_name: str,
    initial_amount: float,
    start_date,
    end_date,
    annual_cdi_rate: float,
    cdi_percentage: float,
    taxable: bool,
    annual_fee: float = 0.0,
    cashflows: list[dict] | None = None,
) -> dict:
    """
    Simula produto CDI com calendário de aportes e resgates.

    Motor:
    - capitalização diária composta;
    - aplicação do CDI apenas em dias úteis;
    - base 252 dias úteis;
    - exclusão de sábados, domingos e feriados nacionais;
    - IR calculado sobre rendimento bruto ao final, por dias corridos;
    - rentabilidades retornadas em decimal, não em percentual cheio.
    """
    if cashflows is None:
        cashflows = []

    start_date = pd.to_datetime(start_date).date()
    end_date = pd.to_datetime(end_date).date()

    cashflows = normalize_cashflows(cashflows)

    effective_annual_rate = get_effective_annual_rate(
        annual_cdi_rate=annual_cdi_rate,
        cdi_percentage=cdi_percentage,
        annual_fee=annual_fee,
    )

    daily_rate = annual_to_business_daily_rate(effective_annual_rate)

    calendar_days = count_calendar_days(start_date, end_date)
    business_days = count_business_days(start_date, end_date)

    holiday_set = get_brazil_national_holidays(start_date, end_date)

    balance = float(initial_amount)
    invested_amount = float(initial_amount)
    total_contributed = float(initial_amount)
    total_withdrawn = 0.0

    daily_rows = []

    current_date = start_date

    while current_date <= end_date:
        aportes = get_cashflow_amount_for_date(
            cashflows=cashflows,
            current_date=current_date,
            movement_type="Aporte",
        )

        resgates = get_cashflow_amount_for_date(
            cashflows=cashflows,
            current_date=current_date,
            movement_type="Resgate",
        )

        if aportes > 0:
            balance += aportes
            invested_amount += aportes
            total_contributed += aportes

        if resgates > 0:
            effective_withdrawal = min(resgates, balance)
            balance -= effective_withdrawal
            invested_amount = max(invested_amount - effective_withdrawal, 0.0)
            total_withdrawn += effective_withdrawal

        rendimento_dia = 0.0

        if current_date > start_date and is_business_day(current_date, holiday_set):
            previous_balance = balance
            balance *= (1 + daily_rate)
            rendimento_dia = balance - previous_balance

        daily_rows.append(
            {
                "Data": current_date,
                "Produto": product_name,
                "Aportes": aportes,
                "Resgates": resgates,
                "Rendimento Bruto no Dia": rendimento_dia,
                "Saldo Bruto": balance,
                "Dia Útil": "Sim" if is_business_day(current_date, holiday_set) else "Não",
            }
        )

        current_date += timedelta(days=1)

    gross_value = balance
    gross_income = gross_value - invested_amount

    tax_value, ir_rate = calculate_tax(
        gross_income=gross_income,
        days=calendar_days,
        taxable=taxable,
    )

    net_value = gross_value - tax_value
    net_income = net_value - invested_amount

    # IMPORTANTE:
    # Daqui para baixo, as rentabilidades ficam em DECIMAL.
    # Exemplo: 9,94% = 0.0994.
    # O app.py já transforma em percentual na visualização.
    period_return = net_income / invested_amount if invested_amount else 0.0

    monthly_return = (
        (1 + period_return) ** (30 / calendar_days) - 1
        if calendar_days > 0
        else 0.0
    )

    annual_return = (
        (1 + period_return) ** (365 / calendar_days) - 1
        if calendar_days > 0
        else 0.0
    )

    daily_df = pd.DataFrame(daily_rows)

    if not daily_df.empty:
        daily_df["Mês"] = pd.to_datetime(daily_df["Data"]).dt.to_period("M").astype(str)

        monthly_rows = (
            daily_df.groupby(["Produto", "Mês"], as_index=False)
            .agg(
                {
                    "Aportes": "sum",
                    "Resgates": "sum",
                    "Rendimento Bruto no Dia": "sum",
                    "Saldo Bruto": "last",
                }
            )
            .rename(
                columns={
                    "Rendimento Bruto no Dia": "Rendimento Bruto no Mês",
                    "Saldo Bruto": "Saldo Bruto Final",
                }
            )
            .to_dict("records")
        )
    else:
        monthly_rows = []

    return {
        "Produto": product_name,
        "% CDI": cdi_percentage,
        "Taxa Efetiva a.a.": effective_annual_rate,
        "Dias Úteis": business_days,
        "Dias Corridos": calendar_days,
        "Valor Inicial": initial_amount,
        "Total Aportado": total_contributed,
        "Total Resgatado": total_withdrawn,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_income,
        "IR": tax_value,
        "Alíquota IR": ir_rate,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_income,
        "Rentabilidade Líquida no Período (%)": period_return,
        "Rentabilidade Líquida ao Mês (%)": monthly_return,
        "Rentabilidade Líquida ao Ano (%)": annual_return,
        "Tributável": "Sim" if taxable else "Não",
        "Evolução Diária": daily_rows,
        "Resumo Mensal": monthly_rows,
    }


def simulate_savings_with_cashflows(
    initial_amount: float,
    start_date,
    end_date,
    selic_rate: float,
    tr_rate: float,
    cashflows: list[dict] | None = None,
) -> dict:
    """
    Simula poupança com calendário de aportes e resgates.

    Motor:
    - rendimento por aniversário mensal simplificado;
    - isenta de IR;
    - rentabilidades retornadas em decimal, não em percentual cheio.
    """
    if cashflows is None:
        cashflows = []

    start_date = pd.to_datetime(start_date).date()
    end_date = pd.to_datetime(end_date).date()

    cashflows = normalize_cashflows(cashflows)

    monthly_rate = get_savings_monthly_rate(selic_rate, tr_rate)
    annual_rate = get_savings_annual_rate(selic_rate, tr_rate)

    calendar_days = count_calendar_days(start_date, end_date)
    business_days = count_business_days(start_date, end_date)

    balance = float(initial_amount)
    invested_amount = float(initial_amount)
    total_contributed = float(initial_amount)
    total_withdrawn = 0.0

    daily_rows = []

    current_date = start_date
    next_anniversary = start_date + relativedelta(months=1)

    while current_date <= end_date:
        aportes = get_cashflow_amount_for_date(
            cashflows=cashflows,
            current_date=current_date,
            movement_type="Aporte",
        )

        resgates = get_cashflow_amount_for_date(
            cashflows=cashflows,
            current_date=current_date,
            movement_type="Resgate",
        )

        if aportes > 0:
            balance += aportes
            invested_amount += aportes
            total_contributed += aportes

        if resgates > 0:
            effective_withdrawal = min(resgates, balance)
            balance -= effective_withdrawal
            invested_amount = max(invested_amount - effective_withdrawal, 0.0)
            total_withdrawn += effective_withdrawal

        rendimento_dia = 0.0

        if current_date == next_anniversary:
            previous_balance = balance
            balance *= (1 + monthly_rate)
            rendimento_dia = balance - previous_balance
            next_anniversary = next_anniversary + relativedelta(months=1)

        daily_rows.append(
            {
                "Data": current_date,
                "Produto": "Poupança",
                "Aportes": aportes,
                "Resgates": resgates,
                "Rendimento Bruto no Dia": rendimento_dia,
                "Saldo Bruto": balance,
                "Dia Útil": "-",
            }
        )

        current_date += timedelta(days=1)

    gross_value = balance
    gross_income = gross_value - invested_amount

    net_value = gross_value
    net_income = net_value - invested_amount

    # Rentabilidade em decimal.
    period_return = net_income / invested_amount if invested_amount else 0.0

    monthly_return = (
        (1 + period_return) ** (30 / calendar_days) - 1
        if calendar_days > 0
        else 0.0
    )

    annualized_return = (
        (1 + period_return) ** (365 / calendar_days) - 1
        if calendar_days > 0
        else annual_rate
    )

    daily_df = pd.DataFrame(daily_rows)

    if not daily_df.empty:
        daily_df["Mês"] = pd.to_datetime(daily_df["Data"]).dt.to_period("M").astype(str)

        monthly_rows = (
            daily_df.groupby(["Produto", "Mês"], as_index=False)
            .agg(
                {
                    "Aportes": "sum",
                    "Resgates": "sum",
                    "Rendimento Bruto no Dia": "sum",
                    "Saldo Bruto": "last",
                }
            )
            .rename(
                columns={
                    "Rendimento Bruto no Dia": "Rendimento Bruto no Mês",
                    "Saldo Bruto": "Saldo Bruto Final",
                }
            )
            .to_dict("records")
        )
    else:
        monthly_rows = []

    return {
        "Produto": "Poupança",
        "% CDI": 0.0,
        "Taxa Efetiva a.a.": annual_rate,
        "Dias Úteis": business_days,
        "Dias Corridos": calendar_days,
        "Valor Inicial": initial_amount,
        "Total Aportado": total_contributed,
        "Total Resgatado": total_withdrawn,
        "Valor Bruto": gross_value,
        "Rendimento Bruto": gross_income,
        "IR": 0.0,
        "Alíquota IR": 0.0,
        "Valor Líquido": net_value,
        "Rendimento Líquido": net_income,
        "Rentabilidade Líquida no Período (%)": period_return,
        "Rentabilidade Líquida ao Mês (%)": monthly_return,
        "Rentabilidade Líquida ao Ano (%)": annualized_return,
        "Tributável": "Não",
        "Evolução Diária": daily_rows,
        "Resumo Mensal": monthly_rows,
    }