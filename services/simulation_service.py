# =========================================================
# SIMULATION SERVICE
# =========================================================

import pandas as pd

from calculations.cdi_calculator import (
    simulate_cdi_product,
    simulate_savings,
)

try:
    from calculations.cashflow_calculator import (
        simulate_product_with_cashflows,
        simulate_savings_with_cashflows,
    )
except Exception:
    simulate_product_with_cashflows = None
    simulate_savings_with_cashflows = None


def _build_comparison_row(item: dict) -> dict:
    """
    Padroniza as colunas esperadas pelo dashboard.
    """

    return {
        "Produto": item.get("Produto"),
        "% CDI": item.get("% CDI", 0.0),
        "Taxa Efetiva a.a. (%)": item.get("Taxa Efetiva a.a.", 0.0),
        "Valor Inicial": item.get("Valor Inicial", 0.0),
        "Total Aportado": item.get("Total Aportado", item.get("Valor Investido", 0.0)),
        "Total Resgatado": item.get("Total Resgatado", 0.0),
        "Valor Bruto": item.get("Valor Bruto", 0.0),
        "Rendimento Bruto": item.get("Rendimento Bruto", 0.0),
        "IR": item.get("IR", 0.0),
        "Alíquota IR (%)": item.get("Alíquota IR", 0.0),
        "Valor Líquido": item.get("Valor Líquido", 0.0),
        "Rendimento Líquido": item.get("Rendimento Líquido", 0.0),
        "Rentab. Líq. Período (%)": item.get("Rentabilidade Líquida no Período (%)", 0.0),
        "Rentab. Líq. Mês (%)": item.get("Rentabilidade Líquida ao Mês (%)", 0.0),
        "Rentab. Líq. Ano (%)": item.get("Rentabilidade Líquida ao Ano (%)", 0.0),
        "Tributável": item.get("Tributável", "Não"),
    }


def run_cdi_simulation(
    initial_amount: float,
    monthly_contribution: float,
    months: int,
    annual_cdi_rate: float,
    selic_rate: float,
    tr_rate: float,
    cdb_percentage: float,
    lci_lca_percentage: float,
    treasury_percentage: float,
    fund_percentage: float,
    fund_annual_fee: float,
    treasury_annual_fee: float,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Executa a simulação comparativa simples do módulo CDI.

    Usada para:
    - Sem aportes adicionais
    - Aporte mensal fixo

    Motor corrigido para juros compostos.
    """

    results = []

    results.append(
        simulate_cdi_product(
            product_name="LCI / LCA",
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=lci_lca_percentage,
            taxable=False,
            annual_fee=0.0,
        )
    )

    results.append(
        simulate_cdi_product(
            product_name="CDB / LC",
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=cdb_percentage,
            taxable=True,
            annual_fee=0.0,
        )
    )

    results.append(
        simulate_cdi_product(
            product_name="Tesouro Selic",
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=treasury_percentage,
            taxable=True,
            annual_fee=treasury_annual_fee,
        )
    )

    results.append(
        simulate_cdi_product(
            product_name="Fundo DI",
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=fund_percentage,
            taxable=True,
            annual_fee=fund_annual_fee,
        )
    )

    results.append(
        simulate_savings(
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            selic_rate=selic_rate,
            tr_rate=tr_rate,
        )
    )

    comparison_rows = []
    evolution_rows = []

    for item in results:
        comparison_rows.append(_build_comparison_row(item))

        for row in item.get("Evolução Mensal", []):
            evolution_rows.append(
                {
                    "Data": row.get("Mês"),
                    "Mês": row.get("Mês"),
                    "Produto": row.get("Produto"),
                    "Saldo Bruto": row.get("Saldo Bruto"),
                }
            )

    comparison_df = pd.DataFrame(comparison_rows)
    evolution_df = pd.DataFrame(evolution_rows)

    comparison_df = comparison_df.sort_values(
        by="Valor Líquido",
        ascending=False,
    ).reset_index(drop=True)

    return comparison_df, evolution_df


def run_cdi_cashflow_simulation(
    initial_amount: float,
    start_date,
    end_date,
    annual_cdi_rate: float,
    selic_rate: float,
    tr_rate: float,
    cdb_percentage: float,
    lci_lca_percentage: float,
    treasury_percentage: float,
    fund_percentage: float,
    fund_annual_fee: float,
    treasury_annual_fee: float,
    cashflows: list[dict],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Executa simulação CDI com calendário de aportes e resgates.

    Mantém compatibilidade com o módulo de cashflow existente.
    """

    if simulate_product_with_cashflows is None or simulate_savings_with_cashflows is None:
            return run_cdi_simulation(
        initial_amount=initial_amount,
        months=12,
        annual_cdi_rate=annual_cdi_rate,
        selic_rate=selic_rate,
        tr_rate=tr_rate,
        cdb_percentage=cdb_percentage,
        lci_lca_percentage=lci_lca_percentage,
        treasury_percentage=treasury_percentage,
        treasury_annual_fee=treasury_annual_fee,
        fund_percentage=fund_percentage,
        fund_annual_fee=fund_annual_fee,
    )

    results = []

    results.append(
        simulate_product_with_cashflows(
            product_name="LCI / LCA",
            initial_amount=initial_amount,
            start_date=start_date,
            end_date=end_date,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=lci_lca_percentage,
            taxable=False,
            annual_fee=0.0,
            cashflows=cashflows,
        )
    )

    results.append(
        simulate_product_with_cashflows(
            product_name="CDB / LC",
            initial_amount=initial_amount,
            start_date=start_date,
            end_date=end_date,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=cdb_percentage,
            taxable=True,
            annual_fee=0.0,
            cashflows=cashflows,
        )
    )

    results.append(
        simulate_product_with_cashflows(
            product_name="Tesouro Selic",
            initial_amount=initial_amount,
            start_date=start_date,
            end_date=end_date,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=treasury_percentage,
            taxable=True,
            annual_fee=treasury_annual_fee,
            cashflows=cashflows,
        )
    )

    results.append(
        simulate_product_with_cashflows(
            product_name="Fundo DI",
            initial_amount=initial_amount,
            start_date=start_date,
            end_date=end_date,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=fund_percentage,
            taxable=True,
            annual_fee=fund_annual_fee,
            cashflows=cashflows,
        )
    )

    results.append(
        simulate_savings_with_cashflows(
            initial_amount=initial_amount,
            start_date=start_date,
            end_date=end_date,
            selic_rate=selic_rate,
            tr_rate=tr_rate,
            cashflows=cashflows,
        )
    )

    comparison_rows = []
    daily_rows = []
    monthly_rows = []

    for item in results:
        comparison_rows.append(_build_comparison_row(item))
        daily_rows.extend(item.get("Evolução Diária", []))
        monthly_rows.extend(item.get("Resumo Mensal", []))

    comparison_df = pd.DataFrame(comparison_rows)
    daily_df = pd.DataFrame(daily_rows)
    monthly_df = pd.DataFrame(monthly_rows)

    comparison_df = comparison_df.sort_values(
        by="Valor Líquido",
        ascending=False,
    ).reset_index(drop=True)

    return comparison_df, daily_df, monthly_df