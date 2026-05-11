import pandas as pd

from calculations.cdi_calculator import simulate_cdi_product, simulate_savings
from calculations.cashflow_calculator import (
    simulate_product_with_cashflows,
    simulate_savings_with_cashflows,
)


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
    Mantida para o modo de aporte mensal fixo.
    """

    results = []

    results.append(
        simulate_savings(
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            selic_rate=selic_rate,
            tr_rate=tr_rate,
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
            annual_fee=0,
        )
    )

    results.append(
        simulate_cdi_product(
            product_name="LCI / LCA",
            initial_amount=initial_amount,
            monthly_contribution=monthly_contribution,
            months=months,
            annual_cdi_rate=annual_cdi_rate,
            cdi_percentage=lci_lca_percentage,
            taxable=False,
            annual_fee=0,
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

    comparison_rows = []
    evolution_rows = []

    for item in results:
        comparison_rows.append(
            {
                "Produto": item["Produto"],
                "% CDI": item["% CDI"],
                "Taxa Efetiva a.a. (%)": item["Taxa Efetiva a.a."],
                "Valor Inicial": initial_amount,
                "Total Aportado": item["Valor Investido"],
                "Total Resgatado": 0.0,
                "Valor Bruto": item["Valor Bruto"],
                "Rendimento Bruto": item["Rendimento Bruto"],
                "IR": item["IR"],
                "Alíquota IR (%)": item["Alíquota IR"],
                "Valor Líquido": item["Valor Líquido"],
                "Rendimento Líquido": item["Rendimento Líquido"],
                "Rentab. Líq. Período (%)": item["Rentabilidade Líquida no Período (%)"],
                "Rentab. Líq. Mês (%)": item["Rentabilidade Líquida ao Mês (%)"],
                "Rentab. Líq. Ano (%)": item["Rentabilidade Líquida ao Ano (%)"],
                "Tributável": item["Tributável"],
            }
        )

        for row in item["Evolução Mensal"]:
            evolution_rows.append(
                {
                    "Data": row["Mês"],
                    "Mês": row["Mês"],
                    "Produto": row["Produto"],
                    "Saldo Bruto": row["Saldo Bruto"],
                }
            )

    comparison_df = pd.DataFrame(comparison_rows)
    evolution_df = pd.DataFrame(evolution_rows)

    comparison_df = comparison_df.sort_values(
        by="Valor Líquido",
        ascending=False
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
    """

    results = []

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

    comparison_rows = []
    daily_rows = []
    monthly_rows = []

    for item in results:
        comparison_rows.append(
            {
                "Produto": item["Produto"],
                "% CDI": item["% CDI"],
                "Taxa Efetiva a.a. (%)": item["Taxa Efetiva a.a."],
                "Valor Inicial": item["Valor Inicial"],
                "Total Aportado": item["Total Aportado"],
                "Total Resgatado": item["Total Resgatado"],
                "Valor Bruto": item["Valor Bruto"],
                "Rendimento Bruto": item["Rendimento Bruto"],
                "IR": item["IR"],
                "Alíquota IR (%)": item["Alíquota IR"],
                "Valor Líquido": item["Valor Líquido"],
                "Rendimento Líquido": item["Rendimento Líquido"],
                "Rentab. Líq. Período (%)": item["Rentabilidade Líquida no Período (%)"],
                "Rentab. Líq. Mês (%)": item["Rentabilidade Líquida ao Mês (%)"],
                "Rentab. Líq. Ano (%)": item["Rentabilidade Líquida ao Ano (%)"],
                "Tributável": item["Tributável"],
            }
        )

        daily_rows.extend(item["Evolução Diária"])
        monthly_rows.extend(item["Resumo Mensal"])

    comparison_df = pd.DataFrame(comparison_rows)
    daily_df = pd.DataFrame(daily_rows)
    monthly_df = pd.DataFrame(monthly_rows)

    comparison_df = comparison_df.sort_values(
        by="Valor Líquido",
        ascending=False
    ).reset_index(drop=True)

    return comparison_df, daily_df, monthly_df