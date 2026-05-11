from services.fund_tax_service import FundTaxService


def print_result(title, result):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)
    print("Produto:", result.product_name)
    print("Tipo:", result.fund_type)
    print("Valor aplicado:", round(result.initial_amount, 2))
    print("Valor bruto final:", round(result.gross_final_amount, 2))
    print("Valor líquido final:", round(result.net_final_amount, 2))
    print("Lucro líquido:", round(result.net_profit, 2))
    print("Impacto taxa adm:", round(result.admin_fee_impact, 2))
    print("Come-cotas:", round(result.come_cotas_tax, 2))
    print("IR no resgate:", round(result.redemption_tax, 2))
    print("IR total:", round(result.total_tax, 2))
    print("Rentabilidade bruta %:", round(result.gross_return_percentage, 2))
    print("Rentabilidade líquida %:", round(result.net_return_percentage, 2))

    print("\nEventos de come-cotas:")
    events_found = False

    for event in result.events:
        if event["come_cotas_tax"] > 0:
            events_found = True
            print(
                "Mês simulação:",
                event["month_index"],
                "| mês calendário:",
                event["month"],
                "| ano:",
                event["year"],
                "| imposto:",
                round(event["come_cotas_tax"], 2),
            )

    if not events_found:
        print("Nenhum evento de come-cotas no período.")


service = FundTaxService()

scenarios = [
    {
        "title": "1. Fundo DI longo prazo | 24 meses | com come-cotas",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.005,
            "months": 24,
            "fund_type": FundTaxService.LONG_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": True,
        },
    },
    {
        "title": "2. Fundo DI curto prazo | 24 meses | com come-cotas",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.005,
            "months": 24,
            "fund_type": FundTaxService.SHORT_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": True,
        },
    },
    {
        "title": "3. Fundo DI longo prazo | 24 meses | sem come-cotas",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.005,
            "months": 24,
            "fund_type": FundTaxService.LONG_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": False,
        },
    },
    {
        "title": "4. Fundo DI longo prazo | 6 meses | com come-cotas",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.005,
            "months": 6,
            "fund_type": FundTaxService.LONG_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": True,
        },
    },
    {
        "title": "5. Fundo DI longo prazo | 36 meses | com come-cotas",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.005,
            "months": 36,
            "fund_type": FundTaxService.LONG_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": True,
        },
    },
    {
        "title": "6. Fundo DI longo prazo | 24 meses | taxa adm zero",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.00,
            "months": 24,
            "fund_type": FundTaxService.LONG_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": True,
        },
    },
    {
        "title": "7. Fundo DI longo prazo | 24 meses | taxa adm 1,50%",
        "params": {
            "initial_amount": 100000,
            "annual_cdi_rate": 0.145,
            "fund_cdi_percentage": 1.00,
            "admin_fee_rate": 0.015,
            "months": 24,
            "fund_type": FundTaxService.LONG_TERM,
            "start_year": 2026,
            "start_month": 1,
            "apply_come_cotas": True,
        },
    },
]

for scenario in scenarios:
    result = service.simulate_fund_di(**scenario["params"])
    print_result(scenario["title"], result)