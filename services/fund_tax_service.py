from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional


@dataclass
class FundSimulationResult:
    product_name: str
    fund_type: str
    initial_amount: float
    gross_final_amount: float
    net_final_amount: float
    gross_profit: float
    net_profit: float
    admin_fee_impact: float
    come_cotas_tax: float
    redemption_tax: float
    total_tax: float
    net_return_percentage: float
    gross_return_percentage: float
    events: List[Dict]


class FundTaxService:
    """
    Serviço de simulação tributária para Fundo DI.

    Regras contempladas:
    - Fundo DI de curto prazo
    - Fundo DI de longo prazo
    - Taxa de administração anual
    - Come-cotas em maio e novembro
    - IR complementar no resgate, quando aplicável

    Observação:
    Esta é uma simulação gerencial simplificada para apoio consultivo.
    Não substitui cálculo fiscal oficial, informe de rendimentos ou análise tributária individualizada.
    """

    LONG_TERM = "Fundo de longo prazo"
    SHORT_TERM = "Fundo de curto prazo"

    def get_final_ir_rate(self, days: int, fund_type: str) -> float:
        """
        Retorna a alíquota final de IR conforme prazo e tipo de fundo.
        """

        if fund_type == self.SHORT_TERM:
            if days <= 180:
                return 0.225
            return 0.20

        if fund_type == self.LONG_TERM:
            if days <= 180:
                return 0.225
            if days <= 360:
                return 0.20
            if days <= 720:
                return 0.175
            return 0.15

        raise ValueError("Tipo de fundo inválido.")

    def get_come_cotas_rate(self, fund_type: str) -> float:
        """
        Retorna a alíquota de come-cotas conforme classificação fiscal do fundo.
        """

        if fund_type == self.SHORT_TERM:
            return 0.20

        if fund_type == self.LONG_TERM:
            return 0.15

        raise ValueError("Tipo de fundo inválido.")

    def is_come_cotas_month(self, month: int) -> bool:
        """
        Come-cotas ocorre em maio e novembro.
        """

        return month in [5, 11]

    def simulate_fund_di(
        self,
        initial_amount: float,
        annual_cdi_rate: float,
        fund_cdi_percentage: float,
        admin_fee_rate: float,
        months: int,
        fund_type: str = LONG_TERM,
        start_year: Optional[int] = None,
        start_month: Optional[int] = None,
        apply_come_cotas: bool = True,
    ) -> FundSimulationResult:
        """
        Simula Fundo DI com taxa de administração e come-cotas.

        Parâmetros:
        - initial_amount: valor aplicado
        - annual_cdi_rate: CDI anual em decimal. Ex.: 0.145 para 14,5%
        - fund_cdi_percentage: percentual do CDI em decimal. Ex.: 1.00 para 100%
        - admin_fee_rate: taxa de administração anual em decimal. Ex.: 0.005 para 0,50%
        - months: prazo em meses
        - fund_type: Fundo de longo prazo ou Fundo de curto prazo
        - start_year/start_month: mês inicial da simulação
        - apply_come_cotas: permite ligar/desligar o come-cotas
        """

        if initial_amount <= 0:
            raise ValueError("O valor aplicado deve ser maior que zero.")

        if months <= 0:
            raise ValueError("O prazo deve ser maior que zero.")

        if start_year is None:
            start_year = date.today().year

        if start_month is None:
            start_month = date.today().month

        balance = float(initial_amount)
        gross_balance_without_admin_fee = float(initial_amount)

        come_cotas_rate = self.get_come_cotas_rate(fund_type)

        total_come_cotas_tax = 0.0
        total_admin_fee_impact = 0.0
        last_taxed_balance = float(initial_amount)

        events: List[Dict] = []

        # Conversões mensais aproximadas
        monthly_cdi_rate = (1 + annual_cdi_rate) ** (1 / 12) - 1
        monthly_fund_gross_rate = monthly_cdi_rate * fund_cdi_percentage

        # Taxa de administração mensal aproximada
        monthly_admin_fee_rate = (1 + admin_fee_rate) ** (1 / 12) - 1

        current_year = start_year
        current_month = start_month

        for month_index in range(1, months + 1):
            opening_balance = balance

            # Rendimento bruto do fundo no mês
            gross_income = balance * monthly_fund_gross_rate
            balance += gross_income

            # Controle de saldo bruto sem taxa de administração
            gross_balance_without_admin_fee *= (1 + monthly_fund_gross_rate)

            # Desconto da taxa de administração
            admin_fee_value = balance * monthly_admin_fee_rate
            balance -= admin_fee_value
            total_admin_fee_impact += admin_fee_value

            event = {
                "month_index": month_index,
                "year": current_year,
                "month": current_month,
                "opening_balance": opening_balance,
                "gross_income": gross_income,
                "admin_fee": admin_fee_value,
                "come_cotas_tax": 0.0,
                "closing_balance": balance,
            }

            # Aplicação do come-cotas em maio e novembro
            if apply_come_cotas and self.is_come_cotas_month(current_month):
                taxable_income_since_last_event = max(
                    0.0,
                    balance - last_taxed_balance
                )

                come_cotas_tax = taxable_income_since_last_event * come_cotas_rate

                balance -= come_cotas_tax
                total_come_cotas_tax += come_cotas_tax

                # Após o come-cotas, este passa a ser o novo saldo-base tributado
                last_taxed_balance = balance

                event["come_cotas_tax"] = come_cotas_tax
                event["closing_balance"] = balance

            events.append(event)

            # Avança o mês
            current_month += 1

            if current_month > 12:
                current_month = 1
                current_year += 1

        # Prazo aproximado em dias
        days = months * 30

        final_ir_rate = self.get_final_ir_rate(days, fund_type)

        # IR total teórico sobre o rendimento total
        gross_profit_for_tax = max(0.0, balance - initial_amount)
        total_tax_due = gross_profit_for_tax * final_ir_rate

        # Complemento no resgate
        redemption_tax = max(0.0, total_tax_due - total_come_cotas_tax)

        net_final_amount = balance - redemption_tax

        gross_final_amount = gross_balance_without_admin_fee
        gross_profit = gross_final_amount - initial_amount
        net_profit = net_final_amount - initial_amount

        total_tax = total_come_cotas_tax + redemption_tax

        gross_return_percentage = (
            gross_profit / initial_amount
        ) * 100

        net_return_percentage = (
            net_profit / initial_amount
        ) * 100

        return FundSimulationResult(
            product_name="Fundo DI",
            fund_type=fund_type,
            initial_amount=initial_amount,
            gross_final_amount=gross_final_amount,
            net_final_amount=net_final_amount,
            gross_profit=gross_profit,
            net_profit=net_profit,
            admin_fee_impact=total_admin_fee_impact,
            come_cotas_tax=total_come_cotas_tax,
            redemption_tax=redemption_tax,
            total_tax=total_tax,
            net_return_percentage=net_return_percentage,
            gross_return_percentage=gross_return_percentage,
            events=events,
        )