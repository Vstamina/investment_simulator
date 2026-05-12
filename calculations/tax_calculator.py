# =========================================================
# TAX CALCULATOR
# =========================================================

def get_ir_rate(days: int) -> float:
    """
    Retorna a alíquota regressiva de IR para renda fixa.

    Tabela regressiva:
    - Até 180 dias: 22,5%
    - De 181 a 360 dias: 20,0%
    - De 361 a 720 dias: 17,5%
    - Acima de 720 dias: 15,0%
    """

    days = int(days or 0)

    if days <= 180:
        return 22.5

    if days <= 360:
        return 20.0

    if days <= 720:
        return 17.5

    return 15.0


def calculate_tax(
    gross_profit: float | None = None,
    gross_income: float | None = None,
    days: int = 0,
    taxable: bool = True,
) -> tuple[float, float]:
    """
    Calcula o IR sobre rendimento positivo.

    Aceita gross_profit e gross_income para compatibilidade com:
    - cdi_calculator.py
    - cashflow_calculator.py

    Retorna:
    - valor do IR
    - alíquota aplicada
    """

    if gross_profit is None:
        gross_profit = gross_income if gross_income is not None else 0.0

    gross_profit = float(gross_profit or 0.0)

    if not taxable:
        return 0.0, 0.0

    if gross_profit <= 0:
        return 0.0, 0.0

    ir_rate = get_ir_rate(days)
    ir_value = gross_profit * (ir_rate / 100)

    return ir_value, ir_rate


def calculate_ir(
    gross_profit: float | None = None,
    gross_income: float | None = None,
    days: int = 0,
    taxable: bool = True,
) -> tuple[float, float]:
    """
    Alias de compatibilidade para chamadas antigas.
    """

    return calculate_tax(
        gross_profit=gross_profit,
        gross_income=gross_income,
        days=days,
        taxable=taxable,
    )