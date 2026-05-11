def get_ir_rate(days: int) -> float:
    """
    Retorna a alíquota de IR regressivo para renda fixa.
    """
    if days <= 180:
        return 0.225
    if days <= 360:
        return 0.20
    if days <= 720:
        return 0.175
    return 0.15


def calculate_tax(gross_profit: float, days: int, taxable: bool = True) -> float:
    """
    Calcula o imposto sobre o rendimento bruto.
    Produtos isentos retornam imposto zero.
    """
    if not taxable or gross_profit <= 0:
        return 0.0

    ir_rate = get_ir_rate(days)
    return gross_profit * ir_rate