import pandas as pd


def generate_consultive_analysis(comparison_df: pd.DataFrame) -> str:
    """
    Gera uma leitura consultiva simples a partir da tabela comparativa.
    """

    if comparison_df.empty:
        return "Não há dados suficientes para gerar a análise consultiva."

    best = comparison_df.iloc[0]
    worst = comparison_df.iloc[-1]

    best_product = best["Produto"]
    best_net_value = best["Valor Líquido"]
    best_net_profit = best["Rendimento Líquido"]
    best_return = best["Rentab. Líq. Período (%)"]

    worst_product = worst["Produto"]
    difference = best["Valor Líquido"] - worst["Valor Líquido"]

    analysis = f"""
A simulação indica que, para as premissas informadas, a alternativa com maior valor líquido projetado é **{best_product}**.

O valor líquido estimado para essa alternativa é de **R$ {best_net_value:,.2f}**, com rendimento líquido aproximado de **R$ {best_net_profit:,.2f}** no período. A rentabilidade líquida projetada no período é de **{best_return:.2f}%**.

A diferença entre a melhor alternativa e a alternativa com menor valor líquido, neste cenário, é de aproximadamente **R$ {difference:,.2f}**.

A comparação é especialmente relevante porque produtos com percentuais diferentes do CDI podem ter resultados líquidos próximos ou até superiores, dependendo da tributação, da isenção fiscal, das taxas e do prazo da aplicação.

Esta simulação deve ser utilizada como apoio à análise consultiva, considerando perfil do cliente, liquidez, objetivo financeiro, horizonte de investimento e adequação do produto.
"""

    return analysis