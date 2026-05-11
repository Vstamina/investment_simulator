from io import BytesIO

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor


# =========================================================
# FUNÇÕES AUXILIARES DE FORMATAÇÃO
# =========================================================

def format_currency(value):
    try:
        value = float(value)
        formatted = f"R$ {value:,.2f}"
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        return formatted
    except Exception:
        return str(value)


def format_percent(value):
    try:
        value = float(value)
        return f"{value:.2f}%".replace(".", ",")
    except Exception:
        return str(value)


def format_number(value):
    try:
        value = float(value)
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def safe_text(value):
    if value is None:
        return ""
    return str(value)


def add_paragraph(document, text, bold=False, font_size=10):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    run = paragraph.add_run(safe_text(text))
    run.bold = bold
    run.font.size = Pt(font_size)

    return paragraph


def add_section_heading(document, number, title):
    paragraph = document.add_paragraph()
    paragraph.style = "Heading 1"

    run = paragraph.add_run(f"{number}. {title}")
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(31, 78, 121)

    return paragraph


def add_subheading(document, title):
    paragraph = document.add_paragraph()
    paragraph.style = "Heading 2"

    run = paragraph.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 78, 121)

    return paragraph


def set_cell_text(cell, text, bold=False, font_size=8):
    cell.text = ""

    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run(safe_text(text))
    run.bold = bold
    run.font.size = Pt(font_size)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_key_value_table(document, rows):
    table = document.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    for label, value in rows:
        row_cells = table.add_row().cells
        set_cell_text(row_cells[0], label, bold=True, font_size=8)
        set_cell_text(row_cells[1], value, font_size=8)

    document.add_paragraph()
    return table


def add_dataframe_table(document, df, max_rows=None, font_size=7):
    if df is None or df.empty:
        add_paragraph(document, "Não há dados disponíveis para esta seção.")
        return None

    table_df = df.copy()

    if max_rows is not None:
        table_df = table_df.head(max_rows)

    table = document.add_table(rows=1, cols=len(table_df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    header_cells = table.rows[0].cells

    for index, column in enumerate(table_df.columns):
        set_cell_text(
            header_cells[index],
            str(column),
            bold=True,
            font_size=font_size
        )

    for _, row in table_df.iterrows():
        row_cells = table.add_row().cells

        for index, value in enumerate(row):
            set_cell_text(
                row_cells[index],
                safe_text(value),
                font_size=font_size
            )

    document.add_paragraph()
    return table


def prepare_dataframe_for_word(df):
    if df is None or df.empty:
        return df

    prepared_df = df.copy()

    for column in prepared_df.columns:
        if prepared_df[column].dtype == "float64" or prepared_df[column].dtype == "int64":
            if any(keyword in column.lower() for keyword in ["valor", "saldo", "aporte", "resgate", "rendimento", "ir", "bruto", "líquido", "liquido"]):
                prepared_df[column] = prepared_df[column].apply(format_currency)
            elif any(keyword in column.lower() for keyword in ["%", "taxa", "rentab", "cdi"]):
                prepared_df[column] = prepared_df[column].apply(format_percent)
            else:
                prepared_df[column] = prepared_df[column].apply(format_number)

    return prepared_df


def create_curve_chart_image(curve_df):
    """
    Gera o gráfico da curva em imagem para inserir no Word.
    Se matplotlib não estiver instalado, retorna None e o relatório segue com a tabela.
    """

    try:
        import matplotlib.pyplot as plt
    except Exception:
        return None

    if curve_df is None or curve_df.empty:
        return None

    try:
        chart_df = curve_df.copy()

        chart_df["Vértice"] = chart_df["Vértice"].astype(str)
        chart_df["Taxa Selic Esperada (%)"] = (
            chart_df["Taxa Selic Esperada (%)"].astype(float)
        )

        x_values = chart_df["Vértice"].tolist()
        y_values = chart_df["Taxa Selic Esperada (%)"].tolist()

        selic_reference = y_values[0]

        fig, ax = plt.subplots(figsize=(8.5, 4.3))

        ax.plot(
            x_values,
            y_values,
            marker="o",
            linewidth=2.5,
            label="Selic esperada"
        )

        ax.axhline(
            y=selic_reference,
            linestyle="--",
            linewidth=1.5,
            label="Selic atual"
        )

        for index, value in enumerate(y_values):
            ax.annotate(
                f"{value:.2f}%",
                (x_values[index], y_values[index]),
                textcoords="offset points",
                xytext=(0, 8),
                ha="center",
                fontsize=9
            )

        ax.set_title("Curva Simplificada de Juros")
        ax.set_xlabel("Horizonte da expectativa")
        ax.set_ylabel("Taxa Selic esperada (%)")
        ax.grid(True, axis="y", alpha=0.3)
        ax.legend(loc="best")

        image_stream = BytesIO()
        plt.tight_layout()
        fig.savefig(image_stream, format="png", dpi=180)
        plt.close(fig)

        image_stream.seek(0)
        return image_stream

    except Exception:
        return None


def infer_curve_reading_from_market_intelligence(market_intelligence):
    """
    Garante que o Word consiga gerar a leitura da curva mesmo que o app
    não tenha salvo movimento_curva, spread_final e leitura_movimento.
    """

    if not market_intelligence:
        return None, None, None

    movimento_curva = market_intelligence.get("movimento_curva")
    spread_final = market_intelligence.get("spread_final")
    leitura_movimento = market_intelligence.get("leitura_movimento")

    if movimento_curva and spread_final is not None and leitura_movimento:
        return movimento_curva, spread_final, leitura_movimento

    curve_df = market_intelligence.get("curve_df")

    if curve_df is None or curve_df.empty:
        return movimento_curva, spread_final, leitura_movimento

    try:
        chart_df = curve_df.copy()

        chart_df["Taxa Selic Esperada (%)"] = (
            chart_df["Taxa Selic Esperada (%)"].astype(float)
        )

        selic_atual_referencia = chart_df["Taxa Selic Esperada (%)"].iloc[0]

        spread_final = (
            chart_df["Taxa Selic Esperada (%)"].iloc[-1]
            - selic_atual_referencia
        )

        limite_neutro = 0.10

        if spread_final > limite_neutro:
            movimento_curva = "curva abrindo"
            leitura_movimento = (
                "A curva está abrindo em relação à Selic atual. "
                "Isso indica que as expectativas de mercado apontam para juros futuros "
                "acima da taxa corrente. A leitura pode favorecer uma conversa consultiva "
                "sobre proteção de taxa, prazo, alternativas prefixadas e adequação ao "
                "perfil de liquidez do cliente."
            )

        elif spread_final < -limite_neutro:
            movimento_curva = "curva fechando"
            leitura_movimento = (
                "A curva está fechando em relação à Selic atual. "
                "Isso indica que as expectativas de mercado apontam para juros futuros "
                "abaixo da taxa corrente. A leitura reforça a importância de avaliar "
                "risco de reinvestimento, horizonte da aplicação e momento de travamento "
                "de taxa em alternativas prefixadas ou híbridas."
            )

        else:
            movimento_curva = "curva praticamente estável"
            leitura_movimento = (
                "A curva está praticamente estável em relação à Selic atual. "
                "Nesse cenário, a análise consultiva deve priorizar liquidez, prazo, "
                "tributação, previsibilidade e aderência ao objetivo financeiro do cliente."
            )

        return movimento_curva, spread_final, leitura_movimento

    except Exception:
        return movimento_curva, spread_final, leitura_movimento


# =========================================================
# FUNÇÃO PRINCIPAL DO RELATÓRIO
# =========================================================

def generate_word_report(
    client_name,
    advisor_name,
    simulation_mode,
    start_date,
    end_date,
    months,
    initial_amount,
    annual_cdi_rate,
    selic_rate,
    tr_rate,
    cdb_percentage,
    lci_lca_percentage,
    treasury_percentage,
    treasury_annual_fee,
    fund_percentage,
    fund_annual_fee,
    comparison_df,
    cashflow_df,
    monthly_df,
    consultive_analysis,
    market_intelligence=None,
    report_options=None,
):
    if report_options is None:
        report_options = {
            "visao_geral": True,
            "premissas": True,
            "comparativo": True,
            "calendario": True,
            "resumo_mensal": True,
            "leitura_consultiva": True,
            "inteligencia_mercado": True,
            "curva_juros": True,
            "leitura_foresight": True,
            "aviso_tecnico": True,
        }

    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    # =========================================================
    # CAPA / CABEÇALHO
    # =========================================================

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run("RELATÓRIO DE SIMULAÇÃO DE INVESTIMENTOS")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(31, 78, 121)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = subtitle.add_run(
        "Módulo CDI | Comparação de renda fixa, tributação e planejamento de movimentações"
    )
    run.font.size = Pt(10)

    client_line = document.add_paragraph()
    client_line.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = client_line.add_run(
        f"Cliente: {client_name} | Assessor: {advisor_name} | "
        f"Período: {start_date} a {end_date}"
    )
    run.font.size = Pt(9)

    document.add_paragraph()

    section_number = 1

    # =========================================================
    # MELHOR ALTERNATIVA
    # =========================================================

    best_product = ""
    best_net_value = 0
    best_net_profit = 0
    best_net_return = 0

    worst_product = ""
    worst_net_value = 0
    difference_best_worst = 0

    if comparison_df is not None and not comparison_df.empty:
        try:
            value_column = "Valor Líquido"
            profit_column = "Rendimento Líquido"
            return_column = "Rentab. Líq. Período (%)"
            product_column = "Produto"

            sorted_df = comparison_df.sort_values(
                by=value_column,
                ascending=False
            )

            best_row = sorted_df.iloc[0]
            worst_row = sorted_df.iloc[-1]

            best_product = safe_text(best_row.get(product_column, ""))
            best_net_value = float(best_row.get(value_column, 0))
            best_net_profit = float(best_row.get(profit_column, 0))
            best_net_return = float(best_row.get(return_column, 0))

            worst_product = safe_text(worst_row.get(product_column, ""))
            worst_net_value = float(worst_row.get(value_column, 0))

            difference_best_worst = best_net_value - worst_net_value

        except Exception:
            pass

    # =========================================================
    # 1. VISÃO GERAL
    # =========================================================

    if report_options.get("visao_geral", True):
        add_section_heading(
            document,
            section_number,
            "Visão Geral da Simulação"
        )
        section_number += 1

        overview_rows = [
            ("Valor inicial", format_currency(initial_amount)),
            ("Melhor alternativa", best_product),
            ("Valor líquido", format_currency(best_net_value)),
            ("Rentabilidade líquida", format_percent(best_net_return)),
            ("Modo de simulação", safe_text(simulation_mode)),
            ("Período/Prazo", f"{start_date} a {end_date}"),
            ("Rendimento líquido projetado", format_currency(best_net_profit)),
        ]

        if cashflow_df is not None and not cashflow_df.empty:
            try:
                total_aportado = cashflow_df[
                    cashflow_df["Tipo"].str.lower() == "aporte"
                ]["Valor"].astype(float).sum()

                total_resgatado = cashflow_df[
                    cashflow_df["Tipo"].str.lower() == "resgate"
                ]["Valor"].astype(float).sum()

                overview_rows.insert(
                    6,
                    ("Total aportado", format_currency(total_aportado))
                )

                overview_rows.insert(
                    7,
                    ("Total resgatado", format_currency(total_resgatado))
                )

            except Exception:
                pass

        add_key_value_table(document, overview_rows)

    # =========================================================
    # 2. PREMISSAS
    # =========================================================

    if report_options.get("premissas", True):
        add_section_heading(
            document,
            section_number,
            "Premissas Utilizadas"
        )
        section_number += 1

        premises_rows = [
            ("CDI anual estimado", format_percent(annual_cdi_rate)),
            ("Selic meta estimada", format_percent(selic_rate)),
            ("TR anual estimada", format_percent(tr_rate)),
            ("CDB / LC", f"{format_percent(cdb_percentage)} do CDI"),
            ("LCI / LCA", f"{format_percent(lci_lca_percentage)} do CDI"),
            ("Tesouro Selic", f"{format_percent(treasury_percentage)} do CDI"),
            ("Custo anual Tesouro Selic", format_percent(treasury_annual_fee)),
            ("Fundo DI", f"{format_percent(fund_percentage)} do CDI"),
            ("Taxa de administração Fundo DI", format_percent(fund_annual_fee)),
        ]

        add_key_value_table(document, premises_rows)

    # =========================================================
    # 3. COMPARATIVO DOS PRODUTOS
    # =========================================================

    if report_options.get("comparativo", True):
        add_section_heading(
            document,
            section_number,
            "Comparativo dos Produtos"
        )
        section_number += 1

        comparison_word_df = prepare_dataframe_for_word(comparison_df)
        add_dataframe_table(
            document,
            comparison_word_df,
            font_size=6
        )

    # =========================================================
    # INTELIGÊNCIA DE MERCADO
    # =========================================================

    if (
        report_options.get("inteligencia_mercado", True)
        and market_intelligence is not None
    ):
        add_section_heading(
            document,
            section_number,
            "Inteligência de Mercado e Foresight"
        )
        section_number += 1

        add_paragraph(
            document,
            "Os dados públicos do Banco Central e as expectativas do Boletim Focus "
            "foram incorporados à simulação para apoiar a leitura macroeconômica "
            "e qualificar a análise consultiva. Esta camada não altera os cálculos "
            "da simulação CDI, mas amplia a contextualização da decisão."
        )

        curve_shape = market_intelligence.get("curve_shape")
        market_reading = market_intelligence.get("reading")

        if curve_shape:
            add_paragraph(
                document,
                f"Classificação da curva: {curve_shape}.",
                bold=True
            )

        if market_reading:
            add_paragraph(
                document,
                safe_text(market_reading)
            )

    # =========================================================
    # CURVA DE JUROS
    # =========================================================

    if (
        report_options.get("curva_juros", True)
        and market_intelligence is not None
    ):
        curve_df = market_intelligence.get("curve_df")

        if curve_df is not None and not curve_df.empty:
            add_section_heading(
                document,
                section_number,
                "Curva Simplificada de Juros"
            )
            section_number += 1

            chart_image = create_curve_chart_image(curve_df)

            if chart_image is not None:
                paragraph = document.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run()
                run.add_picture(chart_image, width=Inches(6.5))

            curve_word_df = prepare_dataframe_for_word(curve_df)

            add_subheading(document, "Tabela técnica da curva")
            add_dataframe_table(
                document,
                curve_word_df,
                font_size=8
            )

    # =========================================================
    # LEITURA FORESIGHT DA CURVA
    # =========================================================

    if (
        report_options.get("leitura_foresight", True)
        and market_intelligence is not None
    ):
        movimento_curva, spread_final, leitura_movimento = (
            infer_curve_reading_from_market_intelligence(market_intelligence)
        )

        if movimento_curva:
            add_section_heading(
                document,
                section_number,
                "Leitura Foresight da Curva"
            )
            section_number += 1

            add_paragraph(
                document,
                f"Em relação à Selic atual, a estrutura observada indica "
                f"{movimento_curva}.",
                bold=True
            )

            if spread_final is not None:
                add_paragraph(
                    document,
                    f"O spread do último vértice frente à taxa corrente é de "
                    f"{float(spread_final):.2f} ponto percentual."
                )

            if leitura_movimento:
                add_paragraph(
                    document,
                    leitura_movimento
                )

    # =========================================================
    # CALENDÁRIO DE MOVIMENTAÇÕES
    # =========================================================

    if report_options.get("calendario", True):
        if cashflow_df is not None and not cashflow_df.empty:
            add_section_heading(
                document,
                section_number,
                "Calendário de Movimentações"
            )
            section_number += 1

            cashflow_word_df = prepare_dataframe_for_word(cashflow_df)
            add_dataframe_table(
                document,
                cashflow_word_df,
                font_size=8
            )

    # =========================================================
    # RESUMO MENSAL
    # =========================================================

    if report_options.get("resumo_mensal", True):
        if monthly_df is not None and not monthly_df.empty:
            add_section_heading(
                document,
                section_number,
                "Resumo Mensal"
            )
            section_number += 1

            monthly_word_df = prepare_dataframe_for_word(monthly_df)

            add_dataframe_table(
                document,
                monthly_word_df,
                font_size=6
            )

    # =========================================================
    # LEITURA CONSULTIVA
    # =========================================================

    if report_options.get("leitura_consultiva", True):
        add_section_heading(
            document,
            section_number,
            "Leitura Consultiva"
        )
        section_number += 1

        if consultive_analysis:
            add_paragraph(
                document,
                consultive_analysis
            )
        else:
            add_paragraph(
                document,
                f"A simulação indica que, para as premissas informadas, "
                f"a alternativa com maior valor líquido projetado é {best_product}."
            )

            add_paragraph(
                document,
                f"O valor líquido estimado para essa alternativa é de "
                f"{format_currency(best_net_value)}, com rendimento líquido aproximado "
                f"de {format_currency(best_net_profit)} no período. A rentabilidade "
                f"líquida projetada no período é de {format_percent(best_net_return)}."
            )

            if difference_best_worst:
                add_paragraph(
                    document,
                    f"A diferença entre a melhor alternativa e a alternativa com menor "
                    f"valor líquido, neste cenário, é de aproximadamente "
                    f"{format_currency(difference_best_worst)}."
                )

            add_paragraph(
                document,
                "A comparação é especialmente relevante porque produtos com percentuais "
                "diferentes do CDI podem ter resultados líquidos próximos ou até superiores, "
                "dependendo da tributação, da isenção fiscal, das taxas e do prazo da aplicação."
            )

            add_paragraph(
                document,
                "Esta simulação deve ser utilizada como apoio à análise consultiva, "
                "considerando perfil do cliente, liquidez, objetivo financeiro, horizonte "
                "de investimento e adequação do produto."
            )

    # =========================================================
    # AVISO TÉCNICO
    # =========================================================

    if report_options.get("aviso_tecnico", True):
        add_section_heading(
            document,
            section_number,
            "Aviso Técnico"
        )
        section_number += 1

        add_paragraph(
            document,
            "Os valores apresentados são estimativas para fins de simulação. "
            "Este relatório não constitui promessa de rentabilidade, oferta, "
            "recomendação individualizada ou garantia de resultado. A análise final "
            "deve considerar perfil do cliente, adequação, liquidez, risco, tributação, "
            "condições de mercado e normas aplicáveis."
        )

    # =========================================================
    # RODAPÉ
    # =========================================================

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = footer.add_run(
        "Relatório gerado automaticamente pelo Simulador Estratégico de Investimentos | "
        "Uso interno e consultivo"
    )
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(100, 100, 100)

    file_stream = BytesIO()
    document.save(file_stream)
    file_stream.seek(0)

    return file_stream