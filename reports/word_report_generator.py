from io import BytesIO
from datetime import datetime
import pandas as pd

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================================================
# FORMATAÇÃO BÁSICA
# =========================================================

PRIMARY_BLUE = "062B5F"
DARK_BLUE = "061B3A"
MEDIUM_BLUE = "0A84FF"
LIGHT_BLUE = "EAF4FF"
SOFT_GRAY = "F3F6FA"
WHITE = "FFFFFF"
TEXT_GRAY = "4F5B6B"
SUCCESS_GREEN = "168A45"
WARNING_ORANGE = "F5A623"
BORDER_GRAY = "D8E2EF"


def format_currency(value: float) -> str:
    try:
        return f"R$ {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def format_percent(value: float) -> str:
    try:
        return f"{float(value):.2f}%".replace(".", ",")
    except Exception:
        return str(value)


def set_cell_background(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_border(cell, color: str = BORDER_GRAY, size: str = "6") -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()

    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    for edge in ("top", "left", "bottom", "right"):
        tag = "w:{}".format(edge)
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_cell_text(
    cell,
    text: str,
    bold: bool = False,
    font_size: int = 9,
    color: str = "000000",
    alignment=WD_ALIGN_PARAGRAPH.LEFT,
) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = alignment
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


# =========================================================
# COMPONENTES VISUAIS DO RELATÓRIO
# =========================================================

def add_cover_header(document: Document, client_name: str, advisor_name: str, period_text: str) -> None:
    table = document.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for row in table.rows:
        for cell in row.cells:
            set_cell_background(cell, DARK_BLUE)
            set_cell_border(cell, DARK_BLUE)

    title_cell = table.cell(0, 0)
    title_cell.merge(table.cell(0, 1))
    set_cell_text(
        title_cell,
        "RELATÓRIO DE SIMULAÇÃO DE INVESTIMENTOS",
        bold=True,
        font_size=17,
        color=WHITE,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    subtitle_cell = table.cell(1, 0)
    subtitle_cell.merge(table.cell(1, 1))
    set_cell_text(
        subtitle_cell,
        "Módulo CDI | Comparação de renda fixa, tributação e planejamento de movimentações",
        bold=False,
        font_size=9,
        color="C9D6E8",
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    set_cell_text(table.cell(2, 0), f"Cliente: {client_name}", bold=True, font_size=9, color=WHITE)
    set_cell_text(table.cell(2, 1), f"Assessor: {advisor_name} | Período: {period_text}", bold=True, font_size=9, color=WHITE)

    document.add_paragraph()


def add_section_heading(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.space_before = Pt(12)
    paragraph.space_after = Pt(4)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor.from_string(PRIMARY_BLUE)


def add_small_note(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.05
    run = paragraph.add_run(text)
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string(TEXT_GRAY)


def add_normal_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.15
    run = paragraph.add_run(text)
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string("111827")


def add_summary_cards(document: Document, cards: list[dict]) -> None:
    table = document.add_table(rows=1, cols=len(cards))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for idx, card in enumerate(cards):
        cell = table.cell(0, idx)
        set_cell_background(cell, LIGHT_BLUE)
        set_cell_border(cell, "B7D7F5")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        cell.text = ""
        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(card["label"])
        r1.bold = True
        r1.font.size = Pt(7)
        r1.font.color.rgb = RGBColor.from_string(TEXT_GRAY)

        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(card["value"])
        r2.bold = True
        r2.font.size = Pt(11)
        r2.font.color.rgb = RGBColor.from_string(PRIMARY_BLUE)

    document.add_paragraph()


def add_key_value_table(document: Document, data: dict) -> None:
    table = document.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for key, value in data.items():
        row = table.add_row()
        set_cell_background(row.cells[0], SOFT_GRAY)
        set_cell_background(row.cells[1], WHITE)
        set_cell_border(row.cells[0])
        set_cell_border(row.cells[1])
        set_cell_text(row.cells[0], str(key), bold=True, font_size=8, color=PRIMARY_BLUE)
        set_cell_text(row.cells[1], str(value), bold=False, font_size=8, color="111827")

    document.add_paragraph()


def add_dataframe_table(
    document: Document,
    df: pd.DataFrame,
    currency_columns: list[str] | None = None,
    percent_columns: list[str] | None = None,
    max_rows: int | None = None,
    highlight_first_row: bool = False,
) -> None:
    if df is None or df.empty:
        add_small_note(document, "Não há dados disponíveis para esta seção.")
        return

    currency_columns = currency_columns or []
    percent_columns = percent_columns or []

    working_df = df.copy()

    if max_rows:
        working_df = working_df.head(max_rows)

    for col in currency_columns:
        if col in working_df.columns:
            working_df[col] = working_df[col].apply(format_currency)

    for col in percent_columns:
        if col in working_df.columns:
            working_df[col] = working_df[col].apply(format_percent)

    table = document.add_table(rows=1, cols=len(working_df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    header_cells = table.rows[0].cells

    for idx, column in enumerate(working_df.columns):
        set_cell_background(header_cells[idx], PRIMARY_BLUE)
        set_cell_border(header_cells[idx], PRIMARY_BLUE)
        set_cell_text(
            header_cells[idx],
            str(column),
            bold=True,
            font_size=7,
            color=WHITE,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )

    for row_index, (_, row_data) in enumerate(working_df.iterrows()):
        row_cells = table.add_row().cells
        is_highlight = highlight_first_row and row_index == 0

        for idx, value in enumerate(row_data):
            fill = "DFF3E8" if is_highlight else WHITE
            text_color = SUCCESS_GREEN if is_highlight else "111827"
            set_cell_background(row_cells[idx], fill)
            set_cell_border(row_cells[idx])
            set_cell_text(
                row_cells[idx],
                str(value),
                bold=is_highlight,
                font_size=7,
                color=text_color,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
            )

    document.add_paragraph()


def add_warning_box(document: Document, text: str) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, "FFF4DE")
    set_cell_border(cell, WARNING_ORANGE)
    set_cell_text(cell, text, bold=False, font_size=8, color="4A3B16")
    document.add_paragraph()


# =========================================================
# FUNÇÃO PRINCIPAL
# =========================================================

def generate_word_report(
    client_name: str,
    advisor_name: str,
    simulation_mode: str,
    start_date,
    end_date,
    months: int,
    initial_amount: float,
    annual_cdi_rate: float,
    selic_rate: float,
    tr_rate: float,
    cdb_percentage: float,
    lci_lca_percentage: float,
    treasury_percentage: float,
    treasury_annual_fee: float,
    fund_percentage: float,
    fund_annual_fee: float,
    comparison_df: pd.DataFrame,
    cashflow_df: pd.DataFrame | None,
    monthly_df: pd.DataFrame | None,
    consultive_analysis: str,
) -> BytesIO:
    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)

    best_product = comparison_df.iloc[0]

    if simulation_mode == "Aportes e resgates por calendário":
        period_text = f"{start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}"
    else:
        period_text = f"{months} meses"

    add_cover_header(document, client_name, advisor_name, period_text)

    add_section_heading(document, "1. Visão Geral da Simulação")

    summary_cards = [
        {"label": "Valor inicial", "value": format_currency(initial_amount)},
        {"label": "Melhor alternativa", "value": str(best_product["Produto"])},
        {"label": "Valor líquido", "value": format_currency(best_product["Valor Líquido"])},
        {"label": "Rentabilidade líquida", "value": format_percent(best_product["Rentab. Líq. Período (%)"])},
    ]

    add_summary_cards(document, summary_cards)

    executive_data = {
        "Modo de simulação": simulation_mode,
        "Período/Prazo": period_text,
        "Total aportado": format_currency(best_product.get("Total Aportado", initial_amount)),
        "Total resgatado": format_currency(best_product.get("Total Resgatado", 0)),
        "Rendimento líquido projetado": format_currency(best_product["Rendimento Líquido"]),
        "Gerado em": datetime.now().strftime("%d/%m/%Y às %H:%M"),
    }

    add_key_value_table(document, executive_data)

    add_section_heading(document, "2. Premissas Utilizadas")

    assumptions_data = {
        "CDI anual estimado": format_percent(annual_cdi_rate),
        "Selic meta estimada": format_percent(selic_rate),
        "TR anual estimada": format_percent(tr_rate),
        "CDB / LC": format_percent(cdb_percentage) + " do CDI",
        "LCI / LCA": format_percent(lci_lca_percentage) + " do CDI",
        "Tesouro Selic": format_percent(treasury_percentage) + " do CDI",
        "Custo anual Tesouro Selic": format_percent(treasury_annual_fee),
        "Fundo DI": format_percent(fund_percentage) + " do CDI",
        "Taxa de administração Fundo DI": format_percent(fund_annual_fee),
    }

    add_key_value_table(document, assumptions_data)

    add_section_heading(document, "3. Comparativo dos Produtos")

    comparison_columns = [
        "Produto",
        "% CDI",
        "Taxa Efetiva a.a. (%)",
        "Total Aportado",
        "Total Resgatado",
        "Valor Bruto",
        "IR",
        "Valor Líquido",
        "Rendimento Líquido",
        "Rentab. Líq. Período (%)",
        "Tributável",
    ]

    comparison_to_report = comparison_df[
        [col for col in comparison_columns if col in comparison_df.columns]
    ].copy()

    add_dataframe_table(
        document,
        comparison_to_report,
        currency_columns=[
            "Total Aportado",
            "Total Resgatado",
            "Valor Bruto",
            "IR",
            "Valor Líquido",
            "Rendimento Líquido",
        ],
        percent_columns=[
            "% CDI",
            "Taxa Efetiva a.a. (%)",
            "Rentab. Líq. Período (%)",
        ],
        highlight_first_row=True,
    )

    if cashflow_df is not None and not cashflow_df.empty:
        add_section_heading(document, "4. Calendário de Movimentações")
        add_dataframe_table(
            document,
            cashflow_df,
            currency_columns=["Valor"],
            highlight_first_row=False,
        )

    if monthly_df is not None and not monthly_df.empty:
        add_section_heading(document, "5. Resumo Mensal")

        monthly_columns = [
            "Produto",
            "Mês",
            "Aportes",
            "Resgates",
            "Rendimento Bruto no Mês",
            "Saldo Bruto Final",
        ]

        monthly_to_report = monthly_df[
            [col for col in monthly_columns if col in monthly_df.columns]
        ].copy()

        add_dataframe_table(
            document,
            monthly_to_report,
            currency_columns=[
                "Aportes",
                "Resgates",
                "Rendimento Bruto no Mês",
                "Saldo Bruto Final",
            ],
            max_rows=90,
            highlight_first_row=False,
        )

    add_section_heading(document, "6. Leitura Consultiva")
    clean_analysis = consultive_analysis.replace("**", "")
    add_normal_paragraph(document, clean_analysis)

    add_section_heading(document, "7. Aviso Técnico")
    add_warning_box(
        document,
        "Os valores apresentados são estimativas para fins de simulação. "
        "Este relatório não constitui promessa de rentabilidade, oferta, recomendação individualizada "
        "ou garantia de resultado. A análise final deve considerar perfil do cliente, adequação, liquidez, "
        "risco, tributação, condições de mercado e normas aplicáveis."
    )

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Relatório gerado automaticamente pelo Simulador Estratégico de Investimentos | Uso interno e consultivo")
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor.from_string(TEXT_GRAY)

    file_stream = BytesIO()
    document.save(file_stream)
    file_stream.seek(0)

    return file_stream
