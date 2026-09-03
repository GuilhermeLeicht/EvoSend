"""
Geração do PDF final de "Quality Performance".

Melhorias v2:
- Visual mais limpo: cabeçalho com faixa colorida, tabela com zebra striping,
  rodapé redesenhado, tipografia consistente.
- Emojis de legenda lidos de PNGs fixos em <base_dir>/Restricted/emojis/.
  Se os arquivos não existirem, a legenda é exibida apenas com texto.
- Corrigido: ano não é mais duplicado no subtítulo.
"""
import os
import traceback

_MESES_PT_TO_EN = {
    "Janeiro":"January","Fevereiro":"February","Março":"March",
    "Abril":"April","Maio":"May","Junho":"June","Julho":"July",
    "Agosto":"August","Setembro":"September","Outubro":"October",
    "Novembro":"November","Dezembro":"December"
}

def _mes_to_en(mes_str: str) -> str:
    """Converte nome do mês PT→EN se necessário."""
    return _MESES_PT_TO_EN.get(mes_str.strip(), mes_str.strip())

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image,
    Table, TableStyle, HRFlowable
)

from utils import criar_pasta_se_nao_existe

# ── Paleta do PDF (espelha ui/theme.py) ──────────────────────────────────────
C_ACCENT      = colors.HexColor("#2D5F4C")   # verde escuro
C_ACCENT_LIGHT= colors.HexColor("#E8F2EE")   # verde muito claro
C_HEADER_TBL  = colors.HexColor("#215D93")   # azul escuro para cabeçalho da tabela
C_ROW_ALT     = colors.HexColor("#F5F7FA")   # linha alternada
C_BORDER      = colors.HexColor("#E5E3DD")
C_TEXT        = colors.HexColor("#1C1C1A")
C_MUTED       = colors.HexColor("#6B6B66")
C_DANGER      = colors.HexColor("#B3261E")
C_GOAL_ROW    = colors.HexColor("#FFF8E1")   # fundo das linhas de meta


def _build_styles():
    styles = getSampleStyleSheet()

    def add(name, **kw):
        styles.add(ParagraphStyle(name=name, **kw))

    add("ESTitle",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=22, fontName="Helvetica-Bold",
        textColor=C_TEXT, spaceAfter=18, spaceBefore=22)

    add("ESSubtitle",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=11, fontName="Helvetica",
        textColor=C_MUTED, spaceAfter=8)

    add("ESCompany",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=17, fontName="Helvetica-Bold",
        textColor=C_ACCENT, wordWrap="CJK", leading=22, spaceAfter=8)

    add("ESSectionHead",
        parent=styles["Normal"], alignment=TA_LEFT,
        fontSize=10, fontName="Helvetica-Bold",
        textColor=C_ACCENT, spaceBefore=6, spaceAfter=4)

    add("ESLegendText",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=9, textColor=C_TEXT, leading=11)

    add("ESTblHeader",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=7, fontName="Helvetica-Bold",
        textColor=colors.white, leading=8)

    add("ESTblMetric",
        parent=styles["Normal"], alignment=TA_LEFT,
        fontSize=5.5, fontName="Helvetica",
        textColor=C_TEXT, leading=7, wordWrap="CJK")

    add("ESTblValue",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=5.5, fontName="Helvetica",
        textColor=C_TEXT, leading=7, wordWrap="CJK")

    add("ESTblYearHeader",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=7, fontName="Helvetica-Bold",
        textColor=C_DANGER, leading=8)

    add("ESFooterTitle",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=9, fontName="Helvetica-Bold",
        textColor=C_ACCENT, spaceBefore=0, spaceAfter=0)

    add("ESFooterText",
        parent=styles["Normal"], alignment=TA_CENTER,
        fontSize=9, fontName="Helvetica",
        textColor=C_TEXT, spaceBefore=0, spaceAfter=0)

    return styles


# ── Logo ──────────────────────────────────────────────────────────────────────

def _add_logo(canvas, doc, logo_path, log_func):
    if not logo_path or not os.path.exists(logo_path):
        return
    try:
        pw, ph = doc.pagesize
        lw, lh = 1.3 * inch, 0.85 * inch
        canvas.drawImage(logo_path,
                         pw - doc.rightMargin - lw,
                         ph - lh - 0.55 * inch,
                         width=lw, height=lh,
                         preserveAspectRatio=True, mask="auto")
    except Exception as e:
        log_func(f"Aviso: não foi possível adicionar logo: {e}")


# ── Cabeçalho do documento ────────────────────────────────────────────────────

def _build_header(story, styles, nome_cliente, mes_str, ano_str):
    """
    Cabeçalho: título, subtítulo (Pomerode — Mês Ano) e nome do fornecedor.
    Corrigido: ano_str é passado separadamente para não duplicar.
    """
    story.append(Paragraph("Quality Performance", styles["ESTitle"]))
    story.append(Paragraph(f"Pomerode — {mes_str} {ano_str}", styles["ESSubtitle"]))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=C_ACCENT, spaceAfter=6))
    story.append(Paragraph(nome_cliente, styles["ESCompany"]))
    story.append(Spacer(1, 0.1 * inch))


# ── Gráficos mensais ──────────────────────────────────────────────────────────

def _build_graphs(story, styles, doc, graph_paths, log_func):
    story.append(Paragraph("Monthly Trend View", styles["ESSectionHead"]))

    if not graph_paths:
        story.append(Paragraph(
            "<i>Nenhum dado de tendência mensal disponível.</i>",
            styles["Normal"]))
        story.append(Spacer(1, 0.05 * inch))
        return

    aw = letter[0] - doc.leftMargin - doc.rightMargin
    iw = (aw - 0.15 * inch) / 2
    ih = iw * (4 / 7)

    rows, cur = [], []
    for i, p in enumerate(graph_paths):
        if not os.path.exists(p):
            log_func(f"Aviso: gráfico não encontrado: {p}")
            continue
        img = Image(p)
        img.drawWidth  = iw
        img.drawHeight = ih
        cur.append(img)
        if len(cur) == 2 or i == len(graph_paths) - 1:
            while len(cur) < 2:
                cur.append(Paragraph("", styles["Normal"]))
            rows.append(cur)
            cur = []

    if rows:
        t = Table(rows, colWidths=[iw, iw])
        t.setStyle(TableStyle([
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",   (0, 0), (0, -1),  0),
            ("RIGHTPADDING",  (1, 0), (1, -1),  0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t)


# ── Legenda de emojis ─────────────────────────────────────────────────────────

def _build_legend(story, styles, doc, emoji_dir, log_func):
    """
    Legenda compacta em linha única: emoji inline + texto, separados por espaço.
    Ocupa apenas uma linha fina no PDF.
    """
    SIZE = 14  # pontos — pequeno, discreto
    
    def _img_or_bullet(filename):
        path = os.path.join(emoji_dir, filename) if emoji_dir else ""
        if path and os.path.exists(path):
            return f'<img src="{path}" width="{SIZE}" height="{SIZE}" valign="middle"/>'
        return "●"

    g = _img_or_bullet("emoji_green.png")
    y = _img_or_bullet("emoji_yellow.png")
    r = _img_or_bullet("emoji_red.png")

    legend_style = ParagraphStyle(
        "LegendLine", parent=styles["Normal"],
        alignment=TA_CENTER, fontSize=8, leading=10,
        textColor=C_MUTED, spaceAfter=4, spaceBefore=2)

    text = (f"{g} Met goal &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"{y} Up to 5% above &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"{r} Exceeded goal")

    aw = letter[0] - doc.leftMargin - doc.rightMargin
    t = Table([[Paragraph(text, legend_style)]], colWidths=[aw])
    t.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND",    (0, 0), (-1, -1), C_ACCENT_LIGHT),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.05 * inch))


# ── Tabela de sumário ─────────────────────────────────────────────────────────

def _build_summary_table(story, styles, doc, data_table_info):
    if not data_table_info:
        return

    header_years  = data_table_info["header_years"]
    header_months = data_table_info["header_months_str"]
    rows_data     = data_table_info["rows"]


    # Linha de cabeçalho
    hdr_cells = [Paragraph("Metric", styles["ESTblHeader"])]
    for yr in header_years:
        hdr_cells.append(Paragraph(str(yr), styles["ESTblYearHeader"]))
    for mo in header_months:
        hdr_cells.append(Paragraph(mo, styles["ESTblHeader"]))
    table_data = [hdr_cells]

    # Linhas de dados
    is_goal_row = []
    for label, values in rows_data.items():
        row = [Paragraph(label, styles["ESTblMetric"])]
        for v in values:
            s = str(v)
            # Reduz fonte progressivamente para evitar quebra de linha
            if len(s) >= 8:
                fs = 4.0
            elif len(s) >= 6:
                fs = 4.5
            elif len(s) >= 5:
                fs = 5.0
            else:
                fs = 5.5
            from reportlab.lib.styles import ParagraphStyle
            cell_style = ParagraphStyle(
                'auto', parent=styles["ESTblValue"],
                fontSize=fs, leading=fs + 1.5, wordWrap=None)
            row.append(Paragraph(s, cell_style))
        table_data.append(row)
        is_goal_row.append("Goal" in label)

    # Larguras de colunas
    aw        = letter[0] - doc.leftMargin - doc.rightMargin
    ncols     = len(hdr_cells)
    w_metric  = aw * 0.12
    w_rest    = (aw - w_metric) / (ncols - 1)
    col_widths = [w_metric] + [w_rest] * (ncols - 1)

    t = Table(table_data, colWidths=col_widths, repeatRows=1,
              rowHeights=None)

    style_cmds = [
        # Cabeçalho
        ("BACKGROUND",    (0, 0), (-1, 0),  C_HEADER_TBL),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  7),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  4),
        ("TOPPADDING",    (0, 0), (-1, 0),  4),
        # Dados
        ("FONTSIZE",      (0, 1), (-1, -1), 6),
        ("TOPPADDING",    (0, 1), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 2),
        ("ALIGN",         (1, 1), (-1, -1), "CENTER"),
        ("ALIGN",         (0, 1), (0, -1),  "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.4, C_BORDER),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]

    # Zebra + linhas de meta
    for i, is_goal in enumerate(is_goal_row):
        data_row = i + 1
        if is_goal:
            style_cmds.append(("BACKGROUND", (0, data_row), (-1, data_row), C_GOAL_ROW))
            style_cmds.append(("TEXTCOLOR",  (0, data_row), (0, data_row),  C_MUTED))
            style_cmds.append(("FONTNAME",   (0, data_row), (-1, data_row), "Helvetica-Oblique"))
        elif data_row % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, data_row), (-1, data_row), C_ROW_ALT))

    t.setStyle(TableStyle(style_cmds))
    story.append(t)


# ── Rodapé ────────────────────────────────────────────────────────────────────

def _build_footer(story, styles, doc, engineer_name, engineer_phone, engineer_email):
    story.append(Spacer(1, 0.15 * inch))
    story.append(HRFlowable(width="40%", thickness=1, color=C_BORDER,
                             hAlign="CENTER", spaceAfter=4))

    content = [Paragraph("QUALITY DEPARTMENT", styles["ESFooterTitle"])]
    if engineer_name:
        content.append(Paragraph(engineer_name,  styles["ESFooterText"]))
    if engineer_phone:
        content.append(Paragraph(engineer_phone, styles["ESFooterText"]))
    if engineer_email:
        content.append(Paragraph(engineer_email, styles["ESFooterText"]))

    aw = letter[0] - doc.leftMargin - doc.rightMargin
    fw = aw * 0.42

    footer_rows = [[item] for item in content]
    t = Table(footer_rows, colWidths=[fw], hAlign="CENTER")
    t.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_BORDER),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, C_BORDER),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND",    (0, 0), (-1, 0),  C_ACCENT_LIGHT),
    ]))
    story.append(t)


# ── Função pública ────────────────────────────────────────────────────────────

def gerar_pdf_com_multiplos_sumarios(nome_cliente, conteudo_texto,
                                      caminhos_imagens_graficos_sumario_individuais,
                                      caminhos_imagens_graficos_mensais,
                                      caminho_saida_pdf, log_func,
                                      mes_referencia_str, ano_referencia_str,
                                      pasta_graficos_temp_para_emojis,
                                      data_table_info, logo_path,
                                      engineer_name, engineer_phone, engineer_email):
    """
    Gera o PDF de desempenho para um fornecedor.
    `mes_referencia_str` deve conter APENAS o nome do mês (ex: "Abril").
    `ano_referencia_str` deve conter APENAS o ano (ex: "2026").
    O subtítulo é montado aqui como "Pomerode — Abril 2026".
    """
    try:
        doc = SimpleDocTemplate(
            caminho_saida_pdf, pagesize=letter,
            leftMargin=0.75 * inch, rightMargin=0.75 * inch,
            topMargin=1.1 * inch, bottomMargin=0.75 * inch)

        styles = _build_styles()
        story  = []

        # Cabeçalho — sem duplicar o ano
        _build_header(story, styles, nome_cliente,
                      mes_referencia_str, ano_referencia_str)

        # Gráficos
        _build_graphs(story, styles, doc,
                      caminhos_imagens_graficos_mensais, log_func)

        # Legenda de emojis (PNGs fixos)
        _build_legend(story, styles, doc,
                      pasta_graficos_temp_para_emojis, log_func)

        # Tabela de sumário
        _build_summary_table(story, styles, doc, data_table_info)

        # Rodapé
        _build_footer(story, styles, doc,
                      engineer_name, engineer_phone, engineer_email)

        def _page_cb(canvas, doc_):
            _add_logo(canvas, doc_, logo_path, log_func)

        doc.build(story, onFirstPage=_page_cb, onLaterPages=_page_cb)
        return True

    except Exception as e:
        log_func(f"Erro ao gerar PDF para '{nome_cliente}': {e}")
        traceback.print_exc()
        return False
