"""
Diagramador de simulados (Word -> layout final)
================================================
Le um .docx "cru" (conteudo certo, sem diagramacao) e gera um .docx
formatado no padrao usado nos simulados (cabecalho/rodape fixos,
"Questao N" em destaque, alternativas com letra em negrito,
tabela de Gabarito Simplificado gerada automaticamente a partir do
gabarito comentado).
"""

import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

AZUL = RGBColor(0x1F, 0x3B, 0x57)
CINZA = RGBColor(0x44, 0x44, 0x44)

QUESTAO_RE = re.compile(r"^Quest(?:ã|a)o\s+(\d+)\b[\s\S]*", re.IGNORECASE)
ALT_RE = re.compile(r"^([A-E])\)\s*(.*)$")
GAB_COMENTADO_Q_RE = re.compile(
    r"Quest(?:ã|a)o\s+(\d+)\s*[—–\-]\s*Alternativa correta:\s*([A-E])", re.IGNORECASE
)
BLOCO_RE = re.compile(r"^BLOCO\s+(\d+)", re.IGNORECASE)
GABARITO_HEAD_RE = re.compile(r"^GABARITO COMENTADO", re.IGNORECASE)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(kwargs[edge].get("sz", 8)))
            el.set(qn("w:color"), kwargs[edge].get("color", "000000"))
            tcBorders.append(el)
    tcPr.append(tcBorders)


def add_box_paragraph(container, title, subtitle):
    table = container.add_table(rows=1, cols=1, width=Cm(17))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    set_cell_border(
        cell,
        top={"sz": 12, "color": "1F3B57"},
        bottom={"sz": 12, "color": "1F3B57"},
        left={"sz": 12, "color": "1F3B57"},
        right={"sz": 12, "color": "1F3B57"},
    )
    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(title)
    r1.bold = True
    r1.font.size = Pt(12)
    r1.font.color.rgb = AZUL

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(subtitle)
    r2.font.size = Pt(9)
    r2.font.color.rgb = CINZA
    return table


def style_runs(paragraph, bold=False, size=11, color=None, italic=False):
    for run in paragraph.runs:
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color


def extract_gabarito(doc):
    gabarito = {}
    for p in doc.paragraphs:
        for m in GAB_COMENTADO_Q_RE.finditer(p.text):
            gabarito[int(m.group(1))] = m.group(2).upper()
    return gabarito


def add_gabarito_table(doc, gabarito, n_blocos=5):
    if not gabarito:
        return
    doc.add_page_break()
    h = doc.add_heading("GABARITO SIMPLIFICADO", level=2)
    style_runs(h, bold=True, size=14, color=AZUL)

    numeros = sorted(gabarito.keys())
    total = len(numeros)
    por_bloco = -(-total // n_blocos)
    table = doc.add_table(rows=por_bloco + 1, cols=n_blocos * 2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for b in range(n_blocos):
        cell = table.cell(0, b * 2)
        cell.merge(table.cell(0, b * 2 + 1))
        cell.text = f"Bloco {b + 1}"
        style_runs(cell.paragraphs[0], bold=True, size=10, color=AZUL)

    for idx, n in enumerate(numeros):
        bloco = idx // por_bloco
        row = idx % por_bloco + 1
        table.cell(row, bloco * 2).text = str(n)
        table.cell(row, bloco * 2 + 1).text = gabarito[n]
        for c in (table.cell(row, bloco * 2), table.cell(row, bloco * 2 + 1)):
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            style_runs(c.paragraphs[0], size=10)


def diagramar(input_stream, titulo=None, subtitulo=None, rodape=None):
    """Recebe um stream (BytesIO) do docx cru e devolve um Document() formatado."""
    src = Document(input_stream)

    if titulo is None or subtitulo is None:
        non_empty = [p.text.strip() for p in src.paragraphs if p.text.strip()]
        titulo = titulo or (non_empty[0] if len(non_empty) > 0 else "")
        subtitulo = subtitulo or (non_empty[1] if len(non_empty) > 1 else "")

    rodape = rodape or (
        "O Concursado de hoje é o Concurseiro que nunca desistiu! "
        "Lista de transmissão do whatsapp @professorfrancelino"
    )

    out = Document()
    sec = out.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    header = sec.header
    header.is_linked_to_previous = False
    for p in list(header.paragraphs):
        p.text = ""
    add_box_paragraph(header, titulo, subtitulo)

    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run(rodape)
    fr.font.size = Pt(8)
    fr.font.color.rgb = CINZA

    skip_first_lines = 2
    count_skipped = 0
    for p in src.paragraphs:
        text = p.text.strip()

        if count_skipped < skip_first_lines and text in (titulo, subtitulo):
            count_skipped += 1
            continue

        if not text:
            continue

        bloco_m = BLOCO_RE.match(text)
        gab_head_m = GABARITO_HEAD_RE.match(text)
        questao_m = QUESTAO_RE.match(text)
        alt_m = ALT_RE.match(text)

        if bloco_m:
            out.add_page_break()
            h = out.add_heading(text, level=1)
            style_runs(h, bold=True, size=14, color=AZUL)
            continue

        if gab_head_m:
            out.add_page_break()
            h = out.add_heading(text, level=2)
            style_runs(h, bold=True, size=13, color=AZUL)
            continue

        if questao_m:
            np = out.add_paragraph()
            np.paragraph_format.space_before = Pt(10)
            np.paragraph_format.space_after = Pt(4)
            lead = f"Questão {questao_m.group(1)}"
            rest = text[len(lead):]
            r1 = np.add_run(lead)
            r1.bold = True
            r1.font.size = Pt(11)
            r2 = np.add_run(rest)
            r2.font.size = Pt(11)
            continue

        if alt_m:
            np = out.add_paragraph()
            np.paragraph_format.left_indent = Cm(0.6)
            np.paragraph_format.space_after = Pt(2)
            letra, resto = alt_m.groups()
            r1 = np.add_run(f"{letra}) ")
            r1.bold = True
            r1.font.size = Pt(10.5)
            r2 = np.add_run(resto)
            r2.font.size = Pt(10.5)
            continue

        np = out.add_paragraph(text)
        np.paragraph_format.space_after = Pt(4)
        style_runs(np, size=10.5)

    gabarito = extract_gabarito(src)
    add_gabarito_table(out, gabarito)

    return out, len(gabarito)
