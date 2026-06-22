"""
Diagramador de simulados (Word -> layout final)
================================================
Le um .docx "cru" (conteudo certo, sem diagramacao) e devolve um novo
documento construido a partir do TEMPLATE_BASE (assets/template_base.docx)
-- um arquivo que ja tem o cabecalho/rodape com a arte definitiva (caixa
flutuante com titulo/ementa + foto + icones), feito direto no Word.
Nunca recriamos cabecalho/rodape via codigo; so trocamos o texto do
titulo/ementa dentro da caixa flutuante do cabecalho e escrevemos o
corpo (blocos, questoes, alternativas, gabarito) do zero.
"""

import os
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

AZUL = RGBColor(0x1F, 0x3B, 0x57)
AZUL_TITULO = RGBColor(0x12, 0x3D, 0x6B)
CINZA = RGBColor(0x44, 0x44, 0x44)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
TEMPLATE_BASE = os.path.join(ASSETS_DIR, "template_base.docx")

QUESTAO_RE = re.compile(r"^Quest(?:ã|a)o\s+(\d+)\b\s*(.*)$", re.IGNORECASE)
ALT_RE = re.compile(r"^([A-E])\)\s*(.*)$")
GAB_COMENTADO_Q_RE = re.compile(
    r"Quest(?:ã|a)o\s+(\d+)\s*[—–\-]\s*Alternativa correta:\s*([A-E])", re.IGNORECASE
)
GAB_TITLE_RE = re.compile(
    r"^(Quest(?:ã|a)o\s+\d+\s*[—–\-]\s*Alternativa correta:\s*[A-E])\s*\*?\s*(.*)$",
    re.IGNORECASE,
)
BLOCO_RE = re.compile(r"^BLOCO\s+(\d+)", re.IGNORECASE)
GABARITO_HEAD_RE = re.compile(r"^GABARITO COMENTADO", re.IGNORECASE)
REVISAO_START_RE = re.compile(r"^Revis(?:ã|a)o de 24 horas$", re.IGNORECASE)
META_GENERICA_RE = re.compile(r"^Meta\s+\d+\s*[—–\-]", re.IGNORECASE)
META_REVISAO_RE = re.compile(r"^Meta de Revis(?:ã|a)o", re.IGNORECASE)
GABARITO_DIVIDER_RE = re.compile(r"^[━_\-\s]*GABARITO[━_\-\s]*$", re.IGNORECASE)
ALT_SPLIT_RE = re.compile(r"\s(?=[A-E]\)\s)")
AZUL_CAIXA = "D9EEFB"  # fundo azul claro da caixa de revisao (ajustavel)


def shade_paragraph(paragraph, fill_hex):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    pPr.append(shd)


def style_runs(paragraph, bold=False, size=11, color=None, italic=False):
    for run in paragraph.runs:
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color


def set_textbox_paragraph_text(paragraph_el, new_text):
    """Troca o texto de um <w:p> mantendo a formatacao do 1o run
    (fonte, tamanho, cor, negrito) -- usado dentro das caixas flutuantes
    do cabecalho, que python-docx nao expoe via paragraphs normal."""
    runs = paragraph_el.findall(qn("w:r"))
    if not runs:
        return
    first_run = runs[0]
    t_elements = first_run.findall(qn("w:t"))
    if t_elements:
        t_elements[0].text = new_text
        for extra_t in t_elements[1:]:
            first_run.remove(extra_t)
    else:
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = new_text
        first_run.append(t)
    for r in runs[1:]:
        r.getparent().remove(r)


def update_header_title(doc, titulo, ementa):
    """Atualiza o titulo/ementa dentro da caixa flutuante azul do
    cabecalho do TEMPLATE_BASE. Existem 2 copias do mesmo texto no XML
    (uma para o Word moderno, outra de fallback para versoes antigas) --
    as duas precisam ser atualizadas para ficarem consistentes."""
    header_el = doc.sections[0].header._element
    boxes = header_el.findall(".//" + qn("w:txbxContent"))
    for box in boxes:
        paragraphs = box.findall(qn("w:p"))
        if len(paragraphs) >= 1:
            set_textbox_paragraph_text(paragraphs[0], titulo)
        if len(paragraphs) >= 2:
            set_textbox_paragraph_text(paragraphs[1], ementa)


def clear_body(doc):
    """Remove todos os paragrafos/tabelas do corpo, preservando section/header/footer."""
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sect_pr:
            body.remove(child)


def extract_gabarito(paragraphs):
    gabarito = {}
    for text in paragraphs:
        for m in GAB_COMENTADO_Q_RE.finditer(text):
            gabarito[int(m.group(1))] = m.group(2).upper()
    return gabarito


def set_table_no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def shade_cell(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, size=10, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color


def split_question(text):
    """Quando a questao e as alternativas A-E vem todas coladas num so
    paragrafo (sem quebra de linha), separa em (numero, enunciado, [(letra, texto), ...])."""
    m = QUESTAO_RE.match(text)
    if not m:
        return None
    numero, resto = m.groups()
    partes = ALT_SPLIT_RE.split(resto)
    enunciado = partes[0].strip()
    alternativas = []
    for chunk in partes[1:]:
        am = ALT_RE.match(chunk.strip())
        if am:
            alternativas.append(am.groups())
    return numero, enunciado, alternativas


def start_blue_box(doc, titulo_box, titulo_size=12):
    """Cria uma caixa de fundo azul claro (usada na 'Revisao de 24 horas'
    e em todos os Gabaritos Comentados) e devolve a celula onde o
    conteudo deve ser escrito."""
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell.width = Cm(17)
    shade_cell(cell, AZUL_CAIXA)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(titulo_box)
    r.bold = True
    r.font.size = Pt(titulo_size)
    r.font.color.rgb = AZUL
    return cell


def add_gabarito_comentado_line(container, text):
    """Trata uma linha dentro do Gabarito Comentado. O padrao no arquivo
    cru e: 'Questao N — Alternativa correta: X * A) <texto da alt A>'
    (titulo e a 1a alternativa colados no mesmo paragrafo); as demais
    alternativas (B-E) e a Fundamentacao Legal vem em paragrafos
    separados. Aqui sempre separamos titulo (negrito) de alternativa
    (linha propria)."""
    gt_m = GAB_TITLE_RE.match(text)
    if gt_m:
        titulo, resto = gt_m.groups()
        tp = container.add_paragraph()
        tp.paragraph_format.space_before = Pt(10)
        tp.paragraph_format.space_after = Pt(3)
        tr = tp.add_run(titulo)
        tr.bold = True
        tr.font.size = Pt(11)

        resto = resto.strip()
        if resto:
            am = ALT_RE.match(resto)
            if am:
                letra, alt_texto = am.groups()
                ap = container.add_paragraph()
                ap.paragraph_format.left_indent = Cm(0.6)
                ap.paragraph_format.space_after = Pt(2)
                ar1 = ap.add_run(f"{letra}) ")
                ar1.bold = True
                ar1.font.size = Pt(10.5)
                ar2 = ap.add_run(alt_texto)
                ar2.font.size = Pt(10.5)
            else:
                container.add_paragraph(resto)
        return

    alt_m = ALT_RE.match(text)
    if alt_m:
        letra, alt_texto = alt_m.groups()
        ap = container.add_paragraph()
        ap.paragraph_format.left_indent = Cm(0.6)
        ap.paragraph_format.space_after = Pt(2)
        ar1 = ap.add_run(f"{letra}) ")
        ar1.bold = True
        ar1.font.size = Pt(10.5)
        ar2 = ap.add_run(alt_texto)
        ar2.font.size = Pt(10.5)
        return

    np = container.add_paragraph(text)
    np.paragraph_format.space_after = Pt(4)
    style_runs(np, size=10)


def add_centered_bold_line(container, text, size=11):
    p = container.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = AZUL


def add_question_block(container, numero, enunciado, alternativas):
    np = container.add_paragraph()
    np.paragraph_format.space_before = Pt(12)
    np.paragraph_format.space_after = Pt(4)
    r1 = np.add_run(f"Questão {numero}")
    r1.bold = True
    r1.font.size = Pt(11)
    r1.add_break(WD_BREAK.LINE)
    r2 = np.add_run(enunciado)
    r2.font.size = Pt(11)

    for letra, texto in alternativas:
        ap = container.add_paragraph()
        ap.paragraph_format.left_indent = Cm(0.6)
        ap.paragraph_format.space_after = Pt(2)
        ar1 = ap.add_run(f"{letra}) ")
        ar1.bold = True
        ar1.font.size = Pt(10.5)
        ar2 = ap.add_run(texto)
        ar2.font.size = Pt(10.5)


def add_gabarito_table(doc, gabarito, n_blocos=5):
    if not gabarito:
        return
    doc.add_page_break()
    h = doc.add_heading("GABARITO SIMPLIFICADO", level=2)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_runs(h, bold=True, size=14, color=AZUL)
    doc.add_paragraph()

    numeros = sorted(gabarito.keys())
    total = len(numeros)
    por_bloco = -(-total // n_blocos)
    chunks = [numeros[i * por_bloco:(i + 1) * por_bloco] for i in range(n_blocos)]

    outer = doc.add_table(rows=1, cols=n_blocos)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    outer.autofit = False
    set_table_no_borders(outer)

    for b in range(n_blocos):
        cell = outer.cell(0, b)
        cell.width = Cm(3.2)

        hp = cell.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = hp.add_run(f"Bloco {b + 1}")
        hr.bold = True
        hr.font.size = Pt(11)
        hr.font.color.rgb = AZUL

        chunk = chunks[b]
        if not chunk:
            continue
        nested = cell.add_table(rows=len(chunk), cols=2)
        nested.autofit = False
        set_table_no_borders(nested)
        for col in nested.columns:
            col.width = Cm(1.5)

        for i, numero in enumerate(chunk):
            c1, c2 = nested.cell(i, 0), nested.cell(i, 1)
            set_cell_text(c1, str(numero), bold=True, size=10)
            set_cell_text(c2, gabarito[numero], bold=True, size=10)
            if i % 2 == 1:
                shade_cell(c1, "EFEFEF")
                shade_cell(c2, "EFEFEF")

        # paragrafo apos a tabela aninhada (exigencia do formato de tabela)
        cell.add_paragraph()


def diagramar(input_stream, titulo=None, subtitulo=None):
    """Recebe um stream (BytesIO) do docx cru e devolve (Document, n_gabarito).

    O documento devolvido E O TEMPLATE_BASE (cabecalho/rodape com a arte
    definitiva, feita no Word) com o corpo preenchido a partir do
    conteudo do arquivo cru."""
    raw = Document(input_stream)

    all_texts = [p.text.strip() for p in raw.paragraphs]
    non_empty = [t for t in all_texts if t]
    titulo = titulo or (non_empty[0] if len(non_empty) > 0 else "")
    subtitulo = subtitulo or (non_empty[1] if len(non_empty) > 1 else "")

    original_paragraph_texts = all_texts

    doc = Document(TEMPLATE_BASE)
    update_header_title(doc, titulo, subtitulo)

    gabarito = extract_gabarito(original_paragraph_texts)

    clear_body(doc)

    skip_first_lines = 2
    count_skipped = 0
    target = doc
    box_mode = None  # None | "revisao" | "gabarito"

    for text in original_paragraph_texts:
        text = text.strip()

        if count_skipped < skip_first_lines and text in (titulo, subtitulo):
            count_skipped += 1
            continue

        if not text:
            continue

        if REVISAO_START_RE.match(text):
            target = start_blue_box(doc, text)
            box_mode = "revisao"
            continue

        if box_mode == "revisao" and GABARITO_DIVIDER_RE.match(text):
            target = start_blue_box(doc, "GABARITO COMENTADO", titulo_size=13)
            box_mode = "gabarito"
            continue

        if box_mode == "revisao":
            if META_REVISAO_RE.match(text) or META_GENERICA_RE.match(text):
                add_centered_bold_line(target, text)
                continue
            q = split_question(text)
            if q:
                add_question_block(target, *q)
                continue
            target.add_paragraph(text)
            continue

        bloco_m = BLOCO_RE.match(text)
        gab_head_m = GABARITO_HEAD_RE.match(text)

        if bloco_m:
            box_mode = None
            target = doc
            doc.add_page_break()
            h = doc.add_heading(text, level=1)
            style_runs(h, bold=True, size=14, color=AZUL)
            continue

        if gab_head_m:
            target = start_blue_box(doc, text, titulo_size=13)
            box_mode = "gabarito"
            continue

        if box_mode == "gabarito":
            add_gabarito_comentado_line(target, text)
            continue

        questao_m = QUESTAO_RE.match(text)
        alt_m = ALT_RE.match(text)

        if questao_m:
            numero, resto = questao_m.groups()
            np = doc.add_paragraph()
            np.paragraph_format.space_before = Pt(12)
            np.paragraph_format.space_after = Pt(4)
            r1 = np.add_run(f"Questão {numero}")
            r1.bold = True
            r1.font.size = Pt(11)
            r1.add_break(WD_BREAK.LINE)
            r2 = np.add_run(resto)
            r2.font.size = Pt(11)
            continue

        if alt_m:
            np = doc.add_paragraph()
            np.paragraph_format.left_indent = Cm(0.6)
            np.paragraph_format.space_after = Pt(2)
            letra, resto = alt_m.groups()
            r1 = np.add_run(f"{letra}) ")
            r1.bold = True
            r1.font.size = Pt(10.5)
            r2 = np.add_run(resto)
            r2.font.size = Pt(10.5)
            continue

        np = doc.add_paragraph(text)
        np.paragraph_format.space_after = Pt(4)
        style_runs(np, size=10.5)

    add_gabarito_table(doc, gabarito)

    return doc, len(gabarito)
