"""
Diagramador de simulados (Word -> layout final)
================================================
Le um .docx "cru" (conteudo certo, sem diagramacao) e devolve o MESMO
documento com:
  - cabecalho/rodape ORIGINAIS preservados (banner, logo e credito que
    ja vem prontos no arquivo cru) -- nada e recriado do zero;
  - uma faixa azul extra no cabecalho com o titulo da materia e a ementa,
    repetindo em todas as paginas (isso elimina a colagem manual);
  - "Questao N" em negrito numa linha e o enunciado na linha de baixo;
  - alternativas A-E com a letra em negrito;
  - tabela de Gabarito Simplificado gerada automaticamente a partir do
    gabarito comentado.
"""

import io
import os
import re
from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

AZUL = RGBColor(0x1F, 0x3B, 0x57)
AZUL_TITULO = RGBColor(0x12, 0x3D, 0x6B)
AZUL_FAIXA = "1C5FA8"   # cor da barra de credito (ajustavel)
CINZA = RGBColor(0x44, 0x44, 0x44)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
ICON_WHATSAPP = os.path.join(ASSETS_DIR, "whatsapp.png")
ICON_INSTAGRAM = os.path.join(ASSETS_DIR, "instagram.png")
CREDITO_ESQ = "O Concursado de hoje é o Concurseiro que nunca desistiu!"
CREDITO_DIR_PRE = "Lista de transmissão do whatsapp "
CREDITO_DIR_POS = " @professorfrancelino"

QUESTAO_RE = re.compile(r"^Quest(?:ã|a)o\s+(\d+)\b\s*(.*)$", re.IGNORECASE)
ALT_RE = re.compile(r"^([A-E])\)\s*(.*)$")
GAB_COMENTADO_Q_RE = re.compile(
    r"Quest(?:ã|a)o\s+(\d+)\s*[—–\-]\s*Alternativa correta:\s*([A-E])", re.IGNORECASE
)
BLOCO_RE = re.compile(r"^BLOCO\s+(\d+)", re.IGNORECASE)
GABARITO_HEAD_RE = re.compile(r"^GABARITO COMENTADO", re.IGNORECASE)


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


def extract_header_image(header):
    """Pega os bytes da imagem do banner ja embutida no cabecalho original."""
    for rel in header.part.rels.values():
        if "image" in rel.reltype:
            return rel.target_part.blob
    return None


def remove_image_paragraphs(header):
    """Remove do cabecalho os paragrafos que contem a imagem/divisoria,
    mantendo a linha de credito do topo intacta."""
    for p in list(header.paragraphs):
        if p._p.findall(qn("w:r") + "/" + qn("w:drawing")) or p.text.strip().startswith("___"):
            p._p.getparent().remove(p._p)
        elif any(run._r.find(qn("w:drawing")) is not None for run in p.runs):
            p._p.getparent().remove(p._p)


def clear_container_paragraphs(container):
    """Remove TODOS os paragrafos/tabelas de um header ou footer."""
    body = container._element
    for child in list(body):
        if child.tag in (qn("w:p"), qn("w:tbl")):
            body.remove(child)
    # garante que sempre exista pelo menos 1 paragrafo (exigencia do formato)
    if not container.paragraphs:
        container.add_paragraph()


def add_credit_band(container):
    """Cria a barra azul de credito (texto + icones) usada no topo do
    cabecalho e no rodape -- mesma arte nos dois lugares."""
    clear_container_paragraphs(container)

    table = container.add_table(rows=1, cols=2, width=Cm(17))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    left_cell, right_cell = table.rows[0].cells
    left_cell.width = Cm(9)
    right_cell.width = Cm(8)

    for cell in (left_cell, right_cell):
        shade_paragraph(cell.paragraphs[0], AZUL_FAIXA)
        cell.paragraphs[0].paragraph_format.space_before = Pt(3)
        cell.paragraphs[0].paragraph_format.space_after = Pt(3)

    lp = left_cell.paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    lr = lp.add_run(CREDITO_ESQ)
    lr.bold = True
    lr.font.size = Pt(9)
    lr.font.color.rgb = BRANCO

    rp = right_cell.paragraphs[0]
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = rp.add_run(CREDITO_DIR_PRE)
    r1.bold = True
    r1.font.size = Pt(9)
    r1.font.color.rgb = BRANCO

    if os.path.exists(ICON_WHATSAPP):
        r1.add_picture(ICON_WHATSAPP, height=Cm(0.35))
    r2 = rp.add_run(" | ")
    r2.bold = True
    r2.font.size = Pt(9)
    r2.font.color.rgb = BRANCO
    if os.path.exists(ICON_INSTAGRAM):
        r2.add_picture(ICON_INSTAGRAM, height=Cm(0.35))
    r3 = rp.add_run(CREDITO_DIR_POS)
    r3.bold = True
    r3.font.size = Pt(9)
    r3.font.color.rgb = BRANCO


def add_title_band(header, titulo, ementa, image_bytes):
    """Adiciona uma tabela com a imagem do banner a esquerda e o
    titulo/ementa a direita, igual a arte do template."""
    table = header.add_table(rows=1, cols=2, width=Cm(17))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    left_cell, right_cell = table.rows[0].cells
    left_cell.width = Cm(6)
    right_cell.width = Cm(11)

    left_p = left_cell.paragraphs[0]
    left_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if image_bytes:
        run = left_p.add_run()
        run.add_picture(io.BytesIO(image_bytes), width=Cm(5.8))

    rp1 = right_cell.paragraphs[0]
    rp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = rp1.add_run(titulo)
    r1.bold = True
    r1.font.size = Pt(13)
    r1.font.color.rgb = AZUL_TITULO

    rp2 = right_cell.add_paragraph()
    rp2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r2 = rp2.add_run(ementa)
    r2.bold = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = AZUL_TITULO


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


def diagramar(input_stream, titulo=None, subtitulo=None):
    """Recebe um stream (BytesIO) do docx cru e devolve (Document, n_gabarito).

    O documento devolvido E O PROPRIO ARQUIVO DE ENTRADA com o corpo
    reescrito -- por isso o cabecalho/rodape originais (banner, logo,
    credito) sao preservados automaticamente."""
    doc = Document(input_stream)

    all_texts = [p.text.strip() for p in doc.paragraphs]
    non_empty = [t for t in all_texts if t]
    titulo = titulo or (non_empty[0] if len(non_empty) > 0 else "")
    subtitulo = subtitulo or (non_empty[1] if len(non_empty) > 1 else "")

    # paragrafos originais (texto) antes de apagar o corpo
    original_paragraph_texts = all_texts

    header = doc.sections[0].header
    footer = doc.sections[0].footer
    image_bytes = extract_header_image(header)

    add_credit_band(header)
    add_title_band(header, titulo, subtitulo, image_bytes)
    add_credit_band(footer)

    gabarito = extract_gabarito(original_paragraph_texts)

    clear_body(doc)

    skip_first_lines = 2
    count_skipped = 0
    for text in original_paragraph_texts:
        text = text.strip()

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
            doc.add_page_break()
            h = doc.add_heading(text, level=1)
            style_runs(h, bold=True, size=14, color=AZUL)
            continue

        if gab_head_m:
            doc.add_page_break()
            h = doc.add_heading(text, level=2)
            style_runs(h, bold=True, size=13, color=AZUL)
            continue

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
