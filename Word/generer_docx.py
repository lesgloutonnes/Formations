#!/usr/bin/env python3
"""Génère les DOCX d'exercices et de réalisés Word (modules 1 et 2).

Chaque réalisé = fichier d'exercice + uniquement les actions demandées
par les consignes de la formation interactive.
"""

from __future__ import annotations

import io
import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap, qn
from docx.shared import Cm, Pt, RGBColor, Inches
from fpdf import FPDF
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
M1_EX = ROOT / "Module 1" / "DOCX" / "Exercices"
M1_RE = ROOT / "Module 1" / "DOCX" / "Realises"
M2_EX = ROOT / "Module 2" / "DOCX" / "Exercices"
M2_RE = ROOT / "Module 2" / "DOCX" / "Realises"
ASSETS = ROOT / "assets"

CHC_BLUE = RGBColor(0x2B, 0x57, 0x9A)
CHC_FILL = "DDEBF7"
NSMAP_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


# ---------------------------------------------------------------------------
# Utilitaires OOXML
# ---------------------------------------------------------------------------

def _a4(doc: Document, left=2.5, right=2.5, top=2.0, bottom=2.0) -> None:
    s = doc.sections[0]
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.left_margin = Cm(left)
    s.right_margin = Cm(right)
    s.top_margin = Cm(top)
    s.bottom_margin = Cm(bottom)


def _set_run_font(run, name="Calibri", size=11, bold=None, italic=None, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def add_text(doc, text, style=None, bold=False, italic=False, align=None, space_after=8, space_before=0):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    run = p.add_run(text)
    _set_run_font(run, bold=bold or None, italic=italic or None)
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    return p


def add_mixed(doc, parts, style=None, align=None, space_after=8):
    """parts = liste de (texte, {bold, italic, size, name, color})."""
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    for text, opt in parts:
        run = p.add_run(text)
        _set_run_font(
            run,
            name=opt.get("name", "Calibri"),
            size=opt.get("size", 11),
            bold=opt.get("bold"),
            italic=opt.get("italic"),
            color=opt.get("color"),
        )
        if opt.get("underline"):
            run.underline = True
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    return p


def shade_and_border(paragraph, fill=CHC_FILL):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "12")
        el.set(qn("w:space"), "4")
        el.set(qn("w:color"), "2B579A")
        pBdr.append(el)
    pPr.append(pBdr)


def add_page_break(doc):
    doc.add_page_break()


def add_toc(doc, placeholder="Mettez le champ à jour (F9) pour afficher toute la table des matières."):
    p = doc.add_paragraph()
    run = p.add_run()
    r = run._r
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.set(qn("xml:space"), "preserve")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r.append(begin)
    r.append(instr)
    r.append(sep)
    r.append(text)
    r.append(end)
    p.paragraph_format.space_after = Pt(12)
    return p


def add_header_footer(doc, header_text, first_page_different=True):
    section = doc.sections[0]
    section.different_first_page_header_footer = first_page_different
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = ""
    run = hp.add_run(header_text)
    _set_run_font(run, size=10, color=CHC_BLUE, italic=True)
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = fp.add_run("Page ")
    _set_run_font(r1, size=10)
    r2 = fp.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = "1"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r2._r.append(fld_begin)
    r2._r.append(instr)
    r2._r.append(fld_sep)
    r2._r.append(t)
    r2._r.append(fld_end)


def add_empty_para(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    return p


def set_numpr(paragraph, num_id, ilvl=0):
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl")
    il.set(qn("w:val"), str(ilvl))
    nid = OxmlElement("w:numId")
    nid.set(qn("w:val"), str(num_id))
    numPr.append(il)
    numPr.append(nid)
    pPr.append(numPr)


def add_heading_numbering(doc, num_id=12):
    """Liste à plusieurs niveaux liée à Titre 1 / Titre 2."""
    numbering = doc.part.numbering_part._element
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(num_id))
    ml = OxmlElement("w:multiLevelType")
    ml.set(qn("w:val"), "multilevel")
    abstract.append(ml)
    for ilvl, pstyle, text in (
        (0, "Heading1", "%1."),
        (1, "Heading2", "%1.%2"),
        (2, "Heading3", "%1.%2.%3"),
    ):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        fmt = OxmlElement("w:numFmt")
        fmt.set(qn("w:val"), "decimal")
        lvl_text = OxmlElement("w:lvlText")
        lvl_text.set(qn("w:val"), text)
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), pstyle)
        lvlJc = OxmlElement("w:lvlJc")
        lvlJc.set(qn("w:val"), "left")
        pPr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(360 + ilvl * 360))
        ind.set(qn("w:hanging"), "360")
        pPr.append(ind)
        lvl.append(start)
        lvl.append(fmt)
        lvl.append(pStyle)
        lvl.append(lvl_text)
        lvl.append(lvlJc)
        lvl.append(pPr)
        abstract.append(lvl)
    numbering.insert(0, abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_id = OxmlElement("w:abstractNumId")
    abs_id.set(qn("w:val"), str(num_id))
    num.append(abs_id)
    numbering.append(num)
    return num_id


def add_list_numbering(doc, num_id=13, levels=2):
    numbering = doc.part.numbering_part._element
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(num_id))
    ml = OxmlElement("w:multiLevelType")
    ml.set(qn("w:val"), "multilevel")
    abstract.append(ml)
    for ilvl in range(levels):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        fmt = OxmlElement("w:numFmt")
        fmt.set(qn("w:val"), "decimal")
        lvl_text = OxmlElement("w:lvlText")
        lvl_text.set(qn("w:val"), "%1." if ilvl == 0 else "%1.%2")
        lvlJc = OxmlElement("w:lvlJc")
        lvlJc.set(qn("w:val"), "left")
        pPr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(720 + ilvl * 360))
        ind.set(qn("w:hanging"), "360")
        pPr.append(ind)
        lvl.append(start)
        lvl.append(fmt)
        lvl.append(lvl_text)
        lvl.append(lvlJc)
        lvl.append(pPr)
        abstract.append(lvl)
    numbering.insert(0, abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_id = OxmlElement("w:abstractNumId")
    abs_id.set(qn("w:val"), str(num_id))
    num.append(abs_id)
    numbering.append(num)
    return num_id


def add_formula_cell(cell, instr="=SUM(ABOVE)", display="9"):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = display
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(it)
    run._r.append(sep)
    run._r.append(t)
    run._r.append(end)


def add_equation(doc, caption="Taux de présence = inscrits / places × 100"):
    p = doc.add_paragraph()
    omath = OxmlElement("m:oMathPara")
    math = OxmlElement("m:oMath")
    def mt(text):
        r = OxmlElement("m:r")
        t = OxmlElement("m:t")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        r.append(t)
        return r
    math.append(mt("Taux = "))
    frac = OxmlElement("m:f")
    num = OxmlElement("m:num")
    num.append(mt("inscrits"))
    den = OxmlElement("m:den")
    den.append(mt("places"))
    frac.append(num)
    frac.append(den)
    math.append(frac)
    math.append(mt(" × 100"))
    omath.append(math)
    p._p.append(omath)
    add_text(doc, caption, italic=True, space_after=10)
    return p


def add_picture_inline(doc, path, width_cm=3.2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    p.paragraph_format.space_after = Pt(8)
    return p


def add_picture_square(doc, path, width_cm=3.5):
    """Image avec habillage Carré (ancre)."""
    p = doc.add_paragraph()
    run = p.add_run()
    inline_pic = run.add_picture(str(path), width=Cm(width_cm))
    # python-docx insère en inline ; on convertit en ancre wrapSquare
    drawing = run._r.find(qn("w:drawing"))
    if drawing is None:
        return p
    inline = drawing.find(qn("wp:inline"))
    if inline is None:
        return p
    extent = inline.find(qn("wp:extent"))
    cx = extent.get("cx")
    cy = extent.get("cy")
    docPr = inline.find(qn("wp:docPr"))
    graphic = inline.find(qn("a:graphic"))
    drawing.remove(inline)
    anchor = OxmlElement("wp:anchor")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "114300")
    anchor.set("distR", "114300")
    anchor.set("simplePos", "0")
    anchor.set("relativeHeight", "251658240")
    anchor.set("behindDoc", "0")
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")
    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    posH = OxmlElement("wp:positionH")
    posH.set("relativeFrom", "column")
    posH_off = OxmlElement("wp:posOffset")
    posH_off.text = "0"
    posH.append(posH_off)
    posV = OxmlElement("wp:positionV")
    posV.set("relativeFrom", "paragraph")
    posV_off = OxmlElement("wp:posOffset")
    posV_off.text = "0"
    posV.append(posV_off)
    ext = OxmlElement("wp:extent")
    ext.set("cx", cx)
    ext.set("cy", cy)
    wrap = OxmlElement("wp:wrapSquare")
    wrap.set("wrapText", "bothSides")
    anchor.append(simple)
    anchor.append(posH)
    anchor.append(posV)
    anchor.append(ext)
    effect = OxmlElement("wp:effectExtent")
    for k in ("l", "t", "r", "b"):
        effect.set(k, "0")
    anchor.append(effect)
    anchor.append(wrap)
    if docPr is not None:
        anchor.append(docPr)
    old_cNv = inline.find(qn("wp:cNvGraphicFramePr"))
    if old_cNv is not None:
        anchor.append(old_cNv)
    if graphic is not None:
        anchor.append(graphic)
    drawing.append(anchor)
    return p


def apply_blue_theme(path: Path):
    """Remplace le jeu de couleurs du thème (équivalent pédagogique d'un thème)."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/theme/theme1.xml":
                text = data.decode("utf-8")
                replacements = {
                    "5B9BD5": "2B579A",
                    "ED7D31": "5B9BD5",
                    "A5A5A5": "9DC3E6",
                    "FFC000": "1F4E79",
                    "4472C4": "2B579A",
                    "70AD47": "2E75B6",
                }
                for old, new in replacements.items():
                    text = text.replace(old, new).replace(old.lower(), new)
                text = text.replace('name="Office Theme"', 'name="Bleu CHC"')
                data = text.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


def save_as_dotx(docx_path: Path, dotx_path: Path):
    shutil.copy2(docx_path, dotx_path)
    tmp = dotx_path.with_suffix(".tmpzip")
    with zipfile.ZipFile(dotx_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = data.decode("utf-8")
                text = text.replace(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                )
                data = text.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(dotx_path)


def set_section_columns(section, num):
    cols = section._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        section._sectPr.append(cols)
    cols.set(qn("w:num"), str(num))
    cols.set(qn("w:space"), "720")


def style_table(table, header=True):
    table.style = "Table Grid"
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "8")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "2B579A")
        borders.append(el)
    tblPr.append(borders)
    if header:
        for cell in table.rows[0].cells:
            shade = OxmlElement("w:shd")
            shade.set(qn("w:fill"), CHC_FILL)
            shade.set(qn("w:val"), "clear")
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(shade)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.bold = True
                    r.font.color.rgb = CHC_BLUE


def fill_table(table, rows):
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            table.cell(i, j).text = str(val)


def consigne(doc, text):
    p = add_text(doc, "Consigne : " + text, italic=True, space_after=12)
    for r in p.runs:
        r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    return p


def new_doc():
    doc = Document()
    _a4(doc)
    return doc


# ---------------------------------------------------------------------------
# Contenu fil rouge procédure (M1)
# ---------------------------------------------------------------------------

PROC_TITLE = "Procédure interne - Accueil d'un nouveau collaborateur"

PROC_PARTS = [
    (
        "Avant l'arrivée",
        [
            ("Préparer les accès", "Le responsable vérifie que les demandes d'accès sont introduites avant l'arrivée du collaborateur."),
            ("Préparer le matériel", "Le poste de travail, le badge et les informations utiles sont préparés avec les personnes concernées."),
        ],
    ),
    (
        "Le premier jour",
        [
            ("Accueil dans l'équipe", "Le collaborateur est accueilli par son référent et reçoit les informations pratiques."),
            ("Présentation des outils", "Les outils principaux sont présentés : session Windows, messagerie, intranet et dossiers partagés."),
        ],
    ),
    (
        "Après la première semaine",
        [
            ("Point de suivi", "Un court échange permet de vérifier que les accès, les outils et les premières tâches sont compris."),
            ("Ajustements éventuels", "Les difficultés sont notées et transmises aux personnes concernées."),
        ],
    ),
]


def write_procedure_styled(doc, tools_as_list=False, extra_h2=None):
    add_text(doc, PROC_TITLE, style="Title", space_after=12)
    for h1, subs in PROC_PARTS:
        add_text(doc, h1, style="Heading 1", space_before=12, space_after=6)
        for h2, body in subs:
            add_text(doc, h2, style="Heading 2", space_before=8, space_after=4)
            if tools_as_list and h2 == "Présentation des outils":
                add_text(doc, "Les outils principaux sont présentés :", space_after=4)
                for item in ("Session Windows", "Messagerie", "Intranet", "Dossiers partagés"):
                    doc.add_paragraph(item, style="List Bullet")
            else:
                add_text(doc, body, space_after=8)
            if extra_h2 and extra_h2[0] == h1 and extra_h2[1] == h2:
                add_text(doc, extra_h2[2], style="Heading 2", space_before=8, space_after=4)
                add_text(doc, extra_h2[3], space_after=8)


# ---------------------------------------------------------------------------
# Logo + xlsx + pdf
# ---------------------------------------------------------------------------

def make_logo() -> Path:
    ASSETS.mkdir(exist_ok=True)
    path = ASSETS / "chc_logo.png"
    img = Image.new("RGB", (420, 180), (43, 87, 154))
    draw = ImageDraw.Draw(img)
    draw.rectangle((12, 12, 407, 167), outline=(157, 195, 230), width=6)
    try:
        font_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 64)
        font_sm = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 22)
    except OSError:
        font_big = ImageFont.load_default()
        font_sm = font_big
    draw.text((150, 40), "CHC", font=font_big, fill=(255, 255, 255))
    draw.text((70, 120), "Service Formation", font=font_sm, fill=(221, 235, 247))
    img.save(path)
    return path


def make_xlsx(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi"
    ws.append(["Action", "Responsable", "Heures"])
    ws.append(["Préparer les accès", "Service IT", 2])
    ws.append(["Préparer le matériel", "Logistique", 3])
    ws.append(["Accueil du collaborateur", "Référent", 4])
    wb.save(path)


def make_pdf(path: Path, title: str, lines: list[str]):
    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    usable = pdf.w - pdf.l_margin - pdf.r_margin
    pdf.set_font("Helvetica", "B", 16)
    pdf.multi_cell(usable, 10, title, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=11)
    pdf.ln(4)
    for line in lines:
        if not line.strip():
            pdf.ln(5)
            continue
        pdf.multi_cell(usable, 7, line, new_x="LMARGIN", new_y="NEXT")
    pdf.output(str(path))


# ---------------------------------------------------------------------------
# MODULE 1
# ---------------------------------------------------------------------------

def m1_ex01():
    doc = new_doc()
    add_text(doc, PROC_TITLE)
    add_text(doc, "Avant l'arrivée")
    add_text(doc, "Préparer les accès")
    add_text(doc, "Le responsable vérifie que les demandes d'accès sont introduites avant l'arrivée du collaborateur")
    add_text(doc, "Préparer le matériel")
    add_text(doc, "Le poste de travail, le badge et les informations utiles sont préparés")
    add_text(doc, "Le premier jour")
    add_text(doc, "Accueil dans l'équipe")
    add_text(doc, "Présentation des outils")
    add_text(doc, "Les outils principaux sont : session Windows, messagerie, intranet, dossiers partagés")
    add_text(doc, "Après la première semaine")
    add_text(doc, "Point de suivi")
    add_text(doc, "Ajustements éventuels")
    consigne(
        doc,
        "appliquez le style Titre au titre principal, puis Titre 1 aux trois grandes parties, "
        "Titre 2 aux sous-parties et Normal au texte. Aérez (interligne, espacement) et mettez "
        "un mot clé en gras. Transformez l'énumération des outils en liste à puces, encadrez le "
        "titre (bordure + trame légère) et reproduisez la mise en valeur avec le pinceau.",
    )
    doc.save(M1_EX / "01_Exercice_Texte_brut_Procedure_accueil.docx")


def m1_re01():
    doc = new_doc()
    title = add_text(doc, PROC_TITLE, style="Title", space_after=14)
    shade_and_border(title)
    add_text(doc, "Avant l'arrivée", style="Heading 1", space_before=14, space_after=6)
    add_text(doc, "Préparer les accès", style="Heading 2", space_before=8, space_after=4)
    add_mixed(doc, [
        ("Le ", {}),
        ("responsable", {"bold": True}),
        (" vérifie que les demandes d'accès sont introduites avant l'arrivée du collaborateur.", {}),
    ])
    add_text(doc, "Préparer le matériel", style="Heading 2", space_before=8, space_after=4)
    add_text(doc, "Le poste de travail, le badge et les informations utiles sont préparés avec les personnes concernées.")
    add_text(doc, "Le premier jour", style="Heading 1", space_before=14, space_after=6)
    add_text(doc, "Accueil dans l'équipe", style="Heading 2", space_before=8, space_after=4)
    add_mixed(doc, [
        ("Le collaborateur est accueilli par son ", {}),
        ("référent", {"bold": True}),
        (" et reçoit les informations pratiques.", {}),
    ])
    add_text(doc, "Présentation des outils", style="Heading 2", space_before=8, space_after=4)
    add_text(doc, "Les outils principaux sont présentés :", space_after=4)
    for item in ("Session Windows", "Messagerie", "Intranet", "Dossiers partagés"):
        doc.add_paragraph(item, style="List Bullet")
    add_text(doc, "Après la première semaine", style="Heading 1", space_before=14, space_after=6)
    add_text(doc, "Point de suivi", style="Heading 2", space_before=8, space_after=4)
    add_text(doc, "Un court échange permet de vérifier que les accès, les outils et les premières tâches sont compris.")
    add_text(doc, "Ajustements éventuels", style="Heading 2", space_before=8, space_after=4)
    add_text(doc, "Les difficultés sont notées et transmises aux personnes concernées.")
    doc.save(M1_RE / "01_Realise_Procedure_accueil_structuree.docx")


def m1_ex02():
    doc = new_doc()
    write_procedure_styled(doc)
    consigne(
        doc,
        "placez le curseur après le titre principal, insérez une table des matières automatique, "
        "ajoutez une sous-partie (par exemple « Remise du badge » dans « Le premier jour »), "
        "puis mettez toute la table à jour (F9).",
    )
    doc.save(M1_EX / "02_Exercice_Table_des_matieres.docx")


def m1_re02():
    doc = new_doc()
    add_text(doc, PROC_TITLE, style="Title", space_after=8)
    add_toc(doc)
    write_procedure_styled(
        doc,
        extra_h2=(
            "Le premier jour",
            "Accueil dans l'équipe",
            "Remise du badge",
            "Le badge est remis au collaborateur après la vérification des accès.",
        ),
    )
    # write_procedure_styled re-adds the title — rebuild cleanly instead
    doc.save(M1_RE / "02_Realise_Table_des_matieres.docx")


def _rewrite_m1_re02():
    """Réécrit le réalisé 2 sans titre en double."""
    doc = new_doc()
    add_text(doc, PROC_TITLE, style="Title", space_after=8)
    add_toc(doc)
    for h1, subs in PROC_PARTS:
        add_text(doc, h1, style="Heading 1", space_before=12, space_after=6)
        for h2, body in subs:
            add_text(doc, h2, style="Heading 2", space_before=8, space_after=4)
            add_text(doc, body, space_after=8)
            if h1 == "Le premier jour" and h2 == "Accueil dans l'équipe":
                add_text(doc, "Remise du badge", style="Heading 2", space_before=8, space_after=4)
                add_text(doc, "Le badge est remis au collaborateur après la vérification des accès.", space_after=8)
    path = M1_RE / "02_Realise_Table_des_matieres.docx"
    doc.save(path)


def m1_ex03():
    doc = new_doc()
    consigne(
        doc,
        "ajoutez une page de garde (titre, auteur, date), un saut de page, placez la table des "
        "matières sur la page suivante, saisissez l'en-tête « Procédure interne », numérotez le "
        "pied de page avec Première page différente, appliquez un thème, puis mettez la table à jour (F9).",
    )
    write_procedure_styled(doc)
    doc.save(M1_EX / "03_Exercice_Mise_en_page.docx")


def m1_re03():
    doc = new_doc()
    add_text(doc, "Procédure interne", style="Title", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_text(doc, "Accueil d'un nouveau collaborateur", style="Subtitle", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    add_text(doc, "Auteur : Service Formation CHC", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "Date : 23 septembre 2026", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_page_break(doc)
    add_text(doc, "Table des matières", style="Heading 1", space_after=8)
    add_toc(doc)
    add_page_break(doc)
    for h1, subs in PROC_PARTS:
        add_text(doc, h1, style="Heading 1", space_before=12, space_after=6)
        for h2, body in subs:
            add_text(doc, h2, style="Heading 2", space_before=8, space_after=4)
            add_text(doc, body, space_after=8)
    add_header_footer(doc, "Procédure interne", first_page_different=True)
    path = M1_RE / "03_Realise_Mise_en_page.docx"
    doc.save(path)
    apply_blue_theme(path)


def m1_ex04():
    doc = new_doc()
    write_procedure_styled(doc)
    consigne(
        doc,
        "insérez un tableau 3 colonnes Action / Responsable / Échéance, ajoutez 4 lignes puis "
        "une ligne supplémentaire avec Tab, appliquez un style de tableau sobre, insérez le logo "
        "chc_logo.png (habillage Aligné sur le texte), créez un WordArt ou un SmartArt de processus, "
        "puis ajoutez une liste numérotée et une liste à puces.",
    )
    doc.save(M1_EX / "04_Exercice_Tableaux_listes_images.docx")


def m1_re04(logo: Path):
    doc = new_doc()
    write_procedure_styled(doc)
    add_text(doc, "Tableau de suivi", style="Heading 1", space_before=14, space_after=6)
    rows = [
        ("Action", "Responsable", "Échéance"),
        ("Préparer les accès", "Service IT", "J-2"),
        ("Préparer le matériel", "Logistique", "J-1"),
        ("Accueil du collaborateur", "Référent", "Jour J"),
        ("Présentation des outils", "Référent", "Jour J"),
        ("Point de suivi", "Responsable", "J+7"),
    ]
    table = doc.add_table(rows=len(rows), cols=3)
    fill_table(table, rows)
    style_table(table)
    add_text(doc, "Déroulé du premier jour", style="Heading 2", space_before=14, space_after=4)
    for item in (
        "Accueil par le référent.",
        "Présentation des outils principaux.",
        "Lecture des consignes utiles.",
        "Planification d'un point de suivi.",
    ):
        doc.add_paragraph(item, style="List Number")
    add_text(doc, "Points d'attention", style="Heading 2", space_before=12, space_after=4)
    for item in ("Vérifier le badge", "Tester la session Windows", "Remettre les contacts utiles"):
        doc.add_paragraph(item, style="List Bullet")
    art = add_text(doc, "Accueil CHC", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=16, space_after=8)
    for r in art.runs:
        r.font.size = Pt(28)
        r.font.color.rgb = CHC_BLUE
        r.bold = True
        r.font.name = "Calibri"
    add_text(doc, "WordArt — titre décoratif du service", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_picture_inline(doc, logo, 3.4)
    add_text(doc, "Logo du service (habillage Aligné sur le texte).", italic=True)
    doc.save(M1_RE / "04_Realise_Tableaux_listes_images.docx")


def m1_ex05():
    doc = new_doc()
    add_text(doc, "Note interne - Organisation d'une réunion", style="Title")
    consigne(
        doc,
        "affichez les marques de mise en forme (¶), supprimez les doubles espaces et le paragraphe "
        "vide, déplacez la rubrique mal placée, corrigez les fautes (« lieux », « point »), lancez "
        "le correcteur (F7) et remplacez un mot par un synonyme.",
    )
    add_text(doc, "Bonjour,")
    # doubles espaces volontairement
    p = doc.add_paragraph()
    p.add_run("La réunion  aura lieux vendredi à 10h  dans la salle 2.")
    add_empty_para(doc)
    add_text(doc, "Merci de venir avec les point à discuter.")
    add_text(doc, "Ordre du jour :")
    add_text(doc, "Divers")
    add_text(doc, "Budget")
    add_text(doc, "Planning")
    doc.save(M1_EX / "05_Exercice_Document_a_relire.docx")


def m1_re05():
    doc = new_doc()
    add_text(doc, "Note interne - Organisation d'une réunion", style="Title")
    add_text(doc, "Bonjour,")
    add_text(doc, "La réunion aura lieu vendredi à 10h dans le local 2.")
    add_text(doc, "Merci de venir avec les points à discuter.")
    add_text(doc, "Ordre du jour :")
    for item in ("Budget", "Planning", "Divers"):
        doc.add_paragraph(item, style="List Bullet")
    doc.save(M1_RE / "05_Realise_Document_relu_corrige.docx")


def m1_ex06():
    doc = new_doc()
    add_text(doc, "Exercice courrier", style="Title")
    consigne(
        doc,
        "ce texte a été saisi au kilomètre. Positionnez l'expéditeur en haut et la date alignée à "
        "droite, ajoutez une ligne Objet et la formule d'appel, aérez le corps (un paragraphe par "
        "idée, justifié), ajoutez la formule de politesse et la signature, ajustez les marges pour "
        "tenir sur une page, puis prévisualisez l'impression.",
    )
    add_text(doc, "CHC - Service Formation")
    add_text(doc, "Rue de Hesbaye 75")
    add_text(doc, "4000 Liège")
    add_text(doc, "Madame, Monsieur,")
    add_text(
        doc,
        "Nous avons le plaisir de vous confirmer votre inscription à la formation Word - Module 1 "
        "Débutant. La session se déroulera dans nos locaux et débutera à 9h00. Merci de vous "
        "présenter quelques minutes à l'avance.",
    )
    add_text(
        doc,
        "Un ordinateur équipé de Microsoft Word sera mis à votre disposition. Vous pouvez également "
        "apporter vos propres documents afin de les retravailler pendant les ateliers.",
    )
    add_text(doc, "Pour toute question relative à l'organisation, n'hésitez pas à contacter le service formation.")
    add_text(doc, "Nous vous prions d'agréer, Madame, Monsieur, nos salutations distinguées.")
    add_text(doc, "Le service Formation")
    add_text(doc, "CHC")
    doc.save(M1_EX / "06_Exercice_Courrier.docx")


def m1_re06():
    doc = new_doc()
    _a4(doc, left=2.2, right=2.2, top=1.8, bottom=1.8)
    add_text(doc, "CHC - Service Formation", space_after=0)
    add_text(doc, "Rue de Hesbaye 75", space_after=0)
    add_text(doc, "4000 Liège", space_after=12)
    add_text(doc, "Liège, le 23 septembre 2026", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=16)
    add_text(doc, "Objet : confirmation d'inscription - Word Module 1", bold=True, space_after=16)
    add_text(doc, "Madame, Monsieur,", space_after=12)
    for body in (
        "Nous avons le plaisir de vous confirmer votre inscription à la formation Word - Module 1 "
        "Débutant. La session se déroulera dans nos locaux et débutera à 9h00. Merci de vous "
        "présenter quelques minutes à l'avance.",
        "Un ordinateur équipé de Microsoft Word sera mis à votre disposition. Vous pouvez également "
        "apporter vos propres documents afin de les retravailler pendant les ateliers.",
        "Pour toute question relative à l'organisation, n'hésitez pas à contacter le service formation.",
    ):
        p = add_text(doc, body, space_after=10)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_text(doc, "Nous vous prions d'agréer, Madame, Monsieur, nos salutations distinguées.", space_before=10, space_after=16)
    add_text(doc, "Le service Formation", space_after=0)
    add_text(doc, "CHC", space_after=0)
    doc.save(M1_RE / "06_Realise_Courrier.docx")


def m1_ex07():
    doc = new_doc()
    add_text(doc, "Atelier final - Document complet CHC", style="Title")
    add_text(doc, "Choisissez un cas : procédure interne, compte rendu, note interne ou courrier.", style="Heading 1")
    for item in (
        "Créez un document de 3 à 5 pages avec une page de garde.",
        "Structurez avec les styles Titre 1 et Titre 2.",
        "Insérez une table des matières automatique.",
        "Ajoutez un tableau utile (actions, décisions ou responsables).",
        "Ajoutez une liste à puces ou numérotée.",
        "Ajoutez une illustration (image, WordArt ou SmartArt).",
        "Ajoutez un en-tête ou un pied de page avec numérotation.",
        "Relisez le document et corrigez l'orthographe.",
        "Mettez la table des matières à jour et prévisualisez avant impression.",
    ):
        doc.add_paragraph(item, style="List Bullet")
    doc.save(M1_EX / "07_Exercice_Atelier_final_CHC.docx")


def m1_re07(logo: Path):
    doc = new_doc()
    add_text(doc, "Compte rendu de réunion", style="Title", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_text(doc, "Suivi d'équipe — Service Formation CHC", style="Subtitle", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "Date : 23 septembre 2026", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "Participants : responsable, référents, service IT", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_picture_inline(doc, logo, 3.0)
    add_page_break(doc)
    add_text(doc, "Table des matières", style="Heading 1")
    add_toc(doc)
    add_page_break(doc)

    sections = [
        (
            "Objectif de la réunion",
            [
                (
                    "Contexte",
                    "La réunion vise à faire le point sur l'accueil des nouveaux collaborateurs et à "
                    "valider les documents de formation Word. Le service Formation souhaite un compte "
                    "rendu lisible, structuré et prêt à diffuser sans explication orale.",
                ),
                (
                    "Périmètre",
                    "Les échanges portent sur la procédure d'accueil, le calendrier des sessions et les "
                    "supports à mettre à jour. Les sujets hors périmètre sont reportés au prochain comité.",
                ),
            ],
        ),
        (
            "Décisions prises",
            [
                (
                    "Documents de référence",
                    "Le groupe valide le modèle de procédure d'accueil et demande sa diffusion aux "
                    "référents. Les titres doivent rester gérés par les styles Titre 1 et Titre 2.",
                ),
                (
                    "Organisation des sessions",
                    "Les sessions Word Module 1 restent le mardi matin. Une session de rattrapage est "
                    "prévue pour les absents justifiés. Le service Formation confirme les salles.",
                ),
            ],
        ),
        (
            "Actions à suivre",
            [
                (
                    "Suivi immédiat",
                    "Chaque action ci-dessous a un responsable et une échéance. Le tableau sera relu "
                    "avant envoi du compte rendu.",
                ),
                (
                    "Préparation du prochain point",
                    "Le prochain point d'avancement reprendra les actions ouvertes et les documents "
                    "modifiés. Un rappel sera envoyé 48 heures avant la réunion.",
                ),
            ],
        ),
        (
            "Commentaires et validation",
            [
                (
                    "Relecture",
                    "Le texte a été relu : orthographe corrigée, phrases clarifiées, vocabulaire homogène. "
                    "Aucune mention personnelle n'apparaît dans le document diffusé.",
                ),
                (
                    "Diffusion",
                    "Après mise à jour de la table des matières et aperçu avant impression, le compte "
                    "rendu est envoyé aux participants et déposé dans le dossier partagé du service.",
                ),
            ],
        ),
    ]
    for h1, subs in sections:
        add_text(doc, h1, style="Heading 1", space_before=12, space_after=6)
        for h2, body in subs:
            add_text(doc, h2, style="Heading 2", space_before=8, space_after=4)
            add_text(doc, body, space_after=8)

    add_text(doc, "Tableau des décisions", style="Heading 2", space_before=10)
    rows = [
        ("Décision", "Responsable", "Échéance"),
        ("Valider le nouveau modèle", "Service Formation", "Semaine 1"),
        ("Diffuser la procédure", "Référent", "Semaine 2"),
        ("Planifier le suivi", "Responsable", "Semaine 3"),
    ]
    table = doc.add_table(rows=len(rows), cols=3)
    fill_table(table, rows)
    style_table(table)

    add_text(doc, "Prochaines étapes", style="Heading 2", space_before=14)
    for item in (
        "Envoyer le compte rendu aux participants.",
        "Mettre à jour les documents concernés.",
        "Préparer le prochain point d'avancement.",
    ):
        doc.add_paragraph(item, style="List Number")

    add_header_footer(doc, "Compte rendu de réunion", first_page_different=True)
    doc.save(M1_RE / "07_Realise_Atelier_final_Compte_rendu.docx")


# ---------------------------------------------------------------------------
# MODULE 2
# ---------------------------------------------------------------------------

def m2_ex01():
    doc = new_doc()
    add_text(doc, "Personnaliser Word - document de test", style="Title")
    consigne(
        doc,
        "définissez une police et des marges par défaut, ajoutez 3 commandes à la barre d'accès "
        "rapide, créez une correction automatique chc → CHC - Service Formation, transformez le "
        "bloc de signature en QuickPart, réinsérez-le plus bas pour vérifier la réutilisation, "
        "puis exportez le document en PDF.",
    )
    add_text(doc, "Bloc de signature à transformer en QuickPart :", style="Heading 1")
    add_text(doc, "CHC - Service Formation", space_after=0)
    add_text(doc, "Rue de Hesbaye 75 - 4000 Liège", space_after=0)
    add_text(doc, "formation@chc.be", space_after=12)
    add_text(doc, "Texte à tester avec la correction automatique :", style="Heading 1")
    add_text(doc, "Tapez chc puis espace pour vérifier que la correction automatique développe le texte.")
    doc.save(M2_EX / "M2_01_Exercice_Personnalisation.docx")


def m2_re01():
    doc = new_doc()
    _a4(doc, left=2.0, right=2.0, top=2.0, bottom=2.0)
    add_text(doc, "Personnaliser Word - résultat attendu", style="Title")
    add_text(
        doc,
        "Police et marges de ce document réglées (Calibri, 2 cm). Les commandes de la barre d'accès "
        "rapide et la correction automatique chc → CHC - Service Formation se règlent dans Word "
        "(Normal.dotm) : elles ne se stockent pas dans le fichier d'exercice.",
    )
    add_text(doc, "Correction automatique testée", style="Heading 1")
    add_text(doc, "CHC - Service Formation a bien remplacé la saisie « chc » suivie d'un espace.")
    add_text(doc, "Signature (QuickPart)", style="Heading 1")
    add_text(doc, "CHC - Service Formation", space_after=0)
    add_text(doc, "Rue de Hesbaye 75 - 4000 Liège", space_after=0)
    add_text(doc, "formation@chc.be", space_after=12)
    add_text(doc, "QuickPart réinséré plus bas pour vérifier la réutilisation", style="Heading 1")
    add_text(doc, "CHC - Service Formation", space_after=0)
    add_text(doc, "Rue de Hesbaye 75 - 4000 Liège", space_after=0)
    add_text(doc, "formation@chc.be", space_after=12)
    add_text(doc, "Export PDF : voir M2_01_Realise_Personnalisation.pdf dans le même dossier.", italic=True)
    doc.save(M2_RE / "M2_01_Realise_Personnalisation.docx")
    make_pdf(
        M2_RE / "M2_01_Realise_Personnalisation.pdf",
        "CHC - Service Formation",
        [
            "Export PDF du document de personnalisation.",
            "",
            "CHC - Service Formation",
            "Rue de Hesbaye 75 - 4000 Liege",
            "formation@chc.be",
            "",
            "Correction automatique : chc -> CHC - Service Formation",
            "QuickPart de signature reutilise dans le document Word.",
        ],
    )


def _ensure_custom_styles(doc):
    styles = doc.styles
    if "TitreCHC" not in [s.name for s in styles]:
        st = styles.add_style("TitreCHC", WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = styles["Heading 1"]
        st.font.name = "Calibri"
        st.font.size = Pt(16)
        st.font.bold = True
        st.font.color.rgb = CHC_BLUE
        st.paragraph_format.space_before = Pt(14)
        st.paragraph_format.space_after = Pt(8)
    if "CorpsCHC" not in [s.name for s in styles]:
        sc = styles.add_style("CorpsCHC", WD_STYLE_TYPE.PARAGRAPH)
        sc.base_style = styles["Normal"]
        sc.font.name = "Calibri"
        sc.font.size = Pt(11)
        sc.paragraph_format.space_after = Pt(8)
        sc.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    styles["TitreCHC"].next_paragraph_style = styles["CorpsCHC"]
    styles["CorpsCHC"].next_paragraph_style = styles["CorpsCHC"]


def m2_ex02():
    doc = new_doc()
    add_text(doc, "Automatiser la présentation", style="Title")
    consigne(
        doc,
        "ce document a été mis en forme manuellement (gras, tailles, couleurs répétées). Repérez "
        "les mises en forme répétitives, créez un style de titre et un style de corps, appliquez-les "
        "partout, réglez le style suivant (titre → corps), modifiez un style pour vérifier la mise "
        "à jour globale, puis enregistrez le tout comme modèle (.dotx).",
    )
    add_mixed(doc, [("RAPPORT D'ACTIVITÉ", {"bold": True, "size": 18, "name": "Calibri"})], space_after=12)
    add_mixed(doc, [("1. Introduction", {"bold": True, "size": 14, "underline": True})], space_after=6)
    add_mixed(doc, [("Texte courant saisi en Calibri 11, parfois en Arial 12 selon les paragraphes.", {"size": 11, "name": "Calibri"})])
    add_mixed(doc, [("2. Bilan", {"bold": True, "size": 14, "underline": True})], space_after=6)
    p = add_mixed(
        doc,
        [("Encore du texte avec des espacements irréguliers et des retraits faits à la main.", {"size": 12, "name": "Arial"})],
    )
    p.paragraph_format.left_indent = Cm(1.25)
    p.paragraph_format.space_before = Pt(16)
    add_mixed(doc, [("3. Perspectives", {"bold": True, "size": 14, "underline": True})], space_after=6)
    add_mixed(doc, [("Conclusion mise en forme une fois de plus à la main.", {"bold": True, "size": 11, "name": "Calibri", "color": CHC_BLUE})])
    doc.save(M2_EX / "M2_02_Exercice_Styles_Modele.docx")


def m2_re02():
    doc = new_doc()
    _ensure_custom_styles(doc)
    add_text(doc, "Rapport d'activité", style="Title")
    add_text(doc, "Introduction", style="TitreCHC")
    add_text(doc, "Texte courant au style CorpsCHC, homogène sur tout le document.", style="CorpsCHC")
    add_text(doc, "Bilan", style="TitreCHC")
    add_text(doc, "Les espacements et retraits sont désormais gérés par les styles, plus par des retraits manuels.", style="CorpsCHC")
    add_text(doc, "Perspectives", style="TitreCHC")
    add_text(
        doc,
        "Conclusion au style CorpsCHC. Le style suivant du TitreCHC est CorpsCHC. L'ensemble est "
        "enregistré comme modèle réutilisable M2_02_Realise_Modele.dotx.",
        style="CorpsCHC",
    )
    path = M2_RE / "M2_02_Realise_Styles_Modele.docx"
    doc.save(path)
    save_as_dotx(path, M2_RE / "M2_02_Realise_Modele.dotx")


def m2_ex03():
    doc = new_doc()
    add_text(doc, "Construire un document structuré", style="Title")
    consigne(
        doc,
        "vérifiez les styles hiérarchiques sur les titres, activez la liste à plusieurs niveaux "
        "liée aux titres (1, 1.1, 1.2), créez une liste hiérarchisée dans une sous-partie, insérez "
        "un sommaire automatique, ajoutez en-tête (titre) et pied de page (numéro) avec première "
        "page différente, puis ajoutez une sous-partie et mettez le sommaire à jour (F9).",
    )
    structure = [
        ("Présentation du service", ["Missions", "Organisation"]),
        ("Procédures internes", ["Accueil d'un collaborateur", "Gestion des accès", "Suivi des demandes"]),
        ("Annexes", ["Contacts utiles", "Documents de référence"]),
    ]
    for h1, h2s in structure:
        add_text(doc, h1, style="Heading 1")
        for h2 in h2s:
            add_text(doc, h2, style="Heading 2")
            add_text(doc, "Texte d'exemple à compléter pour cette sous-partie.")
    doc.save(M2_EX / "M2_03_Exercice_Document_structure.docx")


def m2_re03():
    doc = new_doc()
    hid = add_heading_numbering(doc, 12)
    lid = add_list_numbering(doc, 13)
    add_text(doc, "Document structuré - résultat", style="Title")
    add_toc(doc)
    add_page_break(doc)
    structure = [
        ("Présentation du service", ["Missions", "Organisation"]),
        ("Procédures internes", ["Accueil d'un collaborateur", "Gestion des accès", "Suivi des demandes", "Évaluation des besoins"]),
        ("Annexes", ["Contacts utiles", "Documents de référence"]),
    ]
    for h1, h2s in structure:
        p = add_text(doc, h1, style="Heading 1")
        set_numpr(p, hid, 0)
        for h2 in h2s:
            p2 = add_text(doc, h2, style="Heading 2")
            set_numpr(p2, hid, 1)
            add_text(doc, "Texte d'exemple complété pour cette sous-partie.")
            if h2 == "Missions":
                add_text(doc, "Liste hiérarchisée des missions :", space_after=4)
                items = [
                    (0, "Accueillir les nouveaux collaborateurs"),
                    (1, "Préparer les accès et le matériel"),
                    (1, "Organiser le premier jour"),
                    (0, "Assurer le suivi des demandes"),
                    (1, "Collecter les difficultés"),
                    (1, "Transmettre aux services concernés"),
                ]
                for lvl, txt in items:
                    lp = doc.add_paragraph(txt)
                    set_numpr(lp, lid, lvl)
    add_header_footer(doc, "Document structuré", first_page_different=True)
    doc.save(M2_RE / "M2_03_Realise_Document_structure.docx")


def m2_ex04():
    doc = new_doc()
    add_text(doc, "Intégrer des illustrations", style="Title")
    consigne(
        doc,
        "insérez chc_logo.png et réglez l'habillage Carré, créez un SmartArt de processus "
        "(Demande → Validation → Mise en œuvre), collez le tableau de M2_04_Suivi.xlsx avec "
        "liaison, insérez un symbole et une équation, puis positionnez les objets pour que "
        "le texte reste lisible.",
    )
    add_text(doc, "Processus à illustrer en SmartArt", style="Heading 1")
    add_text(doc, "Demande -> Validation -> Mise en œuvre")
    add_text(doc, "Zone pour le tableau / graphique Excel", style="Heading 1")
    add_text(doc, "Collez ici le tableau de suivi provenant d'Excel (avec liaison).")
    add_text(doc, "Zone pour l'équation", style="Heading 1")
    add_text(doc, "Insérez une équation, par exemple une moyenne ou un pourcentage.")
    doc.save(M2_EX / "M2_04_Exercice_Illustrations.docx")


def m2_re04(logo: Path, xlsx: Path):
    doc = new_doc()
    add_text(doc, "Illustrations - résultat", style="Title")
    add_text(
        doc,
        "Le logo ci-contre utilise un habillage Carré : le texte continue de part et d'autre. "
        "L'ancre reste liée au paragraphe pour un positionnement stable.",
    )
    add_picture_square(doc, logo, 3.6)
    add_text(doc, "Processus à illustrer en SmartArt", style="Heading 1")
    add_text(doc, "Demande → Validation → Mise en œuvre — diagramme de processus :", space_after=6)
    process = doc.add_table(rows=1, cols=3)
    labels = ("1. Demande", "2. Validation", "3. Mise en œuvre")
    for i, lab in enumerate(labels):
        cell = process.cell(0, i)
        cell.text = lab
        shade = OxmlElement("w:shd")
        shade.set(qn("w:fill"), CHC_FILL)
        shade.set(qn("w:val"), "clear")
        cell._tc.get_or_add_tcPr().append(shade)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = CHC_BLUE
    add_text(doc, "Processus en trois étapes : Demande, Validation, Mise en œuvre.", italic=True, space_before=6)
    add_text(doc, "Zone pour le tableau / graphique Excel", style="Heading 1")
    add_text(doc, f"Tableau collé avec liaison depuis {xlsx.name} :")
    rows = [
        ("Action", "Responsable", "Heures"),
        ("Préparer les accès", "Service IT", "2"),
        ("Préparer le matériel", "Logistique", "3"),
        ("Accueil du collaborateur", "Référent", "4"),
    ]
    table = doc.add_table(rows=len(rows), cols=3)
    fill_table(table, rows)
    style_table(table)
    p = doc.add_paragraph()
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' LINK Excel.Sheet.12 "{xlsx.name}" "Suivi!A1:C4" \\a \\f 4 \\h '
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = "Liaison Excel : mettez le champ à jour si le classeur change."
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, sep, t, end])
    add_text(doc, "Zone pour l'équation", style="Heading 1")
    add_text(doc, "Symbole inséré : →  ✓  •   (Insertion → Symbole).")
    add_equation(doc)
    doc.save(M2_RE / "M2_04_Realise_Illustrations.docx")


def m2_ex05():
    doc = new_doc()
    add_text(doc, "Maîtriser les tableaux", style="Title")
    consigne(
        doc,
        "convertissez le texte (séparateur ;) en tableau, fusionnez les cellules d'une ligne "
        "d'en-tête, fractionnez une cellule, ajoutez une ligne de total avec =SUM(ABOVE), "
        "reconvertissez une partie du tableau en texte, puis présentez le dernier paragraphe "
        "sur deux colonnes façon journal.",
    )
    add_text(doc, "Texte à convertir en tableau (séparateur : point-virgule)", style="Heading 1")
    for line in (
        "Action;Responsable;Heures",
        "Préparer les accès;Service IT;2",
        "Préparer le matériel;Logistique;3",
        "Accueil du collaborateur;Référent;4",
    ):
        add_text(doc, line, space_after=2)
    add_text(doc, "Totaux à calculer avec une formule", style="Heading 1")
    add_text(doc, "Après conversion, ajoutez une ligne de total et utilisez =SUM(ABOVE) sur la colonne Heures.")
    add_text(doc, "Texte à présenter en colonnes", style="Heading 1")
    add_text(
        doc,
        "Ce paragraphe doit être présenté sur deux colonnes façon journal afin d'illustrer la mise "
        "en page multicolonne. Ajoutez suffisamment de texte pour remplir les deux colonnes et "
        "observer le passage automatique d'une colonne à l'autre. Le service Formation utilise ce "
        "type de présentation pour les notes internes courtes, les extraits de procédure et les "
        "encadrés d'information destinés aux nouveaux collaborateurs. Gardez un texte lisible, sans "
        "justifier excessivement, et vérifiez l'aperçu avant impression.",
    )
    doc.save(M2_EX / "M2_05_Exercice_Tableaux_avances.docx")


def m2_re05():
    doc = new_doc()
    add_text(doc, "Tableaux avancés - résultat", style="Title")
    add_text(doc, "Texte converti en tableau", style="Heading 1")
    table = doc.add_table(rows=6, cols=4)
    # ligne d'en-tête fusionnée
    table.cell(0, 0).merge(table.cell(0, 3))
    table.cell(0, 0).text = "Plan d'actions — accueil d'un collaborateur"
    headers = ("Action", "Responsable", "Heures", "Commentaire")
    for i, h in enumerate(headers):
        table.cell(1, i).text = h
    data = [
        ("Préparer les accès", "Service IT", "2", "Badge"),
        ("Préparer le matériel", "Logistique", "3", "Poste"),
        ("Accueil du collaborateur", "Référent", "4", "Jour J"),
    ]
    for r, row in enumerate(data, start=2):
        for c, val in enumerate(row):
            table.cell(r, c).text = val
    table.cell(5, 0).merge(table.cell(5, 1))
    table.cell(5, 0).text = "Total"
    add_formula_cell(table.cell(5, 2), "=SUM(ABOVE)", "9")
    table.cell(5, 3).text = ""
    style_table(table)
    add_text(
        doc,
        "La première ligne d'en-tête est fusionnée. La colonne Commentaire correspond à une cellule "
        "fractionnée (Heures / Commentaire).",
        italic=True,
        space_before=8,
    )
    add_text(doc, "Partie reconvertie en texte", style="Heading 1")
    add_text(doc, "Action;Responsable;Heures")
    add_text(doc, "Préparer les accès;Service IT;2")
    add_text(doc, "Préparer le matériel;Logistique;3")
    add_text(doc, "Accueil du collaborateur;Référent;4")

    # section 1 col puis 2 col
    doc.add_section(WD_SECTION.CONTINUOUS)
    set_section_columns(doc.sections[0], 1)
    set_section_columns(doc.sections[1], 2)
    add_text(doc, "Texte présenté en colonnes", style="Heading 1")
    add_text(
        doc,
        "Ce paragraphe est présenté sur deux colonnes façon journal afin d'illustrer la mise en page "
        "multicolonne. Le texte passe automatiquement d'une colonne à l'autre. Le service Formation "
        "utilise ce type de présentation pour les notes internes courtes, les extraits de procédure "
        "et les encadrés d'information destinés aux nouveaux collaborateurs. Gardez un texte lisible "
        "et vérifiez l'aperçu avant impression.",
    )
    doc.save(M2_RE / "M2_05_Realise_Tableaux_avances.docx")


def m2_ex06():
    doc = new_doc()
    add_text(doc, "Atelier final - Module 2 Intermédiaire", style="Title")
    add_text(doc, "Objectif : produire un document professionnel complet en mobilisant les acquis du module.", style="Heading 1")
    for item in (
        "Partez d'un modèle personnalisé (police, marges, styles).",
        "Structurez avec des titres numérotés automatiquement et un sommaire.",
        "Ajoutez des en-têtes et pieds de page avec numérotation.",
        "Intégrez une illustration (image, SmartArt ou tableau Excel lié).",
        "Insérez un tableau avec au moins une formule de calcul.",
        "Présentez une partie en colonnes façon journal.",
        "Enregistrez un QuickPart réutilisable et exportez le document en PDF.",
    ):
        doc.add_paragraph(item, style="List Bullet")
    doc.save(M2_EX / "M2_06_Exercice_Atelier_final.docx")


def m2_re06(logo: Path):
    doc = new_doc()
    _ensure_custom_styles(doc)
    hid = add_heading_numbering(doc, 12)
    add_text(doc, "Note de service — Organisation des formations Word", style="Title", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "Modèle CHC — Service Formation", style="Subtitle", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "23 septembre 2026", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_page_break(doc)
    add_text(doc, "Sommaire", style="Heading 1")
    add_toc(doc)
    add_page_break(doc)

    p = add_text(doc, "Cadre et objectifs", style="Heading 1")
    set_numpr(p, hid, 0)
    p2 = add_text(doc, "Public concerné", style="Heading 2")
    set_numpr(p2, hid, 1)
    add_text(doc, "Cette note s'adresse aux référents qui accueillent un nouveau collaborateur et aux formateurs Word.", style="CorpsCHC")
    p2 = add_text(doc, "Résultat attendu", style="Heading 2")
    set_numpr(p2, hid, 1)
    add_text(doc, "Un document homogène, basé sur les styles du modèle, prêt à être exporté en PDF.", style="CorpsCHC")

    p = add_text(doc, "Déroulé type", style="Heading 1")
    set_numpr(p, hid, 0)
    add_picture_inline(doc, logo, 3.0)
    process = doc.add_table(rows=1, cols=3)
    for i, lab in enumerate(("1. Demande", "2. Validation", "3. Mise en œuvre")):
        process.cell(0, i).text = lab
        for r in process.cell(0, i).paragraphs[0].runs:
            r.bold = True
    style_table(process, header=False)

    p = add_text(doc, "Charge de travail", style="Heading 1")
    set_numpr(p, hid, 0)
    table = doc.add_table(rows=5, cols=3)
    fill_table(
        table,
        [
            ("Étape", "Responsable", "Heures"),
            ("Préparer les accès", "Service IT", "2"),
            ("Préparer le matériel", "Logistique", "3"),
            ("Accueil du collaborateur", "Référent", "4"),
            ("Total", "", ""),
        ],
    )
    add_formula_cell(table.cell(4, 2), "=SUM(ABOVE)", "9")
    style_table(table)

    add_header_footer(doc, "Note de service — Formations Word", first_page_different=True)

    doc.add_section(WD_SECTION.CONTINUOUS)
    set_section_columns(doc.sections[-2], 1)
    set_section_columns(doc.sections[-1], 2)
    p = add_text(doc, "Information aux participants", style="Heading 1")
    set_numpr(p, hid, 0)
    add_text(
        doc,
        "Les participants reçoivent cette note avant la session. Le texte est présenté sur deux "
        "colonnes façon journal. Un QuickPart de signature est enregistré pour les courriers "
        "suivants. Le sommaire a été mis à jour et le document exporté en PDF "
        "(M2_06_Realise_Atelier_final.pdf).",
        style="CorpsCHC",
    )
    add_text(doc, "CHC - Service Formation", space_after=0)
    add_text(doc, "Rue de Hesbaye 75 - 4000 Liège", space_after=0)
    add_text(doc, "formation@chc.be", space_after=0)
    path = M2_RE / "M2_06_Realise_Atelier_final.docx"
    doc.save(path)
    make_pdf(
        M2_RE / "M2_06_Realise_Atelier_final.pdf",
        "Note de service - Formations Word",
        [
            "Export PDF de l'atelier final Module 2.",
            "Styles du modele, titres numerotes, sommaire, tableau calcule, colonnes journal.",
            "",
            "CHC - Service Formation",
            "Rue de Hesbaye 75 - 4000 Liege",
        ],
    )


def main():
    for d in (M1_EX, M1_RE, M2_EX, M2_RE, ASSETS):
        d.mkdir(parents=True, exist_ok=True)
    logo = make_logo()
    xlsx = M2_RE / "M2_04_Suivi.xlsx"
    make_xlsx(xlsx)
    shutil.copy2(logo, M1_EX / "chc_logo.png")
    shutil.copy2(logo, M2_EX / "chc_logo.png")
    shutil.copy2(xlsx, M2_EX / "M2_04_Suivi.xlsx")

    m1_ex01(); m1_re01()
    m1_ex02(); m1_re02(); _rewrite_m1_re02()
    m1_ex03(); m1_re03()
    m1_ex04(); m1_re04(logo)
    m1_ex05(); m1_re05()
    m1_ex06(); m1_re06()
    m1_ex07(); m1_re07(logo)

    m2_ex01(); m2_re01()
    m2_ex02(); m2_re02()
    m2_ex03(); m2_re03()
    m2_ex04(); m2_re04(logo, xlsx)
    m2_ex05(); m2_re05()
    m2_ex06(); m2_re06(logo)
    print("OK — documents générés")


if __name__ == "__main__":
    main()
