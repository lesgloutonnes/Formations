#!/usr/bin/env python3
"""Génère les versions HTML imprimables et PDF des modules Excel 1, 2 et 3."""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

EXCEL_ROOT = Path(__file__).resolve().parent


@dataclass(frozen=True)
class ModuleSpec:
    number: int
    title: str
    kicker: str
    audience: str
    contenu: str
    source: Path
    out_dir: Path
    html_name: str
    pdf_name: str
    footer: str
    sections: tuple[tuple[str, str, str], ...]  # id, label, comment_key
    default_section: str
    tip: str


MODULES = [
    ModuleSpec(
        number=1,
        title="Module 1 : Les Bases",
        kicker="Formation Excel · Support de cours",
        audience="Débutants — prise en main d’Excel",
        contenu="6 chapitres + récapitulatif, avec exercices corrigés",
        source=EXCEL_ROOT / "Module 1" / "WEB_INTERACTIVE" / "Module1_Formation_Interactive.html",
        out_dir=EXCEL_ROOT / "Module 1" / "PDF",
        html_name="Module1_Formation_Imprimable.html",
        pdf_name="Module1_Formation_Imprimable.pdf",
        footer="Excel Module 1 — Les Bases",
        default_section="section-1",
        tip="Astuce : laissez un espace en marge pour vos notes. Les raccourcis clavier (Ctrl+S, Ctrl+Home, F2…) sont indiqués dans chaque chapitre.",
        sections=(
            ("section-1", "1. Introduction à Excel", "SECTION 1"),
            ("section-2", "2. Saisie de données et formats", "SECTION 2"),
            ("section-3", "3. Formules de base et Autofill", "SECTION 3"),
            ("section-4", "4. Gestion des feuilles", "SECTION 4"),
            ("section-5", "5. Création de graphiques", "SECTION 5"),
            ("section-6", "6. Mise en page et impression", "SECTION 6"),
            ("section-recap", "Récapitulatif", "SLIDE FINAL"),
        ),
    ),
    ModuleSpec(
        number=2,
        title="Module 2 : Intermédiaire",
        kicker="Formation Excel · Support de cours",
        audience="Utilisateurs ayant suivi le Module 1",
        contenu="5 chapitres + récapitulatif, avec exercices corrigés",
        source=EXCEL_ROOT / "Module 2" / "WEB_INTERACTIVE" / "Module2_Formation_Interactive.html",
        out_dir=EXCEL_ROOT / "Module 2" / "PDF",
        html_name="Module2_Formation_Imprimable.html",
        pdf_name="Module2_Formation_Imprimable.pdf",
        footer="Excel Module 2 — Intermédiaire",
        default_section="section-0",
        tip="Astuce : entraînez-vous sur un vrai fichier (ventes, planning, listing) plutôt que sur des cellules vides.",
        sections=(
            ("section-0", "1. Références relatives et absolues", "SECTION 0"),
            ("section-1", "2. Formules avancées", "SECTION 1"),
            ("section-2", "3. Gestion des données", "SECTION 2"),
            ("section-3", "4. Tableaux croisés dynamiques", "SECTION 3"),
            ("section-4", "5. Mise en forme conditionnelle", "SECTION 4"),
            ("section-recap", "Récapitulatif", "SLIDE FINAL"),
        ),
    ),
    ModuleSpec(
        number=3,
        title="Module 3 : Macros & VBA avec l’IA",
        kicker="Formation Excel · Support de cours",
        audience="Utilisateurs des Modules 1 et 2 — Excel + Edge + IA autorisée",
        contenu="4 parties + bonus, exercices sur fichiers .xlsx du dossier EXERCICES",
        source=EXCEL_ROOT / "Module 3" / "WEB_INTERACTIVE" / "Module3_Formation_Interactive.html",
        out_dir=EXCEL_ROOT / "Module 3" / "PDF",
        html_name="Module3_Formation_Imprimable.html",
        pdf_name="Module3_Formation_Imprimable.pdf",
        footer="Excel Module 3 — Macros & VBA avec l’IA",
        default_section="section-accueil",
        tip="Astuce : ouvrez une nouvelle conversation IA à chaque nouvelle macro. Les fichiers d’exercices sont dans le dossier EXERCICES à côté de la formation.",
        sections=(
            ("section-accueil", "Accueil", "0 ACCUEIL"),
            ("section-1", "1. Mise en place", "PARTIE 1"),
            ("section-2", "2. Enregistrer & améliorer", "PARTIE 2"),
            ("section-3", "3. Créer avec l’IA", "PARTIE 3"),
            ("section-4", "4. Tableau de bord", "PARTIE 4"),
            ("section-bonus", "Bonus — Idées CHC", "BONUS"),
            ("section-recap", "Récapitulatif", "RECAP"),
        ),
    ),
]


def extract_balanced_div(html: str, start: int) -> tuple[str, int]:
    gt = html.find(">", start)
    if gt == -1:
        raise ValueError("Balise <div> non fermée")
    depth = 1
    pos = gt + 1
    while pos < len(html) and depth > 0:
        nxt_open = html.find("<div", pos)
        nxt_close = html.find("</div>", pos)
        if nxt_close == -1:
            raise ValueError("div non équilibrée")
        if nxt_open != -1 and nxt_open < nxt_close:
            end_tag = html.find(">", nxt_open)
            tag = html[nxt_open : end_tag + 1]
            if not tag.endswith("/>"):
                depth += 1
            pos = end_tag + 1
        else:
            depth -= 1
            pos = nxt_close + 6
    return html[start:pos], pos


def section_from_comment(comment: str, spec: ModuleSpec) -> str | None:
    text = re.sub(r"<!--|=+", " ", comment)
    text = re.sub(r"\s+", " ", text).strip().upper()
    for sid, _label, key in spec.sections:
        if key.upper() in text:
            return sid
    if "SLIDE FINAL" in text or "RÉCAPITULATIF" in text or "RECAPITULATIF" in text:
        return "section-recap"
    return None


def extract_slides(html: str, spec: ModuleSpec) -> list[dict]:
    container_start = html.find('<div class="presentation-container">')
    if container_start == -1:
        raise RuntimeError("Conteneur de présentation introuvable")
    container, _ = extract_balanced_div(html, container_start)

    current_section = spec.default_section
    slides: list[dict] = []
    pos = 0
    comment_re = re.compile(
        r"<!--\s*(?:=+\s*)?(SECTION\s+\d+|PARTIE\s+\d+|Slide Final|RECAP|BONUS|0 ACCUEIL).*?-->",
        re.IGNORECASE | re.DOTALL,
    )
    while True:
        comment = comment_re.search(container[pos:])
        slide_match = re.search(r'<div class="slide\b', container[pos:])
        if not slide_match:
            break

        comment_pos = (pos + comment.start()) if comment else None
        slide_pos = pos + slide_match.start()

        if comment and comment_pos is not None and comment_pos < slide_pos:
            mapped = section_from_comment(comment.group(0), spec)
            if mapped:
                current_section = mapped
            pos = comment_pos + len(comment.group(0))
            continue

        block, end = extract_balanced_div(container, slide_pos)
        classes = re.search(r'class="([^"]+)"', block)
        class_str = classes.group(1) if classes else ""
        if "exercise" in class_str:
            kind = "exercise"
        elif "solution" in class_str:
            kind = "solution"
        else:
            kind = "theory"

        title_m = re.search(r'<div class="slide-title">(.*?)</div>', block, re.DOTALL)
        badge_m = re.search(r'<div class="slide-type-badge[^"]*">(.*?)</div>', block, re.DOTALL)
        subtitle_m = re.search(r'<div class="subtitle">(.*?)</div>', block, re.DOTALL)
        content_m = re.search(
            r'<div class="slide-content[^"]*">(.*)</div>\s*</div>\s*$',
            block,
            re.DOTALL,
        )

        inner = content_m.group(1) if content_m else block
        inner = sanitize_slide_html(inner)

        title = title_m.group(1).strip() if title_m else ""
        badge = badge_m.group(1).strip() if badge_m else ""
        subtitle = subtitle_m.group(1).strip() if subtitle_m else ""
        if subtitle:
            inner = re.sub(
                r'<div class="subtitle">\s*' + re.escape(subtitle) + r"\s*</div>",
                "",
                inner,
                count=1,
            )
        heading = first_plain_heading(inner)
        if heading and title in {"À vous de jouer !", "Correction"}:
            title = heading

        slides.append(
            {
                "section": current_section,
                "kind": kind,
                "title": title,
                "badge": badge,
                "subtitle": subtitle,
                "html": inner.strip(),
            }
        )
        pos = end
    return slides


def first_plain_heading(html: str) -> str:
    match = re.search(r"<h3>(.*?)</h3>", html, re.DOTALL)
    if not match:
        return ""
    text = re.sub(r"<[^>]+>", "", match.group(1))
    return re.sub(r"\s+", " ", text).strip()


def sanitize_slide_html(html: str) -> str:
    html = re.sub(r"<button\b[^>]*>.*?</button>", "", html, flags=re.DOTALL)
    html = re.sub(r'<div class="formateur-note">.*?</div>', "", html, flags=re.DOTALL)
    html = re.sub(r'\sstyle="display:\s*none;?"', "", html)

    def _clean_style(match: re.Match[str]) -> str:
        style = match.group(1)
        keep: list[str] = []
        for decl in style.split(";"):
            decl = decl.strip()
            if not decl or ":" not in decl:
                continue
            prop, value = decl.split(":", 1)
            prop = prop.strip().lower()
            value = value.strip().lower()
            if prop in {"color", "font-size", "max-width", "background", "background-color"}:
                continue
            if prop in {"text-align", "margin-top", "margin-left", "margin-right", "list-style"}:
                keep.append(f"{prop}: {value}")
        if not keep:
            return ""
        return f' style="{"; ".join(keep)}"'

    return re.sub(r'\sstyle="([^"]*)"', _clean_style, html)


PRINT_CSS = r"""
:root {
    --excel: #217346;
    --excel-dark: #185C37;
    --ink: #1c1c1c;
    --muted: #4a4a4a;
    --line: #d9e6dd;
    --paper: #ffffff;
    --soft: #f4f8f5;
    --theory: #217346;
    --exercise: #b8860b;
    --solution: #0d8a52;
}

* { box-sizing: border-box; }

html, body {
    margin: 0;
    padding: 0;
    background: var(--paper);
    color: var(--ink);
    font-family: "Segoe UI", Calibri, "Liberation Sans", Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.45;
    orphans: 3;
    widows: 3;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
}

@page {
    size: A4;
    margin: 18mm 14mm 20mm 14mm;
    @bottom-left {
        content: var(--doc-footer);
        font-size: 8pt;
        color: #66776c;
    }
    @bottom-right {
        content: counter(page);
        font-size: 8pt;
        color: #66776c;
    }
}
@page :first {
    @bottom-left { content: none; }
    @bottom-right { content: none; }
}

.cover {
    padding-top: 28mm;
    page-break-after: always;
    break-after: page;
}

.cover-banner {
    background: var(--excel);
    color: #fff;
    padding: 28px 32px;
    border-radius: 8px 8px 0 0;
}

.cover-banner .kicker {
    letter-spacing: 0.18em;
    text-transform: uppercase;
    font-size: 11pt;
    font-weight: 700;
    opacity: 0.92;
    margin: 0 0 8px;
}

.cover-banner h1 {
    margin: 0;
    font-size: 28pt;
    line-height: 1.15;
}

.cover-body {
    border: 2px solid var(--excel);
    border-top: none;
    border-radius: 0 0 8px 8px;
    padding: 28px 32px 32px;
    background: var(--soft);
}

.cover-body h2 {
    margin: 0 0 12px;
    color: var(--excel-dark);
    font-size: 16pt;
}

.cover-meta {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    margin: 18px 0 8px;
}

.meta-card {
    background: #fff;
    border: 1px solid var(--line);
    border-left: 5px solid var(--excel);
    padding: 12px 14px;
}

.meta-card strong { display: block; color: var(--excel-dark); margin-bottom: 4px; }

.toc {
    page-break-after: always;
    break-after: page;
}

.toc h1, .section-title {
    color: var(--excel-dark);
    border-bottom: 3px solid var(--excel);
    padding-bottom: 8px;
    margin: 0 0 18px;
    font-size: 20pt;
}

.toc ol {
    list-style: none;
    padding-left: 0;
    font-size: 13pt;
}

.toc li {
    margin: 12px 0;
    padding: 8px 12px;
    border-bottom: 1px dotted var(--line);
}

.toc li a { color: var(--ink); text-decoration: none; font-weight: 600; }
.toc li a .toc-num { color: var(--excel); margin-right: 6px; }

.section {
    page-break-before: always;
    break-before: page;
}

.section.first-section {
    page-break-before: auto;
    break-before: auto;
}

.slide {
    border: 1px solid var(--line);
    border-radius: 8px;
    margin: 0 0 14px;
    background: #fff;
    overflow: visible;
}

.slide.theory { border-left: 6px solid var(--theory); }
.slide.exercise { border-left: 6px solid var(--exercise); }
.slide.solution { border-left: 6px solid var(--solution); }

.slide-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 10px 14px;
    border-bottom: 1px solid var(--line);
    background: var(--soft);
}

.slide.exercise .slide-header { background: #fff8e1; }
.slide.solution .slide-header { background: #e8f8ef; }

.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 8.5pt;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    white-space: nowrap;
}

.theory .badge { background: var(--theory); color: #fff; }
.exercise .badge { background: #e6b800; color: #1a1a1a; }
.solution .badge { background: var(--solution); color: #fff; }

.slide-title {
    font-size: 14pt;
    font-weight: 700;
    margin: 0;
    color: var(--ink);
}

.slide-body { padding: 12px 16px 16px; }

.subtitle {
    font-size: 12.5pt;
    color: var(--excel-dark);
    font-weight: 700;
    margin: 0 0 10px;
}

.content-box, .exercise-instructions, .solution-box, .why-box,
.visual-description, .golden-rule, .compare-box, .autofill-demo,
.formula-example, .error-item, .card-item, .a-retenir-box, .file-box,
.objectifs-box, .prompt-box, .aller-plus-loin, .function-card,
.ref-type-card, .dollar-syntax-item, .field-zone, .comparison-item {
    border-radius: 6px;
    padding: 10px 12px;
    margin: 8px 0;
}

.why-box, .visual-description, .golden-rule, .compare-box,
.autofill-demo, .formula-example, .error-item, .card-item,
.function-card, .ref-type-card, .dollar-syntax-item, .field-zone,
.workflow-step, .roadmap-item {
    page-break-inside: avoid;
    break-inside: avoid;
}

.slide-header, .subtitle, h3, h4 {
    break-after: avoid;
    page-break-after: avoid;
}

.content-box, .objectifs-box {
    background: #f3f9f5;
    border-left: 4px solid var(--excel);
}

.exercise-instructions {
    background: #fff8e1;
    border-left: 4px solid #e6b800;
}

.solution-box {
    background: #eefaf3;
    border-left: 4px solid var(--solution);
}

.why-box {
    background: #f6eef8;
    border-left: 4px solid #8e24aa;
}

.visual-description {
    background: #f4f4f4;
    border-left: 4px solid #888;
    font-style: italic;
    color: #333;
}

.visual-description::before {
    content: "À l'écran Excel, vous verrez : ";
    display: block;
    font-weight: 700;
    font-style: normal;
    margin-bottom: 4px;
    color: var(--ink);
}

h3 { margin: 0 0 8px; font-size: 12pt; color: var(--excel-dark); }
.why-box h4 { margin: 0 0 6px; color: #6a1b9a; font-size: 11.5pt; }

ul, ol { margin: 6px 0 8px; padding-left: 22px; }
li { margin: 3px 0; }
p { margin: 6px 0; }

.content-box ul, .objectifs-box ul { list-style: none; padding-left: 0; }
.content-box li, .objectifs-box li { padding-left: 16px; position: relative; }
.content-box li::before, .objectifs-box li::before {
    content: "▸";
    position: absolute;
    left: 0;
    color: var(--excel);
    font-weight: 700;
}
.content-box li ul, .objectifs-box li ul { list-style: disc; padding-left: 22px; margin-top: 4px; }
.content-box li ul li, .objectifs-box li ul li { padding-left: 0; }
.content-box li ul li::before, .objectifs-box li ul li::before { content: none; }

code {
    font-family: "Cascadia Mono", "Consolas", "Courier New", monospace;
    background: #e8f5ec;
    color: #0b5c32;
    padding: 1px 5px;
    border-radius: 3px;
    font-size: 0.95em;
}

pre {
    font-family: "Cascadia Mono", "Consolas", "Courier New", monospace;
    background: #f3f7f4;
    border: 1px solid var(--line);
    color: #0b5c32;
    padding: 10px 12px;
    border-radius: 6px;
    font-size: 9.5pt;
    white-space: pre-wrap;
    margin: 8px 0;
}

.highlight {
    background: #d7eedf;
    padding: 1px 5px;
    border-radius: 3px;
    font-weight: 600;
    color: var(--excel-dark);
}

.duration-badge, .exercise-meta span {
    display: inline-block;
    border: 1px solid #e6b800;
    background: #fff3c4;
    color: #6b5200;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 10pt;
    font-weight: 700;
    margin: 2px 4px 8px 0;
}

.exercise-meta { margin-bottom: 8px; }

.cards-grid, .autofill-grid, .error-grid, .formula-examples, .compare-grid,
.functions-grid, .ref-type-grid, .dollar-syntax-grid, .field-zones,
.comparison-grid, .criteria-showcase, .tcd-two-cols, .roadmap-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    margin: 8px 0;
}

.cards-grid, .functions-grid, .field-zones, .roadmap-grid { grid-template-columns: 1fr 1fr; }
.cards-grid-2 { grid-template-columns: 1fr 1fr; }

.card-item, .autofill-demo, .formula-example, .compare-box,
.function-card, .ref-type-card, .dollar-syntax-item, .field-zone,
.comparison-item, .criteria-box {
    background: #f7fbf8;
    border: 1px solid var(--line);
}

.function-card, .field-zone, .ref-type-card { text-align: center; }
.function-name, .field-zone-title, .card-name { font-weight: 700; color: var(--excel-dark); }
.function-desc, .field-zone-desc, .card-desc, .formula-desc { font-size: 9.5pt; color: var(--muted); }
.function-usage, .formula-code {
    font-family: "Consolas", "Courier New", monospace;
    color: var(--excel-dark);
    font-weight: 700;
    margin-top: 6px;
}

.cell-row { display: flex; gap: 6px; justify-content: center; flex-wrap: wrap; margin: 8px 0; }
.demo-cell {
    min-width: 52px; padding: 6px 8px; border: 1px solid #bbb; border-radius: 4px;
    font-family: "Consolas", "Courier New", monospace; font-size: 9pt; background: #fff;
}
.demo-cell.active { background: #e3f3e9; border-color: var(--excel); color: var(--excel-dark); font-weight: 700; }

.chc-table, .tcd-table, .cell-grid, .mult-table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 10pt;
}
.chc-table th, .tcd-table th, .cell-grid .col-header, .cell-grid .row-header,
.mult-table .col-h, .mult-table .row-h {
    background: var(--excel);
    color: #fff;
    padding: 6px 8px;
    text-align: center;
}
.chc-table td, .tcd-table td, .cell-grid .cell, .mult-table td, .mult-table th {
    border: 1px solid #c5d8cc;
    padding: 5px 6px;
    text-align: center;
    color: var(--ink);
    background: #fff;
}
.tcd-table .tcd-row-total, .tcd-table .tcd-grand-row td {
    background: #e8f8ef;
    color: var(--excel-dark);
    font-weight: 700;
}
.cell-grid .cell-yellow { background: #fff3c4; }
.cell-grid .cell-green { background: #c8eed8; }
.cell-grid .cell-blue { background: #d6eef9; }
.cell-grid .cell-red { background: #ffd6d6; }
.mult-table .formula-cell {
    background: #e8f8ef;
    color: var(--excel-dark);
    font-family: "Consolas", "Courier New", monospace;
    font-weight: 700;
    font-size: 8.5pt;
}
.cell-grid-wrapper { display: flex; gap: 16px; flex-wrap: wrap; align-items: flex-start; }
.cell-answers {
    background: #fff8e1;
    border: 1px solid #e6b800;
    border-radius: 8px;
    padding: 12px;
}
.cell-answer-row { display: flex; align-items: center; gap: 8px; margin: 6px 0; }

.pivot-structure {
    display: grid;
    grid-template-areas: ". . cols" "rows data data" "rows totals grand";
    grid-template-columns: 70px 1fr 1fr;
    grid-template-rows: 36px 1fr 28px;
    gap: 3px;
    max-width: 420px;
    margin: 12px auto;
    min-height: 140px;
}
.pivot-cols { grid-area: cols; background: #ffe082; text-align: center; font-weight: 700; padding: 6px; font-size: 9pt; }
.pivot-rows { grid-area: rows; background: #81c784; color: #113318; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 9pt; writing-mode: vertical-rl; }
.pivot-data { grid-area: data; background: #eceff1; text-align: center; padding: 10px; font-weight: 600; }
.pivot-totals { grid-area: totals; background: #b0bec5; text-align: center; font-weight: 700; font-size: 9pt; }
.pivot-grand { grid-area: grand; background: #ef9a9a; text-align: center; font-weight: 700; font-size: 9pt; }

.tcd-showcase, .mult-table-showcase, .criteria-showcase {
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 12px;
    background: var(--soft);
    margin: 10px 0;
}
.tcd-showcase-title { text-align: center; color: var(--excel-dark); font-weight: 700; margin-bottom: 8px; }
.tcd-source-box {
    background: #fff; border: 1px solid var(--line); border-radius: 6px; padding: 10px;
    font-family: "Consolas", "Courier New", monospace; font-size: 9pt; color: var(--ink);
}
.tcd-config-note { text-align: center; color: #8a6d00; font-weight: 600; margin-top: 8px; }

.a-retenir-box, .aller-plus-loin {
    background: #eefaf3;
    border: 1px solid var(--solution);
}
.a-retenir-box h4, .aller-plus-loin h4 { color: var(--excel-dark); margin: 0 0 6px; }
.file-box, .prompt-box {
    background: #eef6fc;
    border-left: 4px solid #1565c0;
}
.file-box h4, .prompt-box h4 { color: #0d47a1; margin: 0 0 6px; }

.partie-label {
    display: inline-block;
    background: var(--soft);
    border: 1px solid var(--excel);
    color: var(--excel-dark);
    padding: 3px 12px;
    border-radius: 999px;
    font-weight: 700;
    font-size: 9pt;
    margin-bottom: 8px;
}

.workflow-row { display: flex; gap: 8px; flex-wrap: wrap; margin: 10px 0; }
.workflow-step, .roadmap-item {
    background: #f7fbf8;
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 10px;
    text-align: center;
    flex: 1;
    min-width: 110px;
}
.step-name, .roadmap-name { font-weight: 700; color: var(--excel-dark); }
.step-desc { font-size: 9pt; color: var(--muted); margin-top: 4px; }
.roadmap-num {
    display: inline-block;
    background: var(--excel);
    color: #fff;
    width: 24px; height: 24px; line-height: 24px;
    border-radius: 50%;
    font-weight: 700;
    margin-bottom: 4px;
}

.golden-rule { text-align: center; border: 2px solid var(--excel); background: #eef7f1; }
.golden-rule .equals { font-size: 28pt; font-weight: 700; color: var(--excel); margin: 4px 0; }
.compare-box.good { border: 1.5px solid var(--solution); background: #eefaf3; }
.compare-box.bad { border: 1.5px solid #e65100; background: #fff4e8; }
.compare-box.good h4 { color: var(--solution); }
.compare-box.bad h4 { color: #e65100; }

.center-content { text-align: left; }
.icon-large { display: none; }
.slide-body p { color: var(--ink); }

.footer-note {
    margin-top: 24px;
    font-size: 9.5pt;
    color: var(--muted);
    border-top: 1px solid var(--line);
    padding-top: 8px;
}

@media print {
    .slide { box-shadow: none; }
    a { color: inherit; text-decoration: none; }
}
"""


def render_print_html(slides: list[dict], spec: ModuleSpec) -> str:
    first_id = spec.sections[0][0]
    css = PRINT_CSS.replace("var(--doc-footer)", f'"{spec.footer}"')
    parts: list[str] = [
        f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>{spec.footer} (version imprimable)</title>
<style>{css}</style>
</head>
<body>
<article class="cover">
    <div class="cover-banner">
        <p class="kicker">{spec.kicker}</p>
        <h1>{spec.title}</h1>
    </div>
    <div class="cover-body">
        <h2>Version imprimable</h2>
        <p>Ce document reprend l’intégralité du cours interactif Excel Module&nbsp;{spec.number} : théorie, exercices et corrections. Il est prévu pour une impression A4 (recto ou recto-verso).</p>
        <div class="cover-meta">
            <div class="meta-card"><strong>Public</strong>{spec.audience}</div>
            <div class="meta-card"><strong>Contenu</strong>{spec.contenu}</div>
            <div class="meta-card"><strong>Impression conseillée</strong>A4 portrait, couleurs ou niveaux de gris</div>
            <div class="meta-card"><strong>Usage</strong>Cahier de formation, révision, notes en marge</div>
        </div>
        <p>Les blocs jaunes sont des <strong>exercices</strong>. Les blocs verts sont les <strong>corrections</strong>. Les formules et le code apparaissent en <code>police à chasse fixe</code>.</p>
    </div>
</article>

<nav class="toc">
    <h1>Sommaire</h1>
    <ol>
"""
    ]
    for i, (sid, label, _key) in enumerate(spec.sections, start=1):
        parts.append(
            f'        <li><a href="#{sid}"><span class="toc-num">{i}.</span>{label.split(". ", 1)[-1] if label[0].isdigit() else label}</a></li>\n'
        )
    parts.append(f"""    </ol>
    <p class="footer-note">{spec.tip}</p>
</nav>
""")

    current = None
    labels = {sid: label for sid, label, _key in spec.sections}
    for slide in slides:
        if slide["section"] != current:
            if current is not None:
                parts.append("</section>\n")
            current = slide["section"]
            first_class = " first-section" if current == first_id else ""
            parts.append(f'<section class="section{first_class}" id="{current}">\n')
            parts.append(f'    <h1 class="section-title">{labels.get(current, current)}</h1>\n')

        kind = slide["kind"]
        title = slide["title"] or "Sans titre"
        badge = slide["badge"] or (
            "Théorie" if kind == "theory" else "Exercice" if kind == "exercise" else "Solution"
        )
        subtitle = f'<div class="subtitle">{slide["subtitle"]}</div>' if slide["subtitle"] else ""
        parts.append(f"""    <article class="slide {kind}">
        <header class="slide-header">
            <span class="badge">{badge}</span>
            <h2 class="slide-title">{title}</h2>
        </header>
        <div class="slide-body">
            {subtitle}
            {slide["html"]}
        </div>
    </article>
""")

    if current is not None:
        parts.append("</section>\n")
    parts.append("</body>\n</html>\n")
    return "".join(parts)


def find_chrome() -> str:
    for candidate in (
        "/opt/google/chrome/chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/local/bin/google-chrome",
        "/usr/bin/google-chrome",
        "google-chrome-stable",
        "google-chrome",
    ):
        path = Path(candidate) if candidate.startswith("/") else None
        if path and path.exists():
            return str(path)
        found = subprocess.run(["which", candidate], capture_output=True, text=True)
        if found.returncode == 0 and found.stdout.strip():
            return found.stdout.strip()
    raise RuntimeError("Google Chrome est requis pour générer le PDF")


def html_to_pdf(html_path: Path, pdf_path: Path, profile_name: str) -> None:
    chrome = find_chrome()
    uri = html_path.resolve().as_uri()
    profile = Path(f"/tmp/chrome-pdf-profile-{profile_name}")
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [
        chrome,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--no-first-run",
        "--no-default-browser-check",
        f"--user-data-dir={profile}",
        "--allow-file-access-from-files",
        "--no-pdf-header-footer",
        "--virtual-time-budget=8000",
        f"--print-to-pdf={pdf_path}",
        uri,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            "Échec de la génération PDF\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )


def generate_module(spec: ModuleSpec) -> None:
    if not spec.source.exists():
        raise FileNotFoundError(f"Source introuvable : {spec.source}")
    spec.out_dir.mkdir(parents=True, exist_ok=True)
    html = spec.source.read_text(encoding="utf-8")
    slides = extract_slides(html, spec)
    if not slides:
        raise RuntimeError(f"Aucune diapositive extraite pour le module {spec.number}")
    print_html = spec.out_dir / spec.html_name
    print_pdf = spec.out_dir / spec.pdf_name
    print_html.write_text(render_print_html(slides, spec), encoding="utf-8")
    html_to_pdf(print_html, print_pdf, f"excel-m{spec.number}")
    print(
        f"Module {spec.number} : {len(slides)} diapositives → "
        f"{print_html.name} + {print_pdf.name} ({print_pdf.stat().st_size / 1024:.0f} Ko)"
    )


def main(argv: list[str] | None = None) -> int:
    wanted = {int(x) for x in (argv or sys.argv[1:])} if (argv or sys.argv[1:]) else {1, 2, 3}
    for spec in MODULES:
        if spec.number in wanted:
            generate_module(spec)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
