#!/usr/bin/env python3
"""Génère la version HTML imprimable et le PDF du Module 1 Excel."""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "WEB_INTERACTIVE" / "Module1_Formation_Interactive.html"
PRINT_HTML = ROOT / "Module1_Formation_Imprimable.html"
PRINT_PDF = ROOT / "Module1_Formation_Imprimable.pdf"

SECTIONS = [
    ("section-1", "1. Introduction à Excel"),
    ("section-2", "2. Saisie de données et formats"),
    ("section-3", "3. Formules de base et Autofill"),
    ("section-4", "4. Gestion des feuilles"),
    ("section-5", "5. Création de graphiques"),
    ("section-6", "6. Mise en page et impression"),
    ("section-recap", "Récapitulatif"),
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


def extract_slides(html: str) -> list[dict]:
    container_start = html.find('<div class="presentation-container">')
    if container_start == -1:
        raise RuntimeError("Conteneur de présentation introuvable")
    container, _ = extract_balanced_div(html, container_start)

    current_section = SECTIONS[0][0]
    slides: list[dict] = []
    pos = 0
    while True:
        comment = re.search(
            r"<!--\s*(SECTION\s+\d+|Slide Final)[^-]*-->",
            container[pos:],
            re.IGNORECASE,
        )
        slide_match = re.search(r'<div class="slide\b', container[pos:])
        if not slide_match:
            break

        comment_pos = (pos + comment.start()) if comment else None
        slide_pos = pos + slide_match.start()

        if comment and comment_pos is not None and comment_pos < slide_pos:
            text = comment.group(0).upper()
            if "SECTION 1" in text:
                current_section = "section-1"
            elif "SECTION 2" in text:
                current_section = "section-2"
            elif "SECTION 3" in text:
                current_section = "section-3"
            elif "SECTION 4" in text:
                current_section = "section-4"
            elif "SECTION 5" in text:
                current_section = "section-5"
            elif "SECTION 6" in text:
                current_section = "section-6"
            elif "SLIDE FINAL" in text or "RÉCAPITULATIF" in text or "RECAPITULATIF" in text:
                current_section = "section-recap"
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
        if heading:
            if title in {"À vous de jouer !", "Correction"}:
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
    html = re.sub(r'\sstyle="display:\s*none;"', "", html)

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
            if prop in {"color", "font-size", "max-width"}:
                continue
            if prop in {"text-align", "margin-top", "list-style"}:
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
        content: "Excel Module 1 — Les Bases";
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
    font-size: 32pt;
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
    counter-reset: toc;
}

.toc li {
    margin: 12px 0;
    padding: 8px 12px;
    border-bottom: 1px dotted var(--line);
    counter-increment: toc;
}

.toc li a::before {
    content: counter(toc) ". ";
    color: var(--excel);
    font-weight: 700;
}

.doc-header {
    display: none;
}

.section {
    page-break-before: always;
    break-before: page;
}

#section-1 {
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

.slide-body {
    padding: 12px 16px 16px;
}

.subtitle {
    font-size: 12.5pt;
    color: var(--excel-dark);
    font-weight: 700;
    margin: 0 0 10px;
}

.content-box, .exercise-instructions, .solution-box, .why-box,
.visual-description, .golden-rule, .compare-box, .autofill-demo,
.formula-example, .error-item, .card-item {
    border-radius: 6px;
    padding: 10px 12px;
    margin: 8px 0;
}

.why-box, .visual-description, .golden-rule, .compare-box,
.autofill-demo, .formula-example, .error-item, .card-item {
    page-break-inside: avoid;
    break-inside: avoid;
}

.slide-header, .subtitle, h3, h4 {
    break-after: avoid;
    page-break-after: avoid;
}

.content-box {
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

h3 {
    margin: 0 0 8px;
    font-size: 12pt;
    color: var(--excel-dark);
}

.why-box h4 {
    margin: 0 0 6px;
    color: #6a1b9a;
    font-size: 11.5pt;
}

ul, ol { margin: 6px 0 8px; padding-left: 22px; }
li { margin: 3px 0; }
p { margin: 6px 0; }

.content-box ul { list-style: none; padding-left: 0; }
.content-box li { padding-left: 16px; position: relative; }
.content-box li::before {
    content: "▸";
    position: absolute;
    left: 0;
    color: var(--excel);
    font-weight: 700;
}
.content-box li ul { list-style: disc; padding-left: 22px; margin-top: 4px; }
.content-box li ul li { padding-left: 0; }
.content-box li ul li::before { content: none; }

code {
    font-family: "Cascadia Mono", "Consolas", "Courier New", monospace;
    background: #e8f5ec;
    color: #0b5c32;
    padding: 1px 5px;
    border-radius: 3px;
    font-size: 0.95em;
}

.highlight {
    background: #d7eedf;
    padding: 1px 5px;
    border-radius: 3px;
    font-weight: 600;
    color: var(--excel-dark);
}

.duration-badge {
    display: inline-block;
    border: 1px solid #e6b800;
    background: #fff3c4;
    color: #6b5200;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 10pt;
    font-weight: 700;
    margin: 2px 0 8px;
}

.cards-grid, .autofill-grid, .error-grid, .formula-examples, .compare-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    margin: 8px 0;
}

.cards-grid { grid-template-columns: 1fr 1fr 1fr; }
.cards-grid-2 { grid-template-columns: 1fr 1fr; }

.card-item, .autofill-demo, .formula-example, .compare-box {
    background: #f7fbf8;
    border: 1px solid var(--line);
    text-align: center;
}

.card-item .card-icon { font-size: 16pt; }
.card-item .card-name { font-weight: 700; color: var(--excel-dark); margin: 4px 0; }
.card-item .card-desc { font-size: 9.5pt; color: var(--muted); }

.cell-row { display: flex; gap: 6px; justify-content: center; flex-wrap: wrap; margin: 8px 0; }
.demo-cell {
    min-width: 52px;
    padding: 6px 8px;
    border: 1px solid #bbb;
    border-radius: 4px;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9pt;
    background: #fff;
}
.demo-cell.active {
    background: #e3f3e9;
    border-color: var(--excel);
    color: var(--excel-dark);
    font-weight: 700;
}
.autofill-ok { color: var(--solution); font-weight: 700; font-size: 9.5pt; }

.error-item {
    background: #fff1f0;
    border-left: 4px solid #c62828;
}
.error-code {
    font-family: "Consolas", "Courier New", monospace;
    font-weight: 700;
    color: #c62828;
    margin-bottom: 4px;
}

.chc-table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 10.5pt;
}
.chc-table th {
    background: var(--excel);
    color: #fff;
    padding: 7px 8px;
    text-align: left;
}
.chc-table td {
    border: 1px solid #c5d8cc;
    padding: 6px 8px;
}
.chc-table tr:nth-child(even) td { background: var(--soft); }
.chc-table .formula {
    font-family: "Consolas", "Courier New", monospace;
    color: var(--excel-dark);
    font-weight: 700;
}

.formula-code {
    font-family: "Consolas", "Courier New", monospace;
    color: var(--excel-dark);
    font-weight: 700;
    margin-bottom: 4px;
}
.formula-desc { font-size: 9.5pt; color: var(--muted); }

.golden-rule {
    text-align: center;
    border: 2px solid var(--excel);
    background: #eef7f1;
}
.golden-rule .equals {
    font-size: 28pt;
    font-weight: 700;
    color: var(--excel);
    margin: 4px 0;
}

.compare-box.good { border: 1.5px solid var(--solution); background: #eefaf3; }
.compare-box.bad { border: 1.5px solid #e65100; background: #fff4e8; }
.compare-box h4 { margin: 0 0 6px; }
.compare-box.good h4 { color: var(--solution); }
.compare-box.bad h4 { color: #e65100; }

.center-content { text-align: left; }
.icon-large { display: none; }
.slide-body p { color: var(--ink); }
.slide-body [style*="color"] { color: inherit; }

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


def render_print_html(slides: list[dict]) -> str:
    parts: list[str] = []
    parts.append(f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Excel Module 1 — Les Bases (version imprimable)</title>
<style>{PRINT_CSS}</style>
</head>
<body>
<article class="cover">
    <div class="cover-banner">
        <p class="kicker">Formation Excel · Support de cours</p>
        <h1>Module 1 : Les Bases</h1>
    </div>
    <div class="cover-body">
        <h2>Version imprimable</h2>
        <p>Ce document reprend l’intégralité du cours interactif Excel Module&nbsp;1 : théorie, exercices et corrections. Il est prévu pour une impression A4 (recto ou recto-verso).</p>
        <div class="cover-meta">
            <div class="meta-card"><strong>Public</strong>Débutants — prise en main d’Excel</div>
            <div class="meta-card"><strong>Contenu</strong>6 chapitres + récapitulatif, avec exercices corrigés</div>
            <div class="meta-card"><strong>Impression conseillée</strong>A4 portrait, couleurs ou niveaux de gris</div>
            <div class="meta-card"><strong>Usage</strong>Cahier de formation, révision, notes en marge</div>
        </div>
        <p>Les blocs jaunes sont des <strong>exercices</strong>. Les blocs verts sont les <strong>corrections</strong>. Les formules Excel apparaissent en <code>police à chasse fixe</code>.</p>
    </div>
</article>

<nav class="toc">
    <h1>Sommaire</h1>
    <ol>
""")
    toc_labels = [
        ("section-1", "Introduction à Excel"),
        ("section-2", "Saisie de données et formats"),
        ("section-3", "Formules de base et Autofill"),
        ("section-4", "Gestion des feuilles"),
        ("section-5", "Création de graphiques"),
        ("section-6", "Mise en page et impression"),
        ("section-recap", "Récapitulatif"),
    ]
    for sid, label in toc_labels:
        parts.append(f'        <li><a href="#{sid}">{label}</a></li>\n')
    parts.append("""    </ol>
    <p class="footer-note">Astuce : laissez un espace en marge pour vos notes. Les raccourcis clavier (Ctrl+S, Ctrl+Home, F2…) sont indiqués dans chaque chapitre.</p>
</nav>
""")

    current = None
    for slide in slides:
        if slide["section"] != current:
            if current is not None:
                parts.append("</section>\n")
            current = slide["section"]
            label = next(lbl for sid, lbl in SECTIONS if sid == current)
            parts.append(f'<section class="section" id="{current}">\n')
            parts.append(f'    <h1 class="section-title">{label}</h1>\n')

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


def html_to_pdf(html_path: Path, pdf_path: Path) -> None:
    chrome = find_chrome()
    uri = html_path.resolve().as_uri()
    profile = Path("/tmp/chrome-pdf-profile-excel-m1")
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


def main() -> int:
    if not SOURCE.exists():
        print(f"Source introuvable : {SOURCE}", file=sys.stderr)
        return 1
    html = SOURCE.read_text(encoding="utf-8")
    slides = extract_slides(html)
    if not slides:
        print("Aucune diapositive extraite", file=sys.stderr)
        return 1
    PRINT_HTML.write_text(render_print_html(slides), encoding="utf-8")
    html_to_pdf(PRINT_HTML, PRINT_PDF)
    print(f"{len(slides)} diapositives → {PRINT_HTML.name} + {PRINT_PDF.name}")
    print(f"PDF : {PRINT_PDF.stat().st_size / 1024:.0f} Ko")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
