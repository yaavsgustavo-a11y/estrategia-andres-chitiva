#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera el Word del Plan de Preproducción Mensual — sin dependencias externas."""
import os, re, zipfile

BASE = "/projects/sandbox/estrategia-andres-chitiva"
OUT  = "/projects/sandbox/estrategia-andres-chitiva/Preproduccion-Plan-Mensual-Academia-Complejo.docx"

# ── helpers ──────────────────────────────────────────────────────────────────

def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def runs(text):
    out = []
    parts = re.split(r"(\*\*.*?\*\*|`.*?`)", text)
    for p in parts:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            inner = esc(p[2:-2])
            out.append(f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{inner}</w:t></w:r>')
        elif p.startswith("`") and p.endswith("`"):
            inner = esc(p[1:-1])
            out.append(f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>'
                       f'<w:sz w:val="18"/><w:color w:val="C0392B"/></w:rPr>'
                       f'<w:t xml:space="preserve">{inner}</w:t></w:r>')
        else:
            out.append(f'<w:r><w:t xml:space="preserve">{esc(p)}</w:t></w:r>')
    return "".join(out) if out else '<w:r><w:t></w:t></w:r>'

def para(text, style=None, shade=None):
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if shade:
        ppr_parts.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>')
    ppr = f'<w:pPr>{"".join(ppr_parts)}</w:pPr>' if ppr_parts else ''
    return f'<w:p>{ppr}{runs(text)}</w:p>'

def cell(text, header=False, alt=False):
    if header:
        fill = "1F3864"
        inner_runs = []
        for p in re.split(r"(\*\*.*?\*\*)", text):
            if not p:
                continue
            t = p[2:-2] if (p.startswith("**") and p.endswith("**")) else p
            inner_runs.append(
                f'<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr>'
                f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r>'
            )
        body = "".join(inner_runs) if inner_runs else '<w:r><w:t></w:t></w:r>'
    else:
        fill = "E8F0FE" if alt else "FFFFFF"
        body = runs(text)

    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    return (
        f'<w:tc>'
        f'<w:tcPr><w:tcW w:w="0" w:type="auto"/>{shd}'
        f'<w:tcMar>'
        f'<w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        f'<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
        f'</w:tcMar></w:tcPr>'
        f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{body}</w:p>'
        f'</w:tc>'
    )

def table(rows):
    borders = (
        '<w:tblBorders>'
        '<w:top    w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:left   w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:right  w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '</w:tblBorders>'
    )
    tpr = f'<w:tblPr><w:tblW w:w="5000" w:type="pct"/>{borders}<w:tblLook w:val="04A0"/></w:tblPr>'
    out = [f'<w:tbl>{tpr}']
    for i, r in enumerate(rows):
        is_header = (i == 0)
        is_alt = (not is_header) and (i % 2 == 0)
        cells = "".join(cell(c, header=is_header, alt=is_alt) for c in r)
        out.append(f'<w:tr>{cells}</w:tr>')
    out.append('</w:tbl><w:p/>')
    return "".join(out)

def checkbox(line):
    """- [ ] o - [x] → bullet con símbolo unicode."""
    if "- [ ]" in line:
        return para("☐  " + line.replace("- [ ]", "").strip(), style="ListPara")
    if "- [x]" in line or "- [X]" in line:
        return para("☑  " + re.sub(r"- \[[xX]\]", "", line).strip(), style="ListPara")
    return None

def divider():
    return ('<w:p><w:pPr><w:pBdr>'
            '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="C9A227"/>'
            '</w:pBdr><w:spacing w:before="80" w:after="80"/></w:pPr></w:p>')

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# ── MD → Word XML ────────────────────────────────────────────────────────────

def convert(md):
    body = []
    lines = md.split("\n")
    i = 0
    in_code = False
    code_buf = []

    while i < len(lines):
        line = lines[i]

        # bloque código
        if line.strip().startswith("```"):
            if not in_code:
                in_code = True
                code_buf = []
            else:
                in_code = False
                for cl in code_buf:
                    body.append(para(cl if cl else " ", style="Code"))
                body.append('<w:p/>')
            i += 1
            continue
        if in_code:
            code_buf.append(line)
            i += 1
            continue

        # tabla markdown
        if line.strip().startswith("|") and "|" in line.strip()[1:]:
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip())
                i += 1
            rows = []
            for tl in tbl_lines:
                if re.match(r"^\|[\s:\-\|]+\|?$", tl):
                    continue
                cells_raw = [c.strip() for c in tl.strip().strip("|").split("|")]
                rows.append(cells_raw)
            if rows:
                body.append(table(rows))
            continue

        s = line.strip()

        if s == "":
            i += 1
            continue

        # checkboxes
        cb = checkbox(s)
        if cb:
            body.append(cb)
            i += 1
            continue

        # headings
        if s.startswith("# "):
            body.append(para(s[2:], style="DocTitle"))
        elif s.startswith("## "):
            body.append(para(s[3:], style="SectionHead"))
        elif s.startswith("### "):
            body.append(para(s[4:], style="PieceHead"))
        elif s.startswith("#### "):
            body.append(para(s[5:], style="SubHead"))
        elif re.match(r"^-{3,}$", s):
            body.append(divider())
        elif s.startswith("> "):
            body.append(para(s[2:], style="Callout"))
        elif re.match(r"^[-*]\s+", s):
            txt = re.sub(r"^[-*]\s+", "", s)
            body.append(para("•  " + txt, style="ListPara"))
        elif re.match(r"^\d+\.\s+", s):
            body.append(para(s, style="ListPara"))
        else:
            body.append(para(s))

        i += 1

    return "".join(body)

# ── Leer el Markdown ─────────────────────────────────────────────────────────

md_path = os.path.join(BASE, "16-Preproduccion-Plan-Mensual.md")
with open(md_path, encoding="utf-8") as f:
    md_content = f.read()

# ── Construir documento ──────────────────────────────────────────────────────

content = []

# Portada
content.append(
    '<w:p><w:pPr><w:pStyle w:val="DocTitle"/><w:jc w:val="center"/>'
    '<w:spacing w:before="1440" w:after="200"/></w:pPr>'
    '<w:r><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="64"/></w:rPr>'
    '<w:t>PLAN DE PREPRODUCCIÓN MENSUAL</w:t></w:r></w:p>'
)
content.append(
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr>'
    '<w:r><w:rPr><w:b/><w:color w:val="C9A227"/><w:sz w:val="36"/></w:rPr>'
    '<w:t>Academia &amp; Complejo Deportivo</w:t></w:r></w:p>'
)
content.append(
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr>'
    '<w:r><w:rPr><w:color w:val="404040"/><w:sz w:val="26"/></w:rPr>'
    '<w:t>Andrés Chitiva · Pachuca, Hidalgo, México</w:t></w:r></w:p>'
)
content.append(
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="80"/></w:pPr>'
    '<w:r><w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="22"/></w:rPr>'
    '<w:t>Planeación de Contenido Audiovisual · 26 piezas · Academia + Complejo</w:t></w:r></w:p>'
)
content.append(
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="80"/></w:pPr>'
    '<w:r><w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="20"/></w:rPr>'
    '<w:t>Versión 1.0 · Mes 1 de implementación</w:t></w:r></w:p>'
)

# línea decorativa portada
content.append(
    '<w:p><w:pPr><w:jc w:val="center"/><w:pBdr>'
    '<w:bottom w:val="double" w:sz="6" w:space="1" w:color="C9A227"/>'
    '</w:pBdr><w:spacing w:before="240" w:after="240"/></w:pPr></w:p>'
)

content.append(page_break())

# Cuerpo del documento
content.append(convert(md_content))

body_xml = "".join(content)

# ── XMLs del paquete ─────────────────────────────────────────────────────────

document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    f'<w:body>{body_xml}'
    '<w:sectPr>'
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1260" w:header="708" w:footer="708" w:gutter="0"/>'
    '</w:sectPr>'
    '</w:body></w:document>'
)

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

<w:docDefaults>
  <w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
    <w:lang w:val="es-MX"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr>
    <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
  </w:pPr></w:pPrDefault>
</w:docDefaults>

<!-- Normal -->
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/>
</w:style>

<!-- Título principal del documento -->
<w:style w:type="paragraph" w:styleId="DocTitle">
  <w:name w:val="DocTitle"/>
  <w:pPr><w:spacing w:before="240" w:after="160"/></w:pPr>
  <w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="48"/></w:rPr>
</w:style>

<!-- H1: sección principal (ACADEMIA / COMPLEJO) -->
<w:style w:type="paragraph" w:styleId="SectionHead">
  <w:name w:val="heading 1"/>
  <w:pPr>
    <w:spacing w:before="360" w:after="160"/>
    <w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>
    <w:outlineLvl w:val="0"/>
  </w:pPr>
  <w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="32"/></w:rPr>
</w:style>

<!-- H2: semana / subsección -->
<w:style w:type="paragraph" w:styleId="PieceHead">
  <w:name w:val="heading 2"/>
  <w:pPr>
    <w:spacing w:before="280" w:after="120"/>
    <w:shd w:val="clear" w:color="auto" w:fill="E8F0FE"/>
    <w:ind w:left="0"/>
    <w:outlineLvl w:val="1"/>
  </w:pPr>
  <w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="26"/></w:rPr>
</w:style>

<!-- H3: nombre de pieza -->
<w:style w:type="paragraph" w:styleId="SubHead">
  <w:name w:val="heading 3"/>
  <w:pPr>
    <w:spacing w:before="200" w:after="80"/>
    <w:outlineLvl w:val="2"/>
  </w:pPr>
  <w:rPr><w:b/><w:color w:val="C9A227"/><w:sz w:val="24"/></w:rPr>
</w:style>

<!-- Callout / cita -->
<w:style w:type="paragraph" w:styleId="Callout">
  <w:name w:val="Quote"/>
  <w:pPr>
    <w:ind w:left="440" w:right="440"/>
    <w:spacing w:before="80" w:after="80"/>
    <w:shd w:val="clear" w:color="auto" w:fill="FFF9E6"/>
    <w:pBdr>
      <w:left w:val="single" w:sz="12" w:space="8" w:color="C9A227"/>
    </w:pBdr>
  </w:pPr>
  <w:rPr><w:i/><w:color w:val="404040"/><w:sz w:val="22"/></w:rPr>
</w:style>

<!-- Lista -->
<w:style w:type="paragraph" w:styleId="ListPara">
  <w:name w:val="List Paragraph"/>
  <w:pPr><w:ind w:left="440"/><w:spacing w:after="60"/></w:pPr>
</w:style>

<!-- Código -->
<w:style w:type="paragraph" w:styleId="Code">
  <w:name w:val="Code"/>
  <w:pPr>
    <w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>
    <w:spacing w:after="0"/>
    <w:ind w:left="280"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
    <w:sz w:val="18"/>
    <w:color w:val="333333"/>
  </w:rPr>
</w:style>

</w:styles>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>'''

# ── Empaquetar ───────────────────────────────────────────────────────────────

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml",       content_types)
    z.writestr("_rels/.rels",               rels)
    z.writestr("word/document.xml",         document_xml)
    z.writestr("word/styles.xml",           styles_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)

size = os.path.getsize(OUT)
print(f"✅ Documento generado: {OUT}")
print(f"   Tamaño: {size:,} bytes ({size/1024:.1f} KB)")
