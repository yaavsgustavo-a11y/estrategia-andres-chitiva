#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convierte 17-Contenidos-Marketing-Digital.md en un .docx (sin dependencias externas)."""
import os, re, zipfile

BASE = "/projects/sandbox/estrategia-andres-chitiva"
SRC = os.path.join(BASE, "17-Contenidos-Marketing-Digital.md")
OUT = os.path.join(BASE, "Contenidos-Marketing-Andres-Chitiva.docx")

def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

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
            out.append(f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:color w:val="0F6E3D"/></w:rPr><w:t xml:space="preserve">{inner}</w:t></w:r>')
        else:
            out.append(f'<w:r><w:t xml:space="preserve">{esc(p)}</w:t></w:r>')
    return "".join(out) if out else '<w:r><w:t></w:t></w:r>'

def para(text, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return f'<w:p>{ppr}{runs(text)}</w:p>'

def cell(text, header=False):
    shd = '<w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>' if header else ''
    if header:
        inner_runs = []
        parts = re.split(r"(\*\*.*?\*\*)", text)
        for p in parts:
            if not p:
                continue
            t = p[2:-2] if (p.startswith("**") and p.endswith("**")) else p
            inner_runs.append(f'<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
        body = "".join(inner_runs) if inner_runs else '<w:r><w:t></w:t></w:r>'
    else:
        body = runs(text)
    return (f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>{shd}'
            f'<w:tcMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
            f'<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar></w:tcPr>'
            f'<w:p>{body}</w:p></w:tc>')

def table(rows):
    borders = ('<w:tblBorders>'
               '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
               '</w:tblBorders>')
    tpr = f'<w:tblPr><w:tblW w:w="5000" w:type="pct"/>{borders}</w:tblPr>'
    out = [f'<w:tbl>{tpr}']
    for i, r in enumerate(rows):
        cells = "".join(cell(c, header=(i == 0)) for c in r)
        out.append(f'<w:tr>{cells}</w:tr>')
    out.append('</w:tbl>')
    return "".join(out)

def convert(md):
    body = []
    lines = md.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip().startswith("|") and "|" in line.strip()[1:]:
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip()); i += 1
            rows = []
            for tl in tbl_lines:
                if re.match(r"^\|[\s:\-\|]+\|?$", tl):
                    continue
                cells = [c.strip() for c in tl.strip().strip("|").split("|")]
                rows.append(cells)
            if rows:
                body.append(table(rows))
            continue
        s = line.strip()
        if s == "":
            i += 1; continue
        if s.startswith("# "):
            body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
            body.append(para(s[2:], style="Heading1"))
        elif s.startswith("## "):
            body.append(para(s[3:], style="Heading2"))
        elif s.startswith("### "):
            body.append(para(s[4:], style="Heading3"))
        elif s.startswith("#### "):
            body.append(para(s[5:], style="Heading4"))
        elif s.startswith("---") or s.startswith("==="):
            body.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A6A6A6"/></w:pBdr></w:pPr></w:p>')
        elif s.startswith("> "):
            body.append(para(s[2:], style="Quote"))
        elif re.match(r"^[-*]\s+", s):
            body.append(para("•  " + re.sub(r"^[-*]\s+", "", s), style="ListPara"))
        elif re.match(r"^\d+\.\s+", s):
            body.append(para(s, style="ListPara"))
        else:
            body.append(para(s))
        i += 1
    return "".join(body)

with open(SRC, encoding="utf-8") as fh:
    md = fh.read()

# Quitar el primer H1 del markdown (lo reemplazamos por portada) para no duplicar salto
content = []
content.append(para("Plan de Contenidos de Marketing Digital Deportivo", style="Title"))
content.append(para("Academia y Complejo Deportivo Andrés Chitiva", style="Heading1"))
content.append(para("70 propuestas desarrolladas · 10 por línea de producto · Copys, Reels, TikToks, Carruseles y Guiones", style="Quote"))
content.append(para("Pachuca, Hidalgo, México", style="Quote"))

# Eliminar la primera línea de título y subtítulo del md para evitar repetición
md_lines = md.split("\n")
# saltar hasta el primer "> 70 propuestas" incluido, empezar contenido desde el bloque CÓMO USAR
start_idx = 0
for idx, ln in enumerate(md_lines):
    if ln.strip().startswith("## CÓMO USAR"):
        start_idx = idx
        break
md_body = "\n".join(md_lines[start_idx:])
content.append(convert(md_body))

body_xml = "".join(content)

document_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    f'<w:body>{body_xml}'
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
    '</w:sectPr></w:body></w:document>')

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:before="240" w:after="160"/></w:pPr><w:rPr><w:b/><w:color w:val="0F6E3D"/><w:sz w:val="46"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="0F6E3D"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="C9A227"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:pPr><w:spacing w:before="120" w:after="60"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:color w:val="404040"/><w:sz w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="120"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="C9A227"/></w:pBdr></w:pPr><w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="60"/></w:pPr></w:style>
</w:styles>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)

print("OK ->", OUT, os.path.getsize(OUT), "bytes")
