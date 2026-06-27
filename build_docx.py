#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convierte los .md de la estrategia en un solo documento .docx (sin dependencias externas)."""
import os, re, zipfile

BASE = "/projects/sandbox/estrategia-andres-chitiva"
OUT = "/projects/sandbox/Estrategia-Andres-Chitiva.docx"

FILES = [
    "00-INDICE-MAESTRO.md",
    "00-Diagnostico-Fase-0.md",
    "01-Estrategia-Negocio.md",
    "02-Investigacion-Mercado.md",
    "03-Arquitectura-Marca.md",
    "04-Branding-Identidad.md",
    "05-Precios-Paquetes.md",
    "06-Plan-Comercial-Ventas.md",
    "07-Marketing-Digital.md",
    "08-CX-Fidelizacion.md",
    "09-Alianzas-Desarrollo.md",
    "10-Plan-Financiero-Marketing.md",
    "11-KPIs-OKRs.md",
    "12-Estrategia-Lanzamiento.md",
    "13-Implementacion.md",
]

def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def runs(text):
    """Convierte texto inline con **negrita** y `code` en runs de Word."""
    out = []
    # tokenizar por ** y `
    parts = re.split(r"(\*\*.*?\*\*|`.*?`)", text)
    for p in parts:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            inner = esc(p[2:-2])
            out.append(f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{inner}</w:t></w:r>')
        elif p.startswith("`") and p.endswith("`"):
            inner = esc(p[1:-1])
            out.append(f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr><w:t xml:space="preserve">{inner}</w:t></w:r>')
        else:
            out.append(f'<w:r><w:t xml:space="preserve">{esc(p)}</w:t></w:r>')
    return "".join(out) if out else '<w:r><w:t></w:t></w:r>'

def para(text, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return f'<w:p>{ppr}{runs(text)}</w:p>'

def cell(text, header=False):
    shd = '<w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>' if header else ''
    ppr_inner = '<w:pPr></w:pPr>'
    if header:
        # texto blanco y negrita en encabezado
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
            f'<w:p>{ppr_inner}{body}</w:p></w:tc>')

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
    in_code = False
    code_buf = []
    while i < len(lines):
        line = lines[i]
        # bloques de codigo ```
        if line.strip().startswith("```"):
            if not in_code:
                in_code = True; code_buf = []
            else:
                in_code = False
                for cl in code_buf:
                    body.append(para(cl if cl else " ", style="Code"))
            i += 1; continue
        if in_code:
            code_buf.append(line); i += 1; continue
        # tablas
        if line.strip().startswith("|") and "|" in line.strip()[1:]:
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip()); i += 1
            rows = []
            for tl in tbl_lines:
                if re.match(r"^\|[\s:\-\|]+\|?$", tl):  # separador
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
            body.append(para(s[2:], style="Title"))
        elif s.startswith("## "):
            body.append(para(s[3:], style="Heading1"))
        elif s.startswith("### "):
            body.append(para(s[4:], style="Heading2"))
        elif s.startswith("#### "):
            body.append(para(s[5:], style="Heading3"))
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

# Construir cuerpo
content = []
# Portada
content.append(para("Estrategia Comercial y de Marketing", style="Title"))
content.append(para("Proyecto Andrés Chitiva — Academia de Fútbol & Complejo Deportivo", style="Heading1"))
content.append(para("Pachuca, Hidalgo, México · Documento integral de 13 fases", style="Quote"))
content.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

for idx, f in enumerate(FILES):
    path = os.path.join(BASE, f)
    with open(path, encoding="utf-8") as fh:
        md = fh.read()
    content.append(convert(md))
    if idx != len(FILES) - 1:
        content.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

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
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:before="240" w:after="160"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="48"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="C9A227"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="404040"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="120"/></w:pPr><w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="60"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="18"/></w:rPr></w:style>
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
