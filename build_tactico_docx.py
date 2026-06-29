#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convierte el Playbook Tactico (markdown) en un .docx profesional (stdlib)."""
import re, zipfile, os

SRC = "/projects/sandbox/estrategia-andres-chitiva/14-Plan-Tactico-Detallado.md"
OUT = "/projects/sandbox/estrategia-andres-chitiva/Playbook-Tactico-Andres-Chitiva.docx"
NAVY="1F3864"; MUSTARD="C9A227"; GOLD="B8860B"; F_NAVY="EAF0FA"; F_GRAY="F2F2F2"

def esc(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def runs(text,color=None,sz=None,bold=False):
    out=[]; parts=re.split(r"(\*\*.*?\*\*)",text)
    for p in parts:
        if not p: continue
        b=bold; t=p
        if p.startswith("**") and p.endswith("**"): b=True; t=p[2:-2]
        rpr="<w:rPr>"+("<w:b/>" if b else "")+(f'<w:color w:val="{color}"/>' if color else "")+(f'<w:sz w:val="{sz}"/>' if sz else "")+"</w:rPr>"
        out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
    return "".join(out) if out else '<w:r><w:t></w:t></w:r>'

def para(text="",style=None,align=None,color=None,sz=None,bold=False,after=120):
    ppr="<w:pPr>"+(f'<w:pStyle w:val="{style}"/>' if style else "")+f'<w:spacing w:after="{after}"/>'+(f'<w:jc w:val="{align}"/>' if align else "")+"</w:pPr>"
    return f'<w:p>{ppr}{runs(text,color,sz,bold)}</w:p>'

def cell(text,header=False,fill=None):
    f=NAVY if header else fill
    shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{f}"/>' if f else ''
    tcpr=f'<w:tcPr><w:tcW w:w="0" w:type="auto"/>{shd}<w:tcMar><w:top w:w="50" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>'
    body=runs(text,color=("FFFFFF" if header else None),bold=header)
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{body}</w:p></w:tc>'

def table(rows):
    b=('<w:tblBorders>'
       '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
       '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
       '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/></w:tblBorders>')
    out=[f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{b}</w:tblPr>']
    for ri,r in enumerate(rows):
        zfill = None if ri==0 else (F_GRAY if ri%2==0 else None)
        out.append("<w:tr>"+"".join(cell(c,header=(ri==0),fill=zfill) for c in r)+"</w:tr>")
    out.append("</w:tbl>"); out.append(para("",after=80))
    return "".join(out)

def callout(text):
    b=(f'<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="{F_NAVY}"/>'
       f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{NAVY}"/>'
       f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{F_NAVY}"/>'
       f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{F_NAVY}"/></w:tblBorders>')
    tcpr=f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{F_NAVY}"/><w:tcMar><w:top w:w="90" w:type="dxa"/><w:left w:w="150" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tcMar></w:tcPr>'
    inner=f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs(text)}</w:p>'
    return f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{b}</w:tblPr><w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>'+para("",after=80)

def convert(md):
    body=[]; lines=md.split("\n"); i=0
    while i<len(lines):
        line=lines[i]
        if line.strip().startswith("|") and "|" in line.strip()[1:]:
            tl=[]
            while i<len(lines) and lines[i].strip().startswith("|"): tl.append(lines[i].strip()); i+=1
            rows=[]
            for t in tl:
                if re.match(r"^\|[\s:\-\|]+\|?$",t): continue
                rows.append([c.strip() for c in t.strip().strip("|").split("|")])
            if rows: body.append(table(rows))
            continue
        s=line.strip()
        if s=="": i+=1; continue
        if s.startswith("# "): body.append(para(s[2:],style="Title"))
        elif s.startswith("## "): body.append(para(s[3:],style="Heading1"))
        elif s.startswith("### "): body.append(para(s[4:],style="Heading2"))
        elif s.startswith("#### "): body.append(para(s[5:],style="Heading3"))
        elif s.startswith("---"): body.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D9D9D9"/></w:pBdr></w:pPr></w:p>')
        elif s.startswith("> "): body.append(callout(s[2:]))
        elif re.match(r"^[-*]\s+",s): body.append(para("•  "+re.sub(r"^[-*]\s+","",s),style="ListPara"))
        elif re.match(r"^\d+\.\s+",s): body.append(para(s,style="ListPara"))
        else: body.append(para(s))
        i+=1
    return "".join(body)

with open(SRC,encoding="utf-8") as f: md=f.read()
C=[para("",after=200),
   '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tr><w:tc>'
   f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{NAVY}"/><w:tcMar><w:top w:w="360" w:type="dxa"/><w:left w:w="240" w:type="dxa"/><w:bottom w:w="360" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar></w:tcPr>'
   f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="100"/></w:pPr>{runs("PLAYBOOK TÁCTICO DE MARKETING",color="FFFFFF",bold=True,sz="48")}</w:p>'
   f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>{runs("Proyecto Andrés Chitiva — 18 estrategias desarrolladas",color=MUSTARD,bold=True,sz="26")}</w:p>'
   f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>{runs("Departamento de Marketing · 2026",color="D6DCE5",sz="20")}</w:p>'
   '</w:tc></w:tr></w:tbl>',
   '<w:p><w:r><w:br w:type="page"/></w:r></w:p>']
C.append(convert(md))
body="".join(C)

document=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 f'<w:body>{body}<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1418" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')

styles='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="160"/><w:pBdr><w:bottom w:val="single" w:sz="18" w:space="3" w:color="C9A227"/></w:pBdr></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="260" w:after="120"/><w:pBdr><w:bottom w:val="single" w:sz="10" w:space="3" w:color="C9A227"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="140" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="B8860B"/><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="50"/></w:pPr></w:style>
</w:styles>'''

footer='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>
<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">Playbook Tactico — Proyecto Andres Chitiva — Confidencial — Pagina </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>'''

ct='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''
rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
drels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml",ct); z.writestr("_rels/.rels",rels)
    z.writestr("word/document.xml",document); z.writestr("word/styles.xml",styles)
    z.writestr("word/footer1.xml",footer); z.writestr("word/_rels/document.xml.rels",drels)
print("OK ->",OUT,os.path.getsize(OUT),"bytes")
