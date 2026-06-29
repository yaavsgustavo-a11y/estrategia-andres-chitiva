#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aplica formato ejecutivo (portada + indice + pie de pagina) al docx del usuario SIN tocar el contenido."""
import zipfile, shutil, os

SRC = "/projects/sandbox/estrategia-andres-chitiva/editablePlan de Marketing .docx"
OUT = "/projects/sandbox/estrategia-andres-chitiva/Plan-de-Marketing-Andres-Chitiva-EJECUTIVO.docx"
NAVY="1F3864"; MUSTARD="C9A227"

def r(text,color=None,sz=None,bold=False,italic=False):
    rpr="<w:rPr>"
    if bold: rpr+="<w:b/>"
    if italic: rpr+="<w:i/>"
    if color: rpr+=f'<w:color w:val="{color}"/>'
    if sz: rpr+=f'<w:sz w:val="{sz}"/>'
    rpr+="</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'

def pcenter(inner,after=80): return f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="{after}"/></w:pPr>{inner}</w:p>'
PB='<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# Portada: banner navy de ancho completo
banner=(f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'
  f'<w:top w:val="single" w:sz="4" w:color="{NAVY}"/><w:left w:val="single" w:sz="4" w:color="{NAVY}"/>'
  f'<w:bottom w:val="single" w:sz="4" w:color="{NAVY}"/><w:right w:val="single" w:sz="4" w:color="{NAVY}"/></w:tblBorders></w:tblPr>'
  f'<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{NAVY}"/>'
  f'<w:tcMar><w:top w:w="500" w:type="dxa"/><w:left w:w="240" w:type="dxa"/><w:bottom w:w="500" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar></w:tcPr>'
  f'{pcenter(r("PLAN DE MARKETING",color="FFFFFF",bold=True,sz="56"),after=120)}'
  f'{pcenter(r("Proyecto Andrés Chitiva",color=MUSTARD,bold=True,sz="32"),after=60)}'
  f'{pcenter(r("Academia de Fútbol  ·  Complejo Deportivo",color="D6DCE5",sz="24"),after=0)}'
  f'</w:tc></w:tr></w:tbl>')

# Caja confidencial
conf=(f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'
  f'<w:top w:val="single" w:sz="4" w:color="F2F2F2"/><w:left w:val="single" w:sz="24" w:color="7F7F7F"/>'
  f'<w:bottom w:val="single" w:sz="4" w:color="F2F2F2"/><w:right w:val="single" w:sz="4" w:color="F2F2F2"/></w:tblBorders></w:tblPr>'
  f'<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
  f'<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar></w:tcPr>'
  f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>{r("CONFIDENCIAL",color="7F7F7F",bold=True,sz="18")}</w:p>'
  f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{r("Documento de uso interno. Distribución restringida a las partes autorizadas (Dirección, equipo, socios e inversionistas).")}</w:p>'
  f'</w:tc></w:tr></w:tbl>')

cover=('<w:p><w:pPr><w:spacing w:before="600" w:after="200"/></w:pPr></w:p>'
  + banner
  + '<w:p><w:pPr><w:spacing w:after="240"/></w:pPr></w:p>'
  + pcenter(r("Documento ejecutivo preparado por el Departamento de Marketing",color=NAVY,bold=True,sz="24"))
  + pcenter(r("Para revisión y presentación a Dirección  ·  2026",color="595959",sz="20"),after=240)
  + conf + PB)

# Indice (titulo NO usa estilo Heading para no auto-listarse en el TOC)
toc=(f'<w:p><w:pPr><w:spacing w:before="120" w:after="160"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="3" w:color="{MUSTARD}"/></w:pBdr></w:pPr>{r("Contenido",color=NAVY,bold=True,sz="34")}</w:p>'
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  '<w:r><w:rPr><w:i/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">Para ver el índice con números de página: seleccione todo (Ctrl+A) y presione F9, o haga clic derecho y elija &quot;Actualizar campos&quot;.</w:t></w:r>'
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + PB)

inject = cover + toc

footer_xml=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>'
 '<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">Plan de Marketing — Proyecto Andres Chitiva — Confidencial — Pagina </w:t></w:r>'
 '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
 '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r>'
 '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')

zin = zipfile.ZipFile(SRC,"r")
doc = zin.read("word/document.xml").decode("utf-8")
ct  = zin.read("[Content_Types].xml").decode("utf-8")
rels= zin.read("word/_rels/document.xml.rels").decode("utf-8")

# 1) Inyectar portada + indice justo despues de <w:body>
assert "<w:body>" in doc
doc = doc.replace("<w:body>", "<w:body>"+inject, 1)
# 2) Referencia al pie en el sectPr final
doc = doc.replace("<w:sectPr>", '<w:sectPr><w:footerReference w:type="default" r:id="rId6"/>', 1)
# 3) Content type del footer
ct = ct.replace("</Types>", '<Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" PartName="/word/footer1.xml"/></Types>')
# 4) Relationship del footer (rId6 libre)
rels = rels.replace("</Relationships>", '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')

with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        name=item.filename
        if name=="word/document.xml": zout.writestr(name, doc)
        elif name=="[Content_Types].xml": zout.writestr(name, ct)
        elif name=="word/_rels/document.xml.rels": zout.writestr(name, rels)
        else: zout.writestr(name, zin.read(name))
    zout.writestr("word/footer1.xml", footer_xml)
zin.close()
print("OK ->", OUT, os.path.getsize(OUT), "bytes")
