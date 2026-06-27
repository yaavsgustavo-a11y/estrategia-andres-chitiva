#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera un documento ejecutivo profesional .docx (sin dependencias externas)."""
import re, zipfile

OUT = "/projects/sandbox/estrategia-andres-chitiva/Plan-Estrategico-Andres-Chitiva-EJECUTIVO.docx"

# Paleta de marca
NAVY = "1F3864"; MUSTARD = "C9A227"; GOLD = "B8860B"
F_NAVY = "EAF0FA"; F_MUST = "FBF3D9"; F_GREEN = "E6F4EA"; F_RED = "FBE9E9"; F_GRAY = "F2F2F2"
A_GREEN = "2E7D32"; A_RED = "C0392B"; A_GRAY = "7F7F7F"

def esc(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def runs(text, color=None, sz=None, bold=False):
    out=[]; parts=re.split(r"(\*\*.*?\*\*)", text)
    for p in parts:
        if not p: continue
        b = bold; t=p
        if p.startswith("**") and p.endswith("**"): b=True; t=p[2:-2]
        rpr="<w:rPr>"
        if b: rpr+="<w:b/>"
        if color: rpr+=f'<w:color w:val="{color}"/>'
        if sz: rpr+=f'<w:sz w:val="{sz}"/>'
        rpr+="</w:rPr>"
        out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
    return "".join(out) if out else '<w:r><w:t></w:t></w:r>'

def p(text="", style=None, align=None, color=None, sz=None, bold=False, after=120):
    ppr="<w:pPr>"
    if style: ppr+=f'<w:pStyle w:val="{style}"/>'
    ppr+=f'<w:spacing w:after="{after}"/>'
    if align: ppr+=f'<w:jc w:val="{align}"/>'
    ppr+="</w:pPr>"
    return f'<w:p>{ppr}{runs(text,color,sz,bold)}</w:p>'

def h1(text): return p(text, style="Heading1")
def h2(text): return p(text, style="Heading2")
def h3(text): return p(text, style="Heading3")
def bullet(text): return p("•  "+text, style="ListPara")

def divider():
    return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D9D9D9"/></w:pBdr></w:pPr></w:p>'

def page_break(): return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def cell(text, header=False, fill=None, color=None, align=None, bold=False, w=None):
    tcpr="<w:tcPr>"
    if w: tcpr+=f'<w:tcW w:w="{w}" w:type="pct"/>'
    else: tcpr+='<w:tcW w:w="0" w:type="auto"/>'
    f = NAVY if header else fill
    if f: tcpr+=f'<w:shd w:val="clear" w:color="auto" w:fill="{f}"/>'
    tcpr+='<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>'
    c = "FFFFFF" if header else color
    b = True if header else bold
    body = runs(text, color=c, bold=b)
    jc = f'<w:jc w:val="{align}"/>' if align else ''
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>{body}</w:p></w:tc>'

def table(rows, header=True, aligns=None, widths=None):
    borders=('<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
        '</w:tblBorders>')
    out=[f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/>{borders}</w:tblPr>']
    # grid
    n=len(rows[0])
    if widths:
        out.append('<w:tblGrid>'+''.join(f'<w:gridCol w:w="{int(w*94)}"/>' for w in widths)+'</w:tblGrid>')
    for ri,r in enumerate(rows):
        cells=[]
        for ci,c in enumerate(r):
            a = aligns[ci] if aligns else None
            wv = widths[ci]*100 if widths else None
            zebra = None if (ri==0 and header) else (F_GRAY if ri%2==0 else None)
            cells.append(cell(c, header=(ri==0 and header), fill=zebra, align=a))
        out.append("<w:tr>"+"".join(cells)+"</w:tr>")
    out.append("</w:tbl>")
    out.append(p("",after=80))
    return "".join(out)

CALL = {
 "insight":(F_NAVY,NAVY,"INSIGHT CLAVE"),
 "oportunidad":(F_GREEN,A_GREEN,"OPORTUNIDAD"),
 "riesgo":(F_RED,A_RED,"RIESGO"),
 "recomendacion":(F_MUST,GOLD,"RECOMENDACIÓN"),
 "decision":(F_GRAY,A_GRAY,"DECISIÓN"),
}
def callout(text, kind="insight", label=None):
    fill,accent,deflabel = CALL[kind]
    lab = label or deflabel
    borders=('<w:tblBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{accent}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        '</w:tblBorders>')
    tcpr=(f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
          '<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner=(f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>{runs(lab,color=accent,bold=True,sz="18")}</w:p>'
           f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs(text)}</w:p>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{borders}</w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>'+p("",after=80))

def banner(title, subtitle_lines):
    tcpr=(f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{NAVY}"/>'
          '<w:tcMar><w:top w:w="400" w:type="dxa"/><w:left w:w="240" w:type="dxa"/><w:bottom w:w="400" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner=f'<w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr>{runs(title,color="FFFFFF",bold=True,sz="52")}</w:p>'
    for i,s in enumerate(subtitle_lines):
        col = MUSTARD if i==0 else "D6DCE5"
        szv = "30" if i==0 else "22"
        inner+=f'<w:p><w:pPr><w:spacing w:after="60"/><w:jc w:val="center"/></w:pPr>{runs(s,color=col,bold=(i==0),sz=szv)}</w:p>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>')

def toc():
    return ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'+runs("Contenido")+'</w:p>'
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:i/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">Haga clic derecho aqui y elija "Actualizar campos" para generar el indice con numeros de pagina.</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')

# =================== CONTENIDO ===================
C=[]
# Portada
C.append('<w:p><w:pPr><w:spacing w:before="400" w:after="200"/></w:pPr></w:p>')
C.append(banner("PLAN ESTRATÉGICO COMERCIAL Y DE MARKETING",
   ["Proyecto Andrés Chitiva",
    "Academia de Fútbol  ·  Complejo Deportivo",
    "Pachuca, Hidalgo, México  ·  2026"]))
C.append(p("",after=240))
C.append(p("Documento ejecutivo preparado por el Departamento de Marketing", align="center", color=NAVY, bold=True, sz="24"))
C.append(p("Para revisión de Dirección, equipo, socios e inversionistas", align="center", color="595959", sz="20"))
C.append(p("",after=240))
C.append(callout("Documento de carácter confidencial. Contiene estrategia comercial, proyecciones y datos sensibles del proyecto. Distribución restringida a las partes autorizadas.", kind="decision", label="CONFIDENCIAL"))
C.append(page_break())
# TOC
C.append(toc())
C.append(page_break())

# 1. RESUMEN EJECUTIVO
C.append(h1("1. Resumen Ejecutivo"))
C.append(p("El Proyecto Andrés Chitiva integra dos marcas complementarias en Pachuca, Hidalgo: una **Academia de Fútbol** de formación premium y un **Complejo Deportivo** multidisciplinario. Ambas operan hoy en etapa piloto y abren oficialmente al público en la **segunda quincena de agosto de 2026**."))
C.append(p("El presente plan define el posicionamiento, la arquitectura de marca, la estrategia comercial y digital, el modelo de experiencia y retención, las alianzas, el plan financiero de marketing, los indicadores y la hoja de ruta de lanzamiento; todo articulado hacia un objetivo medible."))
C.append(callout("El mercado local ofrece academias genéricas a $400–600 mientras el papá objetivo está dispuesto a pagar hasta $1,200. Ese segmento premium está **prácticamente desocupado**: el único competidor con renombre vive de su nombre y su calidad percibida está en declive. Andrés Chitiva puede tomar ese liderazgo.", kind="oportunidad"))
C.append(callout("Con un valor de vida del alumno (LTV) de ~$15,900 al año (~$30,000 a dos años) y un costo de adquisición estimado de $1,000–1,500, la relación **LTV:CAC ≈ 10:1**. El negocio es altamente rentable en su unidad económica; el presupuesto limita la velocidad, no la viabilidad.", kind="insight"))
C.append(h3("Objetivo central a 12 meses"))
C.append(table([
  ["Indicador","Punto de partida","Meta 12 meses","Escenario ideal"],
  ["Alumnos Academia","~30 (piloto)","90","136 (cupo lleno)"],
  ["Ingreso recurrente mensual","$36,000","$108,000","$163,200"],
  ["Posicionamiento","Desconocido (cerrado)","Referente premium regional","Líder de categoría"],
], aligns=["left","center","center","center"], widths=[0.34,0.22,0.22,0.22]))
C.append(h3("Lo que solicitamos a Dirección"))
C.append(bullet("Confirmar el estatus operativo de **nutrición y servicios médicos** (para validar la promesa de marca)."))
C.append(bullet("Aprobar el **presupuesto de pauta mensual** ($3,500–5,000) con refuerzo para el mes de lanzamiento."))
C.append(bullet("Validar las **promociones propuestas** (Programa Fundadores, pago anual, hermanos, referidos)."))
C.append(bullet("Respaldar el **evento de lanzamiento de agosto** (agenda de Andrés, exjugadores, invitados institucionales)."))
C.append(page_break())

# 2. CONTEXTO Y OPORTUNIDAD
C.append(h1("2. Contexto y Oportunidad"))
C.append(p("El proyecto se compone de dos marcas distintas pero relacionadas, unidas por la figura de Andrés Chitiva, exfutbolista profesional y referente en Hidalgo."))
C.append(table([
  ["","Academia de Fútbol","Complejo Deportivo"],
  ["Qué vende","Formación y desarrollo (ingreso recurrente)","Uso de instalaciones y servicios (transaccional)"],
  ["Cliente","Padres 30–50 años, NSE medio-alto/alto","Rentas, ligas, funcional, familias, empresas"],
  ["Rol estratégico","Motor de marca + activo franquiciable","Motor de utilidad / caja"],
], aligns=["left","left","left"], widths=[0.2,0.4,0.4]))
C.append(p("**Etapa actual:** prueba piloto cerrada al público, con ~30 alumnos. **Apertura oficial:** segunda quincena de agosto de 2026. La marca aún se construye desde cero en redes y adquisición."))
C.append(callout("Pachuca es la 'cuna del fútbol' en México. La cantera institucional de referencia capta talento de forma cerrada/elitista y no compite por el público abierto. Su prestigio futbolístico es **viento de cola cultural** para el proyecto, no competencia directa.", kind="insight"))
C.append(page_break())

# 3. DIAGNÓSTICO (FODA)
C.append(h1("3. Diagnóstico Estratégico"))
C.append(table([
  ["FORTALEZAS","OPORTUNIDADES"],
  ["Nombre de Andrés Chitiva con peso ante el decisor de compra (los padres).","Tier premium local desocupado; competidor histórico en declive."],
  ["Instalaciones nuevas de primer nivel.","Cross-selling Academia↔Complejo (ecosistema)."],
  ["Cámaras con highlights descargables; psicóloga deportiva operando.","Pádel y ligas como flujo de caja adulto."],
  ["Prueba de valor: alumnos que viajan de otras ciudades por la metodología.","Escuelas y empresas como canales de adquisición (sin explotar)."],
], aligns=["left","left"], widths=[0.5,0.5]))
C.append(table([
  ["DEBILIDADES","RIESGOS"],
  ["Diferenciadores aún por demostrar de forma consistente.","Brecha promesa vs. entrega (nutrición/servicios médicos)."],
  ["Marca y redes desde casi cero (~200 seguidores).","Embudo de adquisición en frío sin validar (CAC desconocido)."],
  ["Sin métricas financieras base (las controla Finanzas).","Objeción recurrente: distancia y tráfico."],
  ["Capacidad ociosa del Complejo (horas muertas).","Precio en el tope del rango tolerado por el mercado."],
], aligns=["left","left"], widths=[0.5,0.5]))
C.append(page_break())

# 4. MODELO Y ARQUITECTURA
C.append(h1("4. Modelo de Negocio y Arquitectura de Marca"))
C.append(h2("4.1 Roles de cada marca"))
C.append(p("El **Complejo** es el motor de utilidad (mayor techo de ingreso por múltiples flujos y operación de hasta 17 h/día). La **Academia** es el motor de marca y el activo de mayor valor a futuro: su metodología es la base de la franquicia 'Metodología Chitiva'."))
C.append(callout("El Complejo da de comer hoy; la Academia hace rico mañana. Debe protegerse el presupuesto de la Academia aunque su retorno inmediato sea menor: construye el prestigio y la propiedad intelectual franquiciable.", kind="recomendacion"))
C.append(h2("4.2 Arquitectura de marca y canales"))
C.append(p("Andrés Chitiva es **aval y vocero**, no marca paraguas. Academia y Complejo son marcas hermanas con familia visual compartida (azul marino + amarillo mostaza)."))
C.append(table([
  ["Canal","Estructura","Decisión"],
  ["Facebook","2 cuentas separadas","Se mantienen separadas (público local/adulto, ligas y eventos)"],
  ["Instagram / TikTok","Cuenta compartida","Hub del ecosistema con reglas de diferenciación visual por marca"],
  ["Cuenta personal de Andrés","Independiente","Amplificador / vocería"],
], aligns=["left","left","left"], widths=[0.24,0.28,0.48]))
C.append(page_break())

# 5. MERCADO Y POSICIONAMIENTO
C.append(h1("5. Mercado y Posicionamiento"))
C.append(table([
  ["Competidor","Precio","Oferta","Debilidad"],
  ["Futbol 7 Villas","$500","Solo cancha","Sin formación"],
  ["La Bombonera","$500","Solo cancha","Sin formación"],
  ["UFD","s/d","Instalaciones + renombre","Calidad percibida en declive; vive del nombre"],
  ["Andrés Chitiva","$1,200","Método + instalaciones + acompañamiento","Marca nueva por construir"],
], aligns=["left","center","left","left"], widths=[0.22,0.13,0.33,0.32]))
C.append(callout("El mercado cobra $500 pero está dispuesto a pagar $800–1,200. Andrés Chitiva ocupa ese espacio vacío. El trabajo del marketing no es bajar el precio, sino **justificar el valor**: por qué vale 2–3 veces más.", kind="insight"))
C.append(h2("5.1 Declaración de posicionamiento"))
C.append(callout("La academia premium vigente y completa de Pachuca —método profesional real, instalaciones nuevas y acompañamiento cercano— para el padre que quiere lo mejor para su hijo, dentro y fuera de la cancha.", kind="decision", label="POSICIONAMIENTO"))
C.append(page_break())

# 6. PLATAFORMA DE MARCA
C.append(h1("6. Plataforma de Marca"))
C.append(table([
  ["Elemento","Definición"],
  ["Esencia","Formamos al jugador y a la persona."],
  ["Personalidad","Premium profesional, aspiracional pero no excluyente; cálida y humana."],
  ["Atributos","Profesional · Premium · Nueva · Completa · Humana/Cercana."],
  ["Manifiesto","“Formamos futbolistas inteligentes, apasionados y con valores.”"],
  ["Tagline de apoyo","“Método profesional. Formación de vida.”"],
], aligns=["left","left"], widths=[0.24,0.76]))
C.append(h3("Razones para creer"))
C.append(bullet("Respaldo real de Andrés Chitiva y su cuerpo técnico."))
C.append(bullet("Metodología estructurada de 4 pilares con evaluación individual continua."))
C.append(bullet("Instalaciones nuevas, cámaras con highlights y psicóloga deportiva."))
C.append(bullet("Prueba tangible: alumnos que viajan de otras ciudades por el método."))
C.append(callout("La nutrición y los servicios médicos solo se comunican como beneficio incluido cuando operen al 100%. Hasta entonces se manejan como 'Próximamente' para no arriesgar la confianza premium.", kind="riesgo", label="ALERTA DE PROMESA"))
C.append(page_break())

# 7. PRECIOS Y PAQUETES
C.append(h1("7. Estrategia de Precios y Paquetes"))
C.append(table([
  ["Servicio","Precio"],
  ["Academia — Mensualidad","$1,200"],
  ["Academia — Inscripción","$1,500"],
  ["Pádel — 1 h / 2 h / 3 h","$350 / $600 / $850"],
  ["Entrenamiento Funcional","$800 · $600 estudiante · $700 parejas"],
  ["Ligas — Inscripción / Partido","$600 / $600"],
], aligns=["left","center"], widths=[0.6,0.4]))
C.append(callout("Cualquier flexibilidad en la mensualidad debe ser justificada, acotada y nunca general/permanente. El precio de lista siempre se mantiene creíble: el padre debe sentir que se 'ganó' un beneficio, no que el precio expuesto es falso.", kind="recomendacion", label="REGLA DE INTEGRIDAD DE PRECIO"))
C.append(h3("Palancas promocionales (a validar por Finanzas)"))
C.append(table([
  ["Mecánica","Objetivo"],
  ["Programa 'Generación Fundadora'","Convertir y fidelizar a las familias piloto como embajadores"],
  ["Pago anual anticipado (11x12)","Caja anticipada + retención, sin tocar el precio visible"],
  ["Descuento por hermanos","Mayor valor por familia (LTV)"],
  ["Referidos (beneficio puntual)","Crecimiento con bajo costo de adquisición"],
  ["Beneficio 'Familia Chitiva' (pádel)","Cross-selling + llenado de horas valle del Complejo"],
], aligns=["left","left"], widths=[0.42,0.58]))
C.append(page_break())

# 8. MOTOR COMERCIAL
C.append(h1("8. Motor Comercial y de Ventas"))
C.append(p("Embudo ramificado que coloca **la experiencia antes del precio**, adaptado a prospectos presenciales, digitales/distantes y de decisión rápida."))
C.append(table([
  ["Etapa","Acción clave"],
  ["1. Contacto","Auto-respuesta instantánea + humano en <15–30 min"],
  ["2. Descubrimiento","Detectar motivación y zona del prospecto"],
  ["3. Experiencia","Clase Muestra + recorrido (o assets digitales si es a distancia)"],
  ["4. Valor + Precio","Precio anclado en lo que ya vivió"],
  ["5. Cierre","Beneficios + manejo de objeciones (cupo limitado)"],
  ["6. Seguimiento","Nurture en CRM (Clupik) + lista de espera"],
], aligns=["left","left"], widths=[0.26,0.74]))
C.append(callout("La 'Clase Muestra' es la herramienta de conversión #1. La objeción de distancia se neutraliza con tres argumentos: 'vale el viaje' (prueba de los alumnos foráneos), las cámaras (ver al hijo en remoto) y horarios que esquivan el tráfico.", kind="insight"))
C.append(p("Herramientas: WhatsApp Business (auto-reply + scripts + etiquetas de pipeline) y **Clupik** como CRM y motor de seguimiento. Scripts, manual comercial y manejo de objeciones están documentados y listos para capacitación."))
C.append(page_break())

# 9. DIGITAL
C.append(h1("9. Estrategia Digital y de Contenido"))
C.append(p("Estrategia **content-led**: la fortaleza es el contenido orgánico (equipo audiovisual + cámaras + highlights); la pauta amplifica. Capacidad de producción: 4 reels + 4 posts por semana."))
C.append(callout("Los highlights de las cámaras son el motor orgánico #1: el padre comparte el video de su hijo y genera alcance gratuito con el público objetivo exacto. Conecta adquisición y retención.", kind="oportunidad"))
C.append(table([
  ["Bucket de contenido","Peso","Objetivo"],
  ["Atraer","40%","Crecer audiencia (highlights, jugadas, ambiente)"],
  ["Educar","25%","Autoridad y justificación del premium"],
  ["Conectar","20%","Confianza (testimonios, Andrés, valores)"],
  ["Convertir","15%","Clase Muestra, cupo, CTA a WhatsApp"],
], aligns=["left","center","left"], widths=[0.22,0.12,0.66]))
C.append(p("**Pauta concentrada:** Meta (IG/FB) como canal principal hacia registros de Clase Muestra + retargeting; Google Search para alta intención; TikTok orgánico. Prioridad: validar el CAC en frío antes de escalar."))
C.append(page_break())

# 10. CX
C.append(h1("10. Experiencia del Cliente y Fidelización"))
C.append(callout("En un negocio recurrente, retener vale más que adquirir: un alumno que permanece 2 años vale ~$30,000. Además, la retención es la prueba de resultados que habilita la franquicia.", kind="insight"))
C.append(bullet("**Onboarding premium:** kit de bienvenida (medalla, diploma, termo, folder, 2 uniformes, cuponera de aliados) + primer día especial."))
C.append(bullet("**Demostración de valor mensual:** pruebas físicas + reportes en Clupik + highlight semanal del hijo + contacto del entrenador."))
C.append(bullet("**Comunidad:** torneos internos, convivencias, ceremonia de fin de ciclo, identidad 'Generación Fundadora'."))
C.append(bullet("**Alerta temprana:** caída de asistencia en Clupik dispara contacto proactivo antes de la fuga."))
C.append(bullet("**Medición:** NPS a 30 días y trimestral."))
C.append(page_break())

# 11. ALIANZAS
C.append(h1("11. Alianzas y Desarrollo de Negocio"))
C.append(p("Patrocinadores actuales gracias a Andrés: **Voit, Teqball y Mazda Pachuca**, además de aliados locales."))
C.append(callout("Los clientes de Mazda Pachuca son familias de NSE medio-alto/alto: exactamente el target de la Academia. La promoción cruzada con su base es una vía de adquisición de altísima precisión.", kind="oportunidad", label="OPORTUNIDAD DESTACADA"))
C.append(table([
  ["Frente","Prioridad","Beneficio"],
  ["Escuelas/colegios","Alta","Alumnos + llenado de horas matutinas del Complejo"],
  ["Empresas","Media-alta","Ingreso B2B (ligas, eventos, funcional)"],
  ["Salud/nutrición","Media","Refuerzo de la propuesta integral"],
  ["Medios locales","Alta (lanzamiento)","Prensa ganada vía Andrés"],
], aligns=["left","center","left"], widths=[0.26,0.22,0.52]))
C.append(p("**Franquicia 'Metodología Chitiva' (visión a 3 años):** acción inmediata = documentar el método y registrar resultados como propiedad intelectual."))
C.append(page_break())

# 12. FINANCIERO
C.append(h1("12. Plan Financiero de Marketing"))
C.append(h3("Unidad económica"))
C.append(table([
  ["Métrica","Valor"],
  ["Valor año 1 por alumno","$15,900 ($1,500 inscripción + $1,200 × 12)"],
  ["LTV (~2 años)","~$30,000"],
  ["CAC estimado","$1,000–1,500"],
  ["Relación LTV:CAC","≈ 10:1 (sano ≥ 3:1)"],
], aligns=["left","left"], widths=[0.4,0.6]))
C.append(h3("Matemática del embudo (meta +60 alumnos / 12 meses)"))
C.append(table([
  ["Etapa","Conversión","Necesidad mensual"],
  ["Inscripciones netas","—","~6"],
  ["Asistentes a Clase Muestra","40–50% cierran","~12–15"],
  ["Leads","30–40% asisten","~30–50"],
  ["Inversión Meta","CPL ~$100–130","~$4,000–5,000"],
], aligns=["left","center","center"], widths=[0.4,0.32,0.28]))
C.append(callout("El presupuesto chico limita la VELOCIDAD, no la rentabilidad. Una vez validado el CAC real en los primeros meses, el ratio 10:1 justifica solicitar a Finanzas un presupuesto mayor para escalar.", kind="recomendacion"))
C.append(h3("Proyección de ingreso — Academia"))
C.append(table([
  ["Escenario","Alumnos","Recurrente mensual","Anualizado"],
  ["Hoy","30","$36,000","$432,000"],
  ["Meta 12 meses","90","$108,000","$1,296,000"],
  ["Cupo lleno","136","$163,200","$1,958,400"],
], aligns=["left","center","center","center"], widths=[0.28,0.18,0.27,0.27]))
C.append(page_break())

# 13. KPIs
C.append(h1("13. Indicadores (KPIs y OKRs)"))
C.append(h3("OKRs del ciclo de lanzamiento (Ago 2026 – Ene 2027)"))
C.append(table([
  ["Objetivo","Resultado clave"],
  ["Lanzar la Academia con fuerza","45 alumnos al cierre de agosto; 40–60 registros a Clase Muestra"],
  ["Validar adquisición y construir marca","CPL < $130; CAC < $1,500; ~1,500 seguidores"],
  ["Retener y demostrar valor","Retención > 90%; NPS > 60; highlight semanal por alumno"],
  ["Activar el Complejo","Ocupación de horas valle creciente; equipos en ligas"],
], aligns=["left","left"], widths=[0.38,0.62]))
C.append(p("**Ritmo de revisión:** semanal (leads, respuesta, contenido), mensual (CAC, inscripciones, churn, ocupación) y trimestral (OKRs, NPS, escalamiento de presupuesto)."))
C.append(page_break())

# 14. LANZAMIENTO
C.append(h1("14. Estrategia de Lanzamiento (Agosto 2026)"))
C.append(table([
  ["Fase","Periodo","Foco"],
  ["Expectativa","Med. jul – med. ago","Producir assets, intriga, lista de espera, pre-venta a Fundadores"],
  ["Inauguración","2ª quincena agosto","Evento con medios, autoridad, Andrés y exjugadores; campaña Clase Muestra"],
  ["Conversión","Septiembre+","Convertir leads, retargeting, medir CAC, ritmo continuo"],
], aligns=["left","center","left"], widths=[0.2,0.25,0.55]))
C.append(callout("El evento de inauguración debe tener propósito comercial: cada asistente se captura como lead (registro con QR), con mesa de inscripción y beneficios de lanzamiento. PR + base de datos + inscripciones en un solo momento.", kind="recomendacion"))
C.append(page_break())

# 15. HOJA DE RUTA
C.append(h1("15. Hoja de Ruta e Implementación"))
C.append(table([
  ["Periodo","Acciones clave"],
  ["Fin jun – ~10 jul","Cerrar piloto; configurar WhatsApp Business y Clupik; activar highlights"],
  ["~10 jul","Evaluar piloto; definir reingresos Fundadores; presentar promos a Finanzas"],
  ["Med. jul","Producir assets de video; iniciar campaña de intriga"],
  ["Fin jul","Abrir lista de espera; logística e invitaciones del evento"],
  ["Inicio ago","Pre-venta; capacitar ventas; preparar mesa de inscripción + QR"],
  ["2ª quincena ago","Evento de lanzamiento + apertura + pauta reforzada"],
  ["Septiembre+","Conversión, retargeting, medición de CAC, ritmo continuo"],
], aligns=["left","left"], widths=[0.24,0.76]))
C.append(h3("Responsables"))
C.append(table([
  ["Área","Responsable"],
  ["Estrategia, campañas, presupuesto, KPIs","Marketing"],
  ["Assets, reels, video-tour, highlights","Diseño + Audiovisual"],
  ["Publicación, primer contacto","Community"],
  ["Embudo, scripts, cierre, Clase Muestra","Recepción / Ventas"],
  ["Evento, prensa, vocería","Andrés (+ exjugadores)"],
  ["Precios, presupuesto, costo fijo","Finanzas / Dirección"],
  ["Clase Muestra, evaluaciones, resultados","Cuerpo técnico"],
], aligns=["left","left"], widths=[0.55,0.45]))
C.append(page_break())

# 16. CIERRE
C.append(h1("16. Cierre y Solicitudes a Dirección"))
C.append(p("Este plan posiciona al Proyecto Andrés Chitiva para tomar el liderazgo del segmento premium en Pachuca, con una unidad económica sólida (LTV:CAC ≈ 10:1) y una ruta clara hacia el lanzamiento de agosto y el crecimiento sostenible."))
C.append(p("Para ejecutarlo a tiempo, el Departamento de Marketing solicita a Dirección:"))
C.append(bullet("Confirmación del estatus operativo de nutrición y servicios médicos."))
C.append(bullet("Aprobación del presupuesto de pauta mensual con refuerzo para agosto."))
C.append(bullet("Validación de las promociones propuestas (vía Finanzas)."))
C.append(bullet("Respaldo institucional y de agenda para el evento de lanzamiento."))
C.append(callout("Con producto validado (alumnos que ya viajan por el método), un mercado dispuesto a pagar premium y un competidor en declive, la ventana de oportunidad es ahora. La calidad de la ejecución del lanzamiento de agosto definirá el ritmo de los próximos 12 meses.", kind="insight", label="MENSAJE FINAL"))
C.append(p("",after=200))
C.append(p("Departamento de Marketing — Proyecto Andrés Chitiva — 2026", align="center", color="808080", sz="18"))

body="".join(C)

document_xml=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 f'<w:body>{body}'
 '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
 '<w:pgSz w:w="12240" w:h="15840"/>'
 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1418" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
 '</w:sectPr></w:body></w:document>')

styles_xml='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="3" w:color="C9A227"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="34"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="27"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="B8860B"/><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="60"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="1F3864"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>'''

footer_xml='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>
<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">Plan Estrategico Comercial y de Marketing — Proyecto Andres Chitiva — Confidencial — Pagina </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>'''

content_types='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

doc_rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/footer1.xml", footer_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)

import os
print("OK ->", OUT, os.path.getsize(OUT), "bytes")
