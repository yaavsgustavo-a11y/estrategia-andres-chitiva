#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera el PLAN DE MARKETING COMPLETO Y PROFESIONAL (.docx) del Proyecto Andrés Chitiva.
Documento autónomo, detallado y listo para presentar a equipo, socios e inversionistas.
Sin dependencias externas (solo stdlib: re, zipfile, os).
"""
import re, zipfile, os

OUT = "/projects/sandbox/estrategia-andres-chitiva/Plan-de-Marketing-Andres-Chitiva.docx"

# ============================ PALETA DE MARCA ============================
NAVY = "1F3864"; MUSTARD = "C9A227"; GOLD = "B8860B"
F_NAVY = "EAF0FA"; F_MUST = "FBF3D9"; F_GREEN = "E6F4EA"; F_RED = "FBE9E9"; F_GRAY = "F2F2F2"
F_TEAL = "E3F1F1"
A_GREEN = "2E7D32"; A_RED = "C0392B"; A_GRAY = "7F7F7F"; A_TEAL = "0E7C7B"

# ============================ PRIMITIVAS XML ============================
def esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def runs(text, color=None, sz=None, bold=False, italic=False):
    out = []
    parts = re.split(r"(\*\*.*?\*\*)", str(text))
    for p in parts:
        if not p:
            continue
        b = bold; t = p
        if p.startswith("**") and p.endswith("**"):
            b = True; t = p[2:-2]
        rpr = "<w:rPr>"
        if b: rpr += "<w:b/>"
        if italic: rpr += "<w:i/>"
        if color: rpr += f'<w:color w:val="{color}"/>'
        if sz: rpr += f'<w:sz w:val="{sz}"/>'
        rpr += "</w:rPr>"
        out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
    return "".join(out) if out else '<w:r><w:t></w:t></w:r>'

def p(text="", style=None, align=None, color=None, sz=None, bold=False, italic=False, after=120):
    ppr = "<w:pPr>"
    if style: ppr += f'<w:pStyle w:val="{style}"/>'
    ppr += f'<w:spacing w:after="{after}"/>'
    if align: ppr += f'<w:jc w:val="{align}"/>'
    ppr += "</w:pPr>"
    return f'<w:p>{ppr}{runs(text, color, sz, bold, italic)}</w:p>'

def h1(text): return p(text, style="Heading1")
def h2(text): return p(text, style="Heading2")
def h3(text): return p(text, style="Heading3")
def bullet(text): return p("•  " + text, style="ListPara")
def numbered(text): return p(text, style="ListPara")

def divider():
    return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D9D9D9"/></w:pBdr></w:pPr></w:p>'

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def spacer(after=120):
    return p("", after=after)

# ============================ TABLAS ============================
def cell(text, header=False, fill=None, color=None, align=None, bold=False, sz=None):
    tcpr = '<w:tcPr><w:tcW w:w="0" w:type="auto"/>'
    f = NAVY if header else fill
    if f: tcpr += f'<w:shd w:val="clear" w:color="auto" w:fill="{f}"/>'
    tcpr += ('<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
             '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>'
             '<w:vAlign w:val="center"/></w:tcPr>')
    c = "FFFFFF" if header else color
    b = True if header else bold
    body = runs(text, color=c, bold=b, sz=sz)
    jc = f'<w:jc w:val="{align}"/>' if align else ''
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>{body}</w:p></w:tc>'

def table(rows, header=True, aligns=None, widths=None, zebra_on=True, sz=None, first_col_bold=False):
    borders = ('<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
        '</w:tblBorders>')
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/>{borders}</w:tblPr>']
    if widths:
        out.append('<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{int(w*9400)}"/>' for w in widths) + '</w:tblGrid>')
    for ri, r in enumerate(rows):
        cells = []
        for ci, c in enumerate(r):
            a = aligns[ci] if aligns else None
            is_head = (ri == 0 and header)
            zebra = None
            if not is_head and zebra_on:
                zebra = F_GRAY if ri % 2 == 0 else None
            cb = (first_col_bold and ci == 0 and not is_head)
            cells.append(cell(c, header=is_head, fill=zebra, align=a, bold=cb, sz=sz))
        out.append("<w:tr>" + "".join(cells) + "</w:tr>")
    out.append("</w:tbl>")
    out.append(p("", after=80))
    return "".join(out)

# ============================ CAJAS DE COLOR (CALLOUTS) ============================
CALL = {
 "insight": (F_NAVY, NAVY, "INSIGHT CLAVE"),
 "oportunidad": (F_GREEN, A_GREEN, "OPORTUNIDAD"),
 "riesgo": (F_RED, A_RED, "ALERTA"),
 "recomendacion": (F_MUST, GOLD, "RECOMENDACIÓN"),
 "decision": (F_GRAY, A_GRAY, "NOTA"),
 "dato": (F_TEAL, A_TEAL, "DATO CLAVE"),
}
def callout(text, kind="insight", label=None):
    fill, accent, deflabel = CALL[kind]
    lab = label or deflabel
    borders = ('<w:tblBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{accent}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        '</w:tblBorders>')
    tcpr = (f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
            '<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="160" w:type="dxa"/>'
            '<w:bottom w:w="100" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner = (f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>{runs(lab, color=accent, bold=True, sz="18")}</w:p>'
             f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs(text)}</w:p>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{borders}</w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>' + p("", after=80))

def banner(title, subtitle_lines):
    tcpr = (f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{NAVY}"/>'
            '<w:tcMar><w:top w:w="400" w:type="dxa"/><w:left w:w="240" w:type="dxa"/>'
            '<w:bottom w:w="400" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner = f'<w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr>{runs(title, color="FFFFFF", bold=True, sz="50")}</w:p>'
    for i, s in enumerate(subtitle_lines):
        col = MUSTARD if i == 0 else "D6DCE5"
        szv = "30" if i == 0 else "22"
        inner += f'<w:p><w:pPr><w:spacing w:after="60"/><w:jc w:val="center"/></w:pPr>{runs(s, color=col, bold=(i==0), sz=szv)}</w:p>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>')

def section_band(text):
    """Banda de color para abrir una estrategia/sección destacada."""
    tcpr = (f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{NAVY}"/>'
            '<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/>'
            '<w:bottom w:w="120" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner = f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs(text, color="FFFFFF", bold=True, sz="26")}</w:p>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>' + p("", after=60))

def toc():
    return ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + runs("Contenido") + '</w:p>'
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:i/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">Haga clic derecho aqui y elija "Actualizar campos" para generar el indice con numeros de pagina.</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')

# ============================ CONTENIDO ============================
C = []


# ---------- PORTADA ----------
C.append('<w:p><w:pPr><w:spacing w:before="300" w:after="160"/></w:pPr></w:p>')
C.append(banner("PLAN DE MARKETING",
   ["Complejo Deportivo y Academia de Fútbol Andrés Chitiva",
    "Segmentación · Estrategias · Comercial · Digital · Contenidos · Presupuesto · KPIs",
    "Pachuca, Hidalgo, México  ·  Ciclo Agosto 2026 – Julio 2027"]))
C.append(spacer(220))
C.append(p("Documento profesional preparado para Dirección, equipo de trabajo, socios e inversionistas.", align="center", color=NAVY, bold=True, sz="24"))
C.append(p("Lenguaje claro · Decisiones accionables · Métricas verificables", align="center", color="595959", sz="20"))
C.append(spacer(200))
C.append(callout("Este plan traduce la estrategia del proyecto en acciones concretas, con responsables, presupuesto, entregables y metas medibles mes a mes. Está escrito para entenderse sin tecnicismos y para ejecutarse de inmediato rumbo al lanzamiento de agosto de 2026.", kind="dato", label="CÓMO LEER ESTE DOCUMENTO"))
C.append(p("Documento confidencial — Distribución restringida a las partes autorizadas.", align="center", color="808080", sz="16"))
C.append(page_break())

# ---------- ÍNDICE ----------
C.append(toc())
C.append(page_break())

# ============================================================
# 1. RESUMEN EJECUTIVO
# ============================================================
C.append(h1("1. Resumen Ejecutivo"))
C.append(p("El Proyecto Andrés Chitiva reúne dos negocios complementarios en Pachuca, Hidalgo: una **Academia de Fútbol de formación premium** y un **Complejo Deportivo multidisciplinario**. Ambos operan hoy en etapa piloto cerrada y abren oficialmente al público en la **segunda quincena de agosto de 2026**. Este Plan de Marketing define cómo conseguir y retener clientes para las dos marcas, con un enfoque realista al presupuesto disponible y máximo aprovechamiento de los activos que ya existen."))
C.append(p("La apuesta central es sencilla de explicar: **hay un espacio premium vacío en el mercado y nosotros lo vamos a ocupar**. Mientras la competencia cobra $400–600 por “solo cancha”, las familias de mayor poder adquisitivo están dispuestas a pagar más por una verdadera formación, y no encuentran dónde. Andrés Chitiva, a $1,200, ocupa ese lugar con método, instalaciones nuevas y acompañamiento humano."))
C.append(callout("Un alumno vale ~$15,900 el primer año y ~$30,000 en dos años (LTV). Adquirir uno cuesta un estimado de $1,000–1,500 (CAC). La relación LTV:CAC ≈ 10:1 es excelente (lo sano es 3:1). Conclusión: el presupuesto limita la velocidad de crecimiento, no la rentabilidad.", kind="dato"))
C.append(h3("Lo que este plan entrega"))
C.append(table([
  ["Bloque","Qué encontrará"],
  ["Audiencias","5 buyer personas accionables (Academia y Complejo) con sus motivaciones y objeciones"],
  ["Propuesta de valor","Mensaje central + mensajes y CTA específicos por segmento"],
  ["Estrategias","10 estrategias de marketing con plan de acción, actividades, entregables, KPIs y presupuesto"],
  ["Comercial","Embudo, manual de ventas, manejo de objeciones y scripts (WhatsApp y presencial)"],
  ["Digital y Contenidos","Canales, estructura de pauta, pilares de contenido y calendarios"],
  ["Presupuesto","Asignación estimada por canal (mes base y mes de lanzamiento)"],
  ["KPIs","Tablero con metas mes a mes (agosto 2026 – julio 2027)"],
  ["Medición","Sistema de reportes, cadencia y plantillas de seguimiento"],
], aligns=["left","left"], widths=[0.26,0.74], first_col_bold=True))
C.append(h3("Meta del ciclo (12 meses)"))
C.append(table([
  ["Indicador","Hoy","Meta 12 meses","Escenario ideal"],
  ["Alumnos de la Academia","~30","90","136 (cupo lleno)"],
  ["Ingreso recurrente mensual","$36,000","$108,000","$163,200"],
  ["Comunidad en redes (suma)","~200","~2,200","—"],
  ["Posicionamiento","Cerrado / desconocido","Referente premium de Pachuca","Líder de categoría"],
], aligns=["left","center","center","center"], widths=[0.34,0.2,0.24,0.22], first_col_bold=True))
C.append(page_break())

# ============================================================
# 2. OBJETIVOS DE MARKETING
# ============================================================
C.append(h1("2. Objetivos de Marketing"))
C.append(p("Los objetivos están formulados para ser **específicos, medibles y con fecha** (criterio SMART). Son la brújula de todo el plan: cada estrategia existe para mover al menos uno de ellos."))
C.append(table([
  ["#","Objetivo","Meta","Plazo"],
  ["O1","Lanzar la Academia con fuerza","Cerrar agosto con ~45 alumnos (+15) y 40–60 registros a Clase Muestra","Agosto 2026"],
  ["O2","Triplicar la base de alumnos","Pasar de 30 a 90 alumnos","12 meses (a julio 2027)"],
  ["O3","Validar la adquisición en frío","CPL < $130 y CAC < $1,500 confirmados con datos reales","Meses 1–3"],
  ["O4","Construir marca y comunidad","Crecer de ~200 a ~2,200 seguidores (suma de redes)","12 meses"],
  ["O5","Retener y demostrar valor","Retención mensual > 90% y NPS > 60","Continuo"],
  ["O6","Activar el Complejo","Subir ocupación de horas valle (mañanas) y sumar equipos a ligas","Continuo"],
], aligns=["center","left","left","center"], widths=[0.07,0.31,0.42,0.2], first_col_bold=True))
C.append(callout("Prioridad #1 de los primeros 3 meses: NO escalar el gasto en pauta hasta validar que el costo de adquisición (CAC) es sano. Primero medimos en pequeño, luego invertimos con confianza. Esta disciplina protege el dinero del proyecto.", kind="recomendacion"))
C.append(page_break())


# ============================================================
# 3. SEGMENTACIÓN Y BUYER PERSONAS
# ============================================================
C.append(h1("3. Segmentación de Audiencias y Buyer Personas"))
C.append(p("No le hablamos a “todos”. Hablamos a personas concretas con motivaciones, miedos y formas de decidir distintas. La segmentación nos permite enviar el **mensaje correcto, a la persona correcta, en el canal correcto** — y dejar de gastar esfuerzo en quien no es nuestro cliente."))
C.append(p("Separamos el mercado en dos universos con lógicas de compra diferentes, y dentro de cada uno definimos sus perfiles clave:"))
C.append(table([
  ["Universo","Quién decide","Segmentos","Lógica de compra"],
  ["⚽ Academia","Papá / mamá","1) Desarrollo integral  2) Alto rendimiento","Recurrente (mensualidad); emocional + racional"],
  ["🏟️ Complejo","El usuario directo / la empresa","3) Liga amateur adulta  4) Empresas (B2B)  5) Familia recreativa (pádel/funcional)","Transaccional; por ocasión y conveniencia"],
], aligns=["left","left","left","left"], widths=[0.16,0.18,0.34,0.32], first_col_bold=True))
C.append(callout("Regla de oro de segmentación: NO perseguimos al cliente de $500. Le hablamos al papá que quiere lo mejor para su hijo y no encuentra dónde pagar por ello. Competir por precio destruiría el posicionamiento premium.", kind="insight"))
C.append(p("A continuación, cinco buyer personas listos para usar por los equipos de contenido, pauta y ventas."))
C.append(spacer(60))

def persona(nombre, etiqueta, color_band, datos, motivaciones, objeciones, donde, mensaje):
    out = []
    out.append(section_band(f"{etiqueta}  —  “{nombre}”"))
    out.append(table([["Perfil", "Detalle"]] + datos, aligns=["left","left"], widths=[0.26,0.74], first_col_bold=True))
    out.append(h3("Qué busca de verdad (motivaciones)"))
    for m in motivaciones:
        out.append(bullet(m))
    out.append(h3("Qué lo frena (objeciones)"))
    for o in objeciones:
        out.append(bullet(o))
    out.append(h3("Dónde lo encontramos"))
    out.append(p(donde))
    out.append(callout(mensaje, kind="recomendacion", label="MENSAJE QUE LE FUNCIONA"))
    return "".join(out)

# Persona 1
C.append(persona(
  "Carla, la mamá que quiere formar a su hijo", "ACADEMIA · Segmento 1: Desarrollo integral", NAVY,
  [
    ["Edad / perfil","34–46 años · NSE medio-alto/alto · profesionista o empresaria"],
    ["Hijo(a)","Niño/niña de 6–12 años; busca actividad extraescolar con propósito"],
    ["Vive en","Pachuca / Mineral de la Reforma, zonas residenciales"],
    ["Cómo decide","Investiga, pregunta a otras mamás, visita antes de inscribir"],
    ["Frase típica","“Quiero que aprenda disciplina y haga amigos sanos, no solo que patee un balón.”"],
  ],
  ["Que su hijo gane disciplina, hábitos y confianza.",
   "Un entorno seguro, profesional y bien cuidado.",
   "Acompañamiento humano (entrenadores que conocen a su hijo, psicóloga deportiva).",
   "Sentir que su inversión “se nota” en el comportamiento y ánimo del niño."],
  ["“¿No es muy caro comparado con la escuelita de la colonia?”",
   "Distancia y tráfico para llevar/recoger.",
   "“Son nuevos, ¿de verdad cumplen lo que prometen?”"],
  "Instagram y Facebook (grupos de mamás de Pachuca), recomendación de boca en boca, colegios de NSE medio-alto, y la base de clientes de aliados como Mazda Pachuca.",
  "“No solo formamos futbolistas, formamos personas.” Habla de disciplina, valores y acompañamiento, con la psicóloga deportiva y los reportes de avance como prueba. Invítala a vivir una Clase Muestra: ahí se vende sola."))
C.append(page_break())

# Persona 2
C.append(persona(
  "Roberto, el papá que sueña con un hijo futbolista", "ACADEMIA · Segmento 2: Alto rendimiento", NAVY,
  [
    ["Edad / perfil","36–50 años · NSE medio-alto/alto · apasionado del fútbol"],
    ["Hijo(a)","Niño/joven de 9–17 años con talento y ambición competitiva"],
    ["Vive en","Pachuca, Mineral de la Reforma e incluso otras ciudades"],
    ["Cómo decide","Compara metodología, resultados y respaldo; valora el nombre de Andrés"],
    ["Frase típica","“Quiero que lo formen en serio, con método, no improvisando.”"],
  ],
  ["Proyección real: que su hijo mejore técnica, física y tácticamente.",
   "Metodología estructurada y evaluación individual continua.",
   "Respaldo de alguien que “sí jugó” (Andrés Chitiva) y cuerpo técnico capacitado.",
   "Evidencia: video, datos, highlights y seguimiento del progreso."],
  ["“¿El método de verdad lo lleva a otro nivel o es puro nombre?”",
   "Quiere ver resultados medibles, no promesas.",
   "Distancia (algunos vienen de fuera)."],
  "Instagram/TikTok (highlights y contenido de método), cuenta personal de Andrés, torneos y ligas, y referidos de otras familias de alto rendimiento.",
  "“Método profesional respaldado por Andrés Chitiva, con evaluación individual y cámaras para analizar cada jugada.” Muestra el sistema de highlights y el proof point de los alumnos que viajan de otras ciudades por la metodología."))
C.append(page_break())

# Persona 3
C.append(persona(
  "Memo, el capitán de la liga del fin de semana", "COMPLEJO · Segmento 3: Liga amateur adulta", A_TEAL,
  [
    ["Edad / perfil","25–45 años · organiza o juega en un equipo de amigos/compañeros"],
    ["Necesidad","Cancha de buena calidad, bien ubicada, con buen ambiente y horario nocturno"],
    ["Cómo decide","Precio por partido, disponibilidad, calidad de cancha y experiencia"],
    ["Frase típica","“Necesito una cancha buena, que prendan las luces y se juegue en serio.”"],
  ],
  ["Jugar en instalaciones nuevas y bien mantenidas.",
   "Organización confiable (horarios, arbitraje, tabla de posiciones).",
   "Ambiente y pertenencia: presumir el gol, los highlights del partido.",
   "Que sus compañeros la pasen bien y repitan."],
  ["“¿Hay cupo en el horario que nos sirve (noches/fines)?”",
   "Precio por partido y por temporada.",
   "Que la cancha esté en mal estado o mal organizada."],
  "Facebook (grupos locales de fútbol amateur y ligas), boca en boca entre equipos, y el propio Complejo con sus cámaras y highlights de partidos.",
  "“Tu liga, en instalaciones nuevas, con cámaras que graban tus mejores jugadas para descargar y presumir.” Vende el ambiente, la organización seria y el diferenciador de los highlights; usa las noches y fines de semana como gancho."))
C.append(page_break())

# Persona 4
C.append(persona(
  "Empresa / Recursos Humanos (decisor B2B)", "COMPLEJO · Segmento 4: Empresas", A_TEAL,
  [
    ["Perfil","Responsable de RR.HH., bienestar o un líder de área"],
    ["Necesidad","Actividades de integración, bienestar y deporte para colaboradores"],
    ["Cómo decide","Propuesta clara, factura, logística sencilla y buen trato"],
    ["Frase típica","“Busco algo diferente para integrar al equipo y cuidar su salud.”"],
  ],
  ["Integración y clima laboral (team building).",
   "Bienestar y salud de sus colaboradores.",
   "Una experiencia organizada “llave en mano”, sin complicaciones.",
   "Buena relación costo-beneficio y factura."],
  ["“¿Me resuelven todo o tengo que organizarlo yo?”",
   "Disponibilidad en horarios laborales (mañanas/medio día = horas valle del Complejo).",
   "Requisitos administrativos (facturación, seguros)."],
  "Contacto directo / venta consultiva, LinkedIn y Facebook, alianzas con cámaras empresariales, y la red de contactos de Andrés y patrocinadores como Mazda Pachuca.",
  "“Ligas y torneos corporativos, eventos de integración y entrenamiento funcional para tu equipo, todo organizado por nosotros.” Es un ticket grande que además llena las horas muertas del Complejo: doble beneficio."))
C.append(page_break())

# Persona 5
C.append(persona(
  "Ana y Luis, la pareja que cuida su salud", "COMPLEJO · Segmento 5: Familia recreativa (pádel/funcional)", A_TEAL,
  [
    ["Edad / perfil","28–50 años · activos, buscan deporte social y bienestar"],
    ["Necesidad","Pádel, entrenamiento funcional y planes en pareja/familia"],
    ["Cómo decide","Cercanía, ambiente, precio del paquete y disponibilidad"],
    ["Frase típica","“Queremos hacer ejercicio juntos y que sea divertido.”"],
  ],
  ["Hacer deporte de forma social y divertida.",
   "Bienestar físico sin que se sienta una obligación.",
   "Flexibilidad de horarios y planes (parejas, estudiantes).",
   "Un lugar nuevo y agradable “donde pasan las cosas”."],
  ["“¿Hay cancha de pádel libre cuando puedo ir?” (solo hay 1 cancha).",
   "Precio por hora / mensualidad del funcional.",
   "Constancia: necesitan motivación para repetir."],
  "Instagram/TikTok (contenido aspiracional y de ambiente), Facebook local, promociones de temporada (Año Nuevo, San Valentín parejas) y la cuponera de aliados.",
  "“Entrena funcional y juega pádel en instalaciones nuevas, con planes para parejas y estudiantes.” Aprovecha campañas de temporada (propósitos de Año Nuevo, pádel en pareja) y horarios valle con promociones."))
C.append(page_break())


# ============================================================
# 4. PROPUESTA DE VALOR Y MENSAJES POR SEGMENTO
# ============================================================
C.append(h1("4. Propuesta de Valor Refinada y Mensajes por Segmento"))
C.append(h2("4.1 Propuesta de valor central"))
C.append(callout("Formamos al jugador y a la persona. Rigor profesional dentro de la cancha y desarrollo humano fuera de ella, con método, instalaciones nuevas y acompañamiento cercano — para el papá que quiere lo mejor para su hijo.", kind="decision", label="PROMESA CENTRAL DE LA ACADEMIA"))
C.append(p("**Manifiesto:** “Formamos futbolistas inteligentes, apasionados y con valores.”  ·  **Tagline de apoyo:** “Método profesional. Formación de vida.”"))
C.append(h3("Mapa de valor (qué entregamos vs. qué compra el cliente)"))
C.append(table([
  ["Lo que ofrecemos (racional)","En lo que se convierte para el cliente (emocional)"],
  ["Metodología de 4 pilares + evaluación individual","“Mi hijo mejora de verdad y alguien lo está viendo a él”"],
  ["Psicóloga deportiva y formación en valores","“Va a ser mejor persona, no solo mejor jugador”"],
  ["Instalaciones nuevas + cámaras con highlights","“Está en un lugar seguro y puedo ver su progreso”"],
  ["Respaldo de Andrés Chitiva y cuerpo técnico","“Lo están formando profesionales que sí saben”"],
  ["Comunidad y acompañamiento cercano","“Mi familia pertenece a algo que vale la pena”"],
], aligns=["left","left"], widths=[0.46,0.54]))
C.append(callout("El trabajo del marketing no es bajar el precio: es hacer que el papá SIENTA por qué vale 2–3 veces más. Si no percibe el valor, el precio premium se vuelve la principal objeción. Todo el contenido y la venta deben “mostrar el valor”, no “anunciar el precio”.", kind="insight"))

C.append(h2("4.2 Mensajes y llamados a la acción por segmento"))
C.append(p("Cada segmento recibe un ángulo distinto. Misma marca, diferente puerta de entrada:"))
C.append(table([
  ["Segmento","Mensaje principal","Prueba que lo respalda","Llamado a la acción (CTA)"],
  ["Desarrollo integral","“Más que fútbol: disciplina, valores y confianza para tu hijo.”","Psicóloga deportiva, reportes de avance, ambiente seguro","“Agenda una Clase Muestra gratis”"],
  ["Alto rendimiento","“Método profesional respaldado por Andrés Chitiva.”","Evaluación individual, cámaras/highlights, alumnos que viajan","“Vive una Clase Muestra y compruébalo”"],
  ["Liga amateur","“Tu liga en instalaciones nuevas, con highlights de tus jugadas.”","Canchas nuevas, cámaras, organización seria","“Aparta tu cancha / inscribe a tu equipo”"],
  ["Empresas (B2B)","“Integra y cuida a tu equipo con deporte, llave en mano.”","Funcional, ligas corporativas, eventos organizados","“Solicita una propuesta para tu empresa”"],
  ["Familia recreativa","“Entrena y juega pádel en pareja, en un lugar nuevo.”","Planes parejas/estudiante, ambiente, cercanía","“Reserva tu hora / prueba el funcional”"],
], aligns=["left","left","left","left"], widths=[0.16,0.3,0.3,0.24], first_col_bold=True, sz="18"))
C.append(callout("La nutrición y los servicios médicos solo se comunican como beneficio incluido cuando operen al 100%. Hasta entonces se manejan como “Próximamente”. Prometer lo que aún no se entrega al nivel premium destruiría la confianza.", kind="riesgo", label="REGLA DE PROMESA HONESTA"))
C.append(page_break())


# ============================================================
# 5. ESTRATEGIAS DE MARKETING (10)
# ============================================================
C.append(h1("5. Estrategias de Marketing"))
C.append(p("A continuación, **10 estrategias** de marketing. Superan el mínimo solicitado y cubren todo el embudo: atraer, convertir, retener y expandir. Cada una incluye objetivo, plan de acción paso a paso, actividades específicas, entregables concretos, KPIs con metas, responsable y presupuesto."))
C.append(table([
  ["#","Estrategia","Para qué sirve"],
  ["E1","Contenido orgánico y Highlights","Atraer y construir marca casi sin costo"],
  ["E2","Pauta digital (Performance)","Generar leads medibles y validar el CAC"],
  ["E3","Comercial y conversión (Clase Muestra)","Convertir interesados en alumnos inscritos"],
  ["E4","Lanzamiento de Agosto 2026","Abrir con fuerza: PR + leads + inscripciones"],
  ["E5","Alianzas estratégicas","Acceder a audiencias que ya confían en otros"],
  ["E6","Fidelización y retención (CX)","Que las familias se queden y refieran"],
  ["E7","Referidos y Generación Fundadora","Crecer con el boca a boca de clientes felices"],
  ["E8","Activación del Complejo","Llenar horas valle, ligas y B2B"],
  ["E9","Relaciones Públicas y marca Andrés","Credibilidad y prensa ganada"],
  ["E10","CRM y automatización","No perder ningún lead y dar seguimiento"],
], aligns=["center","left","left"], widths=[0.08,0.34,0.58], first_col_bold=True))
C.append(page_break())

def strategy(num, title, objetivo, publico, responsable, presupuesto, plan, actividades, entregables, kpis):
    out = []
    out.append(section_band(f"Estrategia {num} — {title}"))
    out.append(table([
        ["Objetivo", objetivo],
        ["Público / segmento", publico],
        ["Responsable", responsable],
        ["Presupuesto", presupuesto],
    ], header=False, aligns=["left","left"], widths=[0.22,0.78], first_col_bold=True, zebra_on=False))
    out.append(h3("Plan de acción (paso a paso)"))
    for i, step in enumerate(plan, 1):
        out.append(numbered(f"**{i}.** {step}"))
    out.append(h3("Actividades específicas"))
    for a in actividades:
        out.append(bullet(a))
    out.append(h3("Entregables concretos"))
    for e in entregables:
        out.append(bullet(e))
    out.append(h3("KPIs y metas"))
    out.append(table([["Indicador", "Meta"]] + kpis, aligns=["left","left"], widths=[0.55,0.45], first_col_bold=True))
    return "".join(out)

# ---------- E1 ----------
C.append(strategy(
  "1", "Contenido Orgánico y Highlights (Motor de Marca)",
  "Crecer audiencia y construir credibilidad premium con bajo costo, alimentando el embudo de ventas.",
  "Academia (principal) y Complejo · todos los segmentos en la etapa de atracción",
  "Community Manager + Productor Audiovisual + Diseñador",
  "Bajo (recurso humano ya existente). Producción incluida en nómina del equipo.",
  ["Encender el flujo de **highlights con las cámaras**: sistema simple para generar y entregar un clip por alumno/categoría cada semana.",
   "Definir y calendarizar los **pilares de contenido** (ver capítulo 8) con 4 reels + 4 posts por semana.",
   "Producir los **assets base** antes de agosto: video-tour de instalaciones, reels de método, testimonios de familias piloto.",
   "Aplicar la **señal visual por marca** (Academia vs. Complejo) en cada pieza para evitar confusión.",
   "Coordinar con Andrés el **reposteo desde su cuenta personal** para amplificar sin costo.",
   "Medir semanalmente qué formato funciona y duplicar lo que da resultado."],
  ["Reel semanal de highlight de alumno (cámaras) — el motor #1 de alcance orgánico.",
   "Reel de tip de método / consejo de entrenador o psicóloga (autoridad).",
   "Carrusel de testimonio de familia (confianza).",
   "Contenido de ambiente del Complejo: ligas, pádel, funcional.",
   "Reaprovechar cada reel en TikTok y Facebook (un esfuerzo, tres canales)."],
  ["Parrilla de contenido mensual (16 reels + 16 posts).",
   "Banco de assets: video-tour, 5+ reels de método, 3+ testimonios en video.",
   "Sistema de entrega de highlights al papá (semanal).",
   "Manual de marca visual de 1 página (cómo se ve Academia vs. Complejo)."],
  [["Crecimiento de seguidores (suma de redes)","~200 → ~2,200 en 12 meses"],
   ["Highlights entregados","≥ 1 por alumno por semana"],
   ["Alcance e interacciones","Tendencia mensual al alza"],
   ["Clics a WhatsApp desde redes","Crecientes mes a mes"],
   ["Piezas publicadas","4 reels + 4 posts por semana"]]))
C.append(page_break())

# ---------- E2 ----------
C.append(strategy(
  "2", "Pauta Digital de Resultados (Performance Marketing)",
  "Generar leads medibles hacia la Clase Muestra y validar el costo de adquisición (CAC) antes de escalar.",
  "Academia: papás de desarrollo integral y alto rendimiento (Pachuca y alrededores)",
  "Marketing (gestión de campañas) + Community (respuesta)",
  "$3,500–5,000 MXN/mes (Meta + Google), con refuerzo extra en agosto.",
  ["Concentrar la pauta en **Meta (Instagram/Facebook)** como canal principal; no dispersar el presupuesto chico.",
   "Estructurar 3 niveles: **Alcance** (amplificar mejores reels) → **Retargeting** (a quien interactuó) → **Conversión** (mensajes/registro a Clase Muestra).",
   "Asignar una pequeña parte a **Google Search** para capturar alta intención (“academia de futbol Pachuca”).",
   "Llevar todo el tráfico a una **landing de registro** con píxel para poder retargetear.",
   "Medir CPL y CAC por canal **semanalmente** durante los meses 1–3 (fase de validación).",
   "Solo escalar inversión en las campañas con **CAC probado** y construir el caso de negocio para Finanzas."],
  ["Campaña de video views para construir audiencias (frío).",
   "Campaña de retargeting con testimonios + invitación a Clase Muestra (tibio).",
   "Campaña de mensajes a WhatsApp / registro a Clase Muestra (caliente).",
   "Campaña de Google Search de alta intención (palabras clave locales).",
   "Refuerzo especial de pauta para el evento de lanzamiento de agosto."],
  ["Estructura de campañas documentada (públicos, creativos, presupuestos).",
   "Landing page de registro a Clase Muestra con píxel activo.",
   "Reporte semanal de CPL/CAC por canal.",
   "Biblioteca de creativos probados (lo que funciona vs. lo que no)."],
  [["Costo por lead (CPL)","< $130"],
   ["Costo por inscripción (CAC)","< $1,500 (validar meses 1–3)"],
   ["Relación LTV:CAC","≥ 3:1 (objetivo ~10:1)"],
   ["Leads generados","~30–50 por mes"],
   ["% de inscripciones por fuente","Medido (pauta / referido / orgánico)"]]))
C.append(page_break())


# ---------- E3 ----------
C.append(strategy(
  "3", "Comercial y Conversión (Clase Muestra como punto de conversión)",
  "Convertir el mayor porcentaje posible de interesados en alumnos inscritos, colocando la experiencia antes del precio.",
  "Todos los leads de la Academia (presenciales, digitales y “solo precio”)",
  "Recepción / Ventas + Cuerpo técnico (Clase Muestra)",
  "Bajo (operativo). Incluye material de bienvenida ya existente.",
  ["Responder **rápido**: auto-respuesta instantánea + atención humana en menos de 15–30 minutos.",
   "Hacer **descubrimiento** (motivación, zona, edad) para personalizar el pitch.",
   "Invitar SIEMPRE a la **Clase Muestra gratis** (o entregar valor digital si es a distancia).",
   "Presentar el **precio solo después** de que el papá vivió la experiencia.",
   "Cerrar con **beneficios + cupo limitado** y manejar objeciones con el guion.",
   "Dejar siempre un **siguiente paso agendado** y registrar todo en Clupik."],
  ["Capacitar al equipo con el manual de ventas y los scripts (capítulo 11).",
   "Configurar WhatsApp Business con auto-reply, respuestas rápidas y etiquetas de pipeline.",
   "Definir el formato de la Clase Muestra con los entrenadores.",
   "Activar recordatorios automáticos de la Clase Muestra (reduce ausencias).",
   "Ritual de bienvenida premium al inscribir (kit + primer día especial)."],
  ["Manual comercial + scripts impresos y cargados en WhatsApp.",
   "Pipeline de ventas configurado en Clupik (etapas del embudo).",
   "Formato y guion de la Clase Muestra.",
   "Checklist de cierre e inscripción."],
  [["Tiempo de primera respuesta","< 15–30 min"],
   ["% de leads invitados a Clase Muestra","> 80%"],
   ["Show-rate (asistencia a Clase Muestra)","30–40% de leads"],
   ["% de cierre (inscritos / asistentes)","40–50%"],
   ["Leads sin siguiente paso agendado","→ 0"]]))
C.append(page_break())

# ---------- E4 ----------
C.append(strategy(
  "4", "Lanzamiento de Agosto 2026 (Apertura con propósito comercial)",
  "Abrir oficialmente generando expectativa, prensa, base de datos de leads e inscripciones — no solo una fiesta.",
  "Audiencia tibia + mercado frío + medios + autoridad + aliados",
  "Marketing (coordinación) + Andrés (PR) + Ventas + Audiovisual",
  "Presupuesto especial de lanzamiento (refuerzo de pauta + impresos + logística del evento).",
  ["**Pre-lanzamiento (jul–ago):** campaña de intriga (“algo grande viene a Pachuca”), cuenta regresiva y detrás de cámara.",
   "Abrir **lista de espera / pre-registro** para empezar a llenar el embudo antes de abrir.",
   "Cerrar la **Generación Fundadora** (familias piloto) y activar sus testimonios y referidos.",
   "Producir TODOS los assets y coordinar **invitados**: medios, autoridad municipal, exjugadores, aliados.",
   "**Evento de inauguración** con recorrido, Clase Muestra, activaciones de patrocinadores y mesa de inscripción.",
   "Capturar **cada asistente como lead con registro QR** y dar seguimiento posterior con retargeting."],
  ["Campaña de intriga y cuenta regresiva en redes.",
   "Pre-venta a familias tibias con beneficios de Fundador/lanzamiento.",
   "Kit de prensa + boletín de medios (apoyado en Andrés y exjugadores).",
   "Activaciones de Teqball, Voit y stand de Mazda Pachuca.",
   "Cobertura en vivo del evento (reels, testimonios) = semanas de contenido."],
  ["Plan de evento con programa, logística y lista de invitados.",
   "Sistema de registro QR conectado al CRM.",
   "Kit de prensa y boletín listos.",
   "Calendario de pre-lanzamiento (5 semanas) con responsables."],
  [["Registros a Clase Muestra en agosto","40–60"],
   ["Alumnos al cierre de agosto","~45 (+15)"],
   ["Leads capturados en el evento (QR)","Todos los asistentes"],
   ["Menciones / notas en medios locales","≥ 3"],
   ["Contenido generado del evento","≥ 2 semanas de parrilla"]]))
C.append(page_break())

# ---------- E5 ----------
C.append(strategy(
  "5", "Alianzas Estratégicas (Audiencias prestadas)",
  "Acceder a audiencias que ya confían en otros y llenar las horas valle del Complejo, sin gastar en pauta.",
  "Familias NSE medio-alto (Academia) y demanda corporativa/escolar (Complejo)",
  "Marketing + Andrés (relaciones) + Dirección (acuerdos)",
  "Bajo / intercambio (barter). Activaciones cubiertas por los aliados.",
  ["**Exprimir patrocinadores actuales:** formalizar qué da y recibe cada uno (Voit, Teqball, Mazda Pachuca).",
   "Activar la **promoción cruzada con Mazda Pachuca**: su base de clientes es el target exacto de la Academia.",
   "Abrir conversaciones con **escuelas y colegios** de NSE objetivo (alumnos + mañanas del Complejo).",
   "Desarrollar el frente **empresas (B2B)**: ligas corporativas, eventos, funcional para empleados.",
   "Evolucionar la **cuponera de aliados** a doble vía: que ellos también promuevan la Academia a sus clientes.",
   "Sumar **medios locales** vía el reconocimiento de Andrés para prensa ganada."],
  ["Carpeta de patrocinio con beneficios por nivel.",
   "Activación conjunta con Mazda (presencia de vehículo, rifas, contenido).",
   "Programa “Chitiva en tu escuela” (clínicas y torneos inter-escolares).",
   "Propuesta comercial B2B para empresas (one-pager).",
   "Acuerdos de referido recíproco con comercios aliados."],
  ["Acuerdos formalizados con Voit, Teqball y Mazda.",
   "Pipeline de escuelas y empresas objetivo.",
   "One-pager comercial B2B.",
   "Cuponera de doble vía documentada."],
  [["Activaciones con patrocinadores","≥ 1 por trimestre"],
   ["Escuelas en conversación","≥ 3 (post-piloto)"],
   ["Eventos / ligas corporativas","Crecientes"],
   ["Leads provenientes de alianzas","Medido y creciente"],
   ["Ocupación de horas valle del Complejo","Al alza"]]))
C.append(page_break())

# ---------- E6 ----------
C.append(strategy(
  "6", "Fidelización y Retención (Customer Experience)",
  "Que las familias se queden (LTV ~$30,000) y se conviertan en la mejor prueba de resultados del proyecto.",
  "Familias ya inscritas en la Academia",
  "Cuerpo técnico + Recepción + Marketing (contenido de comunidad)",
  "Bajo (usa Clupik y kit existente).",
  ["Estructurar un **onboarding premium**: bienvenida personalizada + primer día especial + orientación al papá.",
   "Implementar la **demostración de valor mensual**: pruebas físicas, reportes en Clupik, highlight semanal y contacto del entrenador.",
   "Construir **comunidad**: torneos internos, convivencias, ceremonia de fin de ciclo, identidad “Generación Fundadora”.",
   "Activar la **alerta temprana de fuga**: caída de asistencia en Clupik dispara contacto proactivo del entrenador.",
   "Medir **NPS** a los 30 días y luego trimestral; aplicar encuesta de salida a quien se va.",
   "Mitigar distancia/tráfico como retención: carpool entre familias y cámaras para ver al hijo en remoto."],
  ["Kit de bienvenida + ritual de primer día.",
   "Reporte de avance mensual visible para el papá.",
   "Envío semanal de highlights del hijo.",
   "Calendario de momentos de comunidad.",
   "Protocolo de retención proactiva por asistencia."],
  ["Flujo de onboarding documentado.",
   "Tablero de retención y asistencia en Clupik.",
   "Cuestionario NPS y de salida.",
   "Calendario anual de eventos de comunidad."],
  [["Retención mensual","> 90%"],
   ["NPS","> 60"],
   ["Tasa de renovación por ciclo","Alta y creciente"],
   ["Asistencia promedio","Monitoreada (señal temprana)"],
   ["% de familias que refieren","Creciente"]]))
C.append(page_break())


# ---------- E7 ----------
C.append(strategy(
  "7", "Referidos y Generación Fundadora (Boca a boca premium)",
  "Reducir drásticamente el costo de adquisición usando a las familias satisfechas como vendedoras.",
  "Familias actuales y Generación Fundadora (piloto)",
  "Marketing + Recepción / Ventas + Finanzas (aprobación de beneficios)",
  "Bajo (beneficio puntual por referido, no descuento permanente).",
  ["Cerrar la **Generación Fundadora**: reingreso prioritario, distintivo de pertenencia y precio protegido temporal.",
   "Pactar la **contraprestación**: cada familia fundadora aporta un testimonio en video + al menos 1 referido.",
   "Diseñar el **programa de referidos**: beneficio puntual para ambas familias (ej. crédito de 1 mes), nunca recurrente.",
   "Hacer **fácil referir**: mensaje y materiales listos para que el papá comparta con un clic.",
   "Reconocer públicamente a quienes refieren (estatus, pertenencia premium).",
   "Medir el **% de inscripciones por referido** y optimizar el incentivo."],
  ["Paquete Fundador (beneficios + distintivo).",
   "Mecánica de referidos aprobada por Finanzas.",
   "Kit para referir (mensaje + imagen compartible).",
   "Sesión de grabación de testimonios de fundadores."],
  ["Programa Generación Fundadora documentado.",
   "Programa de referidos con mecánica y materiales.",
   "Banco de testimonios en video.",
   "Registro de referidos en Clupik."],
  [["Testimonios de fundadores grabados","≥ 10"],
   ["Referidos por familia fundadora","≥ 1"],
   ["% de inscripciones por referido","Creciente"],
   ["CAC de inscripciones por referido","Muy bajo (cercano a $0)"],
   ["Familias con distintivo Fundador","Todas las del piloto que renueven"]]))
C.append(page_break())

# ---------- E8 ----------
C.append(strategy(
  "8", "Activación del Complejo (Llenar las horas que hoy están vacías)",
  "Convertir la capacidad ociosa (sobre todo mañanas) en ingreso, con públicos y mensajes por franja horaria.",
  "Ligas amateur, empresas, escuelas y familia recreativa (pádel/funcional)",
  "Marketing + Recepción / Ventas + Operación del Complejo",
  "Pauta compartida con E2 + impresos locales según necesidad.",
  ["Segmentar la comunicación **por franja horaria**: mañanas (escuelas, empresas, funcional) vs. noches (ligas after-office, torneos).",
   "Empaquetar y promover las **ligas** (inscripción + cobro por partido) como motor de caja recurrente.",
   "Lanzar **promociones de horas valle** (pádel/funcional) para llenar mañanas y medio día.",
   "Activar el frente **B2B y escolar** (ver E5) para ocupar mañanas con tickets grandes.",
   "Usar el contenido de **ambiente y highlights de ligas** para atraer nuevos equipos.",
   "Explorar a futuro **paquetes/membresías** y temporadas de ligas."],
  ["Calendario de ligas y torneos (con tabla de posiciones).",
   "Promos de temporada: Buen Fin, Año Nuevo (funcional/pádel), San Valentín (pádel parejas).",
   "Beneficio “Familia Chitiva”: horas de pádel valle para alumnos de la Academia.",
   "Contenido de ambiente del Complejo en redes.",
   "Campañas de reserva de cancha por WhatsApp."],
  ["Calendario comercial del Complejo.",
   "Tarifario y paquetes de ligas.",
   "Plan de promociones de horas valle.",
   "Reporte de ocupación por franja horaria."],
  [["Ocupación de horas valle (mañanas)","Al alza mes a mes"],
   ["Equipos / temporada en ligas","Creciente"],
   ["Utilización de la cancha de pádel","Alta en horas pico, valle promovida"],
   ["Eventos / empresas por mes","≥ 1 y creciente"],
   ["Ingreso del Complejo","Tendencia mensual al alza"]]))
C.append(page_break())

# ---------- E9 ----------
C.append(strategy(
  "9", "Relaciones Públicas y Marca Personal de Andrés",
  "Capitalizar el reconocimiento de Andrés Chitiva para ganar credibilidad y prensa sin costo de pauta.",
  "Medios locales, comunidad de Pachuca y target premium",
  "Andrés (vocería) + Marketing (coordinación de PR)",
  "Bajo (prensa ganada). Producción de contenido por el equipo.",
  ["Posicionar a Andrés como **director/arquitecto del método** (no como entrenador diario): protege la promesa vs. la entrega.",
   "Construir un **calendario de PR** apoyado en hitos: lanzamiento, torneos, alianzas, resultados de alumnos.",
   "Coordinar la **cuenta personal de Andrés** como altavoz que amplifica el contenido de marca.",
   "Aprovechar **exjugadores y autoridad municipal** como refuerzo de credibilidad en el lanzamiento.",
   "Generar **historias con valor de prensa**: alumnos que viajan por el método, la “Metodología Chitiva”, casos de éxito.",
   "Documentar la **metodología como IP** desde hoy (base de la futura franquicia)."],
  ["Boletines y kit de prensa para hitos clave.",
   "Agenda de apariciones y vocería de Andrés.",
   "Reposteo coordinado desde la cuenta personal.",
   "Registro sistemático de testimonios y casos de éxito.",
   "Expediente inicial de la “Metodología Chitiva”."],
  ["Calendario de PR del año.",
   "Kit de prensa reutilizable.",
   "Banco de casos de éxito documentados.",
   "Carpeta inicial de IP de la metodología."],
  [["Notas / menciones en medios","≥ 1 por trimestre (pico en lanzamiento)"],
   ["Reposteos desde la cuenta de Andrés","Frecuentes y coordinados"],
   ["Casos de éxito documentados","Crecientes"],
   ["Percepción de marca (encuestas)","“Premium, vigente, cumple”"],
   ["Avance del expediente de IP","Continuo"]]))
C.append(page_break())

# ---------- E10 ----------
C.append(strategy(
  "10", "CRM y Automatización (Que no se pierda ningún lead)",
  "Centralizar prospectos y clientes, automatizar seguimientos y dejar de operar en Excel.",
  "Todos los leads y familias (Academia y Complejo)",
  "Marketing + Recepción / Ventas",
  "Bajo (WhatsApp Business + Clupik ya disponibles); CRM ligero a evaluar.",
  ["**MVP inmediato (para agosto):** WhatsApp Business con etiquetas = etapas del embudo + respuestas rápidas (scripts) + auto-reply.",
   "Usar **Clupik** como base central: miembros, pagos, progreso, comunicación y pipeline de ventas.",
   "Configurar **automatizaciones prioritarias**: auto-reply, recordatorio al asesor por lead sin respuesta, confirmación de Clase Muestra.",
   "Crear una **secuencia de nurturing** para quien no cerró + lista de espera.",
   "Mantener la **base de datos limpia** y capturar siempre nombre, niño, edad, zona y motivación.",
   "Evaluar un **CRM ligero con integración de WhatsApp** cuando el volumen lo justifique."],
  ["Etiquetas de pipeline en WhatsApp Business.",
   "Respuestas rápidas con los scripts cargados.",
   "Automatización de recordatorios de Clase Muestra.",
   "Secuencia de seguimiento (nurture) para no-cierres.",
   "Tablero de leads y conversión."],
  ["WhatsApp Business configurado (etiquetas + scripts + auto-reply).",
   "Clupik configurado como CRM y pipeline.",
   "Secuencias de automatización activas.",
   "Base de datos de leads centralizada."],
  [["Leads registrados en CRM","100% (cero en Excel suelto)"],
   ["Tiempo de respuesta automatizado","Instantáneo (auto-reply)"],
   ["Leads sin seguimiento","→ 0"],
   ["Tasa de show a Clase Muestra","Mejorada por recordatorios"],
   ["Reactivación de no-cierres","Secuencia activa y medida"]]))
C.append(page_break())


# ============================================================
# 6. ESTRATEGIA COMERCIAL (detalle)
# ============================================================
C.append(h1("6. Estrategia Comercial"))
C.append(p("La estrategia comercial define **cómo convertimos interés en inscripciones**. El principio rector es uno: **la experiencia va antes que el precio**. Nunca soltamos un precio “frío”; primero hacemos que la familia sienta el valor."))
C.append(h2("6.1 Embudo comercial ramificado"))
C.append(p("El embudo se adapta a tres tipos de prospecto, porque no todos llegan igual:"))
C.append(table([
  ["Etapa","Qué pasa","Herramienta"],
  ["1. Contacto","Auto-respuesta instantánea + humano en <15–30 min","WhatsApp Business"],
  ["2. Descubrimiento","Preguntas de motivación, zona y edad","Script de descubrimiento"],
  ["3. Ramificación","A) Local → Clase Muestra · B) Distante → assets + videollamada · C) “Solo precio” → valor + precio","Scripts por rama"],
  ["4. Experiencia","Clase Muestra + recorrido (punto de conversión #1)","Cuerpo técnico"],
  ["5. Valor + Precio","Precio anclado en lo que el papá ya vivió","Script de precio"],
  ["6. Cierre","Beneficios + cupo limitado + manejo de objeciones","Script de cierre"],
  ["7. Inscripción","Bienvenida premium (kit + primer día)","Onboarding"],
  ["8. Seguimiento","Nurture en CRM + lista de espera si no cerró","Clupik"],
], aligns=["left","left","left"], widths=[0.18,0.56,0.26], first_col_bold=True))
C.append(callout("La Clase Muestra es la herramienta de conversión #1. Si el papá no quiere tour, se le entrega valor por otra vía (digital o respuesta consultiva): la experiencia es una invitación, no un peaje.", kind="insight"))
C.append(h2("6.2 Objeciones y cómo se resuelven"))
C.append(table([
  ["Objeción","Cómo se responde"],
  ["Distancia / tráfico (la #1)","“Vale el viaje” (familias que vienen de otras ciudades) + cámaras para ver al hijo en remoto + horarios que esquivan el tráfico + carpool"],
  ["Precio (es caro)","Comparar con “solo cancha”: aquí hay método, psicóloga, instalaciones nuevas y evaluación. Ofrecer pago anual y descuento por hermanos al final"],
  ["“Son nuevos / desconfianza”","Instalaciones nuevas + atención personal + respaldo de Andrés + resultados del piloto. La mejor prueba: una Clase Muestra"],
  ["“Lo voy a pensar”","Validar la decisión, apartar el lugar sin compromiso (cupo limitado real) y agendar la Clase Muestra para decidir con información"],
], aligns=["left","left"], widths=[0.26,0.74], first_col_bold=True))
C.append(h2("6.3 Reglas de oro del equipo comercial"))
C.append(bullet("**Experiencia antes que precio.** Nunca soltar el precio sin generar valor."))
C.append(bullet("**Velocidad:** auto-reply instantáneo + humano en menos de 15–30 minutos."))
C.append(bullet("**Nunca liderar con descuento.** Los beneficios se ofrecen al final."))
C.append(bullet("**Siempre capturar datos** en el CRM (nombre, niño, edad, zona, motivación)."))
C.append(bullet("**Siempre dejar un siguiente paso agendado** (nunca “cualquier cosa me avisas”)."))
C.append(bullet("**Tono premium-cálido**, de “tú”, sin presionar, sin sonar a call center."))
C.append(p("El manual de ventas completo y los scripts palabra por palabra están en el **capítulo 11**."))
C.append(page_break())

# ============================================================
# 7. ESTRATEGIA DIGITAL (detalle)
# ============================================================
C.append(h1("7. Estrategia Digital"))
C.append(p("La estrategia digital es **content-led**: nuestra fuerza es el contenido orgánico (equipo audiovisual + cámaras + highlights), y la pauta **amplifica** lo que ya funciona. Es coherente con una marca premium, que se construye por credibilidad y no por “gritar” ofertas."))
C.append(h2("7.1 Rol de cada canal"))
C.append(table([
  ["Canal","Rol","Marca"],
  ["Instagram + TikTok","Hub del ecosistema: reels primero → alcance, marca y descubrimiento","Compartido (con señal visual por marca)"],
  ["Facebook","Comunidad, ligas, eventos, público adulto/local; grupos de papás","2 cuentas (Academia / Complejo)"],
  ["Cuenta personal de Andrés","Amplificador / vocería","Personal (coordinada)"],
  ["WhatsApp Business","Conversión y atención (cierre de ventas)","Compartido"],
  ["Google Search","Captura de alta intención (quien ya busca)","Academia"],
], aligns=["left","left","left"], widths=[0.22,0.5,0.28], first_col_bold=True))
C.append(h2("7.2 Estructura de pauta (presupuesto acotado → concentrar)"))
C.append(p("Con ~$3,500–5,000/mes no alcanza para todos los canales a la vez. Se concentra en Meta y se complementa con Google Search:"))
C.append(table([
  ["Nivel","Temperatura","Objetivo"],
  ["1. Alcance / Video views","Frío","Amplificar mejores reels y construir audiencias"],
  ["2. Retargeting","Tibio","A quien interactuó: testimonios + invitación a Clase Muestra"],
  ["3. Conversión / Mensajes","Caliente","Registro a Clase Muestra → WhatsApp → cierre"],
  ["+ Google Search","Alta intención","Capturar “academia de futbol Pachuca” y similares"],
], aligns=["left","center","left"], widths=[0.28,0.18,0.54], first_col_bold=True))
C.append(callout("El retargeting es la palanca más rentable con presupuesto chico: le habla a quien ya nos conoce y cuesta poco. Es la prioridad táctica de la pauta.", kind="recomendacion"))
C.append(h2("7.3 Embudo digital → comercial"))
C.append(table([
  ["Fase","Acción digital","Conecta con"],
  ["TOFU (atraer)","Reels + highlights → crecer audiencia","Estrategia E1"],
  ["MOFU (nutrir)","Retargeting con valor/testimonios + invitación","Estrategias E2 / E6"],
  ["BOFU (convertir)","Campaña de mensajes → WhatsApp → scripts","Estrategia E3"],
], aligns=["left","left","center"], widths=[0.22,0.56,0.22], first_col_bold=True))
C.append(page_break())

# ============================================================
# 8. ESTRATEGIA DE CONTENIDOS (detalle)
# ============================================================
C.append(h1("8. Estrategia de Contenidos"))
C.append(p("El contenido es el activo que **construye marca y alimenta el embudo** casi sin costo. La regla es **70/30: aportar mucho, vender poco**. Si solo vendemos, la audiencia se va; si aportamos, confía y compra."))
C.append(h2("8.1 Pilares y mezcla de contenido"))
C.append(table([
  ["Bucket","Peso","Objetivo","Ejemplos"],
  ["Atraer","40%","Crecer audiencia","Highlights, jugadas, trends, ambiente de ligas/pádel"],
  ["Educar","25%","Justificar el premium","Tips de método, consejos de entrenador/psicóloga"],
  ["Conectar","20%","Confianza y emoción","Testimonios, detrás de cámara, Andrés, valores"],
  ["Convertir","15%","Llenar el embudo","Clase Muestra, cupo, beneficios, CTA a WhatsApp"],
], aligns=["left","center","left","left"], widths=[0.16,0.1,0.28,0.46], first_col_bold=True))
C.append(callout("Los highlights de las cámaras son el contenido estrella: el papá comparte el video de su hijo y nos consigue alcance gratuito con el público objetivo exacto. Es contenido, prueba de valor y herramienta de retención, todo a la vez.", kind="oportunidad"))
C.append(h2("8.2 Plan semanal (plantilla rotable: 4 reels + 4 posts)"))
C.append(table([
  ["Pieza","Tipo","Bucket","Marca"],
  ["Reel 1","Highlight de alumno (cámaras)","Atraer/Conectar","Academia"],
  ["Reel 2","Tip de método / consejo","Educar","Academia"],
  ["Reel 3","Detrás de cámara / Andrés / valores","Conectar","Academia"],
  ["Reel 4","Ambiente de liga / pádel / trend","Atraer/Convertir","Complejo"],
  ["Post 1","Testimonio de papá (carrusel)","Conectar","Academia"],
  ["Post 2","Instalaciones / diferenciadores","Educar","Mixto"],
  ["Post 3","Clase Muestra / cupo / beneficio","Convertir","Academia"],
  ["Post 4","Liga / torneo / evento / info","Atraer/Convertir","Complejo"],
], aligns=["left","left","center","center"], widths=[0.13,0.42,0.25,0.2], first_col_bold=True))
C.append(h2("8.3 Calendario anual de campañas (estacionalidad México)"))
C.append(table([
  ["Mes","Tema / oportunidad"],
  ["Agosto","🚀 Lanzamiento + regreso a clases (papás inscriben actividades)"],
  ["Septiembre","Conversión post-evento · fiestas patrias (contenido) · ligas"],
  ["Octubre","Consolidación · torneos · contenido de resultados"],
  ["Noviembre","Buen Fin (promos Complejo / pádel / funcional)"],
  ["Diciembre","Cierre de año · ceremonia/posada · clínicas de vacaciones"],
  ["Enero","Propósitos de año nuevo (boom funcional/pádel) · nueva inscripción"],
  ["Febrero","Retención · San Valentín (pádel parejas) · torneos"],
  ["Marzo","Evaluaciones · torneos"],
  ["Abril","Campamento / clínicas de Semana Santa"],
  ["Mayo","Día del niño / madre (contenido familiar)"],
  ["Jun–Jul","Cursos de verano (gran adquisición) · preparar siguiente ciclo"],
], aligns=["center","left"], widths=[0.16,0.84], first_col_bold=True))
C.append(page_break())


# ============================================================
# 9. PRESUPUESTO POR CANAL
# ============================================================
C.append(h1("9. Presupuesto Estimado por Canal"))
C.append(p("El presupuesto parte de la realidad del proyecto: una **bolsa de pauta acotada (~$3,500–5,000/mes)** y un **equipo de contenido ya disponible** (diseñador, audiovisual, community). Por eso la estrategia es content-led: el mayor valor se genera con recursos que ya se pagan."))
C.append(h2("9.1 Mes base (operación normal)"))
C.append(table([
  ["Canal / rubro","Inversión mensual","Objetivo principal"],
  ["Meta (Instagram + Facebook)","$3,500 – $4,000","Leads a Clase Muestra + retargeting"],
  ["Google Search","$800 – $1,000","Alta intención (quien ya busca)"],
  ["Impresos / herramientas","Según necesidad","Lonas, flyers locales, software"],
  ["Producción de contenido","Incluido (nómina)","Reels, posts, highlights, video-tour"],
  ["TikTok","Orgánico (sin pauta aún)","Alcance y descubrimiento"],
  ["TOTAL pauta mensual","≈ $4,300 – $5,000","—"],
], aligns=["left","center","left"], widths=[0.34,0.24,0.42], first_col_bold=True))
C.append(callout("El presupuesto chico limita la VELOCIDAD de crecimiento, no la rentabilidad. Con LTV:CAC ≈ 10:1, una vez validado el CAC real en los meses 1–3, el ratio justifica solicitar a Finanzas más presupuesto para escalar sobre lo que ya funciona.", kind="dato"))
C.append(h2("9.2 Mes de lanzamiento (agosto) — refuerzo"))
C.append(table([
  ["Rubro","Inversión","Objetivo"],
  ["Refuerzo de pauta Meta","Doblar la inversión base","Campaña de registros a Clase Muestra + cobertura del evento"],
  ["Impresos del evento","Bolsa especial","Lonas, señalética, material de Puertas Abiertas"],
  ["Logística del evento","Bolsa especial","Activaciones, registro QR, prensa"],
], aligns=["left","center","left"], widths=[0.3,0.28,0.42], first_col_bold=True))
C.append(h2("9.3 Lógica de escalamiento del presupuesto"))
C.append(bullet("**Meses 1–3:** validar el CAC real del mercado frío (no escalar todavía)."))
C.append(bullet("**Si el CAC se confirma sano:** construir el caso de negocio y pedir más presupuesto a Finanzas (el ratio 10:1 lo respalda)."))
C.append(bullet("**Escalar solo** sobre canales y campañas con CAC probado. Nunca subir el gasto “a ciegas”."))
C.append(page_break())

# ============================================================
# 10. KPIs CON METAS POR MES
# ============================================================
C.append(h1("10. KPIs con Metas por Mes"))
C.append(p("Tablero de metas mensuales para el ciclo de lanzamiento (agosto 2026 – julio 2027). Las cifras son **metas tentativas a validar con Dirección y a ajustar con datos reales** de los primeros meses. La línea base es baja (~30 alumnos, ~200 seguidores, sin pauta previa), por lo que el crecimiento es de construcción, no de explosión."))
C.append(callout("Cómo leer la tabla: “Alumnos” es el acumulado al cierre del mes (parte de 30 y llega a ~90, contemplando bajas naturales). Las inscripciones brutas son mayores que el crecimiento neto porque siempre hay algo de rotación.", kind="dato", label="NOTA METODOLÓGICA"))
C.append(table([
  ["Mes","Leads","Clase Muestra","Inscritos","Alumnos (acum.)","Seguidores","CPL / CAC meta"],
  ["Ago '26","55–60","22–24","15–17","≈ 45","≈ 500","CPL <$130 / CAC <$1,500"],
  ["Sep '26","45","18","10","≈ 52","≈ 700","Validar CAC"],
  ["Oct '26","40","16","9","≈ 58","≈ 850","CAC <$1,500"],
  ["Nov '26","38","15","8","≈ 63","≈ 1,000","CAC <$1,400"],
  ["Dic '26","30","12","6","≈ 66","≈ 1,100","CAC <$1,400"],
  ["Ene '27","45","18","10","≈ 73","≈ 1,300","CAC <$1,300"],
  ["Feb '27","35","14","7","≈ 77","≈ 1,450","CAC <$1,300"],
  ["Mar '27","33","13","7","≈ 81","≈ 1,600","CAC <$1,300"],
  ["Abr '27","30","12","6","≈ 84","≈ 1,750","CAC <$1,250"],
  ["May '27","30","12","6","≈ 87","≈ 1,900","CAC <$1,250"],
  ["Jun '27","28","11","5","≈ 89","≈ 2,050","CAC <$1,200"],
  ["Jul '27","30","12","6","≈ 92","≈ 2,200","CAC <$1,200"],
], aligns=["center","center","center","center","center","center","left"],
   widths=[0.1,0.1,0.13,0.12,0.15,0.13,0.27], sz="16", first_col_bold=True))
C.append(h2("10.1 Metas transversales del ciclo"))
C.append(table([
  ["Indicador","Meta del ciclo"],
  ["Alumnos (cierre julio 2027)","90 (escenario ideal 136)"],
  ["Ingreso recurrente mensual (cierre)","$108,000"],
  ["Retención mensual","> 90%"],
  ["NPS","> 60"],
  ["Comunidad en redes (suma)","≈ 2,200 seguidores"],
  ["Relación LTV:CAC","≥ 3:1 (objetivo ~10:1)"],
], aligns=["left","left"], widths=[0.5,0.5], first_col_bold=True))
C.append(page_break())


# ============================================================
# 11. MANUAL DE VENTAS Y SCRIPTS
# ============================================================
def msgbox(text, who="MENSAJE"):
    """Caja estilo 'mensaje de chat' para scripts."""
    fill = "F4F7FB"; accent = NAVY
    borders = ('<w:tblBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        f'<w:left w:val="single" w:sz="18" w:space="0" w:color="{accent}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{fill}"/>'
        '</w:tblBorders>')
    tcpr = (f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
            '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="140" w:type="dxa"/>'
            '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner = f'<w:p><w:pPr><w:spacing w:after="30"/></w:pPr>{runs(who, color=accent, bold=True, sz="16")}</w:p>'
    inner += f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs(text, italic=True)}</w:p>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{borders}</w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>' + p("", after=70))

C.append(h1("11. Manual de Ventas y Scripts"))
C.append(p("Este manual permite que **cualquier persona del equipo venda bien**, sin depender del talento individual. Los textos están listos para copiar y pegar; solo se ajustan los campos entre [corchetes]. La voz es siempre **premium-cálida, de “tú”, profesional pero humana**."))

C.append(h2("11.1 Script de WhatsApp"))
C.append(h3("Paso 1 — Auto-respuesta instantánea (WhatsApp Business)"))
C.append(msgbox("¡Hola! 👋 Gracias por escribir a la Academia de Fútbol Andrés Chitiva. En unos minutos un asesor te atiende personalmente. Mientras tanto, cuéntanos: ¿para qué edad buscas información? ⚽", who="AUTOMÁTICO"))
C.append(h3("Paso 2 — Primer contacto humano (<15–30 min)"))
C.append(msgbox("¡Hola [Nombre]! Soy [Asesor] de la Academia Andrés Chitiva 😊 Con gusto te ayudo. Para darte la mejor info y no marearte con datos que no aplican, ¿me permites 2–3 preguntas rápidas?", who="ASESOR"))
C.append(h3("Paso 3 — Descubrimiento (calificar y personalizar)"))
C.append(msgbox("1) ¿Cómo se llama tu hijo(a) y qué edad tiene?  2) ¿Por qué zona viven? (para ver los horarios que te queden mejor)  3) ¿Qué es lo que más te gustaría que [niño] lograra: que se desarrolle y agarre disciplina, que haga deporte y amigos, o que lleve su talento a nivel competitivo?", who="ASESOR"))
C.append(p("La respuesta 3 define el pitch. La respuesta 2 anticipa la objeción de distancia.", italic=True, color="595959", sz="18"))
C.append(h3("Paso 4 — Presentación de valor (según motivación)"))
C.append(msgbox("Me encanta lo que buscas, porque es justo nuestra esencia: no solo formamos futbolistas, formamos personas. [Niño] va a desarrollar disciplina, hábitos y confianza, con acompañamiento de nuestra psicóloga deportiva, instalaciones nuevas y entrenadores capacitados. Todo con un método estructurado, nada improvisado.", who="SI BUSCA DESARROLLO / VALORES"))
C.append(msgbox("Perfecto, ahí brillamos. Nuestra metodología está respaldada por Andrés Chitiva (exfutbolista profesional) y trabaja técnica, físico y táctica con evaluación individual continua. Además grabamos entrenamientos y partidos con cámaras para analizar y para que descargues sus mejores jugadas. 🎥", who="SI BUSCA ALTO RENDIMIENTO"))
C.append(h3("Paso 5 — Invitación a la experiencia"))
C.append(msgbox("La mejor forma de que lo sientas es vivirlo: te invito a una Clase Muestra gratis para que [niño] pruebe y tú conozcas las instalaciones. ¿Te queda mejor entre semana o el sábado?", who="RAMA A — LOCAL"))
C.append(msgbox("Para que lo veas desde ya, te paso un video corto de las instalaciones y de cómo entrenamos 🎥. Si quieres, agendamos una videollamada de 10 min para resolver todo. El paso ideal sería una Clase Muestra para que [niño] lo viva. ¿Te late?", who="RAMA B — DISTANTE"))
C.append(msgbox("Claro, te comparto la info 👇 La mensualidad es de $1,200 e inscripción $1,500, e incluye [beneficios según motivación]. Más que el precio, lo que vas a notar es la diferencia en el método y las instalaciones — por eso te invito a una Clase Muestra gratis para que lo compruebes sin compromiso. ¿Agendamos?", who="RAMA C — “SOLO EL PRECIO”"))
C.append(h3("Paso 6 — Precio (después de la experiencia)"))
C.append(msgbox("Ya viste por qué somos diferentes 😊. La inversión es de $1,200 al mes + inscripción de $1,500, e incluye [todo lo que vivió]. Y como vamos arrancando, tenemos beneficios de lanzamiento para las primeras familias. ¿Te explico cómo apartar el lugar de [niño]? Tenemos cupo limitado por categoría.", who="ASESOR"))
C.append(h3("Paso 7 — Cierre"))
C.append(msgbox("Para [niño] tenemos lugar en la categoría [X], pero se están llenando rápido. ¿Te ayudo a apartarlo con la inscripción para asegurarlo? Lo podemos hacer ahora mismo por aquí 💪", who="ASESOR"))
C.append(page_break())

C.append(h2("11.2 Script presencial (Clase Muestra / visita)"))
C.append(p("Estructura de la visita presencial. El objetivo es que el papá **viva** el valor, no que escuche un discurso de venta."))
C.append(table([
  ["Momento","Duración","Qué hacer / decir"],
  ["1. Bienvenida cálida","~2 min","Romper el hielo, agradecer su tiempo, hacerlo sentir esperado"],
  ["2. Descubrimiento","~5 min","Preguntar motivación, expectativas y miedos (igual que en WhatsApp)"],
  ["3. Recorrido del valor","~10 min","Mostrar instalaciones, cámaras y presentar método y entrenadores"],
  ["4. Conexión","Durante el tour","“Como buscabas [X], mira cómo lo trabajamos aquí…”"],
  ["5. Precio con contexto","~3 min","Precio anclado en lo que vio + beneficios de lanzamiento"],
  ["6. Cierre + siguiente paso","~2 min","Apartar lugar / agendar inscripción (cupo limitado real)"],
], aligns=["left","center","left"], widths=[0.24,0.14,0.62], first_col_bold=True))
C.append(h3("Frases presenciales clave"))
C.append(msgbox("Qué bueno que vinieron, los estábamos esperando. Antes de enseñarte todo, cuéntame: ¿qué es lo que más te gustaría que [niño] se llevara de aquí?", who="BIENVENIDA"))
C.append(msgbox("Mira, estas son nuestras canchas y aquí están las cámaras: cada jugada de [niño] queda grabada para analizar y para que tú la veas y la descargues, aunque no te quedes. Y ella es [psicóloga/entrenador], parte del equipo que lo va a acompañar.", who="RECORRIDO"))
C.append(msgbox("Ya viste la diferencia tú mismo. La inversión es de $1,200 al mes; e incluye todo esto que conociste. Como vamos arrancando, las primeras familias tienen beneficios especiales. ¿Te aparto el lugar de [niño] en su categoría? El cupo es limitado y se está llenando.", who="CIERRE PRESENCIAL"))
C.append(h3("Do’s & Don’ts"))
C.append(table([
  ["✅ Sí hacer","❌ No hacer"],
  ["Preguntar antes de hablar","Soltar un catálogo de beneficios sin escuchar"],
  ["Invitar siempre a la Clase Muestra","Cerrar solo por chat sin intentar la experiencia"],
  ["Personalizar según la motivación","Usar el mismo discurso para todos"],
  ["Dejar un siguiente paso agendado","Terminar en “cualquier cosa me avisas”"],
], aligns=["left","left"], widths=[0.5,0.5], first_col_bold=True))
C.append(page_break())


# ============================================================
# 12. SISTEMA DE MEDICIÓN Y REPORTES
# ============================================================
C.append(h1("12. Sistema de Medición y Reportes"))
C.append(p("“Lo que no se mide, no se mejora.” Este sistema define **qué medimos, cada cuánto y quién lo revisa**, usando las herramientas que ya tenemos: **Clupik, Meta Business y WhatsApp**."))
C.append(h2("12.1 Tablero de KPIs por área"))
C.append(table([
  ["Área","Indicadores clave"],
  ["Marketing / Digital","Alcance, crecimiento de seguidores, engagement, clics a WhatsApp, CPL, costo por registro a Clase Muestra"],
  ["Comercial / Ventas","Tiempo de respuesta, % invitados a Clase Muestra, show-rate, % de cierre, inscripciones, CAC, mix por fuente"],
  ["CX / Retención","Retención y churn mensual, renovación por ciclo, NPS, asistencia promedio, % de familias que refieren"],
  ["Financiero","Ingreso recurrente, ingreso del Complejo, LTV:CAC, ROAS del embudo, inscripciones nuevas/mes"],
  ["Complejo","Ocupación por franja horaria, equipos en ligas, uso de pádel, eventos/empresas por mes"],
], aligns=["left","left"], widths=[0.22,0.78], first_col_bold=True))
C.append(h2("12.2 Cadencia de revisión"))
C.append(table([
  ["Frecuencia","Qué se revisa","Quién"],
  ["Semanal","Leads, tiempos de respuesta, Clases Muestra agendadas/asistidas, contenido publicado","Marketing + Ventas"],
  ["Mensual","CAC, inscripciones, churn, ocupación del Complejo, presupuesto vs. resultado","Marketing + Dirección"],
  ["Trimestral","OKRs, NPS, decisiones de escalamiento de presupuesto","Dirección + Finanzas"],
], aligns=["left","left","left"], widths=[0.16,0.62,0.22], first_col_bold=True))
C.append(h2("12.3 Plantilla de reporte mensual (1 página)"))
C.append(p("Cada cierre de mes, Marketing entrega un reporte breve y visual con esta estructura:"))
C.append(bullet("**Resumen del mes:** 3 logros + 3 aprendizajes."))
C.append(bullet("**Embudo:** leads → Clase Muestra → inscritos (con % de conversión)."))
C.append(bullet("**Adquisición:** CPL y CAC por canal vs. meta."))
C.append(bullet("**Crecimiento:** alumnos acumulados vs. rampa planeada (30 → 90)."))
C.append(bullet("**Retención:** churn, NPS y alertas de asistencia."))
C.append(bullet("**Contenido y redes:** alcance, seguidores y mejores piezas."))
C.append(bullet("**Presupuesto:** invertido vs. planeado y resultado obtenido."))
C.append(bullet("**Decisiones / próximos pasos:** qué se ajusta el siguiente mes."))
C.append(callout("La métrica más importante de los primeros 3 meses es el CAC del mercado frío. Es la “luz verde” para escalar el presupuesto. Todo el sistema de medición está diseñado para responder con confianza: ¿estamos adquiriendo alumnos de forma rentable?", kind="insight"))
C.append(page_break())

# ============================================================
# 13. CRONOGRAMA E IMPLEMENTACIÓN
# ============================================================
C.append(h1("13. Cronograma e Implementación"))
C.append(h2("13.1 Hoja de ruta rumbo a agosto"))
C.append(table([
  ["Periodo","Acciones clave"],
  ["Fin jun – ~10 jul","Cerrar piloto; configurar WhatsApp Business y Clupik; activar flujo de highlights"],
  ["~10 jul","Evaluar piloto; definir reingresos (Generación Fundadora); presentar promos a Finanzas"],
  ["Mediados jul","Producir assets de video; iniciar campaña de intriga"],
  ["Fin jul","Abrir lista de espera; logística e invitaciones del evento"],
  ["Inicio ago","Pre-venta a Fundadores/tibios; capacitar ventas; preparar mesa de inscripción + QR"],
  ["2ª quincena ago","🚀 Evento de lanzamiento + apertura + pauta reforzada"],
  ["Septiembre+","Conversión de leads, retargeting, medición de CAC, ritmo continuo"],
], aligns=["left","left"], widths=[0.22,0.78], first_col_bold=True))
C.append(h2("13.2 Responsables (quién hace qué)"))
C.append(table([
  ["Área","Responsable","Entregables"],
  ["Estrategia y coordinación","Marketing","Calendarios, campañas, presupuesto, KPIs"],
  ["Contenido visual","Diseño + Audiovisual","Assets, reels, posts, video-tour, highlights"],
  ["Comunidad y atención","Community","Publicación, respuesta, primer contacto"],
  ["Ventas","Recepción / Ventas","Embudo, scripts, cierre, Clase Muestra"],
  ["PR e imagen","Andrés (+ exjugadores)","Evento, prensa, vocería"],
  ["Precios y aprobaciones","Finanzas / Dirección","Promos, presupuesto, costo fijo"],
  ["Deportivo","Cuerpo técnico","Clase Muestra, evaluaciones, resultados"],
], aligns=["left","left","left"], widths=[0.26,0.24,0.5], first_col_bold=True))
C.append(h2("13.3 Plan de acción inmediato (próximas 2 semanas)"))
C.append(numbered("**1.** WhatsApp Business: auto-reply + scripts cargados + etiquetas de pipeline."))
C.append(numbered("**2.** Clupik: configurar pipeline de ventas, reportes y alertas de asistencia."))
C.append(numbered("**3.** Activar el flujo de highlights con las cámaras."))
C.append(numbered("**4.** Arrancar la producción de assets de video."))
C.append(numbered("**5.** Preparar la propuesta de promociones para Finanzas (Fundadores, anual, hermanos, referidos)."))
C.append(numbered("**6.** Borrador del plan de evento + lista de invitados."))
C.append(page_break())

# ============================================================
# 14. CIERRE
# ============================================================
C.append(h1("14. Cierre y Solicitudes a Dirección"))
C.append(p("Este Plan de Marketing posiciona al Proyecto Andrés Chitiva para **tomar el liderazgo del segmento premium en Pachuca**, con una unidad económica sólida (LTV:CAC ≈ 10:1) y una ruta clara hacia el lanzamiento de agosto y el crecimiento sostenible a 12 meses."))
C.append(p("Para ejecutarlo a tiempo y bien, Marketing solicita a Dirección:"))
C.append(bullet("Confirmar el estatus operativo de **nutrición y servicios médicos** (para validar la promesa de marca)."))
C.append(bullet("Aprobar el **presupuesto de pauta mensual** con refuerzo para agosto."))
C.append(bullet("Validar las **promociones propuestas** (vía Finanzas)."))
C.append(bullet("Respaldar el **evento de lanzamiento** (agenda de Andrés, exjugadores e invitados institucionales)."))
C.append(callout("Tenemos producto validado (alumnos que ya viajan por el método), un mercado dispuesto a pagar premium y un competidor en declive. La ventana de oportunidad es ahora. La calidad con que ejecutemos el lanzamiento de agosto definirá el ritmo de los próximos 12 meses.", kind="insight", label="MENSAJE FINAL"))
C.append(spacer(200))
C.append(p("Departamento de Marketing — Proyecto Andrés Chitiva — 2026", align="center", color="808080", sz="18"))

# ============================================================
# ENSAMBLADO DEL DOCUMENTO
# ============================================================
body = "".join(C)

document_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 f'<w:body>{body}'
 '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
 '<w:pgSz w:w="12240" w:h="15840"/>'
 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1418" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
 '</w:sectPr></w:body></w:document>')

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="3" w:color="C9A227"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="34"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="27"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="B8860B"/><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="360"/><w:spacing w:after="60"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="1F3864"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>'''

footer_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>
<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">Plan de Marketing — Proyecto Andres Chitiva — Confidencial — Pagina </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document_xml)
    z.writestr("word/styles.xml", styles_xml)
    z.writestr("word/footer1.xml", footer_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)

print("OK ->", OUT, os.path.getsize(OUT), "bytes")
