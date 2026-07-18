#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera el documento ejecutivo .docx del Plan de Contingencia de Cierre y Rescate de Leads."""
import re, zipfile, os

OUT = "/projects/sandbox/estrategia-andres-chitiva/Plan-Contingencia-Cierre-Leads-Andres-Chitiva.docx"

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

def page_break(): return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def cell(text, header=False, fill=None, color=None, align=None, bold=False):
    tcpr="<w:tcPr><w:tcW w:w="+'"0" w:type="auto"/>'
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
    if widths:
        out.append('<w:tblGrid>'+''.join(f'<w:gridCol w:w="{int(w*94)}"/>' for w in widths)+'</w:tblGrid>')
    for ri,r in enumerate(rows):
        cells=[]
        for ci,c in enumerate(r):
            a = aligns[ci] if aligns else None
            zebra = None if (ri==0 and header) else (F_GRAY if ri%2==0 else None)
            cells.append(cell(c, header=(ri==0 and header), fill=zebra, align=a))
        out.append("<w:tr>"+"".join(cells)+"</w:tr>")
    out.append("</w:tbl>")
    out.append(p("",after=80))
    return "".join(out)

CALL = {
 "insight":(F_NAVY,NAVY,"INSIGHT CLAVE"),
 "oportunidad":(F_GREEN,A_GREEN,"OPORTUNIDAD"),
 "riesgo":(F_RED,A_RED,"RIESGO / ALERTA"),
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
    inner=f'<w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr>{runs(title,color="FFFFFF",bold=True,sz="46")}</w:p>'
    for i,s in enumerate(subtitle_lines):
        col = MUSTARD if i==0 else "D6DCE5"
        szv = "28" if i==0 else "22"
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
C.append('<w:p><w:pPr><w:spacing w:before="400" w:after="200"/></w:pPr></w:p>')
C.append(banner("PLAN DE CONTINGENCIA DE CIERRE Y RESCATE DE LEADS",
   ["Proyecto Andrés Chitiva",
    "Academia de Fútbol  ·  Complejo Deportivo",
    "Pachuca, Hidalgo, México  ·  Julio 2026"]))
C.append(p("",after=200))
C.append(p("Documento operativo preparado por el Departamento de Marketing", align="center", color=NAVY, bold=True, sz="24"))
C.append(p("Diagnóstico de conversión · Plan de rescate/remarketing · Evaluación del área de ventas", align="center", color="595959", sz="20"))
C.append(p("",after=160))
C.append(callout("El plan de marketing generó buen volumen de leads y conversaciones, pero la conversión a inscripción es baja. Este documento diagnostica el porqué real con datos de la base de leads (461 registros, 16 jun – 13 jul 2026), define un plan de rescate accionable para el área de ventas, y establece un marco objetivo para evaluar a la vendedora.", kind="decision", label="PROPÓSITO"))
C.append(callout("Dato crítico de calendario: la Academia abre oficialmente en la SEGUNDA QUINCENA DE AGOSTO 2026. Al momento de este análisis (julio) estamos a ~4-5 semanas de la apertura. Este factor explica gran parte de los cierres pendientes.", kind="riesgo", label="CONTEXTO DE CALENDARIO"))
C.append(page_break())
C.append(toc())
C.append(page_break())

# 1. DIAGNOSTICO
C.append(h1("1. Diagnóstico con datos (no opinión)"))
C.append(p("Análisis de la base **BASE LEADS CHITIVA JULIO**: 461 leads con datos entre el 16-jun y el 13-jul de 2026."))
C.append(h2("1.1 El embudo real (global)"))
C.append(table([
  ["Resultado","#","%"],
  ["Sin resultado registrado (en limbo)","174","37.7%"],
  ["Perdido — Desinteresado","105","22.8%"],
  ["Interesado / abierto (no cerrado)","55","11.9%"],
  ["Ghost / dejó de responder","47","10.2%"],
  ["Cerrado / en cierre","44","9.5%"],
  ["Otro / en espera","36","7.8%"],
], aligns=["left","center","center"], widths=[0.6,0.2,0.2]))
C.append(h2("1.2 Separando las dos marcas (son ventas distintas)"))
C.append(table([
  ["Segmento","Total","Cerrado/en cierre","Interesado abierto","Sin resultado","Desinteresado"],
  ["ACADEMIA","232","40 (17.2%)","31 (13.4%)","107 (46.1%)","27 (11.6%)"],
  ["COMPLEJO","214","4 (1.9%)","24 (11.2%)","55 (25.7%)","75 (35.0%)"],
], aligns=["left","center","center","center","center","center"], widths=[0.2,0.12,0.19,0.18,0.16,0.15]))
C.append(h2("1.3 Los 4 hallazgos que explican el 'no cierran'"))
C.append(callout("HALLAZGO 1 — El problema #1 NO es el cierre, es el SEGUIMIENTO. El 38% de todos los leads (46% en Academia) no tienen ningún resultado registrado. De 314 leads en etapa 'INFORMES', 104 se quedaron ahí sin desenlace. La mayoría nunca llega al cierre porque se estanca en 'les mandé info' sin un siguiente paso agendado.", kind="insight", label="HALLAZGO 1 · SEGUIMIENTO"))
C.append(callout("HALLAZGO 2 — Timing: la Academia todavía no se puede 'cerrar'. Los leads se generaron en junio-julio, pero la Academia abre en agosto. Muchas conversaciones terminan en 'arranca en agosto' y el lead se enfría en la brecha. Sin producto puente y sin mecánica de reserva, la conversación muere esperando.", kind="riesgo", label="HALLAZGO 2 · TIMING"))
C.append(callout("HALLAZGO 3 — La vendedora SÍ cierra cuando el producto está disponible. La Academia convierte 17.2% vs. 1.9% del Complejo. Donde hay producto vendible hoy y decisor claro (el papá), sí hay cierre. El talento base existe; el problema es sistémico (proceso + timing + oferta), no puramente de habilidad.", kind="oportunidad", label="HALLAZGO 3 · TALENTO"))
C.append(callout("HALLAZGO 4 — No hay pipeline, hay un Excel caótico. La etapa 'INFORMES' aparece escrita de 8 formas distintas y los resultados no están estandarizados. Sin pipeline no hay disciplina de seguimiento ni forma de medir a nadie.", kind="riesgo", label="HALLAZGO 4 · SISTEMA"))
C.append(h2("1.4 El tamaño de la oportunidad (lo recuperable)"))
C.append(table([
  ["Bloque","# leads","Acción"],
  ["Cerrados / en cierre","~44","Ganados o casi — asegurar"],
  ["Desinteresados","~108","Perdidos reales — aprender, no forzar"],
  ["EN LIMBO / ABIERTOS — RECUPERABLES","~309","Objetivo del plan de contingencia"],
  ["  · Academia recuperable","165","31 interesados + 107 sin resultado + 12 ghost + 15 espera"],
  ["  · Complejo recuperable","~135","Reactivar con secuencias"],
], aligns=["left","center","left"], widths=[0.34,0.16,0.5]))
C.append(callout("Hay ~300 leads tibios/en limbo que ya pagaste (CAC hundido) y que se están enfriando por falta de un sistema de seguimiento y por la brecha hasta agosto. El plan de contingencia ataca exactamente eso.", kind="insight", label="CONCLUSIÓN DEL DIAGNÓSTICO"))
C.append(page_break())

# 2. PLAN DE CONTINGENCIA
C.append(h1("2. Plan de Contingencia — 3 movimientos"))
C.append(h2("Movimiento A — TRIAGE: clasificar y reactivar la base en 72 h"))
C.append(p("Antes de enviar nada, ordenar la base (medio día de trabajo). Reclasificar cada lead abierto en 5 estados accionables y asignar acción:"))
C.append(table([
  ["Estado del lead","# aprox (Aca)","Acción inmediata","Prioridad"],
  ["Interesado abierto","31","Reactivar + agendar siguiente paso (clase muestra / curso de verano / reserva)","P1 — hoy"],
  ["Sin resultado / 'informes'","107","Mensaje de reactivación (Secuencia 1) para reclasificar","P2 — 72 h"],
  ["Ghost / dejó de responder","12","Secuencia de reenganche con contenido (Secuencia 2)","P3 — semana"],
  ["En espera / 'lo pienso'","15","Nurture con contenido + fecha concreta","P2"],
  ["Desinteresado","27","Encuesta de salida (1 pregunta) → aprender motivo","P4"],
], aligns=["left","center","left","center"], widths=[0.24,0.13,0.45,0.18]))
C.append(h2("Movimiento B — REMARKETING DIGITAL (Meta + WhatsApp)"))
C.append(bullet("**Públicos personalizados en Meta:** subir los teléfonos de la base → crear audiencia y retargetear con testimonios + tour + invitación a reservar. Le pega exactamente a quien ya te escribió."))
C.append(bullet("**Retargeting de video-viewers:** a quien vio tus reels → anuncio de 'aparta tu cupo de la Generación Fundadora'."))
C.append(bullet("**WhatsApp — envío directo de contenido:** la vendedora manda el activo que le falta al lead para decidir. El remarketing paga por estar presente; el WhatsApp cierra."))
C.append(bullet("**Broadcast / lista de difusión** segmentada (Academia vs. Complejo) con cuenta regresiva a la apertura de agosto."))
C.append(h2("Movimiento C — EL PRODUCTO PUENTE (resuelve el timing de agosto)"))
C.append(callout("No dejes que 'arranca en agosto' mate al lead. El curso de verano / pretemporada intensiva es un producto vendible HOY que hace que el niño viva la metodología antes de la apertura. Ya tienes 62 leads de 'curso de verano' + 13 de 'pretemporada' pidiéndolo.", kind="oportunidad", label="DESBLOQUEO ESTRATÉGICO"))
C.append(p("**Mecánica de conversión en cascada:** Curso de verano (vive el método ahora) → cupo prioritario garantizado en la Academia de agosto → Generación Fundadora (precio congelado + inscripción de fundador)."))
C.append(bullet("**Reserva con anticipo:** replicar la mecánica de 'aparta tu lugar' del Complejo ($200-300 que se abonan a la inscripción) para cerrar el compromiso hoy aunque la Academia abra en agosto."))
C.append(bullet("Para quien no puede el curso de verano: **lista de espera priorizada 'Fundadores'** con fecha concreta de contacto (no 'yo te aviso')."))
C.append(page_break())

# 3. CONTENIDO
C.append(h1("3. Contenido para enviar (videos/fotos)"))
C.append(p("El archivo **Preproduccion PlanMensualAcademia** ya tiene los assets correctos, pero están en estado 'Pendiente'. Priorizar producir estos 6 para el rescate (son munición de venta, no solo de redes):"))
C.append(table([
  ["#","Asset a producir YA","Para qué lead lo usa la vendedora","Objeción que resuelve"],
  ["1","Video tour de instalaciones","Distantes / '¿cómo es el lugar?'","Desconfianza / distancia"],
  ["2","Video '¿Qué esperar al inscribirte?'","Interesado abierto (P1)","'¿vale lo que cobran?'"],
  ["3","Testimonios de papás (piloto)","Ghost + 'lo pienso'","'son nuevos / desconfianza'"],
  ["4","Entrenadores en acción / metodología","Segmento alto rendimiento","'¿mi hijo progresa?'"],
  ["5","Highlights de alumnos (cámaras)","Todos — es el diferenciador #1","'no vale el viaje'"],
  ["6","Presentación del proyecto + Andrés","Papá decisor","autoridad / respaldo"],
], aligns=["center","left","left","left"], widths=[0.06,0.28,0.34,0.32]))
C.append(callout("El contenido se envía SIEMPRE con un objetivo de conversación y un cierre, nunca suelto. Ej: 'Te mando el tour para que lo veas. ¿Te agendo el cupo de [niño] en el curso de verano para que lo viva esta semana?'", kind="recomendacion", label="REGLA DE USO"))
C.append(h2("3.1 Secuencias de remarketing por WhatsApp (copy-paste)"))
C.append(h3("Secuencia 1 — Reactivar 'sin resultado / informes' (los 107)"))
C.append(bullet("**Día 0:** Hola [Nombre], te escribo de la Academia Andrés Chitiva. Te había pasado info para [niño]. Justo abrimos cupos de Generación Fundadora con beneficios para las primeras familias antes de la apertura de agosto. ¿Te gustaría que [niño] viva una sesión del curso de verano para que lo pruebe? ¿Te queda mejor entre semana o sábado?"))
C.append(bullet("**Día 2 (si no responde):** enviar video tour → 'Para que lo veas desde ya. ¿Agendamos?'"))
C.append(bullet("**Día 5:** enviar testimonio de papá → 'Mira lo que dicen las familias que ya están. ¿Te aparto el lugar sin compromiso?'"))
C.append(h3("Secuencia 2 — Reenganchar ghosts (los 12)"))
C.append(bullet("Hola [Nombre], sé que andas ocupado. No quiero que [niño] pierda el cupo de Fundador (son limitados por categoría). Te dejo su mejor opción: [curso de verano/reserva]. ¿Lo aseguramos hoy?"))
C.append(bullet("Si no responde en 3 intentos → lista de newsletter mensual (no insistir, según la regla del plan)."))
C.append(h3("Secuencia 3 — 'Lo voy a pensar' / en espera (los 15)"))
C.append(bullet("Claro, es una decisión importante. ¿Qué te ayudaría a decidir? Mientras, te aparto el lugar sin compromiso porque el cupo es limitado. ¿Te parece si [niño] toma una sesión del curso de verano y ya con eso decides?"))
C.append(h3("Secuencia 4 — Encuesta de salida a desinteresados (los 27)"))
C.append(bullet("Hola [Nombre], gracias por tu tiempo antes. Solo para mejorar: ¿qué fue lo que te hizo no continuar? a) precio b) distancia c) horarios d) esperabas otra cosa. Tu respuesta nos ayuda muchísimo."))
C.append(page_break())

# 4. EVALUACION
C.append(h1("4. Evaluación de la vendedora"))
C.append(callout("Respuesta honesta con los datos actuales: no se puede culpar (ni exonerar) a la vendedora todavía, porque no hay métricas de sus indicadores de proceso. Lo que los datos sí muestran: el mayor leak es seguimiento no ejecutado/registrado (46% sin desenlace) y timing de producto (agosto) — ambos mitad proceso, mitad ejecución. Antes de decidir sobre la persona, hay que medirla bien durante 2-3 semanas.", kind="insight", label="POSTURA BASADA EN DATOS"))
C.append(h2("4.1 Separar los 3 posibles culpables"))
C.append(table([
  ["Si el problema fuera...","La señal en los datos sería...","¿Qué vemos hoy?"],
  ["La vendedora (habilidad)","Muchos leads llegan a precio/cierre y se caen ahí","Poco visible: casi nadie llega al cierre; se caen antes, en el seguimiento"],
  ["El proceso / sistema","Leads sin siguiente paso, sin registro, sin cadencia","Fuerte: 46% sin resultado, etapas caóticas, sin CRM"],
  ["La oferta / timing","Interés alto pero 'espero a agosto' / objeción precio-distancia","Fuerte: producto no abre hasta agosto, sin puente ni reserva"],
], aligns=["left","left","left"], widths=[0.24,0.38,0.38]))
C.append(h2("4.2 Cómo evaluarla objetivamente (2-3 semanas)"))
C.append(bullet("**Auditoría de tiempo de respuesta:** minutos hasta el primer contacto humano (meta del plan: <15-30 min). Cada hora de demora ≈ -30% de conversión."))
C.append(bullet("**Mystery shopping:** 3-5 personas escriben como prospecto por FB/WhatsApp y evalúan con rúbrica. Es la forma más limpia de ver su proceso real."))
C.append(bullet("**Revisión de 15-20 conversaciones reales** de WhatsApp contra el script del plan: ¿califica?, ¿presenta valor antes que precio?, ¿invita a la experiencia?, ¿agenda siguiente paso?, ¿maneja objeciones?"))
C.append(bullet("**Scorecard semanal de KPIs** (los que ya define el plan, hoy no se miden)."))
C.append(table([
  ["KPI del vendedor","Meta","Cómo se mide"],
  ["Tiempo de 1ª respuesta","< 15-30 min","Timestamp WhatsApp"],
  ["% leads con siguiente paso agendado","> 90%","Pipeline"],
  ["% invitados a experiencia (curso/muestra)","> 60%","Pipeline"],
  ["Show-rate de la experiencia","> 50%","Asistencia"],
  ["% de cierre (inscritos / trabajados)","definir base","Pipeline"],
  ["Leads sin resultado registrado","→ 0%","Pipeline"],
], aligns=["left","center","left"], widths=[0.46,0.24,0.3]))
C.append(h2("4.3 Rúbrica de mystery shopping (0-2 cada ítem)"))
C.append(bullet("¿Respondió rápido (<30 min en horario)?"))
C.append(bullet("¿Usó el nombre e hizo preguntas de calificación antes de hablar de precio?"))
C.append(bullet("¿Personalizó según motivación (desarrollo integral vs. alto rendimiento)?"))
C.append(bullet("¿Presentó VALOR antes que PRECIO?"))
C.append(bullet("¿Invitó a una experiencia (curso de verano / clase muestra / reserva)?"))
C.append(bullet("¿Manejó al menos una objeción sin bajar el precio?"))
C.append(bullet("¿Cerró con un siguiente paso concreto y con fecha?"))
C.append(bullet("¿Registró / pidió los datos del lead?"))
C.append(callout("Regla de decisión: si tras darle pipeline + scripts + producto puente + capacitación, sus KPIs siguen bajos (sobre todo respuesta lenta crónica, no agenda siguiente paso, no sube el valor), es tema de persona/ajuste de rol. Si mejoran, era proceso/oferta. El plan comercial original ya lo anticipa: la venta no debe depender del talento individual, sino del playbook.", kind="recomendacion", label="REGLA DE DECISIÓN"))
C.append(page_break())

# 5. QUE HACER ESTA SEMANA
C.append(h1("5. Qué hacer esta semana (7-14 días)"))
C.append(table([
  ["Periodo","Acciones"],
  ["Días 1-3","1) Montar pipeline mínimo (WhatsApp Business con etiquetas = etapas, o Clupik). Dejar el Excel caótico.  2) Triage de los ~300 leads recuperables en los 5 estados.  3) Definir y comunicar la mecánica: curso de verano → cupo Fundador → reserva con anticipo."],
  ["Días 3-7","4) Producir/priorizar los 6 assets (tour, testimonios, highlights primero).  5) Lanzar Secuencia 1 a los 107 'sin resultado' + Secuencia 2 a ghosts.  6) Subir base a Meta y encender retargeting."],
  ["Días 7-14","7) Arrancar mystery shopping + auditoría de tiempo de respuesta.  8) Primer scorecard semanal de KPIs.  9) Encuesta de salida a desinteresados → ajustar oferta/mensaje."],
], aligns=["center","left"], widths=[0.16,0.84]))
C.append(h2("5.1 KPIs de este plan (medir el rescate)"))
C.append(bullet("**Tasa de reactivación:** % de leads en limbo que responden a la Secuencia 1."))
C.append(bullet("**Leads sin resultado registrado:** de 38% → < 5% en 2 semanas."))
C.append(bullet("**Reservas/anticipos capturados** para la apertura de agosto."))
C.append(bullet("**Conversión de curso de verano → inscripción Academia** (la cascada)."))
C.append(bullet("**Tiempo de 1ª respuesta** de la vendedora (tendencia semanal)."))
C.append(page_break())

# 6. CIERRE
C.append(h1("6. Cierre y decisiones a validar"))
C.append(h3("Acciones pendientes del cliente"))
C.append(bullet("Confirmar mecánica y precio de la reserva con anticipo (con Finanzas)."))
C.append(bullet("Definir el curso de verano/pretemporada como producto puente formal (fechas, precio, cupo)."))
C.append(bullet("Asignar responsable de producir los 6 assets antes de fin de semana."))
C.append(bullet("Elegir herramienta de pipeline (WhatsApp Business labels vs. Clupik) y migrar la base."))
C.append(h3("Decisiones a validar"))
C.append(bullet("El leak principal es seguimiento + timing, no (aún) habilidad de la vendedora."))
C.append(bullet("Producto puente (curso de verano) + reserva con anticipo = solución al gap de agosto."))
C.append(bullet("Remarketing = Meta (retargeting a la base) + WhatsApp (envío de contenido con cierre)."))
C.append(bullet("Evaluar a la vendedora con datos (mystery shopping + KPIs + revisión de chats) antes de decidir sobre la persona."))
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
</w:styles>'''

footer_xml='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>
<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">Plan de Contingencia de Cierre y Rescate de Leads — Proyecto Andres Chitiva — Confidencial — Pagina </w:t></w:r>
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

print("OK ->", OUT, os.path.getsize(OUT), "bytes")
