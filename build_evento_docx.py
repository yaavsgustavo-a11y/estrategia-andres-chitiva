#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera el Word ejecutivo del evento de inauguracion."""
import re, zipfile, os

OUT = "/projects/sandbox/estrategia-andres-chitiva/Evento-Inauguracion-Andres-Chitiva.docx"
NAVY="1F3864"; MUSTARD="C9A227"; GOLD="B8860B"
F_NAVY="EAF0FA"; F_MUST="FBF3D9"; F_GREEN="E6F4EA"; F_RED="FBE9E9"; F_GRAY="F2F2F2"
A_GREEN="2E7D32"; A_RED="C0392B"; A_GRAY="7F7F7F"

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


def p(text="",style=None,align=None,color=None,sz=None,bold=False,after=120):
    ppr="<w:pPr>"+(f'<w:pStyle w:val="{style}"/>' if style else "")+f'<w:spacing w:after="{after}"/>'+(f'<w:jc w:val="{align}"/>' if align else "")+"</w:pPr>"
    return f'<w:p>{ppr}{runs(text,color,sz,bold)}</w:p>'

def h1(t): return p(t,style="Heading1")
def h2(t): return p(t,style="Heading2")
def h3(t): return p(t,style="Heading3")
def bullet(t,indent=360): return f'<w:p><w:pPr><w:ind w:left="{indent}"/><w:spacing w:after="60"/></w:pPr>{runs("•  "+t)}</w:p>'
def pb(): return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
def divider(): return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D9D9D9"/></w:pBdr></w:pPr></w:p>'

def cell(text,header=False,fill=None,bold=False,align=None):
    f=NAVY if header else fill
    shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{f}"/>' if f else ""
    tcpr=f'<w:tcPr><w:tcW w:w="0" w:type="auto"/>{shd}<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>'
    c="FFFFFF" if header else None
    b=True if header else bold
    jc=f'<w:jc w:val="{align}"/>' if align else ""
    return f'<w:tc>{tcpr}<w:p><w:pPr><w:spacing w:after="0"/>{jc}</w:pPr>{runs(text,color=c,bold=b)}</w:p></w:tc>'

def table(rows,aligns=None,fills=None):
    b=('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
       '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
       '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
       '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
       '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
       '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/></w:tblBorders>')
    out=[f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{b}</w:tblPr>']
    for ri,row in enumerate(rows):
        cells=[]
        for ci,c in enumerate(row):
            a=aligns[ci] if aligns else None
            zf=None if ri==0 else (F_GRAY if ri%2==0 else None)
            rf=fills[ri] if fills and ri<len(fills) else zf
            cells.append(cell(c,header=(ri==0),fill=rf,align=a))
        out.append("<w:tr>"+"".join(cells)+"</w:tr>")
    out.append("</w:tbl>"); out.append(p("",after=80))
    return "".join(out)


def callout(text,kind="insight",label=None):
    palettes={"insight":(F_NAVY,NAVY,"INSIGHT CLAVE"),"oportunidad":(F_GREEN,A_GREEN,"OPORTUNIDAD"),
               "riesgo":(F_RED,A_RED,"RIESGO"),"recomendacion":(F_MUST,GOLD,"RECOMENDACIÓN"),
               "accion":(F_GREEN,A_GREEN,"ACCIÓN INMEDIATA")}
    fill,accent,deflabel=palettes[kind]; lab=label or deflabel
    b=(f'<w:tblBorders><w:top w:val="single" w:sz="4" w:color="{fill}"/>'
       f'<w:left w:val="single" w:sz="24" w:color="{accent}"/>'
       f'<w:bottom w:val="single" w:sz="4" w:color="{fill}"/>'
       f'<w:right w:val="single" w:sz="4" w:color="{fill}"/></w:tblBorders>')
    tcpr=(f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
          f'<w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="160" w:type="dxa"/>'
          f'<w:bottom w:w="100" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner=(f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>{runs(lab,color=accent,bold=True,sz="18")}</w:p>'
           f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{runs(text)}</w:p>')
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{b}</w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>'+p("",after=80))

def banner(title,lines):
    tcpr=(f'<w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{NAVY}"/>'
          '<w:tcMar><w:top w:w="500" w:type="dxa"/><w:left w:w="240" w:type="dxa"/>'
          '<w:bottom w:w="500" w:type="dxa"/><w:right w:w="240" w:type="dxa"/></w:tcMar></w:tcPr>')
    inner=f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr>{runs(title,color="FFFFFF",bold=True,sz="52")}</w:p>'
    for i,l in enumerate(lines):
        c=MUSTARD if i==0 else "D6DCE5"; s="28" if i==0 else "22"
        inner+=f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>{runs(l,color=c,bold=(i==0),sz=s)}</w:p>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
            f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>')


# =================== CONTENIDO ===================
C=[]

# Portada
C.append(p("",after=300))
C.append(banner("PLAN MAESTRO DE INAUGURACIÓN",
    ["Academia de Fútbol & Complejo Deportivo Andrés Chitiva",
     "Pachuca, Hidalgo, México  ·  2026",
     "Partido de Leyendas  ·  Prensa  ·  Patrocinadores  ·  Apertura Comercial"]))
C.append(p("",after=200))
C.append(p("Documento de planificación preparado por el Departamento de Marketing",align="center",color=NAVY,bold=True,sz="24"))
C.append(p("Para uso interno del equipo  ·  Confidencial",align="center",color="595959",sz="20"))
C.append(p("",after=200))
C.append(callout("Este evento tiene cuatro objetivos concretos: Posicionamiento y PR · Generación de leads · Inscripciones ese mismo día · Banco de contenido para 4–6 semanas de redes sociales. Todo elemento del programa responde a al menos uno de estos objetivos.",kind="recomendacion",label="PROPÓSITO DEL EVENTO"))
C.append(pb())

# 1. Objetivo
C.append(h1("1. Objetivo del Evento"))
C.append(table([
    ["Objetivo","Descripción","Cómo se mide"],
    ["PR y posicionamiento","Cobertura en medios locales; posicionar el proyecto como referente premium","≥ 3 notas publicadas"],
    ["Generación de leads","Capturar datos de cada asistente interesado","≥ 90% de asistentes registrados"],
    ["Inscripciones","Cerrar alumnos ese mismo día con beneficio de lanzamiento","≥ 5–10 inscripciones"],
    ["Contenido","Banco de fotos y video para 4–6 semanas de redes","≥ 10 piezas producidas"],
],aligns=["left","left","left"]))
C.append(pb())

# 2. Concepto
C.append(h1("2. Concepto del Evento"))
C.append(p('**Nombre propuesto:** "La Gran Inauguración Chitiva" / "Nace algo grande en Pachuca"'))
C.append(callout('"Pachuca es la cuna del fútbol en México. Hoy nace en ella un proyecto que va más allá del deporte: una academia que forma futbolistas inteligentes, apasionados y con valores, y un complejo de primer nivel para toda la familia. Bienvenidos al inicio de algo diferente."',kind="insight",label="NARRATIVA CENTRAL"))
C.append(p("**Tono:** premium pero cercano. No solemne. Hay fiesta, hay fútbol, hay comunidad — pero todo con producción cuidada y mensaje claro."))
C.append(p("**Aforo estimado:** ~100 personas. **Formato:** abierto al público con registro previo (genera escasez y captura de leads)."))
C.append(pb())

# 3. Programa del día
C.append(h1("3. Programa del Día (Run of Show)"))
C.append(callout("El partido de leyendas es el climax. Toda la logística del programa debe proteger ese momento y asegurar que el público llegue antes y se quede después.",kind="recomendacion"))
C.append(table([
    ["Hora","Actividad","Responsable","Notas clave"],
    ["T − 1:00 h","Montaje final · prueba de sonido · stands de patrocinadores","Producción + patrocinadores","Todo listo ANTES de abrir puertas"],
    ["T + 0:00","Apertura de puertas · registro de asistentes con QR","Recepción + community","Capturar datos de CADA asistente"],
    ["T + 0:15","Bienvenida · ambientación musical · recorrido libre por instalaciones","Community","Primera impresión premium de las instalaciones"],
    ["T + 0:30","Rueda de prensa (área designada, máx. 20 min)","Andrés + invitados VIP","Declaraciones · fotos · preguntas de medios"],
    ["T + 1:00","Ceremonia oficial de inauguración","Andrés + autoridad + patrocinadores","Corte de listón · momento simbólico · foto grupal"],
    ["T + 1:15","Activaciones abiertas al público","Patrocinadores + equipo","Teqball · demos funcional · tour de canchas · stands"],
    ["T + 1:45","Presentación del proyecto (10 min al micrófono)","Andrés","Filosofía · método · visión · sin PowerPoint largo"],
    ["T + 2:00","Calentamiento visible para el público","Cuerpo técnico + jugadores","Genera expectativa antes del partido"],
    ["T + 2:30","⚽ PARTIDO DE LEYENDAS — 1er tiempo (20 min)","Árbitro + equipos","Clímax del evento · narración por micrófono"],
    ["T + 2:50","Medio tiempo: Clase Muestra de alumnos de la Academia","Entrenadores + alumnos","Muestra del método en vivo — cierra inscripciones"],
    ["T + 3:10","⚽ PARTIDO DE LEYENDAS — 2do tiempo (20 min)","Árbitro + equipos","Continúa la cobertura audiovisual"],
    ["T + 3:30","Premiación simbólica + foto oficial de todos los VIP","Andrés + invitados","Momento de contenido y relaciones públicas"],
    ["T + 3:45","Mesa de inscripciones activa · atención a prospectos","Ventas + recepción","Scripts Fase 6 · beneficio exclusivo del día"],
    ["T + 4:30","Cierre · despedida · cafetería activa","Todos","Las familias se quedan y consumen más tiempo"],
],aligns=["center","left","left","left"]))
C.append(pb())


# 4. Partido de leyendas
C.append(h1("4. Partido de Leyendas — Diseño y Logística"))
C.append(callout("El partido tiene 3 funciones: (1) convoca al público que no vendría por una inauguración, (2) genera contenido de alto valor (goles, jugadas, celebraciones), (3) posiciona a Andrés como protagonista activo — no solo como 'el dueño que da el discurso'.",kind="insight"))
C.append(table([
    ["","Equipo Chitiva (Local)","Equipo Leyendas (Visitante)"],
    ["Capitán","**Andrés Chitiva**","**'Conejo' Pérez** (u otro exjugador invitado)"],
    ["Jugadores","Andrés + alumnos destacados (cat. mayor) + invitados de Andrés","Conejo + alumnos + exjugadores locales/conocidos"],
    ["Uniforme","Uniforme oficial Academia Andrés Chitiva","Uniforme alterno / Complejo"],
],aligns=["left","left","left"]))
C.append(h3("Formato del partido"))
C.append(bullet("2 tiempos de 20 minutos (reloj corrido)."))
C.append(bullet("1 árbitro + 1 asistente (puede ser interno)."))
C.append(bullet("Narración/animación por micrófono durante el partido (community o narrador local invitado)."))
C.append(bullet("Resultado: no importa quién gane — importa el espectáculo y el contenido."))
C.append(p("",after=80))
C.append(callout('El "Conejo" Pérez y otros invitados son INTENCIÓN, no confirmación. Si alguno no puede asistir, el partido se sostiene con Andrés + alumnos + comunidad futbolística local. No depender de UNA sola figura para que el evento funcione.',kind="riesgo",label="RIESGO: INVITADOS EXTERNOS"))
C.append(pb())

# 5. Invitados
C.append(h1("5. Invitados y Convocatoria"))
C.append(h2("5.1 Invitados VIP (gestión inmediata)"))
C.append(table([
    ["Invitado","Rol","Quién gestiona","Plazo máximo"],
    ['"Conejo" Pérez','Capitán equipo visitante + vocería en prensa','Andrés (llamada directa)','72 h desde arranque'],
    ['Otros exfutbolistas','Jugadores partido de leyendas','Andrés + red de contactos','1 semana antes'],
    ['Autoridad municipal','Presencia institucional + corte de listón','Dirección del proyecto','10 días antes'],
    ['Voit (representante)','Stand + activación + foto oficial','Marketing','2 semanas antes'],
    ['Teqball (representante)','Activación + demo en vivo','Marketing','2 semanas antes'],
    ['Mazda Pachuca','Stand/vehículo + cruce de audiencias target','Marketing','2 semanas antes'],
    ['Aliados locales cuponera','Presencia y descuentos para inscritos','Marketing','1 semana antes'],
],aligns=["left","left","left","center"]))
C.append(h2("5.2 Prensa y Medios Locales"))
C.append(table([
    ["Medio","Tipo","Acción"],
    ["El Sol de Hidalgo · Milenio Hidalgo · La Crónica de Hidalgo","Prensa escrita","Enviar kit de prensa + invitación formal (mín. 10 días antes)"],
    ["Radio local deportiva","Audio","Invitar a Andrés para entrevista previa al evento"],
    ["Portales deportivos de Hidalgo","Digital","Boletín + kit de prensa"],
    ["Micro-influencers locales (papás/deporte)","Social","Invitación + acceso VIP al partido"],
    ["Fotógrafo / camarógrafo profesional","Cobertura propia","Contratar para el día — imprescindible"],
],aligns=["left","left","left"]))
C.append(h2("5.3 Público General"))
C.append(bullet("Familias piloto (Generación Fundadora) — invitación personal: son tus mejores embajadores ese día."))
C.append(bullet("Lista de espera y prospectos ya captados."))
C.append(bullet("Difusión en redes sociales (campaña de expectativa previa)."))
C.append(bullet("Grupos de Facebook locales (papás, colonias, deporte amateur)."))
C.append(bullet("Convocatoria abierta con aforo limitado (genera escasez y urgencia)."))
C.append(pb())

# 6. Patrocinadores
C.append(h1("6. Activaciones de Patrocinadores"))
C.append(callout("Los patrocinadores no son solo logos. Cada stand debe tener algo que el asistente pueda HACER o llevarse. Eso eleva la experiencia y genera contenido.",kind="recomendacion"))
C.append(table([
    ["Patrocinador","Activación sugerida","Resultado esperado"],
    ["**Teqball**","Mesa activa con retos para el público durante activaciones","Contenido viral · diferenciación · diversión"],
    ["**Voit**","Stand con producto · foto con balón oficial · posible sorteo","Visibilidad · leads para Voit · imagen premium"],
    ["**Mazda Pachuca**","Vehículo en entrada visible · activación / registro de interés","Halo premium · cruce de audiencias target"],
    ["**Aliados locales**","Mesa de cuponera · descuentos para inscritos ese día","Valor percibido para papás que se inscriban"],
],aligns=["left","left","left"]))
C.append(pb())


# 7. Captación en el evento
C.append(h1("7. Captación Comercial en el Evento"))
C.append(callout("El evento puede tener 100 personas maravilladas con el proyecto y perder el 70% de leads si no hay QR, formulario y seguimiento. La tecnología de captura es tan importante como el partido.",kind="riesgo",label="ERROR CRÍTICO A EVITAR"))
C.append(h2("7.1 Sistema de Registro (capturar a TODOS)"))
C.append(bullet("QR en la entrada → formulario simple: nombre, teléfono, WhatsApp, interés (Academia / Complejo / ambos)."))
C.append(bullet("Máximo 3 campos obligatorios. Rápido, no engorroso."))
C.append(bullet("Leads integrados directo a Clupik con etiqueta 'Evento Inauguración'."))
C.append(h2("7.2 Mesa de Inscripciones"))
C.append(bullet("Activa desde la ceremonia hasta el cierre del evento."))
C.append(bullet("Personal de ventas con script de la Fase 6 memorizado."))
C.append(bullet("**Beneficio exclusivo de inauguración** (a definir con Finanzas): inscripción reducida u otro beneficio puntual — SOLO disponible ese día."))
C.append(bullet("Lista de espera para los que 'quieren pensarlo': capturar igual y hacer seguimiento el día siguiente."))
C.append(h2("7.3 Seguimiento Post-Evento (primeras 24–48 h)"))
C.append(p("Mensaje WhatsApp personalizado a cada lead:"))
C.append(callout('"¡Fue un gusto tenerte en la inauguración! Soy [nombre], del equipo de la Academia Andrés Chitiva. ¿Quieres que te cuente cómo funciona o agendamos una Clase Muestra para [nombre del niño]?"',kind="recomendacion",label="TEMPLATE WHATSAPP POST-EVENTO"))
C.append(bullet("Meta: contactar al 100% de los leads en menos de 24 horas."))
C.append(bullet("Etiqueta en Clupik: 'Lead Inauguración → pendiente de seguimiento'."))
C.append(pb())

# 8. Producción y logística
C.append(h1("8. Producción y Logística"))
C.append(h2("8.1 Escenografía y Ambientación"))
C.append(table([
    ["Elemento","Propósito","Responsable"],
    ["Banner/lona de inauguración (fondo de fotos)","Foto oficial con patrocinadores + contenido de redes","Diseñador"],
    ["Señalética direccional","Experiencia premium desde la entrada","Diseñador"],
    ["Decoración de canchas (banderines, conos de marca)","Visual para video y fotos","Producción"],
    ["Sonido y micrófono","Ceremonias · partido · animación","Contratar o propio"],
    ["Mesa de registro (QR + materiales)","Captación de leads","Recepción"],
    ["Mesa de inscripciones","Cierre de ventas ese día","Ventas"],
    ["Cafetería activa","Ingresos + permanencia extendida de familias","Operación"],
],aligns=["left","left","left"]))
C.append(h2("8.2 Plan de Contenido del Evento (banco para 4–6 semanas)"))
C.append(callout("El evento es también una producción audiovisual. Define ANTES cuáles son las 5 piezas prioritarias y comunícaselas al audiovisual.",kind="recomendacion"))
C.append(table([
    ["Pieza de contenido","Uso posterior","Quién captura"],
    ["Highlight del partido (jugadas + goles)","Reels de alto alcance — semanas 1 y 2","Productor audiovisual"],
    ["Discurso de Andrés al micrófono","Reel de posicionamiento de marca","Audiovisual"],
    ["Reacciones de papás y niños","Testimonios orgánicos auténticos","Community (celular)"],
    ["Rueda de prensa","Clips para medios y redes","Audiovisual"],
    ["Activaciones Teqball/Voit","Contenido dinámico y divertido","Community"],
    ["Foto oficial grupal (todos los VIP)","Imagen de posicionamiento en prensa y redes","Fotógrafo profesional"],
    ["Tour completo de instalaciones (video continuo)","Video-tour definitivo para el embudo de ventas","Audiovisual"],
    ["Clase Muestra de alumnos (medio tiempo)","Pieza de autoridad del método","Audiovisual"],
],aligns=["left","left","left"]))
C.append(pb())

# 9. Comunicación y PR
C.append(h1("9. Comunicación y PR (Antes del Evento)"))
C.append(table([
    ["Semana","Acciones clave"],
    ["Semana −2","Confirmar invitados VIP · enviar kit de prensa a medios · publicar 'algo grande viene' (intriga) · abrir formulario de registro anticipado"],
    ["Semana −1","Publicar invitación oficial con partido de leyendas como gancho · stories de cuenta regresiva · confirmar activaciones de patrocinadores · Andrés publica en su cuenta personal · publicar en grupos de Facebook locales"],
    ["Día −1","Montaje de stands · prueba de sonido · brief final al equipo · story: 'mañana es el día'"],
    ["Día del evento","Cobertura en tiempo real · stories · reels cortos del partido · encuesta '¿ya estás aquí?'"],
    ["Post-evento 0–48 h","Highlight del partido · foto VIP oficial · clip del discurso de Andrés · contactar al 100% de leads"],
],aligns=["left","left"]))
C.append(pb())

# 10. Kit de prensa
C.append(h1("10. Kit de Prensa"))
C.append(h2("Contenido del kit (digital — email y WhatsApp)"))
C.append(table([
    ["#","Pieza","Descripción"],
    ["1","Boletín de prensa","1 página: título gancho · quién es Andrés · qué es el proyecto · qué pasará · cita de Andrés · contacto"],
    ["2","Foto oficial de Andrés","Alta resolución (300 dpi mínimo)"],
    ["3","Fotos de las instalaciones","2–3 fotos de alta resolución, las más impactantes"],
    ["4","Ficha técnica del proyecto","Academia + Complejo en bullet points"],
    ["5","Datos de contacto de prensa","Nombre, email, WhatsApp del responsable"],
],aligns=["center","left","left"]))
C.append(h2("Ángulos noticiosos para los medios"))
C.append(callout('"Exfutbolista profesional abre academia premium en la cuna del fútbol" · "El Conejo Pérez y Andrés Chitiva se enfrentan en partido de leyendas en Pachuca" · "Nuevo complejo deportivo apuesta por el desarrollo integral del futbolista joven"',kind="insight",label="TITULARES PROPUESTOS"))
C.append(pb())


# 11. Checklist maestro
C.append(h1("11. Checklist Maestro por Semanas"))
C.append(h2("✅ Arranque inmediato (esta semana — URGENTE)"))
for t in ["Definir y confirmar la fecha exacta del evento.",
          "Andrés contacta al 'Conejo' Pérez (llamada esta semana).",
          "Contactar a la autoridad municipal para la presencia institucional.",
          "Briefing a patrocinadores (Voit, Teqball, Mazda): confirmar asistencia y activación.",
          "Contratar fotógrafo / camarógrafo profesional para el día.",
          "Diseñar el banner/lona de inauguración.",
          "Crear el formulario de registro con QR.",
          "Publicar primera pieza de intriga en redes sociales."]: C.append(bullet("☐  "+t))
C.append(h2("✅ Semana −2"))
for t in ["Confirmar lista completa de jugadores de ambos equipos.",
          "Enviar kit de prensa a 10–15 medios locales.",
          "Diseño de señalética y materiales del evento.",
          "Publicar invitación oficial (partido de leyendas como gancho).",
          "Briefing al equipo de ventas: scripts, beneficio del día, proceso.",
          "Definir con Finanzas el beneficio exclusivo del día.",
          "Confirmar árbitro del partido.",
          "Planear menú/variedad de la cafetería para el día."]: C.append(bullet("☐  "+t))
C.append(h2("✅ Semana −1"))
for t in ["Follow-up a todos los medios invitados.",
          "Publicar reel de cuenta regresiva (1 por día).",
          "Andrés publica en su cuenta personal.",
          "Brief final a todo el equipo (roles, horarios, posiciones).",
          "Preparar carpetas de patrocinadores (materiales, stand, protocolo).",
          "Preparar mesa de inscripciones (formularios, beneficio, materiales)."]: C.append(bullet("☐  "+t))
C.append(h2("✅ Día −1"))
for t in ["Montaje completo (stands, señalética, canchas).",
          "Prueba de sonido y micrófonos.",
          "Brief final del equipo.",
          "Carga del formulario QR y prueba.",
          "Preparar cafetería con stock suficiente."]: C.append(bullet("☐  "+t))
C.append(h2("✅ Día del Evento"))
for t in ["Apertura de puertas con equipo en posición.",
          "Community cubriendo en tiempo real (stories + reels).",
          "Audiovisual grabando según las 5 piezas prioritarias definidas.",
          "Registro de CADA asistente sin excepción.",
          "Mesa de ventas activa desde la ceremonia.",
          "Leads etiquetados en WhatsApp/Clupik en tiempo real."]: C.append(bullet("☐  "+t))
C.append(pb())

# 12. Roles
C.append(h1("12. Roles y Responsabilidades el Día del Evento"))
C.append(table([
    ["Rol","Persona","Función principal"],
    ["Vocero / anfitrión principal","Andrés Chitiva","Rueda de prensa · ceremonia · partido · discurso"],
    ["MC / Animador","Contratar o community","Narrar el partido · animar el evento · mantener el flujo"],
    ["Registro de asistentes","1–2 personas de recepción","QR · captura de datos · bienvenida cálida"],
    ["Mesa de inscripciones","Personal de ventas","Scripts · cierre · beneficio del día"],
    ["Cobertura audiovisual","Productor audiovisual","Video del partido · discurso · activaciones · tour"],
    ["Cobertura en redes (tiempo real)","Community manager","Stories · reels cortos · interacción"],
    ["Fotografía profesional","Fotógrafo contratado","Fotos oficiales · prensa · banco de contenido"],
    ["Atención a patrocinadores","Marketing","Asegurar que cada stand esté correcto y activo"],
    ["Atención a medios","Responsable de PR","Acompañar a periodistas · facilitar entrevistas"],
    ["Coordinación general","Departamento de Marketing","Supervisar que todo el plan se ejecute según el programa"],
],aligns=["left","left","left"]))
C.append(pb())

# 13. Métricas y presupuesto
C.append(h1("13. Métricas de Éxito del Evento"))
C.append(table([
    ["Métrica","Meta"],
    ["Asistentes registrados","80–100 (con datos capturados)"],
    ["Leads capturados con QR","≥ 90% de los asistentes"],
    ["Inscripciones cerradas ese día","≥ 5–10 alumnos"],
    ["Notas de prensa publicadas","≥ 3 medios locales"],
    ["Alcance en redes el día del evento","Mayor alcance histórico de la cuenta"],
    ["Piezas de contenido producidas","≥ 10 piezas listas para redes"],
    ["Leads contactados en 48 h post-evento","100% (sin excepción)"],
],aligns=["left","center"]))
C.append(h1("14. Presupuesto Estimado (referencia para Finanzas)"))
C.append(table([
    ["Concepto","Estimado MXN","Notas"],
    ["Fotógrafo / camarógrafo profesional","$2,500 – $4,000","4–6 horas de cobertura — NO recortar"],
    ["Lona / banner de inauguración","$800 – $1,500","Según tamaño y calidad de impresión"],
    ["Señalética y materiales impresos","$500 – $1,000","Direccional + decoración de canchas"],
    ["Sonido y micrófono","$1,500 – $3,000","Si no es equipo propio"],
    ["MC / Animador","$1,500 – $2,500","Opcional si lo cubre el community"],
    ["Uniformes","$0","Ya se tienen confirmados"],
    ["Catering","$0","Se sustituye por cafetería activa"],
    ["Activaciones de patrocinadores","$0","Cada patrocinador trae su material"],
    ["**Total estimado**","**$6,800 – $12,000**","Rango según opciones elegidas"],
],aligns=["left","center","left"]))
C.append(p("",after=200))
C.append(callout("El fotógrafo/camarógrafo profesional es el único ítem que no se debe recortar. El contenido del evento es uno de los activos más valiosos: alimenta 4–6 semanas de redes y sirve como banco permanente de imagen del proyecto.",kind="recomendacion",label="NO RECORTAR"))
C.append(p("",after=200))
C.append(p("Departamento de Marketing — Proyecto Andrés Chitiva — 2026",align="center",color="808080",sz="18"))


# =================== ARMADO DEL DOCX ===================
body="".join(C)

document_xml=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 f'<w:body>{body}'
 '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
 '<w:pgSz w:w="12240" w:h="15840"/>'
 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1418" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
 '</w:sectPr></w:body></w:document>')

styles_xml=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="3" w:color="{MUSTARD}"/></w:pBdr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="{NAVY}"/><w:sz w:val="34"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="{NAVY}"/><w:sz w:val="27"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="{GOLD}"/><w:sz w:val="23"/></w:rPr></w:style>
</w:styles>'''

footer_xml='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="D9D9D9"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>
<w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">Plan de Inauguracion — Proyecto Andres Chitiva — Confidencial — Pagina </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''

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

doc_rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

import zipfile, os
with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml",ct); z.writestr("_rels/.rels",rels)
    z.writestr("word/document.xml",document_xml); z.writestr("word/styles.xml",styles_xml)
    z.writestr("word/footer1.xml",footer_xml); z.writestr("word/_rels/document.xml.rels",doc_rels)
print("OK ->",OUT,os.path.getsize(OUT),"bytes")
