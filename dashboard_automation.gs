/**
 * ============================================================
 *  DASHBOARD COMERCIAL — AUTOMATIZACIÓN GOOGLE APPS SCRIPT
 *  Andres Chitiva · Estrategia Comercial
 * ============================================================
 *
 *  INSTALACIÓN:
 *  1. En Google Sheets → Extensiones → Apps Script
 *  2. Pega todo este código y guarda (Ctrl+S)
 *  3. Ejecuta setupTriggers() UNA SOLA VEZ para activar
 *     los triggers automáticos
 *  4. Autoriza los permisos cuando se solicite
 * ============================================================
 */

// ── CONFIGURACIÓN GLOBAL ──────────────────────────────────────────────────────
const CONFIG = {
  HOJA_DB:        "LEADS_DB",
  HOJA_DASHBOARD: "DASHBOARD",
  HOJA_PIVOT:     "DATOS_PIVOT",
  HOJA_CONFIG:    "Config",

  // Columnas en LEADS_DB (1 = col A)
  COL_FECHA:     2,   // B
  COL_NOMBRE:    3,   // C
  COL_ANUNCIO:   8,   // H  ← Anuncio/Fuente original
  COL_RESULTADO: 11,  // K  ← Resultado del seguimiento
  COL_FUENTE:    13,  // M  ← Fuente Limpia (calculada)
  COL_ESTATUS:   14,  // N  ← Estatus Final (calculada)
  COL_SEMANA:    15,  // O  ← Semana ISO
  COL_MES:       16,  // P  ← Mes YYYY-MM

  // Email para reporte semanal (cambia este valor)
  EMAIL_REPORTE: "tu@email.com",

  // Fila donde empiezan los datos (fila 1 = encabezados)
  FILA_INICIO:   2,
};

// ── MAPEOS ────────────────────────────────────────────────────────────────────
const FUENTE_MAP = {
  "FACEBOOK":                    "Facebook",
  "INSTAGRAM":                   "Instagram",
  "TIK TOK":                     "TikTok",
  "TIKTOK":                      "TikTok",
  "TIKTOK":                      "TikTok",
  "CONOCIDO":                    "Referido/Conocido",
  "REFERIDO":                    "Referido/Conocido",
  "WHATSAPP":                    "WhatsApp",
  "CATELOGO DE WHATSAPP":        "WhatsApp",
  "CATALOGO DE WHATSAPP":        "WhatsApp",
  "CERCANIA DE INSTALACIONES":   "Cercanía",
  "CERCANIA":                    "Cercanía",
  "VIDEO":                       "Orgánico",
};

const ESTATUS_MAP = [
  { palabras: ["INSCRITO"],                              resultado: "6 - Inscrito" },
  { palabras: ["PROCESO DE INSCRIPCI", "PROCESO DE INSCRIPCION"],
                                                         resultado: "5 - En proceso inscripción" },
  { palabras: ["CLASE MUESTRA"],                         resultado: "4 - Clase muestra" },
  { palabras: ["PARTIDO AMISTOSO", "TRATO", "CONVENIO", "RENTA"],
                                                         resultado: "3 - Contactado/Activo" },
  { palabras: ["INTERESADO", "ESPERANDO", "EN ESPERA", "ESPERA", "ELECCION", "ELECCIÓN"],
                                                         resultado: "2 - Interesado" },
  { palabras: ["DESINTERESADO", "NO INSCRITO", "SIN LIGA"],
                                                         resultado: "7 - Perdido" },
  { palabras: ["SIN RESPUESTA", "SIN CONTESTACI", "DEJO DE CONTESTAR",
               "NO RESPONDIO", "NO VOLVIO", "NO CONTESTO"],
                                                         resultado: "8 - Sin respuesta" },
];

// ── FUNCIONES DE NORMALIZACIÓN ────────────────────────────────────────────────

/**
 * Convierte el valor de Anuncio/Fuente a una fuente limpia y estandarizada
 */
function normalizarFuente(valor) {
  if (!valor || valor.toString().trim() === "-" || valor.toString().trim() === "") {
    return "Sin datos";
  }
  const v = valor.toString().trim().toUpperCase();
  if (FUENTE_MAP[v]) return FUENTE_MAP[v];
  if (v.includes("FACEBOOK"))   return "Facebook";
  if (v.includes("INSTAGRAM"))  return "Instagram";
  if (v.includes("TIKTOK") || v.includes("TIK TOK")) return "TikTok";
  if (v.includes("WHATSAPP") || v.includes("CATAL")) return "WhatsApp";
  if (v.includes("CONOCIDO") || v.includes("REFERIDO")) return "Referido/Conocido";
  if (v.includes("CERCANIA") || v.includes("CERCANÍA")) return "Cercanía";
  return "Otro";
}

/**
 * Convierte el resultado del seguimiento al estatus del embudo
 */
function normalizarEstatus(valor) {
  if (!valor || valor.toString().trim() === "-" || valor.toString().trim() === "") {
    return "1 - Lead/Informes";
  }
  const v = valor.toString().trim().toUpperCase();
  for (const regla of ESTATUS_MAP) {
    if (regla.palabras.some(p => v.includes(p))) {
      return regla.resultado;
    }
  }
  return "1 - Lead/Informes";
}

/**
 * Obtiene el número de semana ISO (YYYY-Www)
 */
function getSemanaISO(fecha) {
  if (!(fecha instanceof Date)) return "";
  const d = new Date(Date.UTC(fecha.getFullYear(), fecha.getMonth(), fecha.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${week.toString().padStart(2, "0")}`;
}

/**
 * Formatea la fecha como YYYY-MM para agrupación mensual
 */
function getMesAnio(fecha) {
  if (!(fecha instanceof Date)) return "";
  const y = fecha.getFullYear();
  const m = (fecha.getMonth() + 1).toString().padStart(2, "0");
  return `${y}-${m}`;
}

// ── TRIGGER PRINCIPAL — onEdit ────────────────────────────────────────────────

/**
 * Se ejecuta automáticamente cada vez que el usuario edita la hoja.
 * Normaliza fuente, estatus, semana y mes cuando se modifican las columnas clave.
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CONFIG.HOJA_DB) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (row <= CONFIG.FILA_INICIO - 1) return; // No tocar encabezados

    // ── Fecha editada → recalcular Semana y Mes ──────────────────────────
    if (col === CONFIG.COL_FECHA) {
      const fecha = e.range.getValue();
      if (fecha instanceof Date) {
        sheet.getRange(row, CONFIG.COL_SEMANA).setValue(getSemanaISO(fecha));
        sheet.getRange(row, CONFIG.COL_MES).setValue(getMesAnio(fecha));
      }
    }

    // ── Anuncio/Fuente editado → recalcular Fuente Limpia ───────────────
    if (col === CONFIG.COL_ANUNCIO) {
      const fuente = normalizarFuente(e.range.getValue());
      sheet.getRange(row, CONFIG.COL_FUENTE).setValue(fuente);
    }

    // ── Resultado editado → recalcular Estatus Final + color ────────────
    if (col === CONFIG.COL_RESULTADO) {
      const estatus = normalizarEstatus(e.range.getValue());
      const cell = sheet.getRange(row, CONFIG.COL_ESTATUS);
      cell.setValue(estatus);
      aplicarColorEstatus(cell, estatus);
    }

    // ── Actualizar timestamp en Dashboard ───────────────────────────────
    actualizarTimestamp(e.source);

  } catch (err) {
    Logger.log("Error en onEdit: " + err.message);
  }
}

/**
 * Aplica color de fondo a la celda de estatus según el valor
 */
function aplicarColorEstatus(cell, estatus) {
  const colores = {
    "6 - Inscrito":               { bg: "#00B894", fg: "#FFFFFF" },
    "5 - En proceso inscripción": { bg: "#74B9FF", fg: "#FFFFFF" },
    "4 - Clase muestra":          { bg: "#A29BFE", fg: "#FFFFFF" },
    "3 - Contactado/Activo":      { bg: "#0F3460", fg: "#FFFFFF" },
    "2 - Interesado":             { bg: "#FDCB6E", fg: "#2D3436" },
    "7 - Perdido":                { bg: "#E94560", fg: "#FFFFFF" },
    "8 - Sin respuesta":          { bg: "#E17055", fg: "#FFFFFF" },
    "1 - Lead/Informes":          { bg: "#F0F2F5", fg: "#2D3436" },
  };
  const c = colores[estatus];
  if (c) {
    cell.setBackground(c.bg).setFontColor(c.fg).setFontWeight("bold");
  }
}

/**
 * Actualiza la celda de última actualización en el DASHBOARD
 */
function actualizarTimestamp(spreadsheet) {
  const dash = spreadsheet.getSheetByName(CONFIG.HOJA_DASHBOARD);
  if (!dash) return;
  const cell = dash.getRange("C2");
  cell.setValue(new Date());
  cell.setNumberFormat("dd/MM/yyyy HH:mm");
}


// ── NORMALIZACIÓN MASIVA ──────────────────────────────────────────────────────

/**
 * Normaliza TODAS las filas existentes en LEADS_DB de una sola vez.
 * Ejecuta esta función manualmente desde Apps Script cuando importes datos nuevos
 * o la primera vez que instales el script.
 *
 * CÓMO EJECUTAR: Apps Script → selecciona normalizarTodosLosLeads → ▶ Ejecutar
 */
function normalizarTodosLosLeads() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.HOJA_DB);
  if (!sheet) { Browser.msgBox("No se encontró la hoja LEADS_DB"); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.FILA_INICIO) { Browser.msgBox("No hay datos en LEADS_DB"); return; }

  const totalRows = lastRow - CONFIG.FILA_INICIO + 1;
  let procesados = 0;

  // Leer todos los valores en una sola llamada (más eficiente)
  const rangeData = sheet.getRange(CONFIG.FILA_INICIO, 1, totalRows, 16).getValues();

  // Preparar arrays de salida
  const fuentes  = [];
  const estatus  = [];
  const semanas  = [];
  const meses    = [];

  for (let i = 0; i < rangeData.length; i++) {
    const fila      = rangeData[i];
    const fecha     = fila[CONFIG.COL_FECHA - 1];       // col B (idx 1)
    const nombre    = fila[CONFIG.COL_NOMBRE - 1];      // col C (idx 2)
    const anuncio   = fila[CONFIG.COL_ANUNCIO - 1];     // col H (idx 7)
    const resultado = fila[CONFIG.COL_RESULTADO - 1];   // col K (idx 10)

    // Solo procesar filas con nombre real
    if (!nombre || nombre.toString().trim() === "" || nombre.toString().trim() === "-") {
      fuentes.push([""]);
      estatus.push([""]);
      semanas.push([""]);
      meses.push([""]);
      continue;
    }

    fuentes.push([normalizarFuente(anuncio)]);
    estatus.push([normalizarEstatus(resultado)]);
    semanas.push([fecha instanceof Date ? getSemanaISO(fecha) : ""]);
    meses.push([fecha instanceof Date ? getMesAnio(fecha) : ""]);
    procesados++;
  }

  // Escribir resultados en bloque (una sola llamada por columna)
  sheet.getRange(CONFIG.FILA_INICIO, CONFIG.COL_FUENTE,  totalRows, 1).setValues(fuentes);
  sheet.getRange(CONFIG.FILA_INICIO, CONFIG.COL_ESTATUS, totalRows, 1).setValues(estatus);
  sheet.getRange(CONFIG.FILA_INICIO, CONFIG.COL_SEMANA,  totalRows, 1).setValues(semanas);
  sheet.getRange(CONFIG.FILA_INICIO, CONFIG.COL_MES,     totalRows, 1).setValues(meses);

  // Aplicar colores a columna Estatus
  for (let i = 0; i < estatus.length; i++) {
    if (estatus[i][0]) {
      const cell = sheet.getRange(CONFIG.FILA_INICIO + i, CONFIG.COL_ESTATUS);
      aplicarColorEstatus(cell, estatus[i][0]);
    }
  }

  actualizarTimestamp(ss);
  Browser.msgBox(`✅ Normalización completada\n${procesados} leads procesados`);
}

// ── REPORTE SEMANAL POR EMAIL ─────────────────────────────────────────────────

/**
 * Genera y envía un reporte semanal por email con las métricas clave.
 * Para activarlo automáticamente: configura un trigger semanal con setupTriggers()
 */
function enviarReporteSemanal() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.HOJA_DB);
  if (!sheet) return;

  const lastRow   = sheet.getLastRow();
  const totalRows = lastRow - CONFIG.FILA_INICIO + 1;
  if (totalRows <= 0) return;

  const data = sheet.getRange(CONFIG.FILA_INICIO, 1, totalRows, CONFIG.COL_ESTATUS).getValues();

  let total = 0, inscritos = 0, perdidos = 0, interesados = 0, sinResp = 0;
  const fuenteCount = {};
  const hoy = new Date();
  const hace7dias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  let nuevosEstaSemana = 0;

  for (const fila of data) {
    const nombre  = fila[CONFIG.COL_NOMBRE - 1];
    const fecha   = fila[CONFIG.COL_FECHA - 1];
    const estatus = fila[CONFIG.COL_ESTATUS - 1];
    const fuente  = fila[CONFIG.COL_FUENTE - 1];

    if (!nombre || nombre.toString().trim() === "" || nombre.toString().trim() === "-") continue;
    total++;

    if (fecha instanceof Date && fecha >= hace7dias) nuevosEstaSemana++;

    switch (estatus) {
      case "6 - Inscrito":               inscritos++;   break;
      case "7 - Perdido":                perdidos++;    break;
      case "2 - Interesado":             interesados++; break;
      case "8 - Sin respuesta":          sinResp++;     break;
    }

    if (fuente) fuenteCount[fuente] = (fuenteCount[fuente] || 0) + 1;
  }

  const tasa = total > 0 ? ((inscritos / total) * 100).toFixed(1) : "0.0";
  const topFuente = Object.entries(fuenteCount).sort((a, b) => b[1] - a[1])[0];

  const asunto = `📊 Reporte Semanal de Leads — ${Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy")}`;

  const cuerpo = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 REPORTE SEMANAL DE LEADS
Estrategia Comercial — Andres Chitiva
Generado: ${Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📈 RESUMEN GENERAL
  • Total leads acumulados:    ${total}
  • Nuevos esta semana:        ${nuevosEstaSemana}
  • Inscritos totales:         ${inscritos}
  • Tasa de conversión:        ${tasa}%

🔄 ESTADO DEL PIPELINE
  • Interesados activos:       ${interesados}
  • Sin respuesta:             ${sinResp}
  • Perdidos:                  ${perdidos}

📡 MEJOR FUENTE
  • ${topFuente ? topFuente[0] + ": " + topFuente[1] + " leads" : "Sin datos"}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔗 Ver dashboard completo en Google Sheets
${ss.getUrl()}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`;

  MailApp.sendEmail({
    to:      CONFIG.EMAIL_REPORTE,
    subject: asunto,
    body:    cuerpo,
  });

  Logger.log("Reporte enviado a: " + CONFIG.EMAIL_REPORTE);
}

// ── ALERTA DE LEADS SIN SEGUIMIENTO ──────────────────────────────────────────

/**
 * Detecta leads "Interesados" que llevan más de 3 días sin actualización
 * y envía un recordatorio por email.
 * Ideal para configurar como trigger diario.
 */
function alertaLeadsSinSeguimiento() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.HOJA_DB);
  if (!sheet) return;

  const lastRow   = sheet.getLastRow();
  const totalRows = lastRow - CONFIG.FILA_INICIO + 1;
  if (totalRows <= 0) return;

  const data = sheet.getRange(CONFIG.FILA_INICIO, 1, totalRows, CONFIG.COL_ESTATUS).getValues();
  const hoy = new Date();
  const limite = new Date(hoy.getTime() - 3 * 24 * 60 * 60 * 1000); // 3 días

  const leadsSinSeguimiento = [];

  for (const fila of data) {
    const fecha   = fila[CONFIG.COL_FECHA - 1];
    const nombre  = fila[CONFIG.COL_NOMBRE - 1];
    const estatus = fila[CONFIG.COL_ESTATUS - 1];

    if (!nombre || nombre.toString().trim() === "" || nombre.toString().trim() === "-") continue;
    if (estatus !== "2 - Interesado" && estatus !== "3 - Contactado/Activo") continue;
    if (fecha instanceof Date && fecha <= limite) {
      const dias = Math.floor((hoy - fecha) / (1000 * 60 * 60 * 24));
      leadsSinSeguimiento.push({ nombre: nombre.toString().trim(), dias, estatus });
    }
  }

  if (leadsSinSeguimiento.length === 0) return; // Nada que alertar

  let lista = leadsSinSeguimiento
    .sort((a, b) => b.dias - a.dias)
    .map(l => `  • ${l.nombre} — ${l.dias} días — ${l.estatus}`)
    .join("\n");

  MailApp.sendEmail({
    to:      CONFIG.EMAIL_REPORTE,
    subject: `⚠️ ${leadsSinSeguimiento.length} leads sin seguimiento — Acción requerida`,
    body: `
⚠️ ALERTA: LEADS SIN SEGUIMIENTO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Estos leads llevan más de 3 días sin actualización:

${lista}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📌 Acción recomendada: contactar hoy para no perder la oportunidad.
🔗 ${ss.getUrl()}
`,
  });
}

// ── CONFIGURACIÓN DE TRIGGERS AUTOMÁTICOS ────────────────────────────────────

/**
 * ⚙️ EJECUTA ESTA FUNCIÓN UNA SOLA VEZ para activar la automatización completa.
 *
 * Configura:
 *  - Trigger de edición (onEdit automático)
 *  - Reporte semanal (lunes 8am)
 *  - Alerta diaria de leads sin seguimiento (cada día 9am)
 */
function setupTriggers() {
  // Eliminar triggers anteriores para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    ScriptApp.deleteTrigger(t);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. onEdit instalable (más confiable que el simple)
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // 2. Reporte semanal — lunes 8:00am
  ScriptApp.newTrigger("enviarReporteSemanal")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  // 3. Alerta diaria — 9:00am
  ScriptApp.newTrigger("alertaLeadsSinSeguimiento")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Browser.msgBox(
    "✅ Automatización activada\n\n" +
    "• onEdit: activo al guardar cambios\n" +
    "• Reporte semanal: lunes 8:00am\n" +
    "• Alerta seguimiento: diario 9:00am\n\n" +
    "⚠️ Recuerda cambiar CONFIG.EMAIL_REPORTE con tu email."
  );
}

/**
 * Crea un menú personalizado en Google Sheets para acceso rápido
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Dashboard Comercial")
    .addItem("🔄 Normalizar todos los leads", "normalizarTodosLosLeads")
    .addSeparator()
    .addItem("📧 Enviar reporte ahora",        "enviarReporteSemanal")
    .addItem("⚠️ Verificar leads sin seguimiento", "alertaLeadsSinSeguimiento")
    .addSeparator()
    .addItem("⚙️ Activar automatización",      "setupTriggers")
    .addToUi();
}
