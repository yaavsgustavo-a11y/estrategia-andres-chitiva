/**
 * ================================================================
 *  AUTOMATIZACIÓN — BASE LEADS CHITIVA
 *  Archivo: BASE_LEADS_CHITIVA_LIMPIO.xlsx → Google Sheets
 * ================================================================
 *
 *  INSTALACIÓN:
 *  1. En Google Sheets → Extensiones → Apps Script
 *  2. Pega todo este código y guarda (Ctrl+S)
 *  3. Ejecuta setupTriggers() UNA SOLA VEZ
 *  4. Autoriza los permisos
 * ================================================================
 */

const CFG = {
  HOJA_DB:    "LEADS_DB",
  HOJA_DASH:  "DASHBOARD",
  HOJA_PIVOT: "DATOS_PIVOT",
  HOJA_CFG:   "Config",

  // Columnas en LEADS_DB (1=A)
  C_ID:        1,
  C_FECHA:     2,
  C_NOMBRE:    3,
  C_TEL:       4,
  C_CORREO:    5,
  C_INTERES:   6,
  C_MEDIO:     7,
  C_FUENTE:    8,
  C_CANAL:     9,
  C_ETAPA:     10,
  C_RESULTADO: 11,
  C_COMENTARIO:12,
  C_SEMANA:    13,
  C_MES:       14,

  EMAIL: "tu@email.com",   // ← Cambia esto
  FILA_INICIO: 3,          // Los datos empiezan en fila 3
};

// ── Mapas de normalización ─────────────────────────────────────────────────
const FUENTE_MAP = {
  "FACEBOOK":"Facebook","FACEBOK":"Facebook","FACEBO":"Facebook",
  "INSTAGRAM":"Instagram",
  "TIKTOK":"TikTok","TIK TOK":"TikTok",
  "WHATSAPP":"WhatsApp","WHAS":"WhatsApp","WHS":"WhatsApp","WW":"WhatsApp",
  "CATELOGO DE WHATSAPP":"WhatsApp/Catálogo","CATALOGO DE WHATSAPP":"WhatsApp/Catálogo",
  "CONOCIDO":"Referido/Conocido","REFERIDO":"Referido/Conocido",
  "CERCANIA DE INSTALACIONES":"Cercanía","CERCANIA":"Cercanía",
  "VIDEO":"Orgánico","ACADEMIA":"Cercanía","OFICINAS":"Oficinas",
};

const RESULTADO_MAP = [
  { match: ["INSCRITO","COMPLETO"],                                                resultado: "Inscrito" },
  { match: ["PROCESO DE INSCRIPCI","PROCESO DE PAGO"],                             resultado: "En Proceso de Inscripción" },
  { match: ["CLASE MUESTRA","RESPUESTA DEPUES"],                                   resultado: "Clase Muestra" },
  { match: ["PARTIDO AMISTOSO","PROXIMA VISITA","TRATO","CONVENIO","RENTA","ENTRENAMIENTO DE PORTERO"], resultado: "Contactado/Activo" },
  { match: ["ELECCIÓN CURSO","ELECCION CURSO"],                                    resultado: "Interesado" },
  { match: ["INTERESAD"],                                                           resultado: "Interesado" },
  { match: ["ESPERANDO","ESPERA DE","EN ESPERA"],                                  resultado: "En Espera" },
  { match: ["DESINTERESADO","NO INSCRITO","SIN LIGAS","NO HAY AUN","NO ESTA SU","NO TENEMOS","NO TIENE EQUIPO"], resultado: "Desinteresado" },
  { match: ["DEJO DE CONTESTAR","NO RESPONDIO","NO VOLVIO","SIN RESPUESTA","SIN CONTESTACI"], resultado: "Sin Respuesta" },
];

const RESULTADO_COLORS = {
  "Inscrito":                  { bg: "#00B894", fg: "#FFFFFF" },
  "En Proceso de Inscripción": { bg: "#74B9FF", fg: "#FFFFFF" },
  "Clase Muestra":             { bg: "#A29BFE", fg: "#FFFFFF" },
  "Contactado/Activo":         { bg: "#0F3460", fg: "#FFFFFF" },
  "Interesado":                { bg: "#FDCB6E", fg: "#2D3436" },
  "En Espera":                 { bg: "#E17055", fg: "#FFFFFF" },
  "Desinteresado":             { bg: "#E94560", fg: "#FFFFFF" },
  "Sin Respuesta":             { bg: "#BDC3C7", fg: "#2D3436" },
};

// ── Funciones de normalización ─────────────────────────────────────────────
function normFuente(v) {
  if (!v || v.toString().trim() === "-") return "";
  const up = v.toString().trim().toUpperCase();
  if (FUENTE_MAP[up]) return FUENTE_MAP[up];
  if (up.includes("FACEBOOK"))  return "Facebook";
  if (up.includes("INSTAGRAM")) return "Instagram";
  if (up.includes("TIKTOK") || up.includes("TIK TOK")) return "TikTok";
  if (up.includes("WHATSAPP") || up.includes("WHAS") || up.includes("CATAL")) return "WhatsApp";
  if (up.includes("CONOCIDO") || up.includes("REFERIDO")) return "Referido/Conocido";
  if (up.includes("CERCANIA") || up.includes("CERCANÍA")) return "Cercanía";
  return v.toString().trim();
}

function normResultado(v) {
  if (!v || v.toString().trim() === "-") return "";
  const up = v.toString().trim().toUpperCase();
  for (const rule of RESULTADO_MAP) {
    if (rule.match.some(k => up.includes(k))) return rule.resultado;
  }
  return v.toString().trim();
}

function getSemana(fecha) {
  if (!(fecha instanceof Date)) return "";
  const d = new Date(Date.UTC(fecha.getFullYear(), fecha.getMonth(), fecha.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-S${week.toString().padStart(2,"0")}`;
}

function getMes(fecha) {
  if (!(fecha instanceof Date)) return "";
  const m = (fecha.getMonth()+1).toString().padStart(2,"0");
  return `${fecha.getFullYear()}-${m}`;
}

function applyResultadoColor(cell, resultado) {
  const c = RESULTADO_COLORS[resultado];
  if (c) {
    cell.setBackground(c.bg).setFontColor(c.fg).setFontWeight("bold");
  }
}

// ── onEdit — se ejecuta automáticamente al editar ─────────────────────────
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CFG.HOJA_DB) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < CFG.FILA_INICIO) return;

    // Auto-asignar ID si está vacío
    if (!sheet.getRange(row, CFG.C_ID).getValue()) {
      const lastId = getLastId(sheet);
      sheet.getRange(row, CFG.C_ID).setValue(lastId + 1);
    }

    // Fecha → calcular Semana y Mes
    if (col === CFG.C_FECHA) {
      const fecha = e.range.getValue();
      if (fecha instanceof Date) {
        sheet.getRange(row, CFG.C_SEMANA).setValue(getSemana(fecha));
        sheet.getRange(row, CFG.C_MES).setValue(getMes(fecha));
      }
    }

    // Fuente/Anuncio → normalizar
    if (col === CFG.C_FUENTE) {
      const fuente = normFuente(e.range.getValue());
      if (fuente) sheet.getRange(row, CFG.C_FUENTE).setValue(fuente);
    }

    // Resultado → normalizar + color
    if (col === CFG.C_RESULTADO) {
      const res = normResultado(e.range.getValue());
      if (res) {
        const cell = sheet.getRange(row, CFG.C_RESULTADO);
        cell.setValue(res);
        applyResultadoColor(cell, res);
      }
    }

    // Timestamp en dashboard
    const dash = e.source.getSheetByName(CFG.HOJA_DASH);
    if (dash) {
      dash.getRange("S2").setValue(new Date());
      dash.getRange("S2").setNumberFormat("dd/MM/yyyy HH:mm");
    }

  } catch (err) {
    Logger.log("onEdit error: " + err.message);
  }
}

function getLastId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CFG.FILA_INICIO) return 0;
  const ids = sheet.getRange(CFG.FILA_INICIO, CFG.C_ID, lastRow - CFG.FILA_INICIO + 1).getValues()
    .map(r => parseInt(r[0]) || 0);
  return Math.max(...ids, 0);
}

// ── Normalización masiva (ejecutar manualmente la primera vez) ─────────────
function normalizarTodo() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.HOJA_DB);
  if (!sheet) { Browser.msgBox("No se encontró LEADS_DB"); return; }

  const lastRow  = sheet.getLastRow();
  const numRows  = lastRow - CFG.FILA_INICIO + 1;
  if (numRows <= 0) return;

  const data = sheet.getRange(CFG.FILA_INICIO, 1, numRows, CFG.C_MES).getValues();

  const ids      = [], fuentes = [], resultados = [], semanas = [], meses = [];
  let idCounter = 0;

  for (let i = 0; i < data.length; i++) {
    const fila      = data[i];
    const fecha     = fila[CFG.C_FECHA - 1];
    const fuenteRaw = fila[CFG.C_FUENTE - 1];
    const resRaw    = fila[CFG.C_RESULTADO - 1];

    idCounter++;
    ids.push([idCounter]);
    fuentes.push([normFuente(fuenteRaw)]);
    resultados.push([normResultado(resRaw)]);
    semanas.push([fecha instanceof Date ? getSemana(fecha) : ""]);
    meses.push([fecha instanceof Date ? getMes(fecha) : ""]);
  }

  sheet.getRange(CFG.FILA_INICIO, CFG.C_ID,        numRows, 1).setValues(ids);
  sheet.getRange(CFG.FILA_INICIO, CFG.C_FUENTE,    numRows, 1).setValues(fuentes);
  sheet.getRange(CFG.FILA_INICIO, CFG.C_RESULTADO, numRows, 1).setValues(resultados);
  sheet.getRange(CFG.FILA_INICIO, CFG.C_SEMANA,    numRows, 1).setValues(semanas);
  sheet.getRange(CFG.FILA_INICIO, CFG.C_MES,       numRows, 1).setValues(meses);

  // Aplicar colores a resultado
  for (let i = 0; i < resultados.length; i++) {
    const res = resultados[i][0];
    if (res) applyResultadoColor(sheet.getRange(CFG.FILA_INICIO + i, CFG.C_RESULTADO), res);
  }

  Browser.msgBox(`✅ ${numRows} filas normalizadas correctamente.`);
}

// ── Reporte semanal ────────────────────────────────────────────────────────
function reporteSemanal() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.HOJA_DB);
  const pivot = ss.getSheetByName(CFG.HOJA_PIVOT);
  if (!sheet || !pivot) return;

  const total     = pivot.getRange("B3").getValue();
  const inscritos = pivot.getRange("B6").getValue();
  const tasa      = pivot.getRange("B11").getValue();
  const hoy       = new Date();
  const hace7     = new Date(hoy - 7*86400000);

  const data = sheet.getRange(CFG.FILA_INICIO, CFG.C_FECHA,
    sheet.getLastRow() - CFG.FILA_INICIO + 1, 1).getValues();
  const nuevos = data.filter(r => r[0] instanceof Date && r[0] >= hace7).length;

  MailApp.sendEmail({
    to:      CFG.EMAIL,
    subject: `📊 Reporte Semanal Leads — ${Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy")}`,
    body: `
REPORTE SEMANAL DE LEADS — Andres Chitiva
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total leads: ${total}
Nuevos esta semana: ${nuevos}
Inscritos acumulados: ${inscritos}
Tasa de conversión: ${(tasa*100).toFixed(1)}%
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Ver dashboard: ${ss.getUrl()}
    `,
  });
}

// ── Activar triggers ───────────────────────────────────────────────────────
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("onOpen").forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger("reporteSemanal").timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  Browser.msgBox("✅ Automatización activada\n• onEdit activo\n• Reporte semanal: lunes 8am\n\n⚠️ Cambia CFG.EMAIL por tu correo.");
}

// ── Menú personalizado ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Leads Chitiva")
    .addItem("🔄 Normalizar todos los datos",  "normalizarTodo")
    .addItem("📧 Enviar reporte ahora",         "reporteSemanal")
    .addSeparator()
    .addItem("⚙️ Activar automatización",       "setupTriggers")
    .addToUi();
}
