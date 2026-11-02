// Semáforo tributario – servidor en Google Apps Script (ligado a la hoja "Semáforo Tributario Digital")
// Tokens de las API: Configuración del proyecto > Propiedades del script (nunca en el código).

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Semáforo Tributario')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- RUC ----------
function rucValido_(ruc) {
  ruc = String(ruc);
  if (!/^(10|15|16|17|20)\d{9}$/.test(ruc)) return false;
  var pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2], suma = 0;
  for (var i = 0; i < 10; i++) suma += pesos[i] * Number(ruc[i]);
  var dv = 11 - (suma % 11);
  if (dv === 10) dv = 0; else if (dv === 11) dv = 1;
  return dv === Number(ruc[10]);
}

// Consulta el padrón de la SUNAT mediante la API de terceros (el token no se expone al navegador)
function consultarRuc(ruc) {
  if (!rucValido_(ruc)) throw new Error('RUC inválido');
  var token = PropertiesService.getScriptProperties().getProperty('DECOLECTA_TOKEN');
  if (!token) return null;

  var resp = UrlFetchApp.fetch('https://api.decolecta.com/v1/sunat/ruc?numero=' + ruc, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('la API respondió ' + resp.getResponseCode());

  var d = JSON.parse(resp.getContentText());
  return {
    razon: d.razon_social || d.razonSocial || d.nombre || '',
    estado: String(d.estado || '').toUpperCase(),
    condicion: String(d.condicion || '').toUpperCase()
  };
}

// ---------- Obligaciones marcadas como presentadas (confirmación manual del usuario) ----------
var HOJA_PRESENTADAS = 'Presentadas';
var CLAVE_VALIDA = /^[a-z0-9]+:\d{4}(-\d{2})?$/; // p. ej. "f621:2026-10" o "anual:2025"

function hojaPresentadas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName(HOJA_PRESENTADAS);
  if (!h) {
    h = ss.insertSheet(HOJA_PRESENTADAS);
    h.getRange('A:B').setNumberFormat('@'); // RUC y clave como texto
    h.appendRow(['RUC', 'Obligación:periodo', 'Marcada el']);
    h.setFrozenRows(1);
  }
  return h;
}

// Devuelve { "f621:2026-10": "2026-11-02", ... } para el RUC
function obtenerMarcas(ruc) {
  if (!rucValido_(ruc)) throw new Error('RUC inválido');
  var filas = hojaPresentadas_().getDataRange().getValues();
  var tz = Session.getScriptTimeZone(), marcas = {};
  for (var i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === ruc) {
      var f = filas[i][2];
      marcas[String(filas[i][1])] = f instanceof Date ? Utilities.formatDate(f, tz, 'yyyy-MM-dd') : String(f);
    }
  }
  return marcas;
}

// Marca (presentada = true) o desmarca una obligación; devuelve las marcas actualizadas
function marcarPresentada(ruc, clave, presentada) {
  if (!rucValido_(ruc)) throw new Error('RUC inválido');
  if (!CLAVE_VALIDA.test(clave)) throw new Error('Obligación inválida');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var h = hojaPresentadas_(), filas = h.getDataRange().getValues(), fila = -1;
    for (var i = 1; i < filas.length; i++) {
      if (String(filas[i][0]) === ruc && String(filas[i][1]) === clave) { fila = i + 1; break; }
    }
    if (presentada && fila === -1) h.appendRow([ruc, clave, new Date()]);
    if (!presentada && fila !== -1) h.deleteRow(fila);
  } finally {
    lock.releaseLock();
  }
  return obtenerMarcas(ruc);
}
