// Semáforo tributario – servidor en Google Apps Script (ligado a la hoja "Semáforo Tributario Digital")
// Tokens de las API: Configuración del proyecto > Propiedades del script (nunca en el código).
// Las fechas y obligaciones vienen de Calendario.gs (generado desde index.html con "node sincronizar.js").

// La página se publica en GitHub Pages (propiedad URL_SITIO) y llama a este script como API (doPost).
// Este doGet queda como respaldo: sirve las mismas páginas desde Apps Script.
function doGet(e) {
  var p = (e && e.parameter) || {};
  return HtmlService.createHtmlOutputFromFile(p.accion ? 'accion' : 'index')
    .setTitle('Semáforo Tributario')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// API: recibe {"fn": "...", "args": [...]} y responde {"ok": true, "data": ...} o {"ok": false, "error": "..."}
var API_PUBLICA = {
  obtenerMarcas: function (a) { return obtenerMarcas(a[0]); },
  marcarPresentada: function (a) { return marcarPresentada(a[0], a[1], a[2]); },
  consultarRuc: function (a) { return consultarRuc(a[0]); },
  registrarAviso: function (a) { return registrarAviso(a[0], a[1], a[2], a[3]); },
  describirAccion: function (a) { return describirAccion(a[0], a[1], a[2]); },
  ejecutarAccion: function (a) { return ejecutarAccion(a[0], a[1], a[2]); }
};

function doPost(e) {
  var r;
  try {
    var q = JSON.parse(e.postData.contents);
    if (!API_PUBLICA.hasOwnProperty(q.fn)) throw new Error('Función no permitida');
    r = { ok: true, data: API_PUBLICA[q.fn](q.args || []) };
  } catch (x) {
    r = { ok: false, error: String(x && x.message || x) };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- RUC ----------
function rucValido_(ruc) {
  return typeof ruc === 'string' && errorRuc(ruc) === '';
}

// ---------- Padrón de la SUNAT: APIs de terceros en cascada, con caché en la hoja "Padron" ----------
// Cada API necesita su token en Propiedades del script (el token nunca llega al navegador).
// Se prueban en orden; si una no tiene token, falla o se quedó sin cupo, se pasa a la siguiente.
// Para sumar otra API: agregar un objeto aquí con su URL y cómo leer su respuesta.
var APIS_RUC = [
  {
    nombre: 'Decolecta', propiedad: 'DECOLECTA_TOKEN', // plan gratis: 100 consultas al mes
    url: function (ruc) { return 'https://api.decolecta.com/v1/sunat/ruc?numero=' + ruc; },
    leer: function (d) {
      return { razon: d.razon_social, estado: d.estado, condicion: d.condicion,
               agenteRetencion: d.es_agente_retencion === true, buenContribuyente: d.es_buen_contribuyente === true };
    }
  }
];
var DIAS_CACHE_RUC = 15; // los datos del padrón casi no cambian

function hojaPadron_() {
  return hoja_('Padron', ['RUC', 'Razón social', 'Estado', 'Condición', 'Agente de retención', 'Buen contribuyente', 'Fuente', 'Consultado el']);
}

// Devuelve { razon, estado, condicion, agenteRetencion, buenContribuyente, fuente, consultado } o null
function consultarRuc(ruc) {
  if (!rucValido_(ruc)) throw new Error('RUC inválido');
  var h = hojaPadron_(), fila = buscarFila_(h, 0, ruc);
  var guardado = fila ? {
    razon: String(fila.v[1]), estado: String(fila.v[2]), condicion: String(fila.v[3]),
    agenteRetencion: String(fila.v[4]).toLowerCase() === 'true', buenContribuyente: String(fila.v[5]).toLowerCase() === 'true',
    fuente: String(fila.v[6]), consultado: diaDe_(fila.v[7])
  } : null;
  if (guardado && new Date() - aFecha_(fila.v[7]) < DIAS_CACHE_RUC * 86400000) return guardado;

  var nuevo = consultarApis_(ruc);
  if (!nuevo) return guardado; // si ninguna API responde, se usa el dato antiguo (si existe)

  var cuando = ahora_();
  var valores = [ruc, nuevo.razon, nuevo.estado, nuevo.condicion, nuevo.agenteRetencion, nuevo.buenContribuyente, nuevo.fuente, cuando];
  if (fila) h.getRange(fila.n, 1, 1, valores.length).setValues([valores]); else agregarFila_(h, valores);
  nuevo.consultado = cuando.slice(0, 10);
  return nuevo;
}

function consultarApis_(ruc) {
  var props = PropertiesService.getScriptProperties();
  for (var i = 0; i < APIS_RUC.length; i++) {
    var api = APIS_RUC[i], token = props.getProperty(api.propiedad);
    if (!token) continue;
    try {
      var resp = UrlFetchApp.fetch(api.url(ruc), { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
      contarUso_(api.nombre);
      if (resp.getResponseCode() !== 200) {
        Logger.log(api.nombre + ' respondió ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 200));
        continue;
      }
      var d = api.leer(JSON.parse(resp.getContentText()));
      if (!d.razon) continue;
      d.razon = String(d.razon); d.estado = String(d.estado || '').toUpperCase(); d.condicion = String(d.condicion || '').toUpperCase();
      d.fuente = api.nombre;
      return d;
    } catch (e) {
      Logger.log(api.nombre + ' falló: ' + e);
    }
  }
  return null;
}

// Cuenta las consultas de cada API por mes (Propiedades del script: USO_Decolecta_2026-11, ...)
function contarUso_(nombre) {
  var props = PropertiesService.getScriptProperties();
  var clave = 'USO_' + nombre + '_' + Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM');
  props.setProperty(clave, String(Number(props.getProperty(clave) || 0) + 1));
}

// Ejecutar desde el editor para comprobar la conexión con las APIs (usa 1 consulta de cupo)
function probarConsultaRuc() {
  var ruc = '20100070970'; // cámbialo por un RUC real de prueba
  var fila = buscarFila_(hojaPadron_(), 0, ruc);
  if (fila) hojaPadron_().deleteRow(fila.n); // fuerza la consulta a la API
  Logger.log(JSON.stringify(consultarRuc(ruc), null, 2));
}

// ---------- Hojas ----------
// Las fechas se guardan como texto en hora de Lima ("2026-11-02 08:00"): así Sheets no las cambia de zona horaria
function ahora_() {
  return Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm');
}
function aFecha_(v) { // celda (texto o fecha antigua) → Date
  if (v instanceof Date) return v;
  var s = String(v || '');
  return s ? new Date(s.slice(0, 10) + 'T' + (s.slice(11, 16) || '00:00') + ':00-05:00') : null; // Perú: UTC-5 todo el año
}
function diaDe_(v) { // celda → "2026-11-02"
  return v instanceof Date ? Utilities.formatDate(v, 'America/Lima', 'yyyy-MM-dd') : String(v || '').slice(0, 10);
}
function agregarFila_(h, valores) { // como appendRow, pero todo como texto
  h.getRange(h.getLastRow() + 1, 1, 1, valores.length).setNumberFormat('@').setValues([valores]);
}

function hoja_(nombre, encabezados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Las fechas guardadas en la hoja deben leerse en hora de Lima, igual que el script
  if (ss.getSpreadsheetTimeZone() !== 'America/Lima') ss.setSpreadsheetTimeZone('America/Lima');
  var h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, h.getMaxRows(), 2).setNumberFormat('@'); // las dos primeras columnas como texto
    h.appendRow(encabezados);
    h.setFrozenRows(1);
  }
  return h;
}

// ---------- Obligaciones marcadas como presentadas (confirmación manual del usuario) ----------
var CLAVE_VALIDA = /^[a-z0-9]+:\d{4}(-\d{2})?$/; // p. ej. "f621:2026-10" o "anual:2025"

function hojaPresentadas_() {
  return hoja_('Presentadas', ['RUC', 'Obligación:periodo', 'Marcada el']);
}

// Todas las marcas: { ruc: { "f621:2026-10": "2026-11-02", ... } }
function todasLasMarcas_() {
  var filas = hojaPresentadas_().getDataRange().getValues(), r = {};
  for (var i = 1; i < filas.length; i++) {
    var ruc = String(filas[i][0]);
    (r[ruc] = r[ruc] || {})[String(filas[i][1])] = diaDe_(filas[i][2]);
  }
  return r;
}

function obtenerMarcas(ruc) {
  if (!rucValido_(ruc)) throw new Error('RUC inválido');
  return todasLasMarcas_()[ruc] || {};
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
    if (presentada && fila === -1) agregarFila_(h, [ruc, clave, ahora_()]);
    if (!presentada && fila !== -1) h.deleteRow(fila);
  } finally {
    lock.releaseLock();
  }
  return obtenerMarcas(ruc);
}

// ---------- Carpetas en el Drive del proyecto ----------
function carpeta_(nombre) {
  var it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

// ---------- Respaldo semanal de la hoja (copia .xlsx en Drive) ----------
var CARPETA_RESPALDOS = 'Semáforo Tributario - Respaldos';
var RESPALDOS_A_GUARDAR = 12; // los más antiguos van a la papelera

function respaldarDatos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('No se pudo exportar la hoja: ' + resp.getResponseCode());

  var carpeta = carpeta_(CARPETA_RESPALDOS);
  carpeta.createFile(resp.getBlob().setName('Respaldo semáforo ' + ahora_().replace(':', 'h') + '.xlsx'));
  var archivos = [], it = carpeta.getFiles();
  while (it.hasNext()) archivos.push(it.next());
  archivos.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); })
    .slice(RESPALDOS_A_GUARDAR).forEach(function (f) { f.setTrashed(true); });
  Logger.log('Respaldo guardado en la carpeta "' + CARPETA_RESPALDOS + '" de tu Drive.');
}

// Ejecutar UNA vez desde el editor: programa el respaldo de cada lunes a las 7 a. m. y hace el primero ahora
function instalarRespaldoSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'respaldarDatos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('respaldarDatos').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7).inTimezone('America/Lima').create();
  respaldarDatos();
  Logger.log('Respaldo semanal programado: cada lunes a las 7 a. m.');
}

// ---------- Avisos por correo ----------
// Hoja "Avisos": RUC | Correo | Perfil | Token | Estado | Registrado | Confirmado
// Estado: pendiente (falta confirmar el correo), activo o baja
var COL = { ruc: 0, correo: 1, perfil: 2, token: 3, estado: 4, registrado: 5, confirmado: 6 };
var REGIMENES = ['RMT', 'RG', 'RER', 'NRUS'];

function hojaAvisos_() {
  return hoja_('Avisos', ['RUC', 'Correo', 'Perfil', 'Token', 'Estado', 'Registrado', 'Confirmado']);
}

function hojaEnviados_() {
  return hoja_('Enviados', ['Token', 'Obligación:periodo', 'Enviado el']);
}

// Enlaces de los correos: a la página en GitHub Pages (se puede cambiar con la propiedad URL_SITIO)
var URL_SITIO = 'https://proyectosemaforotributario-oss.github.io/semaforo-tributario/';
function urlApp_(params) {
  var sitio = PropertiesService.getScriptProperties().getProperty('URL_SITIO') || URL_SITIO;
  var base = sitio ? sitio.replace(/\/?$/, '/') + (params.accion ? 'accion.html' : '') : ScriptApp.getService().getUrl();
  return base + '?' + Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
}

function limpiarPerfil_(p) {
  p = p || {};
  return {
    regimen: REGIMENES.indexOf(p.regimen) >= 0 ? p.regimen : 'RMT',
    trabajadores: !!p.trabajadores, registros: !!p.registros,
    retencion: !!p.retencion, percepcion: !!p.percepcion, itan: !!p.itan
  };
}

// Busca la fila (1 = encabezado) cuyo valor en la columna coincide
function buscarFila_(h, col, valor) {
  var filas = h.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) if (String(filas[i][col]) === valor) return { n: i + 1, v: filas[i] };
  return null;
}

// Registra el pedido de avisos y envía el correo para confirmarlo
function registrarAviso(ruc, correo, perfil, autoriza) {
  if (!rucValido_(ruc)) throw new Error('RUC inválido');
  correo = String(correo || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) throw new Error('Correo inválido');
  if (autoriza !== true) throw new Error('Falta la autorización');
  var p = JSON.stringify(limpiarPerfil_(perfil));

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var h = hojaAvisos_(), filas = h.getDataRange().getValues();
    for (var i = 1; i < filas.length; i++) {
      if (String(filas[i][COL.ruc]) !== ruc || String(filas[i][COL.correo]) !== correo) continue;
      h.getRange(i + 1, COL.perfil + 1).setValue(p);
      if (filas[i][COL.estado] === 'activo') return { estado: 'activo' };
      // Pendiente o dado de baja: se vuelve a pedir confirmación (sin reenviar más de una vez cada 10 minutos)
      var hace = new Date() - aFecha_(filas[i][COL.registrado]);
      h.getRange(i + 1, COL.estado + 1, 1, 2).setNumberFormat('@').setValues([['pendiente', ahora_()]]);
      if (!(hace < 10 * 60 * 1000)) correoConfirmacion_(ruc, correo, String(filas[i][COL.token]));
      return { estado: 'pendiente' };
    }
    var token = Utilities.getUuid();
    agregarFila_(h, [ruc, correo, p, token, 'pendiente', ahora_(), '']);
    correoConfirmacion_(ruc, correo, token);
    return { estado: 'pendiente' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Correos ----------
function plantillaCorreo_(titulo, cuerpoHtml, pie, color) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.55;color:#16202A;max-width:560px">'
    + '<div style="height:8px;background:' + (color || '#2A3E9E') + ';border-radius:4px;margin-bottom:18px"></div>'
    + '<h1 style="font-size:21px;line-height:1.25;margin:0 0 14px">' + titulo + '</h1>'
    + cuerpoHtml
    + '<p style="margin:22px 0 0;font-size:13px">' + pie + '</p>'
    + '<p style="margin:18px 0 0;font-size:12px;color:#56626C">Prototipo académico del proyecto Semáforo tributario digital (Contabilidad, IESP José Carlos Mariátegui, Moquegua). '
    + 'Las fechas se calculan con la regla de días hábiles del cronograma de la SUNAT; confírmalas en sunat.gob.pe. No usamos tu clave SOL.</p></div>';
}

function boton_(texto, url) {
  return '<a href="' + url + '" style="display:inline-block;padding:10px 16px;border:2px solid #2A3E9E;border-radius:8px;color:#2A3E9E;font-weight:bold;text-decoration:none">' + texto + '</a>';
}

function correoConfirmacion_(ruc, correo, token) {
  var cuerpo = '<p>Hola:</p><p>Pediste recibir avisos del Semáforo Tributario para el RUC <b>' + ruc + '</b>. '
    + 'Te escribiremos cuando una obligación pase a amarillo (faltan 7 días) y cada día que esté en rojo, hasta que la marques como presentada.</p>'
    + '<p>' + boton_('Confirmar avisos', urlApp_({ accion: 'confirmar', t: token })) + '</p>'
    + '<p>Si no lo pediste, ignora este correo: no te enviaremos nada.</p>';
  MailApp.sendEmail({
    to: correo, name: 'Semáforo Tributario',
    subject: 'Confirma tus avisos del Semáforo Tributario',
    htmlBody: plantillaCorreo_('Confirma tus avisos', cuerpo, '<a href="' + urlApp_({ ruc: ruc }) + '" style="color:#2A3E9E">Ver mi semáforo</a>')
  });
}

// ---------- Aviso diario (trigger a las 8 a. m.) ----------
function hoyLima_() {
  return deIso(Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd'));
}

// Regla de los avisos (para lo que no esté marcado como presentado):
// - un aviso cuando la obligación pasa a amarillo (faltan 7 días o menos), una sola vez por periodo;
// - un aviso cada día mientras está en rojo (2 días o menos), incluidos los días después de vencer,
//   hasta DIAS_AVISO_VENCIDA días de atraso (para no escribir sin fin por algo muy antiguo).
var DIAS_AVISO_VENCIDA = 5;

// Qué aviso le toca hoy a una obligación: 'ambar', 'rojo' o '' (ninguno)
function tipoAviso_(it, token, hoyIso, enviados) {
  if (it.hecho || !it.vence) return '';
  var c = colorPorDias(it.dias);
  if (c === 'a') return enviados[token + '|' + it.clave + '|ambar'] ? '' : 'ambar';
  if (c === 'r' && it.dias >= -DIAS_AVISO_VENCIDA) return enviados[token + '|' + it.clave + '|rojo|' + hoyIso] ? '' : 'rojo';
  return '';
}

// Revisa todos los correos activos y envía los avisos que tocan hoy.
// Con enviar = false solo registra en el log lo que enviaría (para probar con otra fecha).
function revisarAvisos_(hoy, enviar) {
  var hoyIso = iso(hoy), marcas = todasLasMarcas_();
  // Hoja Enviados: Token | "obligación:periodo|ambar" o "obligación:periodo|rojo|fecha" | Enviado el
  var hEnv = hojaEnviados_(), enviados = {};
  hEnv.getDataRange().getValues().slice(1).forEach(function (f) { enviados[f[0] + '|' + f[1]] = true; });

  var resumen = [];
  hojaAvisos_().getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[COL.estado] !== 'activo') return;
    var ruc = String(f[COL.ruc]), token = String(f[COL.token]), perfil = limpiarPerfil_(JSON.parse(f[COL.perfil] || '{}'));
    var items = obligaciones(perfil)
      .map(function (ob) { return evaluar(ob, hoy, Number(ruc[10]), marcas[ruc] || {}, perfil.regimen); })
      .map(function (it) { it.tipo = tipoAviso_(it, token, hoyIso, enviados); return it; })
      .filter(function (it) { return it.tipo; })
      .sort(function (a, b) { return a.dias - b.dias; });
    if (!items.length) return;

    resumen.push(ruc + ' → ' + f[COL.correo] + ': ' + items.map(function (it) { return it.clave + ' (' + it.tipo + ')'; }).join(', '));
    if (!enviar) return;
    if (MailApp.getRemainingDailyQuota() < 1) { Logger.log('Se acabó el cupo diario de correos.'); return; }
    correoAviso_(ruc, String(f[COL.correo]), token, items);
    items.forEach(function (it) {
      agregarFila_(hEnv, [token, it.clave + '|' + (it.tipo === 'ambar' ? 'ambar' : 'rojo|' + hoyIso), ahora_()]);
    });
  });
  Logger.log((enviar ? 'Enviados' : 'Se enviarían') + ' (' + hoyIso + '):\n' + (resumen.join('\n') || 'ninguno'));
  return resumen;
}

function textoPlazo_(d) {
  if (d < 0) return 'venció hace ' + (-d) + (d === -1 ? ' día' : ' días');
  if (d === 0) return 'vence hoy';
  if (d === 1) return 'vence mañana';
  return 'vence en ' + d + ' días';
}

// items: obligaciones con su tipo ('ambar' o 'rojo'), ordenadas de la más urgente a la menos
function correoAviso_(ruc, correo, token, items, prueba) {
  var rojo = items.some(function (it) { return it.tipo === 'rojo'; }), una = items.length === 1;
  var titulo = rojo
    ? 'Semáforo en rojo: ' + (una ? items[0].ob.nombre + ' ' + textoPlazo_(items[0].dias) : items.length + ' obligaciones en rojo')
    : 'Semáforo en amarillo: ' + (una ? items[0].ob.nombre + ' vence el ' + larga(items[0].vence) : items.length + ' obligaciones vencen pronto');
  var lista = items.map(function (it) {
    return '<li style="margin:0 0 14px"><b>' + it.ob.nombre + '</b>, ' + it.periodo + '<br>'
      + '<span style="color:' + (it.tipo === 'rojo' ? '#B02A22' : '#8A5D00') + ';font-weight:bold">'
      + textoPlazo_(it.dias).replace(/^./, function (c) { return c.toUpperCase(); }) + '</span> (' + larga(it.vence) + ')<br>'
      + boton_('Ya la presenté', urlApp_({ accion: 'presentada', t: token, c: it.clave })) + '</li>';
  }).join('');
  var cuerpo = (prueba ? '<p style="padding:8px 12px;background:#FFF4D6;border-radius:6px">Correo de prueba: así se verán los avisos reales.</p>' : '')
    + '<p>Hola:</p><p>' + (rojo ? 'Tu semáforo está en <b>rojo</b>' : 'Tu semáforo pasó a <b>amarillo</b>: quedan 7 días o menos') + ' para el RUC <b>' + ruc + '</b>:</p>'
    + '<ul style="padding-left:18px">' + lista + '</ul>'
    + '<p>Si la presentas fuera de plazo, la multa puede llegar a 1 UIT (S/ 5 500 en 2026) en el Régimen General y en el Régimen MYPE Tributario, además de intereses.</p>'
    + '<p>Si ya la presentaste, márcala con «Ya la presenté» y no te volveremos a avisar por ese periodo.'
    + (rojo ? ' Mientras siga en rojo, te escribiremos cada día.' : ' Si no la marcas, te volveremos a escribir cuando pase a rojo.') + '</p>';
  MailApp.sendEmail({
    to: correo, name: 'Semáforo Tributario',
    subject: (prueba ? '[Prueba] ' : '') + titulo,
    htmlBody: plantillaCorreo_(titulo, cuerpo,
      '<a href="' + urlApp_({ ruc: ruc }) + '" style="color:#2A3E9E">Ver mi semáforo</a> &nbsp;|&nbsp; '
      + '<a href="' + urlApp_({ accion: 'baja', t: token }) + '" style="color:#2A3E9E">Dejar de recibir avisos</a>', rojo ? '#D1342B' : '#EAA000')
  });
}

// Lo ejecuta el trigger diario
function enviarAvisosDiarios() {
  revisarAvisos_(hoyLima_(), true);
}

// Ejecutar UNA vez desde el editor para programar el aviso diario de las 8 a. m. (hora de Lima)
function instalarAvisoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarAvisosDiarios') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarAvisosDiarios').timeBased().everyDays(1).atHour(8).inTimezone('America/Lima').create();
  Logger.log('Aviso diario programado a las 8 a. m.');
}

// Envía YA un aviso de ejemplo al último correo con avisos activos, con sus obligaciones pendientes
// que vencen primero. No se registra como enviado, así que no impide el aviso real.
function probarCorreoAviso() {
  var activos = hojaAvisos_().getDataRange().getValues().slice(1).filter(function (f) { return f[COL.estado] === 'activo'; });
  if (!activos.length) { Logger.log('No hay ningún correo con avisos activos.'); return; }
  var f = activos[activos.length - 1], ruc = String(f[COL.ruc]), correo = String(f[COL.correo]);
  var perfil = limpiarPerfil_(JSON.parse(f[COL.perfil] || '{}')), marcas = obtenerMarcas(ruc), hoy = hoyLima_();
  var items = obligaciones(perfil)
    .map(function (ob) { return evaluar(ob, hoy, Number(ruc[10]), marcas, perfil.regimen); })
    .filter(function (it) { return it.vence; });
  if (!items.length) { Logger.log('Ese RUC no tiene obligaciones con fecha para el aviso de prueba.'); return; }
  var pendientes = items.filter(function (it) { return !it.hecho; });
  if (pendientes.length) items = pendientes;
  var primera = Math.min.apply(null, items.map(function (it) { return it.vence.getTime(); }));
  items = items.filter(function (it) { return it.vence.getTime() === primera; });
  items.forEach(function (it) { it.tipo = colorPorDias(it.dias) === 'r' ? 'rojo' : 'ambar'; });
  correoAviso_(ruc, correo, String(f[COL.token]), items, true);
  Logger.log('Aviso de prueba enviado a ' + correo + ': ' + items.map(function (it) { return it.clave; }).join(', '));
}

// Prueba sin enviar correos: cambia la fecha y ejecútala desde el editor; el resultado sale en el registro
function simularAvisos() {
  ['2026-11-08', '2026-11-09', '2026-11-13', '2026-11-14', '2026-11-16', '2026-11-17'].forEach(function (d) {
    revisarAvisos_(deIso(d), false);
  });
}

// ---------- Páginas de los enlaces del correo ----------
// Los enlaces solo muestran una pregunta; la acción ocurre al pulsar el botón
// (así los antivirus que abren enlaces no marcan ni dan de baja por error).
function describirClave_(clave, regimen) {
  var partes = clave.split(':'), id = partes[0], per = partes[1];
  var todas = obligaciones({ regimen: regimen, trabajadores: true, registros: true, retencion: true, percepcion: true, itan: true });
  var ob = todas.filter(function (o) { return o.id === id; })[0];
  var cuando = per.length === 7 ? 'periodo ' + periodoTexto(Number(per.slice(0, 4)), Number(per.slice(5)) - 1) : 'ejercicio ' + per;
  return (ob ? ob.nombre : 'Obligación') + ', ' + cuando;
}

var PREGUNTAS = {
  confirmar: { titulo: 'Confirma tus avisos', boton: 'Confirmar avisos' },
  baja: { titulo: '¿Dejar de recibir avisos?', boton: 'Dejar de recibir avisos' },
  presentada: { titulo: '¿Ya la presentaste?', boton: 'Ya la presenté' },
  deshacer: { titulo: '¿Deshacer la marca?', boton: 'Deshacer' }
};

// La llama accion.html al abrirse: devuelve la pregunta que se muestra antes de pulsar el botón
function describirAccion(accion, token, clave) {
  var p = { accion: String(accion || ''), t: String(token || ''), c: String(clave || '') };
  var datos = { accion: p.accion, t: p.t, c: p.c };
  var fila = PREGUNTAS.hasOwnProperty(p.accion) && p.t ? buscarFila_(hojaAvisos_(), COL.token, p.t) : null;
  if (!fila || ((p.accion === 'presentada' || p.accion === 'deshacer') && !CLAVE_VALIDA.test(datos.c))) {
    datos.titulo = 'Este enlace no funciona';
    datos.texto = 'Puede que esté incompleto o que sea antiguo. Abre el semáforo y vuelve a activar los avisos.';
    datos.enlace = { texto: 'Abrir el semáforo', url: urlApp_({}) };
  } else {
    var ruc = String(fila.v[COL.ruc]), q = PREGUNTAS[p.accion];
    datos.titulo = q.titulo;
    datos.boton = q.boton;
    var regimen = limpiarPerfil_(JSON.parse(fila.v[COL.perfil] || '{}')).regimen;
    // Solo se arma el texto de la acción pedida (describirClave_ necesita una obligación en c)
    datos.texto = {
      confirmar: function () { return 'Te escribiremos a ' + fila.v[COL.correo] + ' cuando una obligación del RUC ' + ruc + ' pase a amarillo y cada día que esté en rojo, hasta que la marques como presentada.'; },
      baja: function () { return 'Ya no te escribiremos a ' + fila.v[COL.correo] + ' sobre el RUC ' + ruc + '. Puedes volver a activarlos cuando quieras.'; },
      presentada: function () { return describirClave_(datos.c, regimen) + ', RUC ' + ruc + '.'; },
      deshacer: function () { return describirClave_(datos.c, regimen) + ' volverá a figurar como pendiente.'; }
    }[p.accion]();
  }
  return datos;
}

// La llama el botón de accion.html
function ejecutarAccion(accion, token, clave) {
  var h = hojaAvisos_(), fila = buscarFila_(h, COL.token, String(token));
  if (!fila || !PREGUNTAS.hasOwnProperty(accion)) throw new Error('Enlace no válido');
  var ruc = String(fila.v[COL.ruc]), verSemaforo = { texto: 'Ver mi semáforo', url: urlApp_({ ruc: ruc }) };

  if (accion === 'confirmar' || accion === 'baja') {
    h.getRange(fila.n, COL.estado + 1).setValue(accion === 'confirmar' ? 'activo' : 'baja');
    if (accion === 'confirmar') h.getRange(fila.n, COL.confirmado + 1).setNumberFormat('@').setValue(ahora_());
    return accion === 'confirmar'
      ? { titulo: 'Avisos activados', texto: 'Te escribiremos cuando una obligación pase a amarillo y cada día que esté en rojo, hasta que la marques como presentada.', enlace: verSemaforo }
      : { titulo: 'Ya no recibirás avisos', texto: 'Si cambias de opinión, vuelve a activarlos desde tu semáforo.', enlace: verSemaforo };
  }
  marcarPresentada(ruc, clave, accion === 'presentada');
  return accion === 'presentada'
    ? { titulo: 'Marcada como presentada', sello: true, texto: 'No te volveremos a avisar por este periodo.', enlace: verSemaforo,
        deshacer: urlApp_({ accion: 'deshacer', t: token, c: clave }) }
    : { titulo: 'Marca deshecha', texto: 'La obligación vuelve a figurar como pendiente en tu semáforo.', enlace: verSemaforo };
}
