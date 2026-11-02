// Semáforo tributario – servidor en Google Apps Script (ligado a la hoja "Semáforo Tributario Digital")
// Tokens de las API: Configuración del proyecto > Propiedades del script (nunca en el código).
// Las fechas y obligaciones vienen de Calendario.gs (generado desde index.html con "node sincronizar.js").

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.accion) return paginaAccion_(p);
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Semáforo Tributario')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- RUC ----------
function rucValido_(ruc) {
  return typeof ruc === 'string' && errorRuc(ruc) === '';
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

// ---------- Hojas ----------
function hoja_(nombre, encabezados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var filas = hojaPresentadas_().getDataRange().getValues();
  var tz = Session.getScriptTimeZone(), r = {};
  for (var i = 1; i < filas.length; i++) {
    var ruc = String(filas[i][0]), f = filas[i][2];
    (r[ruc] = r[ruc] || {})[String(filas[i][1])] = f instanceof Date ? Utilities.formatDate(f, tz, 'yyyy-MM-dd') : String(f);
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
    if (presentada && fila === -1) h.appendRow([ruc, clave, new Date()]);
    if (!presentada && fila !== -1) h.deleteRow(fila);
  } finally {
    lock.releaseLock();
  }
  return obtenerMarcas(ruc);
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

function urlApp_(params) {
  var base = PropertiesService.getScriptProperties().getProperty('URL_APP') || ScriptApp.getService().getUrl();
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
      var hace = new Date() - new Date(filas[i][COL.registrado]);
      h.getRange(i + 1, COL.estado + 1, 1, 2).setValues([['pendiente', new Date()]]);
      if (!(hace < 10 * 60 * 1000)) correoConfirmacion_(ruc, correo, String(filas[i][COL.token]));
      return { estado: 'pendiente' };
    }
    var token = Utilities.getUuid();
    h.appendRow([ruc, correo, p, token, 'pendiente', new Date(), '']);
    correoConfirmacion_(ruc, correo, token);
    return { estado: 'pendiente' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Correos ----------
function plantillaCorreo_(titulo, cuerpoHtml, pie) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.55;color:#16202A;max-width:560px">'
    + '<div style="height:8px;background:#D1342B;border-radius:4px;margin-bottom:18px"></div>'
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
    + 'Te escribiremos un día hábil antes de cada vencimiento que no hayas marcado como presentado.</p>'
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

function siguienteHabil_(d) {
  var s = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
  while (!esHabil(s)) s.setDate(s.getDate() + 1);
  return s;
}

// Revisa a quién le vence algo el día hábil siguiente y le envía el último aviso.
// fechaIso permite simular otro día; con enviar = false solo registra en el log lo que enviaría.
function revisarAvisos_(hoy, enviar) {
  if (!esHabil(hoy)) { Logger.log(iso(hoy) + ' no es día hábil: no se envían avisos.'); return []; }
  var manana = iso(siguienteHabil_(hoy));
  var marcas = todasLasMarcas_();
  var hEnv = hojaEnviados_(), enviados = {};
  hEnv.getDataRange().getValues().slice(1).forEach(function (f) { enviados[f[0] + '|' + f[1]] = true; });

  var resumen = [];
  hojaAvisos_().getDataRange().getValues().slice(1).forEach(function (f) {
    if (f[COL.estado] !== 'activo') return;
    var ruc = String(f[COL.ruc]), token = String(f[COL.token]), perfil = limpiarPerfil_(JSON.parse(f[COL.perfil] || '{}'));
    var items = obligaciones(perfil)
      .map(function (ob) { return evaluar(ob, hoy, Number(ruc[10]), marcas[ruc] || {}, perfil.regimen); })
      .filter(function (it) { return !it.hecho && it.vence && iso(it.vence) === manana && !enviados[token + '|' + it.clave]; });
    if (!items.length) return;

    resumen.push(ruc + ' → ' + f[COL.correo] + ': ' + items.map(function (it) { return it.clave; }).join(', '));
    if (!enviar) return;
    if (MailApp.getRemainingDailyQuota() < 1) { Logger.log('Se acabó el cupo diario de correos.'); return; }
    correoAviso_(ruc, String(f[COL.correo]), token, items);
    items.forEach(function (it) { hEnv.appendRow([token, it.clave, new Date()]); });
  });
  Logger.log((enviar ? 'Enviados' : 'Se enviarían') + ' (' + iso(hoy) + ', vence ' + manana + '):\n' + (resumen.join('\n') || 'ninguno'));
  return resumen;
}

function correoAviso_(ruc, correo, token, items) {
  var vence = larga(items[0].vence), una = items.length === 1;
  var lista = items.map(function (it) {
    return '<li style="margin:0 0 12px"><b>' + it.ob.nombre + '</b>, ' + it.periodo + '<br>'
      + boton_('Ya la presenté', urlApp_({ accion: 'presentada', t: token, c: it.clave })) + '</li>';
  }).join('');
  var cuerpo = '<p>Hola:</p><p>Este es tu último aviso. Para el RUC <b>' + ruc + '</b>, '
    + (una ? 'esta obligación vence' : 'estas obligaciones vencen') + ' el <b>' + vence + '</b>, el próximo día hábil:</p>'
    + '<ul style="padding-left:18px">' + lista + '</ul>'
    + '<p>Si la presentas fuera de plazo, la multa puede llegar a 1 UIT (S/ 5 500 en 2026) en el Régimen General y en el Régimen MYPE Tributario, además de intereses.</p>'
    + '<p>Si ya la presentaste, márcala con «Ya la presenté» y no te volveremos a avisar por ese periodo.</p>';
  MailApp.sendEmail({
    to: correo, name: 'Semáforo Tributario',
    subject: 'Semáforo en rojo: ' + (una ? items[0].ob.nombre + ' vence' : items.length + ' obligaciones vencen') + ' el ' + vence,
    htmlBody: plantillaCorreo_('Semáforo en rojo: ' + (una ? 'tu declaración vence' : 'tus declaraciones vencen') + ' el ' + vence, cuerpo,
      '<a href="' + urlApp_({ ruc: ruc }) + '" style="color:#2A3E9E">Ver mi semáforo</a> &nbsp;|&nbsp; '
      + '<a href="' + urlApp_({ accion: 'baja', t: token }) + '" style="color:#2A3E9E">Dejar de recibir avisos</a>')
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

// Prueba sin enviar correos: cambia la fecha y ejecútala desde el editor; el resultado sale en el registro
function simularAvisos() {
  revisarAvisos_(deIso('2026-11-13'), false);
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

function paginaAccion_(p) {
  var datos = { accion: p.accion, t: p.t || '', c: p.c || '' };
  var fila = PREGUNTAS[p.accion] && p.t ? buscarFila_(hojaAvisos_(), COL.token, String(p.t)) : null;
  if (!fila || ((p.accion === 'presentada' || p.accion === 'deshacer') && !CLAVE_VALIDA.test(datos.c))) {
    datos.titulo = 'Este enlace no funciona';
    datos.texto = 'Puede que esté incompleto o que sea antiguo. Abre el semáforo y vuelve a activar los avisos.';
    datos.enlace = { texto: 'Abrir el semáforo', url: urlApp_({}) };
  } else {
    var ruc = String(fila.v[COL.ruc]), q = PREGUNTAS[p.accion];
    datos.titulo = q.titulo;
    datos.boton = q.boton;
    datos.texto = {
      confirmar: 'Recibirás en ' + fila.v[COL.correo] + ' un aviso un día hábil antes de cada vencimiento del RUC ' + ruc + ' que no hayas marcado como presentado.',
      baja: 'Ya no te escribiremos a ' + fila.v[COL.correo] + ' sobre el RUC ' + ruc + '. Puedes volver a activarlos cuando quieras.',
      presentada: describirClave_(datos.c, limpiarPerfil_(JSON.parse(fila.v[COL.perfil] || '{}')).regimen) + ', RUC ' + ruc + '.',
      deshacer: describirClave_(datos.c, limpiarPerfil_(JSON.parse(fila.v[COL.perfil] || '{}')).regimen) + ' volverá a figurar como pendiente.'
    }[p.accion];
  }
  var t = HtmlService.createTemplateFromFile('pagina');
  t.datosJson = JSON.stringify(datos).replace(/</g, '\\u003c');
  return t.evaluate().setTitle('Semáforo Tributario').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// La llama el botón de pagina.html
function ejecutarAccion(accion, token, clave) {
  var h = hojaAvisos_(), fila = buscarFila_(h, COL.token, String(token));
  if (!fila || !PREGUNTAS[accion]) throw new Error('Enlace no válido');
  var ruc = String(fila.v[COL.ruc]), verSemaforo = { texto: 'Ver mi semáforo', url: urlApp_({ ruc: ruc }) };

  if (accion === 'confirmar' || accion === 'baja') {
    h.getRange(fila.n, COL.estado + 1).setValue(accion === 'confirmar' ? 'activo' : 'baja');
    if (accion === 'confirmar') h.getRange(fila.n, COL.confirmado + 1).setValue(new Date());
    return accion === 'confirmar'
      ? { titulo: 'Avisos activados', texto: 'Te escribiremos un día hábil antes de cada vencimiento que no hayas marcado como presentado.', enlace: verSemaforo }
      : { titulo: 'Ya no recibirás avisos', texto: 'Si cambias de opinión, vuelve a activarlos desde tu semáforo.', enlace: verSemaforo };
  }
  marcarPresentada(ruc, clave, accion === 'presentada');
  return accion === 'presentada'
    ? { titulo: 'Marcada como presentada', sello: true, texto: 'No te volveremos a avisar por este periodo.', enlace: verSemaforo,
        deshacer: urlApp_({ accion: 'deshacer', t: token, c: clave }) }
    : { titulo: 'Marca deshecha', texto: 'La obligación vuelve a figurar como pendiente en tu semáforo.', enlace: verSemaforo };
}
