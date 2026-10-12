const CONFIG = Object.freeze({
  SHEET_NAME: 'Leads',
  TIME_ZONE: 'America/Lima',
  INTERNAL_EMAIL: 'baileyai@realitynear.org',
  ORIGIN: 'JBD Producciones Web',
  MAX_LENGTHS: {
    nombre_completo: 120, fecha_evento: 20, tipo_evento: 80, invitados: 20,
    whatsapp_codigo: 8, whatsapp_numero: 30, correo: 254,
    presupuesto_rango: 80, comentarios: 3000, idioma: 2,
    consentimiento_version: 80, idempotency_key: 80, origen: 80, website: 200
  }
});

const HEADERS = Object.freeze([
  'fecha_recepcion', 'nombre_completo', 'tiene_fecha', 'fecha_evento',
  'tipo_evento', 'invitados', 'whatsapp_codigo', 'whatsapp_numero',
  'correo', 'presupuesto_rango', 'comentarios', 'idioma', 'consentimiento',
  'consentimiento_version', 'idempotency_key', 'origen',
  'aviso_interno_enviado', 'acuse_cliente_enviado'
]);

function doGet() {
  return jsonResponse_({ ok: true, service: 'JBD lead intake' });
}

function doPost(e) {
  try {
    const payload = readPayload_(e);
    if (payload.website) return jsonResponse_({ ok: true });

    const lead = validateAndNormalize_(payload);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getSheet_();
      let row = findByIdempotencyKey_(sheet, lead.idempotency_key);
      if (!row) row = appendLead_(sheet, lead);
      sendPendingEmails_(sheet, row, lead);
      return jsonResponse_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonResponse_({ ok: false, error: 'No fue posible procesar la solicitud.' });
  }
}

function readPayload_(e) {
  if (!e) throw new Error('Solicitud vacía');
  const type = String(e.postData && e.postData.type || '').toLowerCase();
  if (type.indexOf('application/json') !== -1) return JSON.parse(e.postData.contents || '{}');
  return Object.assign({}, e.parameter || {});
}

function validateAndNormalize_(input) {
  const lead = {};
  Object.keys(CONFIG.MAX_LENGTHS).forEach(function (key) {
    lead[key] = cleanText_(input[key], CONFIG.MAX_LENGTHS[key]);
  });
  lead.tiene_fecha = cleanText_(input.tiene_fecha, 2).toLowerCase();
  lead.consentimiento = String(input.consentimiento).toLowerCase() === 'true';
  lead.idioma = ['es', 'en', 'pt'].indexOf(lead.idioma.toLowerCase()) >= 0 ? lead.idioma.toLowerCase() : 'es';

  [
    'nombre_completo', 'tiene_fecha', 'tipo_evento', 'invitados',
    'whatsapp_codigo', 'whatsapp_numero', 'correo', 'presupuesto_rango',
    'consentimiento_version', 'idempotency_key', 'origen'
  ].forEach(function (key) {
    if (!lead[key]) throw new Error('Campo requerido inválido: ' + key);
  });

  if (['si', 'no'].indexOf(lead.tiene_fecha) < 0) throw new Error('tiene_fecha inválido');
  if (lead.tiene_fecha === 'si' && !lead.fecha_evento) throw new Error('fecha_evento requerido');
  if (!/^\+\d{1,4}$/.test(lead.whatsapp_codigo)) throw new Error('Código telefónico inválido');
  if (!/^[0-9 ()-]{5,30}$/.test(lead.whatsapp_numero)) throw new Error('Teléfono inválido');
  if (!/^\d+$/.test(lead.invitados) || Number(lead.invitados) < 1) throw new Error('Invitados inválidos');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(lead.correo)) throw new Error('Correo inválido');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(lead.idempotency_key)) throw new Error('Idempotency key inválida');
  if (!lead.consentimiento) throw new Error('Consentimiento requerido');
  if (lead.origen !== CONFIG.ORIGIN) throw new Error('Origen inválido');
  return lead;
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().replace(/[\u0000-\u001F\u007F]/g, ' ').slice(0, maxLength);
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.setSpreadsheetTimeZone(CONFIG.TIME_ZONE);
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findByIdempotencyKey_(sheet, key) {
  if (sheet.getLastRow() < 2) return 0;
  const column = HEADERS.indexOf('idempotency_key') + 1;
  const match = sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(key).matchEntireCell(true).findNext();
  return match ? match.getRow() : 0;
}

function appendLead_(sheet, lead) {
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
  const values = HEADERS.map(function (header) {
    if (header === 'fecha_recepcion') return timestamp;
    if (header === 'aviso_interno_enviado' || header === 'acuse_cliente_enviado') return false;
    return safeSheetValue_(lead[header]);
  });
  sheet.appendRow(values);
  return sheet.getLastRow();
}

function safeSheetValue_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function sendPendingEmails_(sheet, row, lead) {
  const internalColumn = HEADERS.indexOf('aviso_interno_enviado') + 1;
  const clientColumn = HEADERS.indexOf('acuse_cliente_enviado') + 1;
  if (sheet.getRange(row, internalColumn).getValue() !== true) {
    sendInternalEmail_(lead);
    sheet.getRange(row, internalColumn).setValue(true);
  }
  if (sheet.getRange(row, clientColumn).getValue() !== true) {
    sendClientEmail_(lead);
    sheet.getRange(row, clientColumn).setValue(true);
  }
}

function sendInternalEmail_(lead) {
  const labels = {
    nombre_completo: 'Nombre', tiene_fecha: 'Tiene fecha', fecha_evento: 'Fecha del evento',
    tipo_evento: 'Tipo de evento', invitados: 'Invitados', whatsapp_codigo: 'Código WhatsApp',
    whatsapp_numero: 'WhatsApp', correo: 'Correo', presupuesto_rango: 'Presupuesto',
    comentarios: 'Comentarios', idioma: 'Idioma', consentimiento: 'Consentimiento',
    consentimiento_version: 'Versión del consentimiento', idempotency_key: 'ID de solicitud', origen: 'Origen'
  };
  const keys = Object.keys(labels);
  const plain = keys.map(function (key) { return labels[key] + ': ' + String(lead[key] || ''); }).join('\n');
  const html = '<h2>Nuevo lead web de JBD Producciones</h2><table>' + keys.map(function (key) {
    return '<tr><th align="left" valign="top">' + escapeHtml_(labels[key]) + '</th><td>' + escapeHtml_(lead[key] || '') + '</td></tr>';
  }).join('') + '</table>';
  MailApp.sendEmail({
    to: CONFIG.INTERNAL_EMAIL, replyTo: lead.correo,
    subject: 'Nuevo lead web JBD: ' + lead.nombre_completo + ' – ' + lead.tipo_evento,
    body: plain, htmlBody: html, name: 'JBD Producciones Web'
  });
}

function sendClientEmail_(lead) {
  const messages = {
    es: 'Gracias por cotizar con JBD Producciones. Recibimos tu solicitud y pronto te contactamos.',
    en: 'Thank you for requesting a quote from JBD Producciones. We received your request and will contact you soon.',
    pt: 'Obrigado por solicitar uma cotação à JBD Producciones. Recebemos sua solicitação e entraremos em contato em breve.'
  };
  const subjects = {
    es: 'Recibimos tu solicitud — JBD Producciones',
    en: 'We received your request — JBD Producciones',
    pt: 'Recebemos sua solicitação — JBD Producciones'
  };
  const message = messages[lead.idioma] || messages.es;
  MailApp.sendEmail({
    to: lead.correo, subject: subjects[lead.idioma] || subjects.es,
    body: message, htmlBody: '<p>' + escapeHtml_(message) + '</p>', name: 'JBD Producciones'
  });
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\n/g, '<br>');
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
