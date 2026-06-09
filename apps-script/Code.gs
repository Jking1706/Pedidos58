// Apps Script para registrar pedidos de Sabor +58 en Google Sheets.
// 1. Pega este archivo en un proyecto de Apps Script vinculado a una hoja.
// 2. Despliega como Web app.
// 3. Copia la URL en `window.GOOGLE_SHEETS_WEBAPP_URL` dentro de `index.html`.

const SHEET_NAME = 'Pedidos';
const TIME_ZONE = 'America/Bogota';
const HEADERS = [
  'Fecha',
  'Pedido #',
  'Accion',
  'ID local',
  'Cliente',
  'Telefono',
  'Direccion',
  'Metodo de pago',
  'Notas',
  'Productos',
  'Items JSON',
  'Total',
  'Hora',
  'Estado',
  'Entregado'
];

function doPost(e) {
  try {
    const data = parsePayload_(e);
    const sheet = getOrdersSheet_();
    ensureHeaderRow_(sheet);

    const pedidoKey = String(data.pedidoKey || data.numero || data.idLocal || '').trim();
    if (!pedidoKey) {
      throw new Error('Falta pedidoKey/numero.');
    }

    const row = findOrderRow_(sheet, pedidoKey);
    const values = buildRowValues_(data, pedidoKey, row);

    if (row) {
      sheet.getRange(row, 1, 1, HEADERS.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }

    return jsonResponse_({ ok: true, action: row ? 'updated' : 'created', pedidoKey });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message });
  }
}

function doGet() {
  return jsonResponse_({ ok: true, message: 'Sabor +58 Sheets webhook activo' });
}

function getOrdersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Este script debe estar vinculado a una hoja de calculo.');
  }

  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function ensureHeaderRow_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
}

function findOrderRow_(sheet, pedidoKey) {
  if (sheet.getLastRow() < 2) {
    return 0;
  }

  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][0] || '').trim() === pedidoKey) {
      return index + 2;
    }
  }

  return 0;
}

function buildRowValues_(data, pedidoKey, row) {
  const fecha = String(data.fecha || formatToday_());
  const hora = String(data.hora || getTimeNow_());
  const accion = String(data.accion || (row ? 'update' : 'upsert')).trim();
  const entregado = toBoolean_(data.entregado) ? 'SI' : 'NO';
  const estado = resolveEstado_(data.estado, accion, row, entregado);

  return [
    fecha,
    pedidoKey,
    accion,
    String(data.idLocal || ''),
    String(data.cliente || ''),
    String(data.telefono || ''),
    String(data.direccion || ''),
    String(data.metodoPago || ''),
    String(data.notas || ''),
    String(data.itemsResumen || ''),
    String(data.items || ''),
    Number(data.total || 0),
    hora,
    estado,
    entregado
  ];
}

function resolveEstado_(incomingEstado, accion, row, entregado) {
  const estado = String(incomingEstado || '').trim();
  if (accion === 'delete') return 'Cancelado';
  if (accion === 'delivered') return 'Entregado';
  if (entregado === 'SI') return 'Entregado';
  if (estado) return estado;
  if (row) return 'Editado';
  return 'Nuevo';
}

function parsePayload_(e) {
  if (e && e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }

  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!raw) {
    return {};
  }

  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      return parsed;
    }
  } catch (error) {
    // Fall back to querystring parsing.
  }

  const payload = {};
  const params = new URLSearchParams(raw);
  params.forEach((value, key) => {
    payload[key] = value;
  });

  return payload;
}

function toBoolean_(value) {
  return value === true || value === 'true' || value === 'SI' || value === 'si' || value === '1';
}

function formatToday_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'dd/MM/yyyy');
}

function getTimeNow_() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'HH:mm');
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
