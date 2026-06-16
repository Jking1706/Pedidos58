// Apps Script para registrar pedidos de Sabor +58 en Google Sheets.
// 1. Pega este archivo en un proyecto de Apps Script vinculado a una hoja.
// 2. Despliega como Web app.
// 3. Si usas un backend proxy, configura la URL del Web App en `GOOGLE_SHEETS_WEBAPP_URL` del servidor.

const SHEET_NAME = 'Pedidos_58';
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
  console.log('--- NUEVA PETICIÓN RECIBIDA ---');
  console.log('Contexto:', JSON.stringify((e && e.parameter) || {}));
  if (e && e.postData && e.postData.contents) {
    console.log('Contenido POST:', e.postData.contents);
  } else {
    console.log('No hay contenido POST (e.postData está vacío)');
  }

  try {
    const rawData = parsePayload_(e);
    const data = normalizePayload_(rawData);
    const sheet = getOrdersSheet_();
    ensureHeaderRow_(sheet);

    console.log('doPost raw keys: ' + Object.keys(rawData).join(', '));
    console.log('doPost normalized keys: ' + Object.keys(data).join(', '));

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

function normalizePayload_(data) {
  const itemsSource = pick_(
    data.items,
    data.itemsJson,
    data.itemsJSON,
    data.productosItems,
    data.productos,
    data.producto
  );

  const parsedItems = parseItems_(itemsSource);
  const itemsResumen = firstText_(
    data.itemsResumen,
    data.productosResumen,
    data.productos,
    data.producto,
    summarizeItems_(parsedItems)
  );

  return {
    pedidoKey: firstText_(data.pedidoKey, data.pedido_id, data.pedidoId, data.orderId, data.numero, data.idLocal, data.id),
    numero: firstText_(data.numero, data.orderNumber, data.nro),
    idLocal: firstText_(data.idLocal, data.id, data.pedidoId, data.orderId),
    accion: firstText_(data.accion, data.action, data.operacion),
    cliente: firstText_(data.cliente, data.nombre, data.customerName, data.name),
    telefono: firstText_(data.telefono, data.phone, data.tel, data.customerPhone),
    direccion: firstText_(data.direccion, data.address, data.customerAddress),
    metodoPago: firstText_(data.metodoPago, data.metodo_pago, data.paymentMethod, data.payment, data.metodo, data.pago),
    notas: firstText_(data.notas, data.notes, data.note, data.observaciones, data.comentarios),
    itemsResumen: firstText_(itemsResumen),
    items: normalizeItemsJson_(itemsSource),
    total: Number(pick_(data.total, data.amount, data.monto, 0) || 0),
    hora: firstText_(data.hora, data.time),
    fecha: firstText_(data.fecha, data.date, formatToday_()),
    entregado: toBoolean_(data.entregado) || toBoolean_(data.delivered) || toBoolean_(data.isDelivered) || String(data.estado || '').trim() === 'Entregado',
    estado: firstText_(data.estado, data.status)
  };
}

function firstText_() {
  for (let index = 0; index < arguments.length; index += 1) {
    const text = stringifyValue_(arguments[index]);
    if (text) {
      return text;
    }
  }

  return '';
}

function stringifyValue_(value) {
  if (value === undefined || value === null) {
    return '';
  }

  if (Array.isArray(value)) {
    return summarizeItems_(value) || JSON.stringify(value);
  }

  if (typeof value === 'object') {
    return JSON.stringify(value);
  }

  const text = String(value).trim();
  return text;
}

function parseItems_(value) {
  if (Array.isArray(value)) {
    return value;
  }

  if (value && typeof value === 'object') {
    return [value];
  }

  const text = String(value || '').trim();
  if (!text) {
    return [];
  }

  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      return parsed;
    }
  } catch (error) {
    // Si no es JSON válido, devolvemos vacío y usamos el texto original como resumen.
  }

  return [];
}

function summarizeItems_(items) {
  if (!Array.isArray(items) || items.length === 0) {
    return '';
  }

  return items.map((item) => {
    if (!item || typeof item !== 'object') {
      return String(item);
    }

    const qty = Number(item.qty || item.cantidad || 1);
    const name = String(item.nombre || item.name || item.producto || item.titulo || 'Item').trim();
    return `${qty}x ${name}`;
  }).join(' | ');
}

function normalizeItemsJson_(value) {
  if (Array.isArray(value)) {
    return JSON.stringify(value);
  }

  if (value && typeof value === 'object') {
    return JSON.stringify(value);
  }

  const text = String(value || '').trim();
  if (!text) {
    return '';
  }

  try {
    const parsed = JSON.parse(text);
    return JSON.stringify(parsed);
  } catch (error) {
    return text;
  }
}

function pick_() {
  for (let index = 0; index < arguments.length; index += 1) {
    const value = arguments[index];
    if (value === undefined || value === null) {
      continue;
    }

    if (typeof value === 'string' && value.trim() === '') {
      continue;
    }

    return value;
  }

  return '';
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
