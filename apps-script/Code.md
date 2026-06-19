// Apps Script para registrar pedidos de Sabor +58 en Google Sheets.
// 1. Pega este archivo en un proyecto de Apps Script vinculado a una hoja.
// 2. Despliega como Web app.
// 3. Si usas un backend proxy, configura la URL del Web App en `GOOGLE_SHEETS_WEBAPP_URL` del servidor.

const SHEET_NAME = 'Pedidos_58';
const GASTOS_SHEET_NAME = 'Gastos_58';
const CATALOGO_SHEET_NAME = 'Catalogo_Insumos';
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

const GASTOS_HEADERS = [
  'Fecha',
  'Categoría',
  'Descripción',
  'Monto',
  'Método de Pago',
  'Hora',
  'Producto',
  'Unidad',
  'Precio_Referencia',
  'Cantidad',
  'Precio_Final',
  'Tipo'
];

const CATALOGO_HEADERS = ['Categoría', 'Producto', 'Unidad', 'Precio_Referencia'];

const CATALOGO_SEED = [
  ['Comida', 'Paq Perro Pequeño', 'Paquete x10', 8000],
  ['Comida', 'Paq Perro Grande', 'Paquete x10', 12000],
  ['Comida', 'Paq Hamburguesa Grande', 'Paquete x12', 15000],
  ['Comida', 'Tocineta', 'Kilo', 25000],
  ['Comida', 'Queso Mozzarella', 'Kilo', 22000],
  ['Comida', 'Envases de Salsa', 'Ciento', 5000],
  ['Arepas', 'Harina', 'Kilo', 4500],
  ['Arepas', 'Carne de Res', 'Kilo', 28000],
  ['Arepas', 'Queso Costeño', 'Kilo', 20000],
  ['Servicios', 'Gas', 'Mensual', 60000],
  ['Servicios', 'Papel para envolver', 'Paquete', 12000],
];

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || '').trim().toLowerCase();

  if (action === 'catalogo' || action === 'catalog' || action === 'catalogo_insumos') {
    const sheet = getOrCreateSheet_(CATALOGO_SHEET_NAME);
    ensureHeaderRow_(sheet, CATALOGO_HEADERS);
    seedCatalogoSheetIfNeeded_(sheet);
    const items = getCatalogoInsumos_(sheet);
    return jsonResponse_({
      ok: true,
      action: 'catalogo',
      sheet: CATALOGO_SHEET_NAME,
      total: items.length,
      categories: uniqueValues_(items.map((item) => item.categoria)),
      items,
    });
  }

  return jsonResponse_({ ok: true, message: 'Sabor +58 Sheets webhook activo' });
}

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
    const tipoEntrada = String(data.tipoEntrada || 'pedido').trim().toLowerCase();

    console.log('doPost raw keys: ' + Object.keys(rawData).join(', '));
    console.log('doPost normalized keys: ' + Object.keys(data).join(', '));
    console.log('tipoEntrada:', tipoEntrada);

    if (tipoEntrada === 'gasto') {
      const gastosSheet = getOrCreateSheet_(GASTOS_SHEET_NAME);
      ensureHeaderRow_(gastosSheet, GASTOS_HEADERS);
      gastosSheet.appendRow(buildGastoRowValues_(data, 'manual'));

      return jsonResponse_({ ok: true, action: 'created', tipoEntrada: 'gasto', sheet: GASTOS_SHEET_NAME });
    }

    if (tipoEntrada === 'gasto_carrito' || tipoEntrada === 'gastos_carrito' || tipoEntrada === 'gasto_lote') {
      const gastosSheet = getOrCreateSheet_(GASTOS_SHEET_NAME);
      ensureHeaderRow_(gastosSheet, GASTOS_HEADERS);

      const items = Array.isArray(data.itemsArray) ? data.itemsArray : parseItems_(rawData.items || data.items);
      if (!items.length) {
        throw new Error('El carrito de gastos está vacío.');
      }

      const fecha = firstText_(data.fecha, formatToday_());
      const hora = firstText_(data.hora, getTimeNow_());
      const metodoPago = firstText_(data.metodoPago, data.metodo_pago, data.paymentMethod, data.payment);

      items.forEach((item) => {
        const payloadItem = {
          ...data,
          ...item,
          tipoRegistro: 'catalogo',
          fecha: firstText_(item.fecha, fecha),
          hora: firstText_(item.hora, hora),
          metodoPago: firstText_(item.metodoPago, metodoPago),
        };

        gastosSheet.appendRow(buildGastoRowValues_(payloadItem, 'catalogo'));
      });

      return jsonResponse_({
        ok: true,
        action: 'created',
        tipoEntrada: 'gasto_carrito',
        sheet: GASTOS_SHEET_NAME,
        itemsProcesados: items.length,
      });
    }

    const sheet = getOrCreateSheet_(SHEET_NAME);
    ensureHeaderRow_(sheet, HEADERS);

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

function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Este script debe estar vinculado a una hoja de calculo.');
  }

  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function getOrdersSheet_() {
  return getOrCreateSheet_(SHEET_NAME);
}

function ensureHeaderRow_(sheet, headers) {
  const currentLastColumn = Math.max(sheet.getLastColumn(), headers.length);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const current = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];
  const hasAnyValue = current.some((value) => String(value || '').trim() !== '');

  if (!hasAnyValue) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const next = current.slice();
  let changed = false;

  headers.forEach((header, index) => {
    if (!String(next[index] || '').trim()) {
      next[index] = header;
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([next.slice(0, headers.length)]);
  }
}

function normalizePayload_(data) {
  const amount = Number(pick_(data.total, data.monto, data.valor, data.precio, data.amount, 0) || 0);
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
    tipoEntrada: firstText_(data.tipoEntrada, data.tipo_entrada, data.entryType, data.type, 'pedido'),
    pedidoKey: firstText_(data.pedidoKey, data.pedido_id, data.pedidoId, data.orderId, data.numero, data.idLocal, data.id),
    numero: firstText_(data.numero, data.orderNumber, data.nro),
    idLocal: firstText_(data.idLocal, data.id, data.pedidoId, data.orderId),
    accion: firstText_(data.accion, data.action, data.operacion),
    cliente: firstText_(data.cliente, data.nombre, data.customerName, data.name),
    telefono: firstText_(data.telefono, data.phone, data.tel, data.customerPhone),
    direccion: firstText_(data.direccion, data.address, data.customerAddress),
    metodoPago: firstText_(data.metodoPago, data.metodo_pago, data.paymentMethod, data.payment, data.metodo, data.pago),
    categoria: firstText_(data.categoria, data.category, data.categoriaGasto, data.tipoGasto, data.rubro),
    descripcion: firstText_(data.descripcion, data.detalle, data.glosa, data.concepto, data.observacion, data.notas, data.notes, data.comentarios),
    notas: firstText_(data.notas, data.notes, data.note, data.observaciones, data.comentarios),
    itemsResumen: firstText_(itemsResumen),
    items: normalizeItemsJson_(itemsSource),
    itemsArray: parsedItems,
    total: amount,
    monto: amount,
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

function buildGastoRowValues_(data, tipoRegistro) {
  const tipo = String(tipoRegistro || data.tipoRegistro || data.tipoEntrada || 'manual').trim().toLowerCase();
  const esCatalogo = ['catalogo', 'gasto_carrito', 'gastos_carrito', 'gasto_lote'].includes(tipo);
  const fecha = String(data.fecha || formatToday_());
  const categoria = String(data.categoria || data.category || data.rubro || '').trim();
  const descripcion = String(data.descripcion || (esCatalogo ? '' : data.notas || '')).trim();
  const cantidad = esCatalogo ? Math.max(1, Math.round(parseNumberLoose_(pick_(data.cantidad, data.qty, 1), 1)) || 1) : '';
  const precioReferencia = esCatalogo ? parseNumberLoose_(pick_(data.precioReferencia, data.precio_referencia, data.precioRef, data.precioOriginal, 0), 0) : '';
  const precioFinal = esCatalogo ? parseNumberLoose_(pick_(data.precioFinal, data.precio_final, data.precio, 0), precioReferencia || 0) : '';
  const monto = esCatalogo
    ? parseNumberLoose_(pick_(data.monto, data.total, Number(cantidad || 0) * Number(precioFinal || precioReferencia || 0), 0), 0)
    : parseNumberLoose_(pick_(data.total, data.monto, data.valor, data.precio, 0), 0);
  const producto = esCatalogo ? String(data.producto || data.nombre || '').trim() : '';
  const unidad = esCatalogo ? String(data.unidad || '').trim() : '';
  const metodoPago = String(data.metodoPago || '').trim();
  const hora = String(data.hora || getTimeNow_());

  return [
    fecha,
    categoria,
    descripcion,
    monto,
    metodoPago,
    hora,
    producto,
    unidad,
    precioReferencia || '',
    cantidad || '',
    precioFinal || '',
    esCatalogo ? 'catalogo' : 'manual'
  ];
}

function parseNumberLoose_(value, fallback) {
  if (value === undefined || value === null || value === '') {
    return fallback || 0;
  }

  if (typeof value === 'number' && !Number.isNaN(value)) {
    return value;
  }

  const text = String(value).trim();
  if (!text) {
    return fallback || 0;
  }

  const normalized = text.replace(/\s/g, '').replace(/\./g, '').replace(/,/g, '.');
  const parsed = Number(normalized);
  return Number.isNaN(parsed) ? (fallback || 0) : parsed;
}

function normalizeText_(value) {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    const key = normalizeText_(header);
    if (key && map[key] === undefined) {
      map[key] = index;
    }
  });
  return map;
}

function findHeaderIndex_(headerMap, names, fallbackIndex) {
  for (let index = 0; index < names.length; index += 1) {
    const key = normalizeText_(names[index]);
    if (headerMap[key] !== undefined) {
      return headerMap[key];
    }
  }
  return fallbackIndex;
}

function seedCatalogoSheetIfNeeded_(sheet) {
  if (sheet.getLastRow() >= 2) {
    return;
  }

  sheet.getRange(2, 1, CATALOGO_SEED.length, CATALOGO_HEADERS.length).setValues(CATALOGO_SEED);
}

function getCatalogoInsumos_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headers = values.shift();
  const headerMap = buildHeaderMap_(headers);
  const categoriaIndex = findHeaderIndex_(headerMap, ['categoría', 'categoria'], 0);
  const productoIndex = findHeaderIndex_(headerMap, ['producto', 'insumo'], 1);
  const unidadIndex = findHeaderIndex_(headerMap, ['unidad', 'presentacion', 'presentación'], 2);
  const precioIndex = findHeaderIndex_(headerMap, ['precio_referencia', 'precio referencia', 'precio_ref', 'precio'], 3);

  return values
    .map((row, index) => {
      const categoria = String(row[categoriaIndex] || '').trim();
      const producto = String(row[productoIndex] || '').trim();
      const unidad = String(row[unidadIndex] || '').trim();
      const precioReferencia = parseNumberLoose_(row[precioIndex], 0);

      if (!categoria || !producto) {
        return null;
      }

      return {
        id: `cat-${index + 2}-${normalizeText_(categoria)}-${normalizeText_(producto)}`,
        categoria,
        producto,
        unidad,
        precioReferencia,
      };
    })
    .filter(Boolean);
}

function uniqueValues_(values) {
  return Array.from(new Set(values.filter((value) => String(value || '').trim() !== '')));
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
