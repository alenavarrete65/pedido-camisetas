// ═══ Propiedades del script (Configuración del proyecto → Propiedades de secuencia de comandos) ═══
//   ADMIN_PASSWORD      → la contraseña con la que entras como admin
//   TELEGRAM_BOT_TOKEN  → token que te da @BotFather
//   TELEGRAM_CHAT_ID    → tu chat id (en grupos empieza por -100…)

const DEFAULTS = {
  products: {
    'Versión Fan Adulto': 15, 'Versión Player Adulto': 20, 'Versión Mujer': 15,
    'Versión Niño': 20, 'Versión Bebé': 10, 'Versión Retro': 20,
    'Conjunto Entrenamiento Con Mangas': 25, 'Conjunto Entrenamiento Sin Mangas': 25,
    'Chándal Media Cremallera': 30, 'Chándal Cremallera Completa': 37.5,
    'NBA Bordado': 20, 'NBA Termoimpresión': 22.5, 'NFL': 25
  },
  extras: { nombre: 1, numero: 2, parches: 1 },
  bigSize: 1
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function getPrices_() {
  const raw = PropertiesService.getScriptProperties().getProperty('PRICES');
  const s = raw ? JSON.parse(raw) : {};
  return {
    products: Object.assign({}, DEFAULTS.products, s.products),
    extras: Object.assign({}, DEFAULTS.extras, s.extras),
    bigSize: s.bigSize != null ? s.bigSize : DEFAULTS.bigSize,
    costs: s.costs || {}
  };
}

function doGet() {
  const p = getPrices_();
  delete p.costs; // los costes nunca se publican
  return json_({ ok: true, prices: p });
}

function checkAdmin_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 5) throw new Error('Demasiados intentos. Espera 10 minutos.');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real || pw !== real) {
    cache.put('fails', String(fails + 1), 600);
    Utilities.sleep(1000);
    throw new Error('Contraseña incorrecta');
  }
  cache.remove('fails');
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'login') { checkAdmin_(d.password); return json_({ ok: true }); }
    if (d.action === 'setPrices') { checkAdmin_(d.password); return setPrices_(d.prices); }
    if (d.action === 'listOrders') { checkAdmin_(d.password); return listOrders_(); }
    if (d.action === 'setStatus') { checkAdmin_(d.password); return setStatus_(d.row, d.estado); }
    if (d.action === 'closeLote') { checkAdmin_(d.password); return closeLote_(); }
    return saveOrder_(d);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function setPrices_(p) {
  const clean = { products: {}, extras: {}, costs: {}, bigSize: 0 };
  const val = (v, name) => {
    v = Number(v);
    if (!(v >= 0 && v <= 1000)) throw new Error('Precio no válido: ' + name);
    return v;
  };
  Object.keys(DEFAULTS.products).forEach(k => clean.products[k] = val(p.products[k], k));
  Object.keys(DEFAULTS.extras).forEach(k => clean.extras[k] = val(p.extras[k], k));
  Object.keys(DEFAULTS.products).forEach(k => clean.costs[k] = val(p.costs && p.costs[k] != null ? p.costs[k] : 0, 'coste ' + k));
  clean.bigSize = val(p.bigSize, 'suplemento talla');
  PropertiesService.getScriptProperties().setProperty('PRICES', JSON.stringify(clean));
  return json_({ ok: true });
}

function saveOrder_(d) {
  const cache = CacheService.getScriptCache();
  if (d.id && cache.get('o_' + d.id)) return json_({ ok: true, duplicate: true });

  // El precio se recalcula aquí: nunca nos fiamos del que manda el navegador
  const P = getPrices_();
  const base = P.products[d.producto];
  if (base === undefined) throw new Error('Producto no válido');
  const qty = Math.min(99, Math.max(1, parseInt(d.cantidad) || 1));
  const ids = (d.extras_ids || []).filter(id => P.extras[id] !== undefined);
  const unit = base + (['3XL', '4XL'].indexOf(d.size_raw) >= 0 ? P.bigSize : 0) +
    ids.reduce((a, id) => a + P.extras[id], 0);
  const totalNum = Math.round(unit * qty * 100) / 100;           // número real, para poder sumar en la hoja
  const costNum = Math.round((P.costs[d.producto] || 0) * qty * 100) / 100; // coste de este pedido (se guarda fijo)
  const totalTxt = String(totalNum).replace('.', ',') + '€';     // texto, solo para Telegram

  // Evita fórmulas maliciosas en la hoja (=, +, -, @) y textos enormes
  const t = v => { v = String(v == null ? '' : v).slice(0, 200); return /^[=+\-@]/.test(v) ? "'" + v : v; };
  const row = [t(d.fecha), t(d.equipo), t(d.equipacion), t(d.temporada), t(d.camiseta), t(d.talla), qty,
    t(d.nombre_camiseta), t(d.numero_camiseta), t(d.parches), t(d.contacto_nombre), t(d.contacto_telefono),
    t(d.extras_resumen), totalNum, t(d.id), '', 'Pendiente', currentLote_(), costNum];

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Pedidos') || ss.insertSheet('Pedidos');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Fecha', 'Equipo', 'Equipación', 'Temporada', 'Producto', 'Talla', 'Cantidad', 'Nombre',
        'Número', 'Parches', 'Cliente', 'Teléfono', 'Extras', 'Total', 'ID', 'Telegram', 'Estado', 'Lote', 'Coste']);
    }
    if (sh.getRange(1, 17).getValue() === '') sh.getRange(1, 17, 1, 3).setValues([['Estado', 'Lote', 'Coste']]);
    sh.appendRow(row);
    const r = sh.getLastRow();
    sh.getRange(r, 14).setNumberFormat('#,##0.00 "€"'); // formato dinero
    sh.getRange(r, 19).setNumberFormat('#,##0.00 "€"');

    const msg = '🛒 NUEVO PEDIDO\n' +
      '👤 ' + row[10] + ' · ' + row[11] + '\n' +
      '⚽ ' + row[1] + ' — ' + row[2] + ' (' + row[3] + ')\n' +
      '👕 ' + row[4] + ' · ' + row[5] + ' × ' + qty + '\n' +
      '✏️ ' + (row[12] || 'Sin extras') + '\n' +
      '💰 ' + totalTxt;
    const status = sendTelegram_(msg);
    sh.getRange(r, 16).setValue(status); // aquí verás si Telegram falló y por qué
  } finally {
    lock.releaseLock();
  }
  if (d.id) cache.put('o_' + d.id, '1', 600);
  return json_({ ok: true });
}

function sendTelegram_(text) {
  try {
    const p = PropertiesService.getScriptProperties();
    const token = p.getProperty('TELEGRAM_BOT_TOKEN'), chat = p.getProperty('TELEGRAM_CHAT_ID');
    if (!token || !chat) return 'ERROR: faltan TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID';
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chat, text: text })
    });
    return res.getResponseCode() === 200 ? 'OK' : 'ERROR ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 150);
  } catch (err) {
    return 'ERROR: ' + err.message;
  }
}

// Ejecútala a mano desde el editor: autoriza los permisos y comprueba Telegram
function testTelegram() { Logger.log(sendTelegram_('✅ Prueba de Telegram desde Apps Script')); }

// Ejecútala UNA vez a mano: convierte los totales antiguos (texto) en números con formato dinero
function arreglarTotales() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rng = sh.getRange(2, 14, n, 1);
  const vals = rng.getValues().map(r => {
    if (typeof r[0] === 'number') return [r[0]];
    const num = parseFloat(String(r[0]).replace('€', '').replace(/\s/g, '').replace(',', '.'));
    return [isNaN(num) ? r[0] : num];
  });
  rng.setValues(vals);
  rng.setNumberFormat('#,##0.00 "€"');
}

// ═══ Pedidos, estados y lotes ═══
function currentLote_() { return Number(PropertiesService.getScriptProperties().getProperty('LOTE') || 1); }

function listOrders_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const base = { ok: true, lote: currentLote_(), costs: getPrices_().costs, orders: [] };
  if (!sh || sh.getLastRow() < 2) return json_(base);
  const tz = Session.getScriptTimeZone();
  base.orders = sh.getRange(2, 1, sh.getLastRow() - 1, 19).getValues().map((r, i) => ({
    row: i + 2,
    fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0]),
    equipo: r[1], equipacion: r[2], temporada: r[3], producto: r[4], talla: r[5], cantidad: r[6],
    extras: r[12], cliente: r[10], telefono: r[11],
    total: Number(r[13]) || 0, estado: r[16] || 'Pendiente', lote: Number(r[17]) || 1, coste: Number(r[18]) || 0
  }));
  return json_(base);
}

function setStatus_(row, estado) {
  if (['Pendiente', 'Pedido', 'Anulado'].indexOf(estado) < 0) throw new Error('Estado no válido');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  row = parseInt(row);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Fila no válida');
  sh.getRange(row, 17).setValue(estado);
  return json_({ ok: true });
}

function closeLote_() {
  const n = currentLote_() + 1;
  PropertiesService.getScriptProperties().setProperty('LOTE', String(n));
  return json_({ ok: true, lote: n });
}
