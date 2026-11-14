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
    if (d.action === 'setLoteInfo') { checkAdmin_(d.password); return setLoteInfo_(d.lote, d.tracking, d.entrega); }
    if (d.action === 'setOrderLote') { checkAdmin_(d.password); return setOrderLote_(d.row, d.lote); }
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
    t(d.extras_resumen), totalNum, t(d.id), ''];

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Pedidos') || ss.insertSheet('Pedidos');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Fecha', 'Equipo', 'Equipación', 'Temporada', 'Producto', 'Talla', 'Cantidad', 'Nombre',
        'Número', 'Parches', 'Cliente', 'Teléfono', 'Extras', 'Total', 'ID', 'Telegram']);
    }
    const [cE, cL, cC] = metaCols_(sh); // columnas Estado / Lote / Coste (se buscan por nombre)
    sh.appendRow(row);
    const r = sh.getLastRow();
    sh.getRange(r, 14).setNumberFormat('#,##0.00 "€"'); // formato dinero
    sh.getRange(r, cE).setValue('Pendiente');
    sh.getRange(r, cL).setValue(currentLote_());
    sh.getRange(r, cC).setValue(costNum).setNumberFormat('#,##0.00 "€"');

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

// Las columnas de la app se buscan por su nombre en la fila 1 (y se crean al final si no existen),
// así nunca pisan datos que ya tuvieras en tu hoja.
const META = ['Estado pedido', 'Lote pedido', 'Coste pedido'];
function metaCols_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 16);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let next = lastCol;
  return META.map(name => {
    const i = head.indexOf(name);
    if (i >= 0) return i + 1;
    next++;
    sh.getRange(1, next).setValue(name);
    return next;
  });
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace('€', '').replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function listOrders_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const costs = getPrices_().costs;
  const base = { ok: true, lote: currentLote_(), lotes: lotesMeta_(), costs: costs, orders: [] };
  if (!sh || sh.getLastRow() < 2) return json_(base);
  const [cE, cL, cC] = metaCols_(sh);
  const tz = Session.getScriptTimeZone();
  base.orders = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().map((r, i) => {
    const qty = Number(r[6]) || 1;
    const stored = num_(r[cC - 1]);
    return {
      row: i + 2,
      fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0]),
      equipo: r[1], equipacion: r[2], temporada: r[3], producto: r[4], talla: r[5], cantidad: qty,
      extras: r[12], cliente: r[10], telefono: r[11],
      total: num_(r[13]),
      estado: r[cE - 1] || 'Pendiente',
      lote: Number(r[cL - 1]) || 1,
      // si el pedido no tiene coste guardado, se calcula con el coste actual del producto
      coste: stored > 0 ? stored : Math.round((costs[r[4]] || 0) * qty * 100) / 100
    };
  });
  return json_(base);
}

function setStatus_(row, estado) {
  if (['Pendiente', 'Pedido', 'Anulado'].indexOf(estado) < 0) throw new Error('Estado no válido');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  row = parseInt(row);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Fila no válida');
  const cE = metaCols_(sh)[0];
  sh.getRange(row, cE).setValue(estado);
  return json_({ ok: true });
}

function setOrderLote_(row, lote) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  row = parseInt(row); lote = parseInt(lote);
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Fila no válida');
  if (!(lote >= 1 && lote <= currentLote_())) throw new Error('Lote no válido');
  sh.getRange(row, metaCols_(sh)[1]).setValue(lote);
  return json_({ ok: true });
}

// Datos de cada lote: { "1": { cerrado:'dd/MM/yyyy', tracking:'...', entrega:'yyyy-MM-dd' } }
function lotesMeta_() {
  const raw = PropertiesService.getScriptProperties().getProperty('LOTES');
  return raw ? JSON.parse(raw) : {};
}
function saveLotesMeta_(m) { PropertiesService.getScriptProperties().setProperty('LOTES', JSON.stringify(m)); }

function closeLote_() {
  const old = currentLote_(), n = old + 1;
  const m = lotesMeta_();
  m[old] = Object.assign({}, m[old], { cerrado: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy') });
  saveLotesMeta_(m);
  PropertiesService.getScriptProperties().setProperty('LOTE', String(n));
  return json_({ ok: true, lote: n, lotes: m });
}

function setLoteInfo_(lote, tracking, entrega) {
  lote = parseInt(lote);
  if (!(lote >= 1 && lote <= currentLote_())) throw new Error('Lote no válido');
  entrega = String(entrega || '');
  if (entrega && !/^\d{4}-\d{2}-\d{2}$/.test(entrega)) throw new Error('Fecha no válida');
  const m = lotesMeta_();
  m[lote] = Object.assign({}, m[lote], { tracking: String(tracking || '').trim().slice(0, 100), entrega: entrega });
  saveLotesMeta_(m);
  return json_({ ok: true, lotes: m });
}

// ═══ Ejecútala UNA sola vez a mano desde el editor ═══
// Cierra el lote actual con todo lo que ya está "Pedido" y deja lo "Pendiente"
// en un lote nuevo abierto, para que lo cierres tú cuando quieras.
function separarPedidosYa() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SEPARADO')) { Logger.log('Ya se ejecutó antes. No hago nada.'); return; }
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  if (!sh || sh.getLastRow() < 2) { Logger.log('No hay pedidos.'); return; }
  const actual = currentLote_(), nuevo = actual + 1;
  const [cE, cL] = metaCols_(sh);
  const n = sh.getLastRow() - 1;
  const rngE = sh.getRange(2, cE, n, 1), rngL = sh.getRange(2, cL, n, 1);
  const est = rngE.getValues(), lot = rngL.getValues();
  let cerrados = 0, pendientes = 0;
  for (let i = 0; i < n; i++) {
    if ((Number(lot[i][0]) || 1) !== actual) continue;           // solo toca el lote actual
    const e = est[i][0] || 'Pendiente';
    if (e === 'Pendiente') { lot[i][0] = nuevo; pendientes++; }   // pasa al lote nuevo (abierto)
    else { lot[i][0] = actual; if (e === 'Pedido') cerrados++; }  // Pedido / Anulado se quedan en el lote cerrado
  }
  rngL.setValues(lot);
  const m = lotesMeta_();
  m[actual] = Object.assign({}, m[actual], { cerrado: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy') });
  saveLotesMeta_(m);
  props.setProperty('LOTE', String(nuevo));
  props.setProperty('SEPARADO', '1');
  Logger.log('Lote ' + actual + ' cerrado con ' + cerrados + ' pedidos. Lote ' + nuevo + ' abierto con ' + pendientes + ' pendientes.');
}
