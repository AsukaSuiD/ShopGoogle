// Code.gs
/** @OnlyCurrentDoc */
const SHEET_ID = '';                           // можно оставить пустым для текущей таблицы
const BUSINESS_TZ = 'Europe/Moscow';           // единая бизнес-таймзона
function getTz_(){ return BUSINESS_TZ; }       // всегда используем одну TZ

/* ====== КЭШ (5–10 минут) ====== */
function cache_(){ return CacheService.getScriptCache(); }
function cacheGetJson_(key){
  try{ const v = cache_().get(key); return v ? JSON.parse(v) : null; }catch(e){ return null; }
}
function cachePutJson_(key, obj, sec){
  try{ cache_().put(key, JSON.stringify(obj), sec || 300); }catch(e){}
}
function cacheDel_(keys){
  try{ cache_().removeAll(keys); }catch(e){
    (keys||[]).forEach(k=>{ try{ cache_().remove(k); }catch(_){} });
  }
}

const NALICHIE_META_KEY = 'NALICHIE::META';
const NALICHIE_STOCK_PREFIX = 'NALICHIE::STOCK::';
const NALICHIE_TTL_SEC = 600;
const NALICHIE_STOCK_CHUNK = 50;

function invalidateNalichieCache_(){
  const meta = cacheGetJson_(NALICHIE_META_KEY);
  const keys = new Set([NALICHIE_META_KEY]);
  if (meta && Array.isArray(meta.stockChunkKeys)){
    meta.stockChunkKeys.forEach(k => { if (k) keys.add(k); });
  }
  cacheDel_(Array.from(keys));
}

function invalidateDailyReportCache_(){
  cacheDel_(['DAILY_BOOT_V1']);
}
/** Вставка HTML-парциалов в шаблон Index.html */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function doGet() {
  const data    = getNalichieBootstrap();
  const diagBoot= getDiagnosticsBootstrap();  // быстрый прелоад диагностики
  const preBoot = getPreordersBootstrap();    // быстрый прелоад предзаказов

  const t = HtmlService.createTemplateFromFile('Index');
  t.NALICHIE_JSON   = JSON.stringify(data);
  t.DIAG_BOOT_JSON  = JSON.stringify(diagBoot);
  t.PRE_BOOT_JSON   = JSON.stringify(preBoot);
  return t.evaluate().setTitle('Магазины App');
}

/*
Структура листов и КЛЮЧИ row1 (строго):
- "Справочники" (колоночный формат): store, pre_day, city, staff_id, staff_name, staff_color,
  stock_statuses, preorder_statuses, diagnostic_statuses, condition, appearance, payments, item_type, complect
- "Каталог моделей": model_name, memory, color [, pre_price]
- "Склад": city, condition, model_name, memory, color, imei, sale_price, stock_statuses, note
- "Склад аксессуаров": store, model_name, sku, sale_price, qty
- "Продажи": id, date, store, staff, item_type, condition, model_name, memory, color, imei_or_sku, total, payments, sdacha, customer, phone, zarplata, note
- "Смены": id, date_vyhoda, vremya_vyhoda, store, staff_id, staff, pre_day [, ...]
- "Предзаказы": id, date, store, staff, preorder_statuses, model_name, memory, color, pre_price, prepay, payments, customer, phone, zarplata, note, pre_imei
- "Диагностика": id, purchase_date, intake_date, issued_date, store, staff, issued_staff, model_name, memory, color, imei, complect, neispravnost, appearance, diagnostic_statuses, diag_pay, payments, customer, phone_klienta, note
*/

/* ============================== ПРЕЛОАД НАЛИЧИЯ ============================== */
function preloadNalichie_(ss) {
  if (!ss) {
    ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  }

  // Справочники
  const dictTable = readSheetObjectsStrict_(ss, 'Справочники');
  const D = buildDictsColumnar_(dictTable);
  const cities     = D.cities;
  const conditions = D.conditions;
  const stores     = D.stores;
  const staff      = D.staff;     // [{code, name}]
  const payments   = D.payments;
  const complects  = D.complects;
  const preorder_statuses = D.preorder_statuses;
  const diagnostic_statuses = D.diagnostic_statuses;
  const appearance = D.appearance;

  // Карта цветов сотрудников (имя -> цвет)
  const staffColors = {};
  (dictTable || []).forEach(r => {
    const name = safeStr_(r['staff_name']);
    const color = safeStr_(r['staff_color']);
    if (name && color && !staffColors[name]) staffColors[name] = color;
  });

  // Каталог моделей (с поддержкой pre_price, если есть такой столбец)
  const catalogRows = readSheetObjectsStrict_(ss, 'Каталог моделей');
  const models   = uniqNotEmpty_(catalogRows.map(r => r['model_name']));
  const memories = uniqNotEmpty_(catalogRows.map(r => r['memory']));
  const colors   = uniqNotEmpty_(catalogRows.map(r => r['color']));
  const hasPrePrice = (catalogRows._headers || []).includes('pre_price');
  const catalogTable = (catalogRows || []).map(r => ({
    model_name: safeStr_(r['model_name']),
    memory    : safeStr_(r['memory']),
    color     : safeStr_(r['color']),
    pre_price : hasPrePrice ? safeStr_(r['pre_price']) : ''
  })).filter(x => x.model_name);

  // Склад (телефоны)
  const stockRows = readSheetObjectsStrict_(ss, 'Склад');
  const cityMap      = toMap_(cities, 'code', 'name');
  const conditionMap = toMap_(conditions, 'code', 'name');

  const stock = stockRows.map(r => {
    const cityCode = safeStr_(r['city']);
    const condCode = safeStr_(r['condition']);
    return {
      city_code:       cityCode,
      city_name:       cityMap[cityCode] || cityCode,
      condition_code:  condCode,
      condition_name:  conditionMap[condCode] || condCode,
      item_type:       inferItemType_(r),
      model_name:      safeStr_(r['model_name']),
      memory:          safeStr_(r['memory']),
      color:           safeStr_(r['color']),
      imei:            safeStr_(r['imei']),
      sale_price:      r['sale_price'] ?? '',
      stock_statuses:  safeStr_(r['stock_statuses']),
      note:            safeStr_(r['note'])
    };
  }).filter(x =>
    x.city_code || x.condition_code || x.model_name || x.memory || x.color || x.imei
  );

  // Склад аксессуаров
  const accRows = readSheetObjectsStrict_(ss, 'Склад аксессуаров');
  const accessoryStock = (accRows || []).map(r => {
    const qtyNum = Number(r['qty']);
    return {
      store:      safeStr_(r['store']),
      model_name: safeStr_(r['model_name']),
      sku:        safeStr_(r['sku']),
      sale_price: r['sale_price'] ?? '',
      qty:        isNaN(qtyNum) ? 0 : qtyNum,
      note:       safeStr_(r['note'])
    };
  }).filter(x => x.model_name || x.sku);

  // Подписи колонок (row2) — только для UI
  function labelOfStock(key) {
    const labels = stockRows._labels || {};
    return labels[key] || ({
      city: 'Город', condition: 'Состояние', model_name: 'Модель/Название',
      memory: 'Память', color: 'Цвет', imei: 'IMEI', sale_price: 'Цена продажи',
      stock_statuses: 'Статус склада', note: 'Примечание'
    }[key] || key);
  }
  function labelOfAcc(key) {
    const labels = accRows._labels || {};
    return labels[key] || ({
      model_name:'Модель/Название', sku:'SKU', sale_price:'Цена продажи', qty:'Кол-во', note:'Примечание'
    }[key] || key);
  }

  return {
    dicts: {
      cities, conditions, stores, staff, payments,
      complects,
      preorder_statuses,
      diagnostic_statuses, appearance,
      staffColors
    },
    catalog: { models, memories, colors, table: catalogTable },
    stock,
    titles: {
      city: labelOfStock('city'),
      condition: labelOfStock('condition'),
      model_name: labelOfStock('model_name'),
      memory: labelOfStock('memory'),
      color: labelOfStock('color'),
      imei: labelOfStock('imei'),
      sale_price: labelOfStock('sale_price'),
      stock_statuses: labelOfStock('stock_statuses'),
      note: labelOfStock('note')
    },
    accessoryStock,
    accTitles: {
      model_name: labelOfAcc('model_name'),
      sku: labelOfAcc('sku'),
      sale_price: labelOfAcc('sale_price'),
      qty: labelOfAcc('qty'),
      note: labelOfAcc('note')
    }
  };
}

function getNalichieBootstrap(){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const meta = cacheGetJson_(NALICHIE_META_KEY);
  if (meta && Array.isArray(meta.stockChunkKeys)){
    if (meta.stockChunkKeys.length){
      const stock = [];
      let miss = false;
      for (let i = 0; i < meta.stockChunkKeys.length; i++){
        const key = meta.stockChunkKeys[i];
        const chunk = cacheGetJson_(key);
        if (!Array.isArray(chunk) || !chunk.length){
          miss = true;
          break;
        }
        Array.prototype.push.apply(stock, chunk);
      }
      if (!miss){
        const { stockChunkKeys, ...rest } = meta;
        return { ...rest, stock };
      }
    } else {
      const { stockChunkKeys, ...rest } = meta;
      return { ...rest, stock: [] };
    }
  }

  invalidateNalichieCache_();
  const data = preloadNalichie_(ss);
  const stock = Array.isArray(data.stock) ? data.stock : [];
  const stockChunkKeys = [];
  if (stock.length){
    for (let i = 0; i < stock.length; i += NALICHIE_STOCK_CHUNK){
      const chunk = stock.slice(i, i + NALICHIE_STOCK_CHUNK);
      const key = `${NALICHIE_STOCK_PREFIX}${String(stockChunkKeys.length).padStart(4,'0')}`;
      cachePutJson_(key, chunk, NALICHIE_TTL_SEC);
      stockChunkKeys.push(key);
    }
  }
  const { stock: _ignored, ...rest } = data;
  const metaPayload = { ...rest, stockChunkKeys };
  cachePutJson_(NALICHIE_META_KEY, metaPayload, NALICHIE_TTL_SEC);
  return data;
}

/* Если в складе есть явный столбец item_type — используем */
function inferItemType_(row) {
  const keys = Object.keys(row || {});
  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i]).toLowerCase().replace(/\s+/g, '');
    if (k === 'item_type' || k === 'itemtype' || k === 'типтовара' || k === 'вид') return safeStr_(row[keys[i]]);
  }
  return '';
}

/* ===================== ПРОДАЖИ ===================== */
function createSale(doc) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();
    if (!doc) throw new Error('Пустой документ продажи');

    const saleId   = genSaleId_(tz);                         // SALE-YYYYMMDDHHMMSS
    const dateOut  = toDdMmYyyy_(safeStr_(doc.date), tz);    // DD.MM.YYYY

    const store    = safeStr_(doc.store);
    const staff    = safeStr_(doc.staff);
    const itemType = 'Телефон';
    const condition= safeStr_(doc.condition);
    const model    = safeStr_(doc.model_name);
    const memory   = safeStr_(doc.memory);
    const color    = safeStr_(doc.color);
    const imei     = safeStr_(doc.imei_or_sku);
    const totalStr = safeStr_(doc.total);
    const total    = totalStr ? Number(totalStr.toString().replace(',', '.')) : 0;
    const sdacha   = safeStr_(doc.sdacha);
    const customer = safeStr_(doc.customer);
    const phone    = safeStr_(doc.phone);
    const zp       = safeStr_(doc.zarplata);
    const note     = safeStr_(doc.note);
    const payments = Array.isArray(doc.payments) ? doc.payments : [];

    if (!dateOut)  throw new Error('Укажите дату');
    if (!store)    throw new Error('Укажите магазин');
    if (!staff)    throw new Error('Укажите сотрудника');
    if (!condition)throw new Error('Укажите состояние');
    if (!imei)     throw new Error('Нет IMEI/SKU');
    if (!(total >= 0)) throw new Error('Некорректная сумма "Итого"');

    // Обновить склад: статус -> "Продан" (+ condition)
    const stockUpd = updateStockByImei_(ss, imei, { status: 'Продан', condition });
    if (!stockUpd || !stockUpd.updated) {
      throw new Error(stockUpd?.reason === 'already_sold' ? 'Позиция уже продана' : 'Позиция на складе не найдена');
    }

    // Запись
    const saleRow = {
      id: saleId, date: dateOut, store, staff,
      item_type: itemType, condition,
      model_name: model, memory, color, imei_or_sku: imei,
      total,
      payments: serializePayments_(payments), sdacha,
      customer, phone, zarplata: zp, note
    };
    appendSale_(ss, saleRow);
    invalidateNalichieCache_();
    invalidateDailyReportCache_();

    return { ok: true, id: saleId, imei, newStatus: 'Продан' };
  } finally {
    lock.releaseLock();
  }
}

function createAccessorySale(doc) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    if (!doc) throw new Error('Пустой документ продажи аксессуара');

    const saleId  = genSaleId_(tz);
    const dateOut = toDdMmYyyy_(safeStr_(doc.date), tz);

    const store   = safeStr_(doc.store);
    const staff   = safeStr_(doc.staff);
    const itemType= 'Аксессуар';
    const model   = safeStr_(doc.model_name);
    const sku     = safeStr_(doc.imei_or_sku) || safeStr_(doc.sku);
    const total   = Number((safeStr_(doc.total) || '0').toString().replace(',', '.'));
    const sdacha  = safeStr_(doc.sdacha);
    const zp      = safeStr_(doc.zarplata);
    const payments= Array.isArray(doc.payments) ? doc.payments : [];

    if (!dateOut) throw new Error('Укажите дату');
    if (!store)   throw new Error('Укажите магазин');
    if (!staff)   throw new Error('Укажите сотрудника');
    if (!model)   throw new Error('Выберите аксессуар');
    if (!sku)     throw new Error('Нет SKU аксессуара');
    if (!(total >= 0)) throw new Error('Некорректная сумма "Итого"');

    const upd = decrementAccessoryQtyBySku_(ss, sku, 1);
    if (!upd || !upd.updated) {
      throw new Error(upd?.reason === 'not_found' ? 'Аксессуар не найден на складе' : 'Не удалось списать количество');
    }

    const saleRow = {
      id: saleId, date: dateOut, store, staff,
      item_type: itemType, condition: '',
      model_name: model, memory: '', color: '',
      imei_or_sku: sku, total,
      payments: serializePayments_(payments), sdacha, customer: '', phone: '',
      zarplata: zp, note: ''
    };
    appendSale_(ss, saleRow);
    invalidateNalichieCache_();
    invalidateDailyReportCache_();

    return { ok: true, id: saleId, sku, decremented: 1 };
  } finally {
    lock.releaseLock();
  }
}

function createServiceSale(doc) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    if (!doc) throw new Error('Пустой документ продажи услуги');

    const saleId  = genSaleId_(tz);
    const dateOut = toDdMmYyyy_(safeStr_(doc.date), tz);

    const store   = safeStr_(doc.store);
    const staff   = safeStr_(doc.staff);
    const itemType= 'Услуга';
    const model   = safeStr_(doc.model_name);
    const payments= Array.isArray(doc.payments) ? doc.payments : [];
    const sdacha  = safeStr_(doc.sdacha);
    const zp      = safeStr_(doc.zarplata);

    if (!dateOut) throw new Error('Укажите дату');
    if (!store)   throw new Error('Укажите магазин');
    if (!staff)   throw new Error('Укажите сотрудника');
    if (!model)   throw new Error('Укажите название услуги');

    const total = sumPayments_(payments);

    const saleRow = {
      id: saleId, date: dateOut, store, staff,
      item_type: itemType, condition: '',
      model_name: model, memory: '', color: '',
      imei_or_sku: '', total,
      payments: serializePayments_(payments), sdacha, customer: '', phone: '',
      zarplata: zp, note: ''
    };
    appendSale_(ss, saleRow);
    invalidateDailyReportCache_();

    return { ok: true, id: saleId, total };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== ОБНОВЛЕНИЕ СКЛАДОВ ===================== */
function updateStockByImei_(ss, imei, payload) {
  const sh = ss.getSheetByName('Склад');
  if (!sh) throw new Error('Лист "Склад" не найден');
  const values = sh.getDataRange().getValues();
  if (values.length < 3) return { updated: false };

  const headers = (values[0] || []).map(safeStr_);
  const H = {}; headers.forEach((h, i) => H[h] = i);

  const idxImei   = H['imei'];
  const idxStatus = H['stock_statuses'];
  const idxCond   = H['condition'];
  if (idxImei == null || idxStatus == null) throw new Error('В "Склад" нет колонок imei/stock_statuses');

  const target = String(imei).trim();
  for (let r = 2; r < values.length; r++) {
    const rowImei = String(values[r][idxImei] ?? '').trim();
    if (!rowImei) continue;
    if (rowImei === target) {
      const prev = String(values[r][idxStatus] ?? '').trim();
      if (prev === 'Продан') return { updated: false, reason: 'already_sold' };
      sh.getRange(r + 1, idxStatus + 1).setValue(payload.status || 'Продан');
      if (idxCond != null && payload.condition) {
        sh.getRange(r + 1, idxCond + 1).setValue(payload.condition);
      }
      return { updated: true, row: r + 1 };
    }
  }
  return { updated: false, reason: 'not_found' };
}

function decrementAccessoryQtyBySku_(ss, sku, dec) {
  const sh = ss.getSheetByName('Склад аксессуаров');
  if (!sh) throw new Error('Лист "Склад аксессуаров" не найден');
  const values = sh.getDataRange().getValues();
  if (values.length < 3) return { updated: false };

  const headers = (values[0] || []).map(safeStr_);
  const H = {}; headers.forEach((h, i) => H[h] = i);

  const idxSku = H['sku'];
  const idxQty = H['qty'];
  if (idxSku == null || idxQty == null) throw new Error('В "Склад аксессуаров" нет колонок sku/qty');

  const target = String(sku).trim();
  for (let r = 2; r < values.length; r++) {
    const rowSku = String(values[r][idxSku] ?? '').trim();
    if (!rowSku) continue;
    if (rowSku === target) {
      const prevQty = Number(values[r][idxQty] ?? 0);
      if (isNaN(prevQty)) return { updated: false, reason: 'qty_nan' };
      if (prevQty <= 0)   return { updated: false, reason: 'no_stock' };
      const newQty = Math.max(0, prevQty - (dec || 1));
      sh.getRange(r + 1, idxQty + 1).setValue(newQty);
      return { updated: true, row: r + 1, from: prevQty, to: newQty };
    }
  }
  return { updated: false, reason: 'not_found' };
}

/* ===================== "ПРОДАЖИ": запись ===================== */
function appendSale_(ss, rowObj) {
  const meta = ensureSalesSheet_(ss);
  const sh = meta.sheet;

  const line = meta.headers.map(h => rowObj[h]);
  const idxDate = meta.headers.indexOf('date');
  const idxTotal = meta.headers.indexOf('total');
  if (idxDate >= 0) line[idxDate] = rowObj['date'];
  if (idxTotal >= 0) {
    const n = Number(rowObj['total']);
    line[idxTotal] = isNaN(n) ? rowObj['total'] : n;
  }
  sh.appendRow(line);
}

function ensureSalesSheet_(ss) {
  const headers = [
    'id','date','store','staff',
    'item_type','condition',
    'model_name','memory','color','imei_or_sku',
    'total','payments','sdacha',
    'customer','phone','zarplata','note'
  ];
  const labels = [
    'ID','Дата','Магазин','Сотрудник',
    'Тип товара','Вид',
    'Модель/Название','Память','Цвет','IMEI/SKU',
    'Сумма ₽','Способы оплаты','Сдача ₽',
    'Клиент','Телефон','Зарплата ₽','Примечание'
  ];

  let sh = ss.getSheetByName('Продажи');
  if (!sh) sh = ss.insertSheet('Продажи');
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < headers.length) {
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(2,1,1,labels.length).setValues([labels]);
  }
  return { sheet: sh, headers };
}

/* ===================== Форматы / утилиты сумм ===================== */
function genSaleId_(tz) { return 'SALE-' + Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss'); }
function toDdMmYyyy_(dateStr, tz) {
  const s = safeStr_(dateStr);
  if (!s) return Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${m[3]}.${m[2]}.${m[1]}`;
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'dd.MM.yyyy');
  return s;
}
function serializePayments_(payments) {
  if (!Array.isArray(payments) || !payments.length) return '';
  const parts = [];
  payments.forEach(p => {
    const method = safeStr_(p?.method || '');
    if (!method) return;
    const amount = safeStr_(p.amount || '');
    parts.push(method + ':' + amount);
  });
  return parts.join('; ');
}
function sumPayments_(payments) {
  if (!Array.isArray(payments) || !payments.length) return 0;
  let s = 0;
  payments.forEach(p => {
    const a = Number((p?.amount ?? 0).toString().replace(',', '.'));
    if (!isNaN(a)) s += a;
  });
  return s;
}

/* ===================== Справочники (row1-строго) ===================== */
function buildDictsColumnar_(rows) {
  const list = (col) => uniqNotEmpty_((rows || []).map(r => safeStr_(r[col]))).map(v => ({ code: v, name: v }));

  const cities     = list('city');
  const conditions = list('condition');
  const stores     = list('store');
  const payments   = list('payments');
  const item_types = list('item_type');
  const appearance = list('appearance');
  const complects  = list('complect');
  const preorder_statuses = list('preorder_statuses');
  const diagnostic_statuses = list('diagnostic_statuses');

  const staff = [];
  const seen = new Set();
  (rows || []).forEach(r => {
    const id   = safeStr_(r['staff_id']);
    const name = safeStr_(r['staff_name']);
    if (name && !seen.has(name)) {
      seen.add(name);
      staff.push({ code: id || name, name: name });
    }
  });

  return { cities, conditions, stores, staff, payments, item_types, appearance, complects, preorder_statuses, diagnostic_statuses };
}

/* ===================== Общие хелперы ===================== */
function readSheetObjectsStrict_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return Object.assign([], { _headers: [], _labels: {} });

  const values = sh.getDataRange().getValues();
  if (!values || !values.length) return Object.assign([], { _headers: [], _labels: {} });

  const headers   = (values[0] || []).map(safeStr_); // row1 — ключи
  const labelsRow = (values[1] || []).map(safeStr_); // row2 — подписи (только для UI)

  const labelsMap = {};
  for (let i = 0; i < headers.length; i++) labelsMap[headers[i]] = labelsRow[i] || '';

  const rows = [];
  for (let i = 2; i < values.length; i++) {          // данные с row3
    const row = values[i];
    if (!row) continue;
    const allEmpty = row.every(c => c === '' || c === null);
    if (allEmpty) continue;
    const obj = {};
    for (let j = 0; j <= headers.length - 1; j++) obj[headers[j]] = row[j];
    rows.push(obj);
  }
  rows._headers = headers;
  rows._labels  = labelsMap;
  return rows;
}
function readShiftsObjects_(ss) {
  const sh = ss.getSheetByName('Смены');
  if (!sh) return Object.assign([], { _headers: [], _labels: {} });

  const rng      = sh.getDataRange();
  const values   = rng.getValues();         // «сырые» значения
  const displays = rng.getDisplayValues();  // как видно в таблице (то, что нам нужно для времени)

  const headers = (values[0] || []).map(safeStr_);
  const labels  = (values[1] || []).map(safeStr_);
  const labelsMap = {};
  for (let i = 0; i < headers.length; i++) labelsMap[headers[i]] = labels[i] || '';

  const rows = [];
  for (let i = 2; i < values.length; i++) {
    const vRow = values[i];
    if (!vRow) continue;
    const allEmpty = vRow.every(c => c === '' || c === null);
    if (allEmpty) continue;

    const dRow = displays[i] || []; // строка «видимых» значений
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j];
      // ВАЖНО: для vremya_vyhoda берём строку отображения, без преобразований
      obj[key] = (key === 'vremya_vyhoda') ? safeStr_(dRow[j]) : vRow[j];
    }
    rows.push(obj);
  }
  rows._headers = headers;
  rows._labels  = labelsMap;
  return rows;
}
function uniqNotEmpty_(arr) { const out = []; const seen = Object.create(null); (arr || []).forEach(v => { const s = safeStr_(v); if (s && !seen[s]) { seen[s] = true; out.push(s); } }); return out; }
function toMap_(arr, key, val) { const m = {}; (arr || []).forEach(o => { if (o && o[key] != null) m[String(o[key])] = o[val]; }); return m; }
function safeStr_(v) { if (v == null) return ''; return String(v).trim(); }

/* ======================================================================= */
/* =============================   СМЕНЫ   =============================== */
/* ======================================================================= */

function getShiftsHeaders_(ss) {
  const sh = ss.getSheetByName('Смены');
  if (!sh) throw new Error('Лист "Смены" не найден');
  const headers = (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || []).map(safeStr_);
  return { sheet: sh, headers };
}
function ensureShiftsHeaders_(ss) {
  const sh = ss.getSheetByName('Смены');
  if (!sh) throw new Error('Лист "Смены" не найден');

  const lc = Math.max(1, sh.getLastColumn());
  const headers = (sh.getRange(1, 1, 1, lc).getValues()[0] || []).map(safeStr_);
  const labels = (sh.getRange(2, 1, 1, lc).getValues()[0] || []).map(safeStr_);

  // Обязательные ключи row1 (добавили device_type)
  const must = ['id', 'date_vyhoda', 'vremya_vyhoda', 'store', 'staff_id', 'staff', 'pre_day', 'device_type'];
  const labelMap = {
    id: 'ID',
    date_vyhoda: 'Дата',
    vremya_vyhoda: 'Время',
    store: 'Магазин',
    staff_id: 'Код сотрудника',
    staff: 'Сотрудник',
    pre_day: 'Премия, ₽',
    device_type: 'Тип устройства'
  };

  const cur = headers.slice();
  const missing = must.filter(k => cur.indexOf(k) < 0);
  if (missing.length) {
    // дописываем недостающие ключи в конец row1
    sh.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]);
    // и подписи (row2) для этих ключей
    const newLabels = missing.map(k => labelMap[k] || k);
    sh.getRange(2, labels.length + 1, 1, newLabels.length).setValues([newLabels]);
    cur.push(...missing);
  }
  return { sheet: sh, headers: cur };
}


/** ID вида SHF-YYYYMM-#### (по месяцу) */
function nextShiftId_(ss) {
  const { sheet: sh, headers } = getShiftsHeaders_(ss);
  const idIdx = headers.indexOf('id');
  if (idIdx < 0) throw new Error('В "Смены" нет колонки id');

  const tz = getTz_();
  const ym = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  const vals = sh.getDataRange().getValues();
  let maxN = 0;
  for (let r = 2; r < vals.length; r++) {
    const v = safeStr_(vals[r][idIdx]);
    const m = v.match(/^SHF-(\d{6})-(\d{4})$/);
    if (m && m[1] === ym) {
      const n = parseInt(m[2], 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
  }
  const next = String(maxN + 1).padStart(4, '0');
  return `SHF-${ym}-${next}`;
}

/** Профиль сотрудника из "Справочники" (по имени, без привязки к магазину) */
function findStaffProfile_(ss, staffName) {
  const rows = readSheetObjectsStrict_(ss, 'Справочники') || [];
  const name = safeStr_(staffName);
  const profile = { staff_id: '', pre_day: '', staff_color: '' };

  rows.forEach(r => {
    if (safeStr_(r['staff_name']) === name) {
      if (!profile.staff_id)   profile.staff_id   = safeStr_(r['staff_id']);
      if (!profile.pre_day)    profile.pre_day    = safeStr_(r['pre_day']);
      if (!profile.staff_color)profile.staff_color= safeStr_(r['staff_color']);
    }
  });
  return profile;
}

function getStaffColorMap_(ss) {
  const map = {};
  const rows = readSheetObjectsStrict_(ss, 'Справочники') || [];
  rows.forEach(r => {
    const name  = safeStr_(r['staff_name']);
    const color = safeStr_(r['staff_color']);
    if (name && color && !map[name]) map[name] = color;
  });
  return map;
}
function getStaffPreDayMap_(ss) {
  const map = {};
  const rows = readSheetObjectsStrict_(ss, 'Справочники') || [];
  rows.forEach(r => {
    const name  = safeStr_(r['staff_name']);
    const pre   = safeStr_(r['pre_day']);
    if (name && pre && !map[name]) map[name] = pre;
  });
  return map;
}
function getStorePreDayMap_(ss) {
  const map = {};
  const rows = readSheetObjectsStrict_(ss, 'Справочники') || [];
  rows.forEach(r => {
    const store = safeStr_(r['store']);
    const pre   = safeStr_(r['pre_day']);
    if (store && pre && map[store] == null) map[store] = pre;
  });
  return map;
}

/** Время/дата «сейчас» в TZ бизнеса (для шапки UI) */
function getShiftsBootstrap() {
  const tz = getTz_();
  const now = new Date();
  return {
    nowDate: Utilities.formatDate(now, tz, 'dd.MM.yyyy'),
    nowTime: Utilities.formatDate(now, tz, 'HH:mm:ss')
  };
}

/** Чек-ин смены: пишем текущие дату/время в TZ бизнеса */
function createShiftCheckIn(doc) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();

  if (!doc) throw new Error('Пустой документ смены');
  const store = safeStr_(doc.store);
  const staff = safeStr_(doc.staff);
  if (!store) throw new Error('Выберите магазин');
  if (!staff) throw new Error('Выберите сотрудника');

  const { sheet: sh, headers } = ensureShiftsHeaders_(ss);

  // профайл нужен только для staff_id/цвета; pre_day берём ТОЛЬКО по store
  const prof           = findStaffProfile_(ss, staff) || {};
  const storePreDayMap = getStorePreDayMap_(ss) || {};
  const preDay         = safeStr_(storePreDayMap[store] || '');

  const id = nextShiftId_(ss);
  const now = new Date();
  const date_vyhoda   = Utilities.formatDate(now, tz, 'dd.MM.yyyy');
  const vremya_vyhoda = Utilities.formatDate(now, tz, 'HH:mm');

  const rowObj = {
    id,
    date_vyhoda,
    vremya_vyhoda,
    store,
    staff_id: safeStr_(prof.staff_id),
    staff,
    pre_day: preDay,
    device_type: safeStr_(doc.device_type) // ← вот это добавили
  };
  const line = headers.map(h => (h in rowObj ? rowObj[h] : ''));
  sh.appendRow(line);

  invalidateDailyReportCache_();
  bumpShiftsCache_();
  return { ok:true, id, date_vyhoda, vremya_vyhoda, store, staff, staff_id: rowObj.staff_id, pre_day: preDay };
}

/* ===================== УЧЁТ СМЕН (row1-строго) ===================== */
function getShiftsLedger(params) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();
  params = params || {};

  const dateFromISO = safeStr_(params.dateFrom);
  const dateToISO   = safeStr_(params.dateTo);
  const fltStore    = safeStr_(params.store);
  const fltStaff    = safeStr_(params.staff);

  const shifts = readShiftsObjects_(ss) || [];
  const sales  = readSheetObjectsStrict_(ss, 'Продажи') || [];
  const preordersSheet = ss.getSheetByName('Предзаказы');
  const preorders = preordersSheet ? (readSheetObjectsStrict_(ss, 'Предзаказы') || []) : null;

  const staffColorMap = getStaffColorMap_(ss);
  const storePreDayMap = getStorePreDayMap_(ss);

  const fromMs = dateFromISO ? new Date(dateFromISO + 'T00:00:00').getTime() : null;
  const toMs   = dateToISO   ? new Date(dateToISO   + 'T23:59:59').getTime() : null;

  const key3 = (dateStr, store, staff) => [safeStr_(dateStr), safeStr_(store), safeStr_(staff)].join('||');

  // Индекс продаж по дате/магазину/сотруднику (row1-ключи)
  const salesByKey = {};
  (sales || []).forEach(r => {
    const dateRaw  = r['date'];
    const storeRaw = r['store'];
    const staffRaw = r['staff'];
    const k = key3(normalizeDateKey_(dateRaw, tz), storeRaw, staffRaw);
    (salesByKey[k] || (salesByKey[k] = [])).push(r);
  });

  // Индекс предзаказов (если есть отдельный лист)
  let preByKey = null;
  if (preorders) {
    preByKey = {};
    preorders.forEach(r => {
      const dateRaw  = r['date'];
      const storeRaw = r['store'];
      const staffRaw = r['staff'];
      const k = key3(normalizeDateKey_(dateRaw, tz), storeRaw, staffRaw);
      (preByKey[k] || (preByKey[k] = [])).push(r);
    });
  } else {
    // если предзаказы пишутся в "Продажи" — всё уже в salesByKey
  }

  const rowsOut = [];
  const totals = {
    sales_total:0,
    sales_phones:0,
    sales_accessories:0,
    sales_services:0,
    checks_total:0,
    preorders_sum:0,
    salary_total:0,
    positions: {
      phones: 0,
      accessories: 0,
      services: 0,
      preorders: 0
    }
  };

  for (const sh of (shifts || [])) {
    const dateKey = normalizeDateKey_(sh['date_vyhoda'], tz);
    const store   = safeStr_(sh['store']);
    const staff   = safeStr_(sh['staff']);
    const timeStr = safeStr_(sh['vremya_vyhoda']); // без преобразований!
    const id      = safeStr_(sh['id']);

    if (!dateKey) continue;

    const ms = ddmmyyyyToMs_(dateKey);
    if (fromMs && ms < fromMs) continue;
    if (toMs   && ms > toMs)   continue;
    if (fltStore && store !== fltStore) continue;
    if (fltStaff && staff !== fltStaff) continue;

    const k = key3(dateKey, store, staff);
    const daySales = salesByKey[k] || [];

    // Продажи: суммы по типам + ЗП с продаж
    let sumsPhones=0, sumsAcc=0, sumsSvc=0, checks=0, salaryFromSales=0;
    let cntPhones=0, cntAccessories=0, cntServices=0, cntPreorders=0;
    const preorderRowsFromSales = [];

    daySales.forEach(r=>{
      const t   = safeStr_(r['item_type']);
      const total = toNumberSafe_(r['total']);
      const zp    = toNumberSafe_(r['zarplata']);
      if (t === 'Телефон') {
        sumsPhones += total;
        cntPhones += 1;
        if (total>0) checks++;
        salaryFromSales += zp;
      }
      else if (t === 'Аксессуар') {
        sumsAcc    += total;
        cntAccessories += 1;
        if (total>0) checks++;
        salaryFromSales += zp;
      }
      else if (t === 'Услуга') {
        sumsSvc    += total;
        cntServices += 1;
        if (total>0) checks++;
        salaryFromSales += zp;
      }
      else if (t === 'Предзаказ') {
        preorderRowsFromSales.push(r);
      }
    });

    const salesTotal = sumsPhones + sumsAcc + sumsSvc;
    const avg = checks ? salesTotal / checks : 0;

    // Предзаказы: только сумма в ЭТОТ день (+ ЗП по ним)
    let preorderSum = 0;
    let salaryFromPreorders = 0;

    if (preByKey) {
      const arr = preByKey[k] || [];
      arr.forEach(r=>{
        const prepay = toNumberSafe_(r['prepay']);
        if (prepay > 0) {
          preorderSum += prepay;
        } else {
          preorderSum += sumPaymentsFromString_(safeStr_(r['payments']));
        }
        salaryFromPreorders += toNumberSafe_(r['zarplata']);
        cntPreorders += 1;
      });
    } else {
      // предзаказы идут строками в "Продажи"
      preorderRowsFromSales.forEach(r=>{
        preorderSum += sumPaymentsFromString_(safeStr_(r['payments']));
        salaryFromPreorders += toNumberSafe_(r['zarplata']);
        cntPreorders += 1;
      });
    }

    const preDay = toNumberSafe_( safeStr_(sh['pre_day']) || storePreDayMap[store] || '' );
    const salaryTotal = round2_(salaryFromSales + salaryFromPreorders + preDay);

    rowsOut.push({
      shift_id: id,
      date: dateKey,
      store,
      staff,
      pre_day: preDay,
      startTime: timeStr,
      staff_color: staffColorMap[staff] || '',
      device_type: safeStr_(sh['device_type']),
      positions: {
        phones: cntPhones,
        accessories: cntAccessories,
        services: cntServices,
        preorders: cntPreorders
      },
      sales: {
        phones: round2_(sumsPhones),
        accessories: round2_(sumsAcc),
        services: round2_(sumsSvc),
        total: round2_(salesTotal),
        avg: round2_(avg),
        checks
      },
      preorders: { total: round2_(preorderSum) },
      salary: {
        from_sales: round2_(salaryFromSales),
        from_preorders: round2_(salaryFromPreorders),
        total: round2_(salaryTotal)
      }
    });

    totals.sales_total += salesTotal;
    totals.sales_phones += sumsPhones;
    totals.sales_accessories += sumsAcc;
    totals.sales_services += sumsSvc;
    totals.checks_total += checks;
    totals.preorders_sum += preorderSum;
    totals.salary_total += salaryTotal;
    totals.positions.phones += cntPhones;
    totals.positions.accessories += cntAccessories;
    totals.positions.services += cntServices;
    totals.positions.preorders += cntPreorders;
  }

  totals.sales_total   = round2_(totals.sales_total);
  totals.sales_phones  = round2_(totals.sales_phones);
  totals.sales_accessories = round2_(totals.sales_accessories);
  totals.sales_services = round2_(totals.sales_services);
  totals.preorders_sum = round2_(totals.preorders_sum);
  totals.salary_total  = round2_(totals.salary_total);

  rowsOut.sort((a, b) => {
    const d = ddmmyyyyToMs_(b.date) - ddmmyyyyToMs_(a.date);
    if (d) return d;
    const toMin = s => {
      const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s || '');
      if (!m) return -1;
      return (+m[1]) * 60 + (+m[2]) + (m[3] ? (+m[3]) / 60 : 0);
    };
    return toMin(b.startTime) - toMin(a.startTime);
  });
  return { rows: rowsOut, totals };
}
function bumpShiftsCache_(){cachePutJson_('SHIFTS_LEDGER_BUMP', String(Date.now()), 86400);}

/* Быстрый кэш отчёта по сменам (по параметрам) */
function getShiftsLedgerCached(params){
  const bump = cacheGetJson_('SHIFTS_LEDGER_BUMP') || 0;
  const key = 'SHIFTS_LEDGER_V2::' + bump + '::' +
  Utilities.base64EncodeWebSafe(JSON.stringify(params||{}));
  let hit = cacheGetJson_(key);
  if (hit) return hit;
  const data = getShiftsLedger(params||{});
  cachePutJson_(key, data, 300);
  return data;
}

/* ====== нормализация дат/времени для учёта ====== */
function ddmmyyyyToMs_(s) {
  const m = String(s || '').match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!m) return NaN;
  const d = new Date(`${m[3]}-${m[2]}-${m[1]}T00:00:00`);
  return d.getTime();
}
function toNumberSafe_(v) {
  if (v == null || v === '') return 0;
  const n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
/** payments строка: поддерживает "Метод:Сумма" И "Метод Сумма" */
function sumPaymentsFromString_(s) {
  if (!s) return 0;
  let sum = 0;
  String(s).split(/[;,\n]/).forEach(part => {
    const p = String(part || '').trim();
    if (!p) return;
    // 1) "Метод:Сумма"
    let m = p.match(/^(.+?):\s*([0-9]+(?:[.,][0-9]+)?)$/);
    // 2) "Метод Сумма"
    if (!m) m = p.match(/^(.+?)\s+([0-9]+(?:[.,][0-9]+)?)$/);
    const val = m ? m[2] : '';
    const n = Number(String(val || '').replace(',', '.'));
    if (!isNaN(n)) sum += n;
  });
  return sum;
}

function isDate_(v){ return Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v); }

/** Любое значение → "DD.MM.YYYY" */
function normalizeDateKey_(v, tz){
  if (isDate_(v)) return Utilities.formatDate(v, tz, 'dd.MM.yyyy');
  const s = safeStr_(v);
  if (!s) return '';
  let m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/); if (m) return `${m[1]}.${m[2]}.${m[3]}`;
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);       if (m) return `${m[3]}.${m[2]}.${m[1]}`;
  // серийная дата Sheets (дни от 1899-12-30)
  const n = Number(s.replace(',', '.'));
  if (!isNaN(n) && n > 59 && n < 100000) {
    const base = new Date(Date.UTC(1899, 11, 30));
    base.setUTCDate(base.getUTCDate() + Math.floor(n));
    return Utilities.formatDate(base, tz, 'dd.MM.yyyy');
  }
  const d = new Date(s); if (!isNaN(d)) return Utilities.formatDate(d, tz, 'dd.MM.yyyy');
  return s;
}

function normalizeTimeKey_(v, tz) {
  if (v === null || v === undefined || v === '') return '';

  // Если это Date — форматируем в бизнес-TZ без секунд
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz || BUSINESS_TZ, 'HH:mm');
  }

  // Поддерживаем "HH:mm" и "HH:mm:ss" (и "HH.mm")
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[:.](\d{2})(?::(\d{2}))?$/);
  if (!m) return '';

  const hh = Math.min(Math.max(parseInt(m[1], 10), 0), 23);
  const mm = Math.min(Math.max(parseInt(m[2], 10), 0), 59);
  const pad = n => ('0' + n).slice(-2);
  return `${pad(hh)}:${pad(mm)}`; // ← только HH:mm
}

/* ======================================================================= */
/* ============================= ПРЕДЗАКАЗ =============================== */
/* ======================================================================= */

/** Старый ID (не используется для новых) */
function genPreId_(tz){ return 'PRE-' + Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss'); }

/** Следующий ID предзаказа в формате PRE-YYYYMM-0001 (нумерация по месяцу) */
function nextPreId_(ss, tz){
  const { sheet: sh, headers } = getPreordersHeaders_(ss);
  const idIdx = headers.indexOf('id');
  if (idIdx < 0) throw new Error('В "Предзаказы" нет колонки id');

  const ym = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  const vals = sh.getDataRange().getValues();
  let maxN = 0;

  for (let r = 2; r < vals.length; r++) {
    const v = String(vals[r][idIdx] || '');
    const m = v.match(/^PRE-(\d{6})-(\d{4})$/);
    if (m && m[1] === ym) {
      const n = parseInt(m[2], 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
  }
  const next = String(maxN + 1).padStart(4, '0');
  return `PRE-${ym}-${next}`;
}

/** Дефолтный статус для предзаказа */
function getDefaultPreorderStatus_(ss){
  const dict = buildDictsColumnar_(readSheetObjectsStrict_(ss, 'Справочники'));
  const arr = dict.preorder_statuses || [];
  if (arr.some(x=>safeStr_(x.name)==='Ожидание')) return 'Ожидание';
  return safeStr_(arr[0]?.name) || 'Ожидание';
}

/** Получаем лист и заголовки "Предзаказы" */
function getPreordersHeaders_(ss){
  const sh = ss.getSheetByName('Предзаказы');
  if (!sh) throw new Error('Лист "Предзаказы" не найден');
  const headers = (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || []).map(safeStr_);
  return { sheet: sh, headers };
}

/** Универсальная запись строки в "Предзаказы" по row1 */
function appendPreorder_(ss, rowObj){
  const { sheet: sh, headers } = getPreordersHeaders_(ss);
  const line = headers.map(h => {
    if (h === 'date') return rowObj['date'];
    if (h === 'pre_price' || h === 'prepay' || h === 'zarplata') {
      const n = Number(rowObj[h]); return isNaN(n) ? rowObj[h] : n;
    }
    return rowObj[h];
  });
  sh.appendRow(line);
}

/** In-place запись IMEI в первую строку данного id */
function setPreorderImeiInPlace_(ss, id, imei){
  const { sheet: sh, headers } = getPreordersHeaders_(ss);
  const idIdx = headers.indexOf('id');
  const imeiIdx = headers.indexOf('pre_imei');
  if (idIdx < 0 || imeiIdx < 0) throw new Error('В "Предзаказы" нет колонок id/pre_imei');

  const vals = sh.getDataRange().getValues();
  for (let r = 2; r < vals.length; r++) {
    if (safeStr_(vals[r][idIdx]) === id) {
      sh.getRange(r + 1, imeiIdx + 1).setValue(imei);
      return { ok:true, row:r+1 };
    }
  }
  return { ok:false, reason:'not_found' };
}

/** Создание предзаказа */
function createPreorder(doc){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();
    if (!doc) throw new Error('Пустой документ предзаказа');

    const id = nextPreId_(ss, tz);
    const date = toDdMmYyyy_(safeStr_(doc.date), tz);

    const store   = safeStr_(doc.store);
    const staff   = safeStr_(doc.staff);
    const status  = safeStr_(doc.preorder_statuses) || getDefaultPreorderStatus_(ss);
    const model   = safeStr_(doc.model_name);
    const memory  = safeStr_(doc.memory);
    const color   = safeStr_(doc.color);
    const pre_price = safeStr_(doc.pre_price);
    const paymentsArr = Array.isArray(doc.payments) ? doc.payments : [];
    const payments  = serializePayments_(paymentsArr);
    const prepay    = String(sumPayments_(paymentsArr));   // ← предоплата всегда = сумма оплат
    const customer  = safeStr_(doc.customer);
    const phone     = safeStr_(doc.phone);
    const zp        = safeStr_(doc.zarplata);
    const note      = safeStr_(doc.note);
    const pre_imei  = safeStr_(doc.pre_imei);

    if (!date)  throw new Error('Укажите дату');
    if (!store) throw new Error('Выберите магазин');
    if (!staff) throw new Error('Выберите сотрудника');

    appendPreorder_(ss, {
      id, date, store, staff,
      preorder_statuses: status,
      model_name: model, memory, color,
      pre_price, prepay, payments,
      customer, phone, zarplata: zp, note,
      pre_imei
    });

    cacheDel_(['PRE_BOOT_V2']);        // инвалидация кэша
    invalidateDailyReportCache_();
    return { ok:true, id };
  } finally {
    lock.releaseLock();
  }
}

/** Доплата по предзаказу (совместимость со старым UI) */
function addPreorderPayment(doc){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    const id      = safeStr_(doc.id);
    const date    = toDdMmYyyy_(safeStr_(doc.date), tz);
    const store   = safeStr_(doc.store);
    const staff   = safeStr_(doc.staff);
    const paymentsArr = Array.isArray(doc.payments) ? doc.payments : [];
    const payments= serializePayments_(paymentsArr);
    const prepay  = String(sumPayments_(paymentsArr));
    const note    = safeStr_(doc.note);

    if (!id)   throw new Error('Нет id предзаказа');
    if (!date) throw new Error('Укажите дату');
    if (!store) throw new Error('Выберите магазин');
    if (!staff) throw new Error('Выберите сотрудника');

    appendPreorder_(ss, { id, date, store, staff, prepay, payments, note });
    cacheDel_(['PRE_BOOT_V2']);
    invalidateDailyReportCache_();
    return { ok:true, id };
  } finally {
    lock.releaseLock();
  }
}

/** Статус предзаказа */
function updatePreorderStatus(doc){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    const id = safeStr_(doc.id);
    const newStatus = safeStr_(doc.preorder_statuses);
    const pre_imei  = safeStr_(doc.pre_imei);

    if (!id) throw new Error('Нет id предзаказа');
    if (!newStatus) throw new Error('Не указан статус');

    const staffColorMap = getStaffColorMap_(ss);
    const aggregated = aggregatePreorderById_(ss, id, null, staffColorMap);

    if (newStatus === 'Завершен' || newStatus === 'Завершён') {
      const need = round2_(toNumberSafe_(aggregated.pre_balance));
      if (need > 0.009) {
        throw new Error('Для завершения с остатком используйте операцию «Завершить» (finalizePreorder).');
      }
      const currentImei = safeStr_(aggregated.pre_imei);
      if (!currentImei && !pre_imei) {
        throw new Error('Для статуса «Завершен» необходимо указать IMEI');
      }
      if (pre_imei && pre_imei !== currentImei) {
        const r = setPreorderImeiInPlace_(ss, id, pre_imei);
        if (!r.ok) throw new Error('Не найдена исходная строка предзаказа для записи IMEI');
      }
      appendPreorder_(ss, { id, date: toDdMmYyyy_('', tz), preorder_statuses: 'Завершен' });
    } else {
      appendPreorder_(ss, { id, date: toDdMmYyyy_('', tz), preorder_statuses: newStatus });
    }
    cacheDel_(['PRE_BOOT_V2']);
    invalidateDailyReportCache_();
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

/** Завершение предзаказа: одна строка "Завершен" */
function finalizePreorder(doc){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    const id       = safeStr_(doc.id);
    const date     = toDdMmYyyy_(safeStr_(doc.date), tz);
    const store    = safeStr_(doc.store);
    const staff    = safeStr_(doc.staff);
    const pre_imei = safeStr_(doc.pre_imei);

    const paymentsArr = Array.isArray(doc.payments) ? doc.payments : [];
    const paymentsStr = serializePayments_(paymentsArr);
    const paySum      = sumPayments_(paymentsArr);

    if (!id)    throw new Error('Нет id предзаказа');
    if (!date)  throw new Error('Укажите дату завершения');
    if (!store) throw new Error('Выберите магазин');
    if (!staff) throw new Error('Выберите сотрудника');

    const staffColorMap = getStaffColorMap_(ss);
    const agg = aggregatePreorderById_(ss, id, null, staffColorMap);
    if (!agg || !agg.id)     throw new Error('Предзаказ не найден');
    if (agg.completed_at)    throw new Error('Предзаказ уже завершён');

    const price = toNumberSafe_(agg.pre_price);
    if (!(price > 0))        throw new Error('Не указана цена предзаказа');

    const needRaw = toNumberSafe_(agg.pre_balance);
    const need    = needRaw < 0 ? 0 : round2_(needRaw);

    const currentImei = safeStr_(agg.pre_imei);
    if (!currentImei && !pre_imei) throw new Error('IMEI обязателен для завершения');
    if (pre_imei && pre_imei !== currentImei) {
      const r = setPreorderImeiInPlace_(ss, id, pre_imei);
      if (!r.ok) throw new Error('Не найдена исходная строка предзаказа для записи IMEI');
    }

    if (need > 0) {
      if (!(paySum > 0)) throw new Error('Требуется внести доплату: ' + need);
      if (Math.abs(paySum - need) > 0.01) {
        throw new Error('Сумма оплат должна совпадать с остатком: ' + need);
      }
    } else {
      if (paySum > 0) throw new Error('Оплата не требуется: остаток уже 0');
    }

    appendPreorder_(ss, {
      id, date, store, staff,
      preorder_statuses: 'Завершен',
      prepay: String(need > 0 ? paySum : 0),
      payments: need > 0 ? paymentsStr : ''
    });

    cacheDel_(['PRE_BOOT_V2']);
    invalidateDailyReportCache_();
    const updated = aggregatePreorderById_(ss, id, null, staffColorMap);
    return { ok:true, aggregate: updated };
  } finally {
    lock.releaseLock();
  }
}

/** Просто записать/поправить IMEI (без создания строки) */
function upsertPreorderImei(doc){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

    const id = safeStr_(doc.id);
    const imei = safeStr_(doc.pre_imei);
    if (!id) throw new Error('Нет id предзаказа');
    if (!imei) throw new Error('Введите IMEI');

    const r = setPreorderImeiInPlace_(ss, id, imei);
    if (!r.ok) throw new Error('Не найдена исходная строка предзаказа');

    cacheDel_(['PRE_BOOT_V2']);
    invalidateDailyReportCache_();
    return { ok:true };
  } finally {
    lock.releaseLock();
  }
}

/** Агрегация по одному id (совместимо со старыми вызовами) */
function aggregatePreorderById_(ss, id, rowsOrOptions, staffColorMapOpt){
  let rows = null;
  let staffColorMap = staffColorMapOpt;

  if (Array.isArray(rowsOrOptions)) {
    rows = rowsOrOptions;
  } else if (rowsOrOptions && typeof rowsOrOptions === 'object') {
    if (Array.isArray(rowsOrOptions.rows)) rows = rowsOrOptions.rows;
    if (rowsOrOptions.staffColorMap && !staffColorMap) staffColorMap = rowsOrOptions.staffColorMap;
  }

  if (!Array.isArray(rows)) {
    rows = readSheetObjectsStrict_(ss, 'Предзаказы') || [];
  }
  const tz = getTz_();
  const group = (rows || []).filter(r => safeStr_(r['id']) === id);
  if (!group.length) return {};

  const toMs = s => ddmmyyyyToMs_(normalizeDateKey_(s, tz));
  const sortByDate = [...group].sort((a,b)=> (toMs(a['date'])||0) - (toMs(b['date'])||0));

  const prepaySum = group.reduce((s,r)=> s + toNumberSafe_(r['prepay']||0), 0);
  const zpSum     = group.reduce((s,r)=> s + toNumberSafe_(r['zarplata']||0), 0);

  const first = sortByDate[0];
  const last  = sortByDate[sortByDate.length-1];

  const pre_price_raw = safeStr_(first['pre_price']) || safeStr_(last['pre_price']);
  const pre_imeiRow   = group.find(r=>safeStr_(r['pre_imei']));
  const pre_imei      = pre_imeiRow ? pre_imeiRow['pre_imei'] : '';

  const paymentsExtra = group.reduce((s,r)=> s + (toNumberSafe_(r['prepay'])>0 ? 0 : sumPaymentsFromString_(safeStr_(r['payments']))), 0);

  const prepayTotal   = round2_(prepaySum + paymentsExtra);
  const pre_price_num = toNumberSafe_(pre_price_raw);
  const pre_balance   = round2_(pre_price_num - prepayTotal);

  const completedRow = [...sortByDate].filter(r => {
    const st = safeStr_(r['preorder_statuses']).toLowerCase();
    return st === 'завершен' || st === 'завершён';
  }).pop();
  const completed_at = completedRow ? normalizeDateKey_(completedRow['date'], tz) : '';
  const completed_by = completedRow ? safeStr_(completedRow['staff']) : '';

  if (!staffColorMap) {
    staffColorMap = getStaffColorMap_(ss);
  }
  const creatorStaffName = safeStr_(first['staff']);
  const staff_color = staffColorMap[creatorStaffName] || '';

  return {
    id,
    date: normalizeDateKey_(first['date'], tz),
    store: safeStr_(last['store']) || safeStr_(first['store']),
    staff: creatorStaffName,
    model_name: safeStr_(first['model_name']) || safeStr_(last['model_name']),
    memory: safeStr_(first['memory']) || safeStr_(last['memory']),
    color: safeStr_(first['color']) || safeStr_(last['color']),
    pre_imei: safeStr_(pre_imei) || '',
    pre_price: pre_price_raw,
    prepay: prepayTotal,
    pre_balance,
    payments: safeStr_(last['payments']),
    customer: safeStr_(last['customer']) || safeStr_(first['customer']),
    phone: safeStr_(last['phone']) || safeStr_(first['phone']),
    zarplata: round2_(zpSum),
    note: safeStr_(last['note']),
    preorder_statuses: safeStr_(last['preorder_statuses']) || 'Ожидание',
    completed_at, completed_by,
    staff_color
  };
}

/* Список предзаказов (агрегировано по id) — старый серверный способ */
function getPreorders(filters){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();
  filters = filters || {};

  const rows = readSheetObjectsStrict_(ss, 'Предзаказы') || [];
  if (!rows.length) return { rows: [], totals: { count:0, prepay:0, zarplata:0 } };

  const byId = {};
  (rows || []).forEach(r => {
    const id = safeStr_(r['id']);
    if (!id) return;
    (byId[id] || (byId[id] = [])).push(r);
  });

  const staffColorMap = getStaffColorMap_(ss);
  let list = Object.keys(byId).map(id => aggregatePreorderById_(ss, id, { rows: byId[id], staffColorMap }));

  // Фильтры
  const fDateFrom = safeStr_(filters.dateFrom);
  const fDateTo   = safeStr_(filters.dateTo);
  const fromMs = fDateFrom ? new Date(fDateFrom + 'T00:00:00').getTime() : null;
  const toMs   = fDateTo   ? new Date(fDateTo   + 'T23:59:59').getTime() : null;

  const fStore  = safeStr_(filters.store);
  const fStaff  = safeStr_(filters.staff);
  const fStat   = safeStr_(filters.preorder_statuses);
  const fModel  = safeStr_(filters.model_name);
  const fMemory = safeStr_(filters.memory);
  const fColor  = safeStr_(filters.color);
  const fCust   = safeStr_(filters.customer).toLowerCase();
  const fPhone  = safeStr_(filters.phone).toLowerCase();

  const filtered = list.filter(it=>{
    const ms = ddmmyyyyToMs_(it.date);
    if (fromMs && ms < fromMs) return false;
    if (toMs   && ms > toMs)   return false;
    if (fStore  && it.store !== fStore) return false;
    if (fStaff  && it.staff !== fStaff) return false;
    if (fStat   && it.preorder_statuses !== fStat) return false;
    if (fModel  && it.model_name !== fModel) return false;
    if (fMemory && it.memory !== fMemory) return false;
    if (fColor  && it.color !== fColor) return false;
    if (fCust   && !safeStr_(it.customer).toLowerCase().includes(fCust)) return false;
    if (fPhone  && !safeStr_(it.phone).toLowerCase().includes(fPhone)) return false;
    return true;
  });

  filtered.sort((a,b) => ddmmyyyyToMs_(b.date) - ddmmyyyyToMs_(a.date));

  const totals = {
    count: filtered.length,
    prepay: round2_(filtered.reduce((s,x)=>s + toNumberSafe_(x.prepay||0), 0)),
    zarplata: round2_(filtered.reduce((s,x)=>s + toNumberSafe_(x.zarplata||0), 0))
  };
  return { rows: filtered, totals };
}

/* ======================================================================= */
/* ============================ ДИАГНОСТИКА ============================== */
/* ======================================================================= */

/** Заголовки листа "Диагностика" (строгаем по row1) */
function getDiagnosticsHeaders_(ss){
  const sh = ss.getSheetByName('Диагностика');
  if (!sh) throw new Error('Лист "Диагностика" не найден');
  const headers = (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || []).map(safeStr_);
  const must = ['id','intake_date','store','staff','diagnostic_statuses','diag_pay','payments','issued_date','issued_staff'];
  const miss = must.filter(k => headers.indexOf(k) < 0);
  if (miss.length) throw new Error('В "Диагностика" отсутствуют колонки: ' + miss.join(', '));
  return { sheet: sh, headers };
}

/** Следующий ID: DIAG-YYYYMM-0001 */
function nextDiagId_(ss, tz){
  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const idIdx = headers.indexOf('id');
  const ym = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  const values = sh.getDataRange().getValues();
  let maxN = 0;
  for (let r = 2; r < values.length; r++){
    const v = safeStr_(values[r][idIdx]);
    const m = v.match(/^DIAG-(\d{6})-(\d{4})$/);
    if (m && m[1] === ym){
      const n = parseInt(m[2], 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
  }
  return `DIAG-${ym}-${String(maxN+1).padStart(4,'0')}`;
}

/** Универсальное добавление строки в лист "Диагностика" по row1 */
function appendDiagnostic_(ss, rowObj){
  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const line = headers.map(h => {
    if (h === 'intake_date' || h === 'purchase_date') return rowObj[h];
    if (h === 'diag_pay') {
      const n = Number(rowObj[h]); return isNaN(n) ? rowObj[h] : n;
    }
    return rowObj[h];
  });
  sh.appendRow(line);
}

/** Найти индекс ПЕРВОЙ строки (1-based) по id — "живая" строка */
function findFirstDiagRowIndex_(ss, id){
  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const idIdx = headers.indexOf('id');
  const values = sh.getDataRange().getValues();
  for (let r = 2; r < values.length; r++){
    if (safeStr_(values[r][idIdx]) === id){
      return r + 1; // 1-based
    }
  }
  return -1;
}

/** In-place обновление ПЕРВОЙ строки заявки (не создаёт новых строк) */
function setDiagnosticInPlace_(ss, id, patch){
  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const rowIdx = findFirstDiagRowIndex_(ss, id);
  if (rowIdx < 0) return { ok:false, reason:'not_found' };

  const H = {}; headers.forEach((h,i)=> H[h]=i+1); // 1-based cols
  const up = (k, v) => { if (H[k]) sh.getRange(rowIdx, H[k]).setValue(v); };

  Object.keys(patch || {}).forEach(k => {
    if (headers.indexOf(k) >= 0) {
      up(k, patch[k]);
    }
  });
  return { ok:true, row: rowIdx };
}

/** Дефолтный статус для диагностики */
function getDefaultDiagnosticStatus_(ss){
  const dict = buildDictsColumnar_(readSheetObjectsStrict_(ss, 'Справочники'));
  const arr = dict.diagnostic_statuses || [];
  const pref = ['Принят','В работе','В ремонте'];
  const hit = arr.map(x=>safeStr_(x.name)).find(n => pref.indexOf(n) >= 0);
  return hit || safeStr_(arr[0]?.name) || 'Принят';
}

/** Агрегация по одному id (две строки: base + "Выдан") */
function aggregateDiagnosticById_(ss, id, opts){
  opts = opts || {};
  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const H = {}; headers.forEach((h,i)=>H[h]=i);

  let values = opts.values;
  if (!values){
    values = sh.getDataRange().getValues();
  }

  const groupRows = Array.isArray(opts.groupRows) ? opts.groupRows : null;
  const rows = [];

  if (groupRows && groupRows.length){
    groupRows.forEach((item, idx)=>{
      if (!item) return;
      let rowArr = null;
      let rowIdx = null;
      if (Array.isArray(item)){
        rowArr = item;
        rowIdx = idx + 1; // относительный порядок
      } else if (Array.isArray(item.row)){
        rowArr = item.row;
        if (item.idx != null) rowIdx = item.idx;
        else if (item.rowIndex != null) rowIdx = item.rowIndex;
        else if (item.position != null) rowIdx = item.position;
        else rowIdx = idx + 1;
      }
      if (!rowArr) return;
      if (safeStr_(rowArr[H['id']]) !== id) return;
      const obj = {};
      headers.forEach((h, i) => obj[h] = rowArr[i]);
      obj._rowIndex = rowIdx != null ? rowIdx : idx + 1;
      rows.push(obj);
    });
  }

  if (!rows.length){
    for (let r = 2; r < values.length; r++){
      const row = values[r];
      if (safeStr_(row[H['id']]) !== id) continue;
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj._rowIndex = r + 1;
      rows.push(obj);
    }
  }

  if (!rows.length) return {};

  rows.sort((a,b)=> a._rowIndex - b._rowIndex);
  const first = rows[0];
  const tz = getTz_();
  const normalizeDate = s => normalizeDateKey_(s, tz);

  const findLastFilled = key => {
    for (let i = rows.length - 1; i >= 0; i--){
      const val = rows[i][key];
      if (val == null) continue;
      if (typeof val === 'string' && val.trim() === '') continue;
      return { value: val, row: rows[i] };
    }
    return null;
  };

  const diagPayInfo = findLastFilled('diag_pay');
  const diagPay = diagPayInfo ? round2_(toNumberSafe_(String(diagPayInfo.value).replace(/\s+/g,''))) : 0;
  const paymentsInfo = findLastFilled('payments');
  const payments = paymentsInfo ? safeStr_(paymentsInfo.value) : '';

  const statusInfo = findLastFilled('diagnostic_statuses');
  const status = statusInfo ? safeStr_(statusInfo.value) : (safeStr_(first['diagnostic_statuses']) || 'Принят');

  const issuedDateInfo = findLastFilled('issued_date');
  const issuedDate = issuedDateInfo ? normalizeDate(issuedDateInfo.value) : '';
  const issuedStaffInfo = findLastFilled('issued_staff');
  const issuedStaff = issuedStaffInfo ? safeStr_(issuedStaffInfo.value) : '';
  const fallbackIssuedRow = rows.slice().reverse().find(r => safeStr_(r['diagnostic_statuses']).toLowerCase() === 'выдан');
  const completedAt = issuedDate || (fallbackIssuedRow ? normalizeDate(fallbackIssuedRow['intake_date']) : '');
  const completedBy = issuedStaff || (fallbackIssuedRow ? safeStr_(fallbackIssuedRow['staff']) : '');

  const staffColorMap = opts.staffColorMap || getStaffColorMap_(ss);
  const staffColor = staffColorMap[safeStr_(first['staff'])] || '';
  const phone = safeStr_(first['phone_klienta']);

  return {
    id: safeStr_(first['id']),
    date: normalizeDate(first['intake_date']),
    purchase_date: normalizeDate(first['purchase_date']),
    store: safeStr_(first['store']),
    staff: safeStr_(first['staff']),
    staff_color: staffColor,

    model_name: safeStr_(first['model_name']),
    memory: safeStr_(first['memory']),
    color: safeStr_(first['color']),
    imei: safeStr_(first['imei']),
    complect: safeStr_(first['complect']),
    neispravnost: safeStr_(first['neispravnost']),
    appearance: safeStr_(first['appearance']),

    customer: safeStr_(first['customer']),
    phone: phone,
    note: safeStr_(first['note']),

    diag_pay: diagPay,
    payments: payments,

    diagnostic_statuses: status,
    issued_date: issuedDate || completedAt,
    issued_staff: issuedStaff || completedBy,
    completed_at: completedAt,
    completed_by: completedBy
  };
}

/* Список диагностики (агрегировано по id) — старый серверный способ */
function getDiagnostics(filters){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz_();
  filters = filters || {};

  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const H = {}; headers.forEach((h,i)=>H[h]=i);
  const values = sh.getDataRange().getValues();

  const groups = {};
  for (let r = 2; r < values.length; r++){
    const row = values[r];
    const id = safeStr_(row[H['id']]);
    if (!id) continue;
    (groups[id] || (groups[id] = [])).push({ row, idx: r+1 });
  }

  const staffColorMap = getStaffColorMap_(ss);

  const list = Object.keys(groups).map(id => aggregateDiagnosticById_(ss, id, {
    values,
    groupRows: groups[id],
    staffColorMap
  }));

  const f = {
    dateFrom: safeStr_(filters.dateFrom),
    dateTo:   safeStr_(filters.dateTo),
    store:    safeStr_(filters.store),
    staff:    safeStr_(filters.staff),
    status:   safeStr_(filters.diagnostic_statuses),
    model:    safeStr_(filters.model_name),
    memory:   safeStr_(filters.memory),
    color:    safeStr_(filters.color),
    imei:     safeStr_(filters.imei),
    customer: safeStr_(filters.customer).toLowerCase(),
    phone:    safeStr_(filters.phone).toLowerCase()
  };
  const fromMs = f.dateFrom ? new Date(f.dateFrom + 'T00:00:00').getTime() : null;
  const toMs   = f.dateTo   ? new Date(f.dateTo   + 'T23:59:59').getTime() : null;

  const filtered = list.filter(x=>{
    const ms = ddmmyyyyToMs_(x.date);
    if (fromMs && ms < fromMs) return false;
    if (toMs   && ms > toMs)   return false;
    if (f.store  && x.store !== f.store) return false;
    if (f.staff  && x.staff !== f.staff) return false;
    if (f.status && x.diagnostic_statuses !== f.status) return false;
    if (f.model  && x.model_name !== f.model) return false;
    if (f.memory && x.memory !== f.memory) return false;
    if (f.color  && x.color !== f.color) return false;
    if (f.imei   && !safeStr_(x.imei).includes(f.imei)) return false;
    if (f.customer && !safeStr_(x.customer).toLowerCase().includes(f.customer)) return false;
    if (f.phone && !safeStr_(x.phone).toLowerCase().includes(f.phone)) return false;
    return true;
  });

  filtered.sort((a,b) => ddmmyyyyToMs_(b.date) - ddmmyyyyToMs_(a.date));

  const totals = {
    count: filtered.length,
    diag_pay_sum: round2_(filtered.reduce((s,x)=> s + toNumberSafe_(x.diag_pay||0), 0))
  };
  return { rows: filtered, totals };
}

/** Создание заявки (статус "Принят"), одна живая строка */
function createDiagnostic(doc){
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();
    if (!doc) throw new Error('Пустой документ диагностики');

    const id = nextDiagId_(ss, tz);
    const intake = toDdMmYyyy_(safeStr_(doc.intake_date || doc.date), tz);
    const purchase = toDdMmYyyy_(safeStr_(doc.purchase_date), tz);
    const store  = safeStr_(doc.store);
    const staff  = safeStr_(doc.staff);
    if (!intake) throw new Error('Укажите дату приёма');
    if (!store)  throw new Error('Выберите магазин');
    if (!staff)  throw new Error('Выберите сотрудника');

    const status = getDefaultDiagnosticStatus_(ss); // «Принят»
    const complectValue = Array.isArray(doc.complect)
      ? doc.complect.map(x=>safeStr_(x)).filter(Boolean).join(', ')
      : safeStr_(doc.complect);

    appendDiagnostic_(ss, {
      id,
      purchase_date: purchase,
      intake_date: intake,
      store,
      staff,
      diagnostic_statuses: status,
      model_name: safeStr_(doc.model_name),
      memory: safeStr_(doc.memory),
      color: safeStr_(doc.color),
      imei: safeStr_(doc.imei),
      complect: complectValue,
      neispravnost: safeStr_(doc.neispravnost),
      appearance: safeStr_(doc.appearance),
      customer: safeStr_(doc.customer),
      phone_klienta: safeStr_(doc.phone),
      note: safeStr_(doc.note)
    });

    cacheDel_(['DIAG_BOOT_V2']);
    return { ok:true, id };
  } finally { lock.releaseLock(); }
}

/** In-place статусы: "Принят", "В ремонте", "Готов", "Выдан" */
function updateDiagnosticStatus(doc){
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    const id = safeStr_(doc.id);
    const newSt = safeStr_(doc.diagnostic_statuses);
    if (!id) throw new Error('Нет id диагностики');
    if (!newSt) throw new Error('Не указан статус');

    const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
    const rowIdx = findFirstDiagRowIndex_(ss, id);
    if (rowIdx < 0) throw new Error('Живая строка диагностики не найдена');

    const H = {}; headers.forEach((h,i)=>H[h]=i+1);
    const rowValues = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0] || [];
    const currentStatus = safeStr_(rowValues[H['diagnostic_statuses']-1]);
    const currentIssued = currentStatus.toLowerCase() === 'выдан';
    const nextIssued = newSt.toLowerCase() === 'выдан';

    if (currentIssued && !nextIssued){
      throw new Error('Статус "Выдан" изменить нельзя');
    }

    const patch = { diagnostic_statuses: newSt };

    if (safeStr_(doc.note)) patch.note = safeStr_(doc.note);
    if (safeStr_(doc.purchase_date)) patch.purchase_date = toDdMmYyyy_(safeStr_(doc.purchase_date), tz);

    const dateStr = safeStr_(doc.intake_date || doc.date);
    if (dateStr && !nextIssued) patch.intake_date = toDdMmYyyy_(dateStr, tz);

    if (nextIssued){
      const issuedDateRaw = safeStr_(doc.issued_date || doc.date || doc.intake_date);
      if (!issuedDateRaw) throw new Error('Укажите дату выдачи');
      patch.issued_date = toDdMmYyyy_(issuedDateRaw, tz);

      const issuedStaff = safeStr_(doc.issued_staff || doc.staff);
      if (!issuedStaff) throw new Error('Укажите сотрудника выдачи');
      patch.issued_staff = issuedStaff;
    } else {
      if (safeStr_(doc.staff)) patch.staff = safeStr_(doc.staff);
      if (safeStr_(doc.issued_date)) patch.issued_date = toDdMmYyyy_(safeStr_(doc.issued_date), tz);
      if (safeStr_(doc.issued_staff)) patch.issued_staff = safeStr_(doc.issued_staff);
    }

    Object.keys(patch).forEach(key => {
      if (H[key]) sh.getRange(rowIdx, H[key]).setValue(patch[key]);
    });

    cacheDel_(['DIAG_BOOT_V2']);
    return { ok:true };
  } finally { lock.releaseLock(); }
}

function updateDiagnosticPayment(doc){
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try{
    const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const tz = getTz_();

    const id = safeStr_(doc.id);
    if (!id) throw new Error('Нет id диагностики');

    const paymentsRaw = doc.payments;
    const paymentsArr = Array.isArray(paymentsRaw) ? paymentsRaw : null;
    const paymentsStr = paymentsArr ? serializePayments_(paymentsArr) : safeStr_(paymentsRaw);

    let diagPayValue = safeStr_(doc.diag_pay);
    if (!diagPayValue && paymentsArr && paymentsArr.length){
      const sum = sumPayments_(paymentsArr);
      if (sum || sum === 0) diagPayValue = String(sum);
    }
    let diagPayPatchValue = '';
    if (diagPayValue !== ''){
      const normalizedPay = diagPayValue.replace(/\s+/g,'').replace(',', '.');
      const diagPayNum = Number(normalizedPay);
      diagPayPatchValue = isNaN(diagPayNum) ? diagPayValue : diagPayNum;
    }
    const patch = {
      diag_pay: diagPayPatchValue,
      payments: paymentsStr
    };

    const dateStr = safeStr_(doc.intake_date || doc.date);
    if (dateStr) patch.intake_date = toDdMmYyyy_(dateStr, tz);
    if (safeStr_(doc.staff)) patch.staff = safeStr_(doc.staff);

    const res = setDiagnosticInPlace_(ss, id, patch);
    if (!res.ok) throw new Error('Живая строка диагностики не найдена');

    cacheDel_(['DIAG_BOOT_V2']);
    const aggregate = aggregateDiagnosticById_(ss, id);
    return { ok:true, aggregate };
  } finally { lock.releaseLock(); }
}

/* ======================= БЫСТРЫЕ ПРЕЛОАДЫ (кэшируемые) ======================= */

/* Диагностика: быстрый агрегат одной проходкой */
function computeDiagnosticsAggregate_(ss){
  const tz = getTz_();
  const { sheet: sh, headers } = getDiagnosticsHeaders_(ss);
  const H={}; headers.forEach((h,i)=>H[h]=i);
  const values = sh.getDataRange().getValues();

  const staffColorMap = getStaffColorMap_(ss);

  const groups = {};
  for (let r=2;r<values.length;r++){
    const row = values[r];
    const id  = safeStr_(row[H['id']]);
    if (!id) continue;
    (groups[id] || (groups[id]=[])).push({ row, idx:r+1 });
  }

  const list = [];
  Object.keys(groups).forEach(id=>{
    const arr = groups[id].slice().sort((a,b)=>a.idx-b.idx).map(x=>x.row);
    const first = arr[0];
    const intake   = normalizeDateKey_(first[H['intake_date']], tz);
    const purchRaw = H['purchase_date']!=null ? first[H['purchase_date']] : '';
    const purchase = normalizeDateKey_(purchRaw, tz);

    const findLastFilled = key => {
      const col = H[key];
      if (col == null) return null;
      for (let i=arr.length-1; i>=0; i--){
        const val = arr[i][col];
        if (val == null) continue;
        if (typeof val === 'string' && val.trim() === '') continue;
        return { value: val, row: arr[i] };
      }
      return null;
    };

    const diagPayInfo = findLastFilled('diag_pay');
    const diagPay = diagPayInfo ? round2_(toNumberSafe_(String(diagPayInfo.value).replace(/\s+/g,''))) : 0;
    const paymentsInfo = findLastFilled('payments');
    const payments = paymentsInfo ? safeStr_(paymentsInfo.value) : '';

    const statusInfo = findLastFilled('diagnostic_statuses');
    const status = statusInfo ? safeStr_(statusInfo.value) : (safeStr_(first[H['diagnostic_statuses']]) || 'Принят');

    const issuedDateInfo = findLastFilled('issued_date');
    const issuedDate = issuedDateInfo ? normalizeDateKey_(issuedDateInfo.value, tz) : '';
    const issuedStaffInfo = findLastFilled('issued_staff');
    const issuedStaff = issuedStaffInfo ? safeStr_(issuedStaffInfo.value) : '';
    const issuedRow = arr.slice().reverse().find(r=>safeStr_(r[H['diagnostic_statuses']]).toLowerCase()==='выдан');
    const completedAt = issuedDate || (issuedRow ? normalizeDateKey_(issuedRow[H['intake_date']], tz) : '');
    const completedBy = issuedStaff || (issuedRow ? safeStr_(issuedRow[H['staff']]) : '');

    list.push({
      id: safeStr_(first[H['id']]),
      date: intake,
      purchase_date: purchase,
      store: safeStr_(first[H['store']]),
      staff: safeStr_(first[H['staff']]),
      staff_color: staffColorMap[safeStr_(first[H['staff']])] || '',

      model_name: safeStr_(first[H['model_name']]),
      memory: safeStr_(first[H['memory']]),
      color: safeStr_(first[H['color']]),
      imei: safeStr_(first[H['imei']]),
      complect: safeStr_(first[H['complect']]),
      neispravnost: safeStr_(first[H['neispravnost']]),
      appearance: safeStr_(first[H['appearance']]),

      customer: safeStr_(first[H['customer']]),
      phone: safeStr_(first[H['phone_klienta']]),
      note: safeStr_(first[H['note']]),

      diag_pay: diagPay,
      payments: payments,

      diagnostic_statuses: status,
      issued_date: issuedDate || completedAt,
      issued_staff: issuedStaff || completedBy,
      completed_at: completedAt,
      completed_by: completedBy
    });
  });

  list.sort((a,b)=> ddmmyyyyToMs_(b.date) - ddmmyyyyToMs_(a.date));

  const totals = {
    count: list.length,
    diag_pay_sum: round2_(list.reduce((s,x)=>s+toNumberSafe_(x.diag_pay||0),0))
  };

  return { rows: list, totals, updatedAt: Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm') };
}

function getDiagnosticsBootstrap(){
  const ss  = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const key = 'DIAG_BOOT_V2';
  let data  = cacheGetJson_(key);
  if (data) return data;
  data = computeDiagnosticsAggregate_(ss);
  cachePutJson_(key, data, 600); // 10 минут
  return data;
}

/* Предзаказы: быстрый агрегат одной проходкой */
function computePreordersAggregate_(ss){
  const tz = getTz_();
  const { sheet: sh, headers } = getPreordersHeaders_(ss);
  const H={}; headers.forEach((h,i)=>H[h]=i);
  const values = sh.getDataRange().getValues();
  const staffColorMap = getStaffColorMap_(ss);

  const groups={};
  for (let r=2;r<values.length;r++){
    const row = values[r];
    const id  = safeStr_(row[H['id']]);
    if (!id) continue;
    (groups[id]||(groups[id]=[])).push({row, idx:r+1});
  }

  const out=[];
  Object.keys(groups).forEach(id=>{
    const arr=groups[id].slice().sort((a,b)=>a.idx-b.idx).map(x=>x.row);
    const first=arr[0], last=arr[arr.length-1];

    const prepaySum = arr.reduce((s,r)=> s+toNumberSafe_(r[H['prepay']]||0), 0);
    const paymentsExtra = arr.reduce((s,r)=>{
      const hasPre = toNumberSafe_(r[H['prepay']])>0;
      return s + (hasPre ? 0 : sumPaymentsFromString_(safeStr_(r[H['payments']])));
    },0);
    const totalPrepay = round2_(prepaySum + paymentsExtra);

    const prePriceRaw = safeStr_(first[H['pre_price']]) || safeStr_(last[H['pre_price']]);
    const prePriceNum = toNumberSafe_(prePriceRaw);
    const preBalance  = round2_(prePriceNum - totalPrepay);

    let preImei = '';
    for (let i=0;i<arr.length;i++){ const v=safeStr_(arr[i][H['pre_imei']]); if (v){ preImei=v; break; } }

    const completedRow = arr.filter(r=>{
      const st = safeStr_(r[H['preorder_statuses']]).toLowerCase();
      return st==='завершен' || st==='завершён';
    }).pop();
    const completedAt = completedRow ? normalizeDateKey_(completedRow[H['date']], tz) : '';
    const completedBy = completedRow ? safeStr_(completedRow[H['staff']]) : '';

    const creator = safeStr_(first[H['staff']]);

    out.push({
      id,
      date: normalizeDateKey_(first[H['date']], tz),
      store: safeStr_(last[H['store']]) || safeStr_(first[H['store']]),
      staff: creator,
      staff_color: staffColorMap[creator] || '',
      model_name: safeStr_(first[H['model_name']]) || safeStr_(last[H['model_name']]),
      memory: safeStr_(first[H['memory']]) || safeStr_(last[H['memory']]),
      color: safeStr_(first[H['color']]) || safeStr_(last[H['color']]),
      pre_imei: preImei,
      pre_price: prePriceRaw,
      prepay: totalPrepay,
      pre_balance: preBalance,
      payments: safeStr_(last[H['payments']]),
      customer: safeStr_(last[H['customer']]) || safeStr_(first[H['customer']]),
      phone: safeStr_(last[H['phone']]) || safeStr_(first[H['phone']]),
      zarplata: round2_(arr.reduce((s,r)=>s+toNumberSafe_(r[H['zarplata']]||0),0)),
      note: safeStr_(last[H['note']]),
      preorder_statuses: safeStr_(last[H['preorder_statuses']]) || 'Ожидание',
      completed_at: completedAt, completed_by: completedBy
    });
  });

  out.sort((a,b)=> ddmmyyyyToMs_(b.date) - ddmmyyyyToMs_(a.date));

  const totals = {
    count: out.length,
    prepay: round2_(out.reduce((s,x)=>s+toNumberSafe_(x.prepay||0),0)),
    zarplata: round2_(out.reduce((s,x)=>s+toNumberSafe_(x.zarplata||0),0))
  };
  return { rows: out, totals, updatedAt: Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm') };
}

function getPreordersBootstrap(){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const key='PRE_BOOT_V2';
  let data = cacheGetJson_(key);
  if (data) return data;
  data = computePreordersAggregate_(ss);
  cachePutJson_(key, data, 600);
  return data;
}