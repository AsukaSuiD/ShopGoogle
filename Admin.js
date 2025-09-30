/** @OnlyCurrentDoc *
 * Сервер «Админ»: PIN-авторизация + операции со складом (телефоны).
 * Требует утилиты из Code.gs: SHEET_ID, getTz_(), cachePutJson_/cacheGetJson_/cacheDel_,
 * readSheetObjectsStrict_, safeStr_, uniqNotEmpty_.
 */

/* ===================== ПИН-КОД / АВТОРИЗАЦИЯ ===================== */
const ADMIN_PIN_HASH_PROP = 'ADMIN_PIN_SHA256';
const ADMIN_PIN_SALT_PROP = 'ADMIN_PIN_SALT';
const ADMIN_TOKEN_TTL_SEC = 60 * 60; // 1 час

function sha256Hex_(s){
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/** Одноразовая установка/смена PIN. Вызывается из IDE вручную: adminSetPin('1234') */
function adminSetPin(newPin){
  newPin = safeStr_(newPin);
  if (!newPin) throw new Error('Введите новый PIN');
  const props = PropertiesService.getScriptProperties();
  const salt = Utilities.getUuid().replace(/-/g,'');
  const hash = sha256Hex_(salt + '::' + newPin);
  props.setProperty(ADMIN_PIN_SALT_PROP, salt);
  props.setProperty(ADMIN_PIN_HASH_PROP, hash);
  return { ok:true };
}

/** Есть ли настроенный PIN */
function adminIsPinSet(){
  const props = PropertiesService.getScriptProperties();
  return !!props.getProperty(ADMIN_PIN_HASH_PROP);
}

/** Проверка PIN */
function adminCheckPin_(pin){
  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty(ADMIN_PIN_SALT_PROP) || '';
  const hash = props.getProperty(ADMIN_PIN_HASH_PROP) || '';
  if (!salt || !hash) return false;
  const hex = sha256Hex_(salt + '::' + safeStr_(pin));
  return hex === hash;
}

/** Логин по PIN -> выдать токен (в кэше) */
function adminLogin(pin){
  if (!adminIsPinSet()) throw new Error('PIN не настроен. Сначала вызовите adminSetPin(pin) в IDE.');
  if (!adminCheckPin_(pin)) throw new Error('Неверный PIN');

  const token = Utilities.getUuid();
  cachePutJson_('ADMIN_TOKEN::' + token, { ok:true }, ADMIN_TOKEN_TTL_SEC);
  return { ok:true, token, ttl: ADMIN_TOKEN_TTL_SEC };
}

/** Выход */
function adminLogout(token){
  cacheDel_(['ADMIN_TOKEN::' + safeStr_(token)]);
  return { ok:true };
}

/** Валидация токена */
function adminValidateToken(token){
  return !!cacheGetJson_('ADMIN_TOKEN::' + safeStr_(token));
}

/** Состояние для UI (опционально принимает token) */
function adminGetAuthState(token){
  return {
    isPinSet: adminIsPinSet(),
    authed: token ? adminValidateToken(token) : false,
    ttl: ADMIN_TOKEN_TTL_SEC
  };
}

/* ===================== ВСПОМОГАТЕЛЬНОЕ: СКЛАД/КАТАЛОГ/СПРАВОЧНИКИ ===================== */
function getStockHeaders_(ss){
  const sh = ss.getSheetByName('Склад');
  if (!sh) throw new Error('Лист "Склад" не найден');
  const headers = (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || []).map(safeStr_);
  return { sheet: sh, headers };
}

/** Следующий ID формата STK-YYYYMM-0001 (если в "Склад" есть колонка id) */
function nextStockId_(ss){
  const { sheet: sh, headers } = getStockHeaders_(ss);
  const idIdx = headers.indexOf('id');
  if (idIdx < 0) return ''; // нет колонки id — оставляем пусто
  const tz = (typeof getTz_ === 'function') ? getTz_() : 'Europe/Moscow';
  const ym = Utilities.formatDate(new Date(), tz, 'yyyyMM');

  const vals = sh.getDataRange().getValues();
  let maxN = 0;
  for (let r = 2; r < vals.length; r++){
    const v = safeStr_(vals[r][idIdx]);
    const m = v.match(/^STK-(\d{6})-(\d{4})$/);
    if (m && m[1] === ym){
      const n = parseInt(m[2], 10);
      if (!isNaN(n) && n > maxN) maxN = n;
    }
  }
  return `STK-${ym}-${String(maxN+1).padStart(4,'0')}`;
}

/** Добавление строки по ключам row1 листа "Склад" */
function appendStock_(ss, rowObj){
  const { sheet: sh, headers } = getStockHeaders_(ss);
  const line = headers.map(h => {
    let val = rowObj[h];
    if (h === 'sale_price') {
      const n = Number(val); val = isNaN(n) ? val : n;
    }
    return val;
  });
  sh.appendRow(line);
}

/** Обновление первой строки по IMEI (только указанные поля) */
function patchStockByImei_(ss, imei, patch){
  const { sheet: sh, headers } = getStockHeaders_(ss);
  const values = sh.getDataRange().getValues();
  const H = {}; headers.forEach((h,i)=> H[h]=i);
  const idxImei = H['imei'];
  if (idxImei == null) throw new Error('В "Склад" отсутствует колонка imei');

  const target = safeStr_(imei).replace(/\s+/g,'');
  for (let r = 2; r < values.length; r++){
    const v = safeStr_(values[r][idxImei]).replace(/\s+/g,'');
    if (!v) continue;
    if (v === target){
      Object.keys(patch || {}).forEach(k=>{
        if (H[k] == null) return;
        const col = H[k] + 1; // 1-based
        let val = patch[k];
        if (k === 'sale_price') {
          const n = Number(val); val = isNaN(n) ? val : n;
        }
        sh.getRange(r + 1, col).setValue(val);
      });
      return { ok:true, row: r+1 };
    }
  }
  return { ok:false, reason: 'not_found' };
}

/** Каталог для UI (строки из "Каталог моделей") с фолбэком на pre_price */
function adminGetCatalog(token){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Каталог моделей');
  if (!sh) throw new Error('Лист "Каталог моделей" не найден');

  const vals = sh.getDataRange().getValues();
  const head = (vals[0] || []).map(safeStr_);
  const H = {}; head.forEach((h,i)=> H[h]=i);

  const rows = [];
  for (let r=1; r<vals.length; r++){
    const v = vals[r];
    const model = safeStr_(v[H['model_name']]);
    if (!model) continue;
    const mem   = safeStr_(v[H['memory']]);
    const color = safeStr_(v[H['color']]);

    const sale = H['sale_price']!=null ? v[H['sale_price']] : (H['pre_price']!=null ? v[H['pre_price']] : '');
    rows.push({
      model_name: model,
      memory: mem,
      color: color,
      sale_price: sale,
      pre_price: H['pre_price']!=null ? v[H['pre_price']] : ''
    });
  }
  return { ok:true, rows };
}

/** Поиск цены в каталоге по комбинации (sale_price || pre_price) */
function adminLookupCatalogPrice(token, model, memory, color){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Каталог моделей');
  if (!sh) throw new Error('Лист "Каталог моделей" не найден');

  const vals = sh.getDataRange().getValues();
  const head = (vals[0] || []).map(safeStr_);
  const H = {}; head.forEach((h,i)=> H[h]=i);

  const m = safeStr_(model), mem = safeStr_(memory), c = safeStr_(color);
  for (let r=1; r<vals.length; r++){
    const v = vals[r];
    if (safeStr_(v[H['model_name']])===m &&
        safeStr_(v[H['memory']])===mem &&
        safeStr_(v[H['color']])===c){
      const sale = H['sale_price']!=null ? v[H['sale_price']] : (H['pre_price']!=null ? v[H['pre_price']] : '');
      return { ok:true, sale_price: sale };
    }
  }
  return { ok:false };
}

/** Справочники для Админки: города, состояния, статусы склада — напрямую из листа "Справочники" */
function adminGetCoreDicts(token){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const rows = readSheetObjectsStrict_(ss, 'Справочники') || [];

  const list = (col) => uniqNotEmpty_((rows || []).map(r => safeStr_(r[col]))).map(v => ({ code: v, name: v }));
  return {
    ok: true,
    cities:         list('city'),
    conditions:     list('condition'),
    stock_statuses: list('stock_statuses')
  };
}

/* ===================== API: СКЛАД для Админа ===================== */

/** Добавить позицию склада (телефон) */
function adminStockAdd(token, doc){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

  const imeiRaw = safeStr_(doc.imei);
  const imei = imeiRaw.replace(/\s+/g,'');

  const row = {
    id: nextStockId_(ss), // может остаться пустым, если нет колонки id
    city: safeStr_(doc.city),
    condition: safeStr_(doc.condition),
    model_name: safeStr_(doc.model_name),
    memory: safeStr_(doc.memory),
    color: safeStr_(doc.color),
    imei: imei,
    sale_price: safeStr_(doc.sale_price),
    stock_statuses: safeStr_(doc.stock_statuses) || 'В наличии',
    note: safeStr_(doc.note)
  };

  if (!row.imei) throw new Error('IMEI обязателен');

  const rows = readSheetObjectsStrict_(ss, 'Склад') || [];
  if (rows.some(r => safeStr_(r['imei']).replace(/\s+/g,'') === row.imei)){
    throw new Error('Такой IMEI уже есть на складе');
  }

  if (!row.sale_price && row.model_name && row.memory && row.color){
    try{
      const p = adminLookupCatalogPrice(token, row.model_name, row.memory, row.color);
      if (p && p.ok && p.sale_price != null) row.sale_price = p.sale_price;
    }catch(e){ /* no-op */ }
  }

  ['city','condition','model_name','memory','color'].forEach(k=>{
    if (!row[k]) throw new Error('Заполните поле: ' + k);
  });

  appendStock_(ss, row);
  invalidateNalichieCache_();
  return { ok:true, imei: row.imei, id: row.id };
}

/** Батч-добавление позиций склада
 * docs: [{ city,condition,model_name,memory,color,imei,sale_price,stock_statuses,note }, ...]
 */
function adminStockAddMany(token, docs){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  if (!Array.isArray(docs) || !docs.length) throw new Error('Список пуст');

  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

  const catalogPriceMap = (()=>{
    const map = new Map();
    try{
      const shCatalog = ss.getSheetByName('Каталог моделей');
      if (!shCatalog) return map;

      const vals = shCatalog.getDataRange().getValues();
      if (!vals.length) return map;

      const head = (vals[0] || []).map(safeStr_);
      const H = {}; head.forEach((h,i)=>H[h]=i);

      const idxModel = H['model_name'];
      const idxMemory = H['memory'];
      const idxColor = H['color'];
      if (idxModel == null || idxMemory == null || idxColor == null) return map;

      for (let r=1; r<vals.length; r++){
        const v = vals[r];
        const model = safeStr_(v[idxModel]);
        if (!model) continue;
        const memory = safeStr_(v[idxMemory]);
        const color = safeStr_(v[idxColor]);
        const price = (H['sale_price']!=null)
          ? v[H['sale_price']]
          : (H['pre_price']!=null ? v[H['pre_price']] : '');
        const key = [model, memory, color].join('|');
        map.set(key, price);
      }
    }catch(e){ /* no-op */ }
    return map;
  })();

  const { sheet: sh, headers } = getStockHeaders_(ss);

  const vals = sh.getDataRange().getValues();
  const H = {}; headers.forEach((h,i)=>H[h]=i);
  const idxImei = H['imei'];
  if (idxImei == null) throw new Error('В "Склад" отсутствует колонка imei');

  const existing = new Set();
  for (let r=2; r<vals.length; r++){
    const v = safeStr_(vals[r][idxImei]).replace(/\s+/g,'');
    if (v) existing.add(v);
  }

  const idxId = H['id'];
  const tz = (typeof getTz_ === 'function') ? getTz_() : 'Europe/Moscow';
  const ym = Utilities.formatDate(new Date(), tz, 'yyyyMM');
  let maxN = 0;
  if (idxId != null){
    for (let r=2; r<vals.length; r++){
      const s = safeStr_(vals[r][idxId]);
      const m = s.match(/^STK-(\d{6})-(\d{4})$/);
      if (m && m[1] === ym){
        const n = parseInt(m[2], 10);
        if (!isNaN(n) && n > maxN) maxN = n;
      }
    }
  }

  const results = {
    ok: true,
    total_requested: docs.length,
    total_added: 0,
    added: [],                  // [{ imei, id }]
    skipped_duplicates: [],     // [imei,...]
    invalid: []                 // [{ imei, error }]
  };

  const seenInput = new Set();
  const rowsToWrite = [];

  for (let i=0; i<docs.length; i++){
    const d = docs[i] || {};
    const imei = safeStr_(d.imei).replace(/\s+/g,'');
    if (!imei){ results.invalid.push({ imei:'', error:'IMEI пуст' }); continue; }

    if (seenInput.has(imei)){ results.invalid.push({ imei, error:'Дубль в списке' }); continue; }
    seenInput.add(imei);

    if (existing.has(imei)){ results.skipped_duplicates.push(imei); continue; }

    const row = {
      id: '',
      city:       safeStr_(d.city),
      condition:  safeStr_(d.condition),
      model_name: safeStr_(d.model_name),
      memory:     safeStr_(d.memory),
      color:      safeStr_(d.color),
      imei:       imei,
      sale_price: safeStr_(d.sale_price),
      stock_statuses: safeStr_(d.stock_statuses) || 'В наличии',
      note:       safeStr_(d.note)
    };

    for (const k of ['city','condition','model_name','memory','color']){
      if (!row[k]){ results.invalid.push({ imei, error:`Не заполнено поле: ${k}` }); row._skip = true; break; }
    }
    if (row._skip) continue;

    if (!row.sale_price){
      const key = [row.model_name, row.memory, row.color].join('|');
      if (catalogPriceMap.has(key)){
        const price = catalogPriceMap.get(key);
        if (price != null) row.sale_price = price;
      }
    }

    if (idxId != null){
      maxN += 1;
      row.id = `STK-${ym}-${String(maxN).padStart(4,'0')}`;
    }

    rowsToWrite.push(row);
  }

  if (!rowsToWrite.length){
    return results;
  }

  const lines = rowsToWrite.map(r=>{
    return headers.map(h=>{
      if (h === 'sale_price'){
        const n = Number(r[h]); return isNaN(n) ? r[h] : n;
      }
      return r[h];
    });
  });

  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, lines.length, headers.length).setValues(lines);

  results.total_added = rowsToWrite.length;
  rowsToWrite.forEach(r => results.added.push({ imei: r.imei, id: r.id || '' }));
  invalidateNalichieCache_();
  return results;
}

/** Поиск по складу для редактирования */
function adminStockSearch(token, filters){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

  const list = readSheetObjectsStrict_(ss, 'Склад') || [];
  const f = {
    imeiQuery: safeStr_(filters?.imeiQuery).replace(/\s+/g,''),
    model: safeStr_(filters?.model_name),
    city: safeStr_(filters?.city),
    condition: safeStr_(filters?.condition)
  };

  const res = (list || []).filter(r=>{
    const okImei = f.imeiQuery ? safeStr_(r['imei']).replace(/\s+/g,'').includes(f.imeiQuery) : true;
    const okModel = f.model ? safeStr_(r['model_name']) === f.model : true;
    const okCity  = f.city  ? safeStr_(r['city']) === f.city : true;
    const okCond  = f.condition ? safeStr_(r['condition']) === f.condition : true;
    return okImei && okModel && okCity && okCond;
  }).slice(0, 50);

  return { ok:true, rows: res };
}

/** Обновить строку склада по IMEI (телефон) */
function adminStockUpdate(token, payload){
  if (!adminValidateToken(token)) throw new Error('Нет доступа (требуется PIN)');
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

  const imei = safeStr_(payload?.imei).replace(/\s+/g,'');
  if (!imei) throw new Error('Укажите IMEI позиции, которую правим');

  const allowed = ['city','condition','model_name','memory','color','imei','sale_price','stock_statuses','note','id'];
  const patch = {};
  allowed.forEach(k => {
    if (payload.hasOwnProperty(k)) {
      patch[k] = (k === 'imei') ? safeStr_(payload[k]).replace(/\s+/g,'') : payload[k];
    }
  });

  const r = patchStockByImei_(ss, imei, patch);
  if (!r.ok) throw new Error('Позиция не найдена');
  invalidateNalichieCache_();
  return { ok:true, row: r.row };
}

/* ===================== Bootstrap для вкладки "Админ" ===================== */
function adminGetBootstrap(token){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(sh => sh.getName());
  return {
    ok: true,
    sheets,
    tz: (typeof getTz_ === 'function') ? getTz_() : 'Europe/Moscow',
    version: 'admin-ux-1',
    auth: adminGetAuthState(token)
  };
}

/** Удобно вызвать один раз в IDE, чтобы выставить PIN */
function adminInitPin() {
  // поставь свой PIN вместо 1234
  adminSetPin('742611');
}