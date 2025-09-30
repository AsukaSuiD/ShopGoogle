/** ===== SortStockAuto.gs — автосорт ТОЛЬКО листа "Склад" =====
 * Сортирует "Склад" (строки с 3-й) по правилам:
 * 1) stock_statuses: "В наличии" → "Продан" → прочие (как в «Справочники»)
 * 2) city: порядок из «Справочники»
 * 3) condition: порядок из «Справочники»
 * 4) model_name → memory → color: порядок из «Каталог моделей» (с fallback по числу памяти)
 * Безопасно: добавляются 6 временных колонок-ключей, range.sort(), затем удаляются.
 */

const SSA_SHEET_STOCK   = 'Склад';
const SSA_SHEET_DICT    = 'Справочники';
const SSA_SHEET_CATALOG = 'Каталог моделей';
const SSA_HEADER_ROW    = 1;
const SSA_DATA_START    = 3;

/** Один раз запусти, чтобы поставить триггер */
function SSA_setupAutoSort(){
  SSA_removeAutoSort();
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('SSA_onEdit').forSpreadsheet(ss).onEdit().create();
  // меню для ручного запуска (не сортирует автоматически)
  ScriptApp.newTrigger('SSA_addMenuOpen').forSpreadsheet(ss).onOpen().create();
}

/** Снять триггеры этого файла */
function SSA_removeAutoSort(){
  ScriptApp.getProjectTriggers()
    .filter(t => ['SSA_onEdit','SSA_addMenuOpen'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** Меню с ручным запуском */
function SSA_addMenuOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('Сервис (Склад)')
      .addItem('Сортировать "Склад" сейчас','SSA_sortStockNow')
      .addToUi();
  }catch(_){}
}

/** Ручной запуск */
function SSA_sortStockNow(){
  SSA_sortInPlace_();
}

/** Инсталлимый onEdit — реагируем ТОЛЬКО на правки на листе "Склад" (с 3-й строки) */
function SSA_onEdit(e){
  const sh = e && e.range && e.range.getSheet ? e.range.getSheet() : null;
  if (!sh || sh.getName() !== SSA_SHEET_STOCK) return;
  if (e && e.range && e.range.getRow() < SSA_DATA_START) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;
  try{
    SSA_sortInPlace_();
  } finally {
    lock.releaseLock();
  }
}

/* ================= внутренняя сортировка без потери формул ================= */

function SSA_sortInPlace_(){
  const ss = SpreadsheetApp.getActive();
  const shStock = ss.getSheetByName(SSA_SHEET_STOCK);
  const shDict  = ss.getSheetByName(SSA_SHEET_DICT);
  const shCat   = ss.getSheetByName(SSA_SHEET_CATALOG);
  if (!shStock || !shDict || !shCat) throw new Error('Проверь имена листов: Склад / Справочники / Каталог моделей');

  const lastRow = shStock.getLastRow();
  const lastCol = shStock.getLastColumn();
  if (lastRow < SSA_DATA_START) return;

  const hdr = shStock.getRange(SSA_HEADER_ROW,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  const col = (names)=>{
    const a = Array.isArray(names)? names : [names];
    for (const n of a){ const i = hdr.indexOf(n); if (i !== -1) return i; }
    throw new Error('Нет колонки: '+JSON.stringify(a));
  };
  const idx = {
    city:   col('city'),
    cond:   col('condition'),
    model:  col('model_name'),
    memory: col('memory'),
    color:  col('color'),
    status: col(['stock_statuses','stock_status'])
  };

  // --- словари («Справочники») ---
  const dHdr = shDict.getRange(1,1,1,shDict.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const dCol = (name)=>{ const i=dHdr.indexOf(name); return i===-1? -1 : i+1; };
  const Lcity = getDictList_(shDict, dCol('city'));
  const Lcond = getDictList_(shDict, dCol('condition'));
  const Lstat = getDictList_(shDict, dCol('stock_statuses'));

  const norm = s=>String(s||'').trim().toLowerCase();
  const mapOrder = (arr)=>{ const m=new Map(); arr.forEach((v,i)=>m.set(norm(v), i)); return m; };

  const oCity = mapOrder(Lcity);
  const oCond = mapOrder(Lcond);
  const oStat = new Map([[norm('В наличии'),0],[norm('Продан'),1]]);
  let base = 2; Lstat.forEach(v=>{ const n=norm(v); if(!oStat.has(n)) oStat.set(n, base++); });

  // --- каталог («Каталог моделей») ---
  const cHdr = shCat.getRange(1,1,1,shCat.getLastColumn()).getValues()[0].map(s=>String(s||'').trim());
  const ciModel = cHdr.indexOf('model_name');
  const ciMem   = cHdr.indexOf('memory');
  const ciCol   = cHdr.indexOf('color');
  if (ciModel===-1 || ciMem===-1 || ciCol===-1) throw new Error('В «Каталог моделей» нужны: model_name, memory, color');

  const cData = shCat.getRange(2,1,Math.max(0,shCat.getLastRow()-1),shCat.getLastColumn()).getValues();
  const oModel = new Map();
  const oMemByModel  = new Map(); // model -> Map(memory->ord)
  const oColByModMem = new Map(); // model||mem -> Map(color->ord)

  const memMap = (m)=>{ const k=norm(m); if(!oMemByModel.has(k)) oMemByModel.set(k,new Map()); return oMemByModel.get(k); };
  const colMap = (m,mem)=>{ const k=norm(m)+'||'+norm(mem); if(!oColByModMem.has(k)) oColByModMem.set(k,new Map()); return oColByModMem.get(k); };

  let mIdx = 0;
  cData.forEach(r=>{
    const m=r[ciModel], mem=r[ciMem], colr=r[ciCol];
    if(!m) return;
    const mk = norm(m);
    if(!oModel.has(mk)) oModel.set(mk, mIdx++);
    if(mem){ const mm=memMap(m); const key=norm(mem); if(!mm.has(key)) mm.set(key, mm.size); }
    if(mem && colr){ const cm=colMap(m,mem); const ck=norm(colr); if(!cm.has(ck)) cm.set(ck, cm.size); }
  });

  // --- подготовка ключей ---
  const dataRange = shStock.getRange(SSA_DATA_START, 1, lastRow - SSA_DATA_START + 1, lastCol);
  const values = dataRange.getValues();

  const keys = values.map(row=>{
    const statusW = oStat.has(norm(row[idx.status])) ? oStat.get(norm(row[idx.status])) : 999;
    const cityW   = oCity.has(norm(row[idx.city]))   ? oCity.get(norm(row[idx.city]))   : 999;
    const condW   = oCond.has(norm(row[idx.cond]))   ? oCond.get(norm(row[idx.cond]))   : 999;

    const m = row[idx.model], mem = row[idx.memory], colr = row[idx.color];
    const mo = oModel.has(norm(m)) ? oModel.get(norm(m)) : 1e6;

    let memW = 1e9;
    const mm = oMemByModel.get(norm(m));
    if (mm && mm.has(norm(mem))) {
      memW = mm.get(norm(mem));
    } else {
      const sMem = String(mem||'').toLowerCase();
      const mNum = sMem.match(/(\d+(?:[.,]\d+)?)/);
      let num = mNum ? parseFloat(mNum[1].replace(',', '.')) : 999999;
      if (/тб|tb/.test(sMem)) num *= 1024;
      memW = 1e9 + num;
    }

    const cm = oColByModMem.get(norm(m)+'||'+norm(mem));
    const colW = (cm && cm.has(norm(colr))) ? cm.get(norm(colr)) : 1e9;

    return [statusW, cityW, condW, mo, memW, colW];
  });

  // --- сортировка через временные колонки ---
  const tempFirst = lastCol + 1;
  shStock.insertColumnsAfter(lastCol, 6);
  shStock.getRange(SSA_DATA_START, tempFirst, lastRow - SSA_DATA_START + 1, 6).setValues(keys);

  const sortRange = shStock.getRange(SSA_DATA_START, 1, lastRow - SSA_DATA_START + 1, lastCol + 6);
  sortRange.sort([
    {column: tempFirst+0, ascending:true},
    {column: tempFirst+1, ascending:true},
    {column: tempFirst+2, ascending:true},
    {column: tempFirst+3, ascending:true},
    {column: tempFirst+4, ascending:true},
    {column: tempFirst+5, ascending:true},
  ]);

  shStock.deleteColumns(tempFirst, 6);
}

/* helpers */
function getDictList_(sheet, colIndex1based){
  if (colIndex1based === -1) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, colIndex1based, last-1, 1).getValues()
    .map(r=>String(r[0]||'').trim()).filter(Boolean);
}
