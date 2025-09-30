/** DailyReport.gs — серверная логика ежедневного отчёта.
 * Не трогаем Code.gs. Используем его хелперы: getTz_(), readSheetObjectsStrict_(), normalizeDateKey_(),
 * toNumberSafe_(), round2_(), sumPaymentsFromString_(), getStaffColorMap_(), buildDictsColumnar_().
 */

/* ===== Локальная утилита: парсер "Метод:Сумма" / "Метод Сумма" → {method: amount} ===== */
function DR_parsePaymentsObject_(s) {
  const out = {};
  if (!s) return out;
  String(s).split(/[;,\n]/).forEach(part => {
    const p = String(part || '').trim();
    if (!p) return;
    // 1) "Метод:Сумма"
    let m = p.match(/^(.+?):\s*([0-9]+(?:[.,][0-9]+)?)$/);
    // 2) "Метод Сумма"
    if (!m) m = p.match(/^(.+?)\s+([0-9]+(?:[.,][0-9]+)?)$/);
    if (!m) return;
    const method = String(m[1] || '').trim();
    const n = Number(String(m[2] || '').replace(',', '.'));
    if (!method || isNaN(n)) return;
    out[method] = (out[method] || 0) + n;
  });
  return out;
}
function DR_addIntoMap_(agg, obj) {
  Object.keys(obj || {}).forEach(k => { agg[k] = (agg[k] || 0) + (obj[k] || 0); });
}
function DR_addSalary_(map, key, n) {
  if (!key) return;
  const v = Number(n || 0);
  if (isNaN(v)) return;
  map[key] = (map[key] || 0) + v;
}

/* ===== Основной агрегатор по дням ===== */
function DR_computeDailyAggregate_(ss) {
  const tz = getTz_();
  const dictRows = readSheetObjectsStrict_(ss, 'Справочники');
  const dicts = buildDictsColumnar_(dictRows);
  const staffColors = getStaffColorMap_(ss);

  const sales = readSheetObjectsStrict_(ss, 'Продажи') || [];
  const pre   = readSheetObjectsStrict_(ss, 'Предзаказы') || [];
  const shifts= readSheetObjectsStrict_(ss, 'Смены') || [];

  const days = {}; // dateKey -> {sales:[], preorders:[], chips:Set, ...}
  const ensureDay = (dateKey) => {
    const k = String(dateKey || '');
    if (!k) return null;
    if (!days[k]) {
      days[k] = {
        date: k,
        staffStores: [],             // [{store, staff, staff_color}]
        sales: [],                   // сырьевые строки "Продажи"
        preorders: [],               // сырьевые строки "Предзаказы" (помесячные/догоняющие — как в листе)
        totals: {
          sales_total: 0,
          preorders_paid: 0,
          payments_by_method: {},    // {method: sum}
          salary_by_store: {},       // {store: sum}
          salary_by_staff: {},       // {staff: sum}
        }
      };
      days[k]._chips = new Set();
    }
    return days[k];
  };

  // --- чипсы (кто отметился) из "Смены"
  (shifts || []).forEach(r => {
    const dateKey = normalizeDateKey_(r['date_vyhoda'], tz);
    const d = ensureDay(dateKey); if (!d) return;
    const store = String(r['store'] || '').trim();
    const staff = String(r['staff'] || '').trim();
    if (!store || !staff) return;
    const key = store + '||' + staff;
    if (!d._chips.has(key)) {
      d._chips.add(key);
      d.staffStores.push({ store, staff, staff_color: staffColors[staff] || '' });
    }
  });

  // --- ПРОДАЖИ
  (sales || []).forEach(r => {
    const dateKey = normalizeDateKey_(r['date'], tz);
    const d = ensureDay(dateKey); if (!d) return;

    const row = {
      id: String(r['id'] || ''),
      date: dateKey,
      store: String(r['store'] || ''),
      staff: String(r['staff'] || ''),
      item_type: String(r['item_type'] || ''),
      condition: String(r['condition'] || ''),
      model_name: String(r['model_name'] || ''),
      memory: String(r['memory'] || ''),
      color: String(r['color'] || ''),
      imei_or_sku: String(r['imei_or_sku'] || ''),
      total: toNumberSafe_(r['total']),
      payments: String(r['payments'] || ''),
      payments_map: DR_parsePaymentsObject_(String(r['payments'] || '')),
      sdacha: String(r['sdacha'] || ''),
      customer: String(r['customer'] || ''),
      phone: String(r['phone'] || ''),
      zarplata: toNumberSafe_(r['zarplata']),
      note: String(r['note'] || '')
    };
    d.sales.push(row);

    // День — итоги
    d.totals.sales_total += row.total || 0;
    DR_addIntoMap_(d.totals.payments_by_method, row.payments_map);
    DR_addSalary_(d.totals.salary_by_store, row.store, row.zarplata);
    DR_addSalary_(d.totals.salary_by_staff, row.staff, row.zarplata);
  });

  // --- ПРЕДЗАКАЗЫ (берём сырьевые строки по датам; "paid_row" — что реально внесено в эту дату)
  (pre || []).forEach(r => {
    const dateKey = normalizeDateKey_(r['date'], tz);
    const d = ensureDay(dateKey); if (!d) return;

    const paymentsStr = String(r['payments'] || '');
    // paid_row: если prepay задан – это «внесено этой строкой», иначе — парсим payments
    const prepayNum = toNumberSafe_(r['prepay']);
    const paid_row = prepayNum > 0 ? prepayNum : sumPaymentsFromString_(paymentsStr);

    const row = {
      id: String(r['id'] || ''),
      date: dateKey,
      store: String(r['store'] || ''),
      staff: String(r['staff'] || ''),
      preorder_statuses: String(r['preorder_statuses'] || ''),
      model_name: String(r['model_name'] || ''),
      memory: String(r['memory'] || ''),
      color: String(r['color'] || ''),
      pre_imei: String(r['pre_imei'] || ''),
      pre_price: toNumberSafe_(r['pre_price']),
      prepay: prepayNum,
      payments: paymentsStr,
      payments_map: DR_parsePaymentsObject_(paymentsStr),
      customer: String(r['customer'] || ''),
      phone: String(r['phone'] || ''),
      zarplata: toNumberSafe_(r['zarplata']),
      note: String(r['note'] || ''),
      _paid_row: paid_row
    };
    d.preorders.push(row);

    d.totals.preorders_paid += paid_row || 0;
    DR_addIntoMap_(d.totals.payments_by_method, row.payments_map);
    DR_addSalary_(d.totals.salary_by_store, row.store, row.zarplata);
    DR_addSalary_(d.totals.salary_by_staff, row.staff, row.zarplata);
  });

  // --- свод по всем дням
  const dayList = Object.values(days).sort((a,b) => {
    // DD.MM.YYYY → ms
    const A = a.date.split('.'); const B = b.date.split('.');
    const ams = new Date(`${A[2]}-${A[1]}-${A[0]}T00:00:00`).getTime();
    const bms = new Date(`${B[2]}-${B[1]}-${B[0]}T00:00:00`).getTime();
    return bms - ams;
  });

  const totalsAll = {
    sales_total: 0,
    preorders_paid: 0,
    payments_by_method: {},
    salary_by_store: {},
    salary_by_staff: {}
  };
  dayList.forEach(d => {
    totalsAll.sales_total += d.totals.sales_total;
    totalsAll.preorders_paid += d.totals.preorders_paid;
    DR_addIntoMap_(totalsAll.payments_by_method, d.totals.payments_by_method);
    DR_addIntoMap_(totalsAll.salary_by_store, d.totals.salary_by_store);
    DR_addIntoMap_(totalsAll.salary_by_staff, d.totals.salary_by_staff);
  });

  // Округлим слегка
  const roundMap = (m) => {
    const o = {}; Object.keys(m || {}).forEach(k => o[k] = round2_(m[k] || 0)); return o;
  };

  return {
    dicts: {
      stores: dicts.stores || [],
      staff:  dicts.staff  || [],
      item_types: dicts.item_types || [],
      payments: dicts.payments || [],
      staffColors
    },
    days: dayList.map(d => ({
      date: d.date,
      staffStores: d.staffStores,
      sales: d.sales,
      preorders: d.preorders,
      totals: {
        sales_total: round2_(d.totals.sales_total),
        preorders_paid: round2_(d.totals.preorders_paid),
        payments_by_method: roundMap(d.totals.payments_by_method),
        salary_by_store: roundMap(d.totals.salary_by_store),
        salary_by_staff: roundMap(d.totals.salary_by_staff)
      }
    })),
    totals: {
      sales_total: round2_(totalsAll.sales_total),
      preorders_paid: round2_(totalsAll.preorders_paid),
      payments_by_method: roundMap(totalsAll.payments_by_method),
      salary_by_store: roundMap(totalsAll.salary_by_store),
      salary_by_staff: roundMap(totalsAll.salary_by_staff)
    },
    updatedAt: Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm')
  };
}

/* ===== Быстрый кэш как в других вкладках ===== */
function getDailyReportBootstrap() {
  const key = 'DAILY_BOOT_V1';
  let hit = cacheGetJson_(key);
  if (hit) {
    if (!hit._cacheToken) {
      hit._cacheToken = 'v' + Date.now();
      cachePutJson_(key, hit, 600);
    }
    return hit;
  }
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const data = DR_computeDailyAggregate_(ss);
  data._cacheToken = 'v' + Date.now();
  cachePutJson_(key, data, 600); // 10 минут
  return data;
}