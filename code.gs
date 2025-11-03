/* ==================== 1) OKR TANIMLARI ==================================== */
const OKR_METRICS = [
  { name: "Decrease number of WHs with availability below 85%", target: { G10: 0, GMore: 0 }, unit: "", behavior: "decrease_to_zero_numeric" },
  { name: "Decrease the number of pending + rejected pallet backlog (Ops)", target: { G10: 7, GMore: 7 }, unit: "%" },
  { name: "Achieve LMD variable cost per order target for G10 and GMore (Courier + Fleet)", target: { G10: 55.75, GMore: 55.75 }, unit: "₺" },
  { name: "Decrease the problematic order ratio (G10 and GMore)", target: { G10: 1.0, GMore: 2.0 }, unit: "%" },
  { name: "Increase throughput for G10 and GMore", target: { G10: 3.96, GMore: 3.96 }, unit: "" },
  { name: "Achieve Waste + Waste A&M ratio target", target: { G10: 3.80, GMore: 3.80 }, unit: "%" },
  { name: "Decrease missed order ratio", target: { G10: 0.70, GMore: 2.0 }, unit: "%" },
  { name: "Decrease lateness percent", target: { G10: 6.20, GMore: 9.0 }, unit: "%" },
  { name: "Decrease non-agreed / problematic shipment ratios", target: { G10: 0.30, GMore: 0.80 }, unit: "%" },
  { name: "Increase stock accuracy in dark stores", target: { G10: 87.50, GMore: 84.50 }, unit: "%" },
  { name: "Increase Franchise Satisfaction Score", target: { G10: 10, GMore: 10 }, unit: "pp" },
  { name: "Increase GMore fresh order penetration", target: { G10: null, GMore: 42.2 }, unit: "%" },
  { name: "Decrease GMore customer quality complaint rate (problematic order) F&V", target: { G10: null, GMore: 1.13 }, unit: "%" }
];

/* ==================== 2) GENEL YARDIMCILAR ================================= */
function listAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheets().map(s => s.getName());
  } catch (e) {
    console.error('listAllSheets error:', e);
    return [];
  }
}

function doGet() {
  try {
    listAllSheets();
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Getir OKR Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return HtmlService.createHtmlOutput('<h1>Hata</h1><p>' + e + '</p>');
  }
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/* ---------- Sheet okuma ---------- */
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      const target = sheetName.toLowerCase().trim();
      sheet = ss.getSheets().find(s => {
        const n = s.getName().toLowerCase().trim();
        return n === target || n.includes(target) || target.includes(n);
      });
    }
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    return data.length > 1 ? data.slice(1) : null;
  } catch (e) {
    console.error('getSheetData error:', sheetName, e);
    return null;
  }
}

/* ---------- Domain normalize ---------- */
function _normDomain(d) {
  return String(d || '').trim()
    .replace('Getir10', 'G10').replace('Getir 10', 'G10')
    .replace('Getir More', 'GMore').replace('GetirMore', 'GMore');
}

/* ==================== WAREHOUSE MANAGERS YARDIMCILARI ====================== */

// "Temmuz 2025" ya da Date -> { key:"2025-07", date: Date, label:"Temmuz 2025" }
function _parseTRMonthYear(v) {
  if (v instanceof Date) {
    const d = new Date(v.getFullYear(), v.getMonth(), 1);
    const TR = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'];
    return { key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`, date: d, label: `${TR[d.getMonth()]} ${d.getFullYear()}` };
  }
  const s = String(v || '').toLowerCase()
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/ı/g, 'i').replace(/\s+/g, ' ').trim();
  const map = { ocak: 0, subat: 1, mart: 2, nisan: 3, mayis: 4, haziran: 5, temmuz: 6, agustos: 7, eylul: 8, ekim: 9, kasim: 10, aralik: 11 };
  const m = s.match(/^(ocak|subat|mart|nisan|mayis|haziran|temmuz|agustos|eylul|ekim|kasim|aralik)\s+(\d{4})$/);
  if (!m) return null; const mo = map[m[1]], y = +m[2];
  const d = new Date(y, mo, 1);
  const TR = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'];
  return { key: `${y}-${String(mo + 1).padStart(2, '0')}`, date: d, label: `${TR[mo]} ${y}` };
}

// W label (W27) -> 27
function _weekNum(weekLabel) {
  const n = parseInt(String(weekLabel || '').replace(/\D/g, ''), 10);
  return Number.isFinite(n) ? n : null;
}
// W27 ve sonrası mı?
function _isWeek27Plus(weekLabel) {
  const n = _weekNum(weekLabel);
  return n == null ? true : n >= 27;
}
function _isWeek31Plus(weekLabel) {
  const n = _weekNum(weekLabel);
  return n == null ? true : n >= 31;
}
function _hasManagerFilter(filters) {
  return !!(
    (filters?.sahaYoneticisi && filters.sahaYoneticisi !== 'all') ||
    (filters?.operasyonMuduru && filters.operasyonMuduru !== 'all') ||
    (filters?.bolgeMuduru && filters.bolgeMuduru !== 'all')
  );
}
// Target sheet sayısal dönüşüm
function _toNum(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v; // Sheet gerçek sayı döndürmüşse dokunma

  let s = String(v).trim();
  // işaret ve ayırıcıları dışındakileri temizle
  s = s.replace(/[^\d.,\-]/g, '');

  // eksi işaretini ayır
  let sign = '';
  if (s.startsWith('-')) { sign = '-'; s = s.slice(1); }

  const lastDot = s.lastIndexOf('.');
  const lastComma = s.lastIndexOf(',');

  if (lastDot === -1 && lastComma === -1) {
    // hiç ayıraç yok → tüm . , yok → düz tamsayı
    s = sign + s.replace(/[.,]/g, '');
  } else {
    // sağdaki son ayıraç ondalık, diğerleri binlik
    const decPos = Math.max(lastDot, lastComma);
    const intPart = s.slice(0, decPos).replace(/[.,]/g, '');
    const fracPart = s.slice(decPos + 1).replace(/[.,]/g, '');
    s = sign + intPart + '.' + fracPart;
  }

  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

// İSİM ve BAŞLIK normalize yardımcıları
function _normName(s) {
  return String(s || '')
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim();
}
function _normHeaderText(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/[^\w%&()\/+.\-]/g, '') // metin dışı karakterleri at
    .trim();
}

let personTargetsCache = null;
function getPersonTargets() {
  if (personTargetsCache) return personTargetsCache;

  const data = getSheetData('Target');
  const out = {};
  if (!data || !data.length) { personTargetsCache = out; return out; }

  // 1) Başlık satırını "en çok eşleşen" satır şeklinde bul
  const wanted = OKR_METRICS.map(m => _normHeaderText(m.name));
  let headerRow = null, headerIdx = -1, maxHits = -1;

  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i] || [];
    const hits = row.reduce((acc, cell) => {
      const t = _normHeaderText(cell);
      return acc + (wanted.includes(t) ? 1 : 0);
    }, 0);
    if (hits > maxHits) {
      maxHits = hits;
      headerRow = row;
      headerIdx = i;
    }
  }
  // Fallback: "increase/decrease/achieve" içeren ilk satır
  if (!headerRow) {
    headerIdx = data.findIndex(r => (r || []).some(c => /increase|decrease|achieve/i.test(String(c))));
    headerRow = headerIdx >= 0 ? data[headerIdx] : (data[0] || []);
  }

  // 2) Kolon → metrik eşleşmesi
  const colToMetric = {};
  for (let c = 1; c < headerRow.length; c++) {
    const hh = _normHeaderText(headerRow[c]);
    const mm = OKR_METRICS.find(m => _normHeaderText(m.name) === hh);
    if (mm) colToMetric[c] = mm.name;
  }

  // 3) Başlığın altındaki satırlardan kişi → hedefler haritasını kur
  for (let r = headerIdx + 1; r < data.length; r++) {
    const row = data[r] || [];
    const rawName = row[0];
    const name = _normName(rawName);
    if (!name) continue;

    // Grup başlıklarını atla
    const nlow = name.toLowerCase();
    if (nlow === 'ops manager' || nlow === 'bölge müdürü' || nlow === 'saha yöneticisi') continue;

    const rec = {};
    Object.keys(colToMetric).forEach(k => {
      const col = +k;
      const v = _toNum(row[col]);
      if (v != null) rec[colToMetric[col]] = v;
    });
    if (Object.keys(rec).length) out[name] = rec;
  }

  personTargetsCache = out;
  return out;
}

// Seçilen yöneticiye göre hedefi çöz (yoksa domain hedefi)
function resolveTargetValue(domain, filters, metricName) {
  const personTargets = getPersonTargets();
  let person =
    (filters.sahaYoneticisi && filters.sahaYoneticisi !== 'all' && filters.sahaYoneticisi) ||
    (filters.operasyonMuduru && filters.operasyonMuduru !== 'all' && filters.operasyonMuduru) ||
    (filters.bolgeMuduru && filters.bolgeMuduru !== 'all' && filters.bolgeMuduru) || null;

  if (person) {
    const key = _normName(person);
    if (personTargets[key] && personTargets[key][metricName] != null) {
      return +personTargets[key][metricName];
    }
  }
  const m = OKR_METRICS.find(x => x.name === metricName);
  return (m && m.target) ? m.target[domain] : null;
}


let warehouseManagersCache = null;

function getWarehouseManagers() {
  if (warehouseManagersCache) return warehouseManagersCache;

  try {
    const data = getSheetData('Warehouse Managers');
    const managers = {};
    if (!data) { warehouseManagersCache = managers; return managers; }

    data.forEach(row => {
      const warehouse = String(row[0] || '').trim();      // A: WH
      if (!warehouse) return;

      const domain = _normDomain(row[1]);              // B: Domain
      const sahaYoneticisi = String(row[2] || '').trim();      // C: Saha Yöneticisi
      const operasyonMuduru = String(row[3] || '').trim();      // D: Ops Manager
      const bolgeMuduru = String(row[4] || '').trim();      // E: Bölge Müdürü  ← YENİ

      managers[warehouse] = {
        domain,
        sahaYoneticisi,
        operasyonMuduru,
        bolgeMuduru
      };
    });

    warehouseManagersCache = managers;
    return managers;
  } catch (e) {
    console.error('getWarehouseManagers error:', e);
    return {};
  }
}

// dd.mm.yyyy ya da Date -> Date
function _parseTRDate(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!m) return null;
  const d = +m[1], mo = +m[2] - 1, y = +m[3];
  const dt = new Date(y, mo, d);
  return isNaN(dt.getTime()) ? null : dt;
}

// ISO week number (Mon=1) -> 1..53
function _isoWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}
function _isoWeekLabel(date) { return 'W' + _isoWeekNumber(date); }

function getWarehousesByFilter(filters) {
  const managers = getWarehouseManagers();
  const filtered = new Set();

  Object.entries(managers).forEach(([wh, info]) => {
    if (filters?.domain && filters.domain !== 'all' && info.domain !== filters.domain) return;
    if (filters?.sahaYoneticisi && filters.sahaYoneticisi !== 'all' && info.sahaYoneticisi !== filters.sahaYoneticisi) return;
    if (filters?.operasyonMuduru && filters.operasyonMuduru !== 'all' && info.operasyonMuduru !== filters.operasyonMuduru) return;
    if (filters?.bolgeMuduru && filters.bolgeMuduru !== 'all' && info.bolgeMuduru !== filters.bolgeMuduru) return;
    if (filters?.warehouse && filters.warehouse !== 'all' && filters.warehouse !== wh) return;
    filtered.add(wh);
  });

  return filtered;
}

function applyManagerFilters(warehouse, domain, filters) {
  // Yönetici filtresi seçilmiş mi?
  const hasManagerFilter =
    (filters?.sahaYoneticisi && filters.sahaYoneticisi !== 'all') ||
    (filters?.operasyonMuduru && filters.operasyonMuduru !== 'all') ||
    (filters?.bolgeMuduru && filters.bolgeMuduru !== 'all');

  // Yönetici filtresi yoksa klasik kontroller
  if (!hasManagerFilter) {
    if (filters?.warehouse && filters.warehouse !== 'all' && filters.warehouse !== warehouse) return false;
    if (filters?.domain && filters.domain !== 'all' && filters.domain !== domain) return false;
    return true;
  }

  // Yönetici filtresi varsa, izinli depolar setinden kontrol et
  const allowedWarehouses = getWarehousesByFilter(filters);
  return allowedWarehouses.has(warehouse);
}

/* ==================== 3) METRİK HESAPLAMALARI ============================== */
/* --- A) WH availability <85% (adet) --------------------------------------- */
function calculateWHBelow85Data(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], avail = r[3], total = r[4];
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weeklyData: {}, whSet: {} };
    if (!res[dom].weeklyData[week]) { res[dom].weeklyData[week] = { belowCount: 0 }; res[dom].whSet[week] = new Set(); }

    if (avail != null && total != null && total > 0) {
      const ratio = avail / total;
      if (ratio < 0.85) {
        const key = String(wh);
        if (!res[dom].whSet[week].has(key)) {
          res[dom].weeklyData[week].belowCount++;
          res[dom].whSet[week].add(key);
        }
      }
    }
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weeklyData)
      .filter(_isWeek27Plus)
      .sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;
    const latest = weeks[weeks.length - 1];
    const cur = d.weeklyData[latest].belowCount;

    const data = {
      domain: d.domain,
      currentValue: cur,
      targetValue: resolveTargetValue(d.domain, filters, "Decrease number of WHs with availability below 85%"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const val = d.weeklyData[w].belowCount;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    out.push(data);
  });
  return out;
}

/* --- B) Lateness % --------------------------------------------------------- */
function calculateLatenessData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], late = +r[3] || 0, order = +r[4] || 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek31Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { late: 0, order: 0 };
    res[dom].weekly[week].late += late;
    res[dom].weekly[week].order += order;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek31Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totL = 0, totO = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Decrease lateness percent"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const val = ww.order > 0 ? (ww.late / ww.order) * 100 : 0;
      totL += ww.late; totO += ww.order;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = totO > 0 ? (totL / totO) * 100 : 0;
    out.push(data);
  });
  return out;
}

/* --- C) Stock accuracy % --------------------------------------------------- */
function calculateStockAccuracyData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], A = +r[3] || 0, B = +r[4] || 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { A: 0, B: 0 };
    res[dom].weekly[week].A += A;
    res[dom].weekly[week].B += B;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totA = 0, totB = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Increase stock accuracy in dark stores"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const val = ww.A > 0 ? (ww.B / ww.A) * 100 : 0;
      totA += ww.A; totB += ww.B;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = totA > 0 ? (totB / totA) * 100 : 0;
    out.push(data);
  });
  return out;
}

/* --- D) Throughput --------------------------------------------------------- */
function calculateThroughputData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], orders = +r[3] || 0, hrs = +r[4] || 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { orders: 0, hrs: 0 };
    res[dom].weekly[week].orders += orders;
    res[dom].weekly[week].hrs += hrs;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totO = 0, totH = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Increase throughput for G10 and GMore"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const val = ww.hrs > 0 ? (ww.orders / ww.hrs) : 0;
      totO += ww.orders; totH += ww.hrs;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = totH > 0 ? (totO / totH) : 0;
    out.push(data);
  });
  return out;
}

/* --- E) Missed order ratio % ---------------------------------------------- */
function calculateMissedOrderData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], delivered = +r[3] || 0, missed = +r[4] || 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek31Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { del: 0, miss: 0 };
    res[dom].weekly[week].del += delivered;
    res[dom].weekly[week].miss += missed;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek31Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totD = 0, totM = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Decrease missed order ratio"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const total = ww.del + ww.miss;
      const val = total > 0 ? (ww.miss / total) * 100 : 0;
      totD += ww.del; totM += ww.miss;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    const totalAll = totD + totM;
    data.currentValue = totalAll > 0 ? (totM / totalAll) * 100 : 0;
    out.push(data);
  });
  return out;
}

/* --- F) Problematic order ratio % (genel) ---------------------------------- */
function calculateProblematicOrderRatioData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], orders = +r[3] || 0, fb = +r[4] || 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { orders: 0, fb: 0 };
    res[dom].weekly[week].orders += orders;
    res[dom].weekly[week].fb += fb;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totO = 0, totFb = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Decrease the problematic order ratio (G10 and GMore)"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const val = ww.orders > 0 ? (ww.fb / ww.orders) * 100 : 0;
      totO += ww.orders; totFb += ww.fb;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = totO > 0 ? (totFb / totO) * 100 : 0;
    out.push(data);
  });
  return out;
}

/* --- NEW) Pending + rejected pallet backlog (Ops) ------------------------- */
/* Sheet kolonları:
   B: Tarih (dd.mm.yyyy) | C: Sorumlu | D: Günlük Sevk Edilen Palet | E: Backlog
   Oran = E / D * 100  (haftalık toplamlardan)
   Yalnızca C == "Operasyon" ve B >= 01.07.2025.
*/
function calculatePalletBacklogOpsData(sheetData, filters) {
  const START = new Date(2025, 6, 1); // 01.07.2025
  const weekly = {}; // week -> { ship, back }
  let totalShip = 0, totalBack = 0;

  sheetData.forEach(r => {
    const dateVal = _parseTRDate(r[1]);
    const sorumlu = String(r[2] || '').trim().toLowerCase();
    const shipped = +r[3] || 0;
    const backlog = +r[4] || 0;
    if (!dateVal || dateVal < START) return;
    if (sorumlu !== 'operasyon') return;

    const week = _isoWeekLabel(dateVal);
    if (!_isWeek27Plus(week)) return;

    if (!weekly[week]) weekly[week] = { ship: 0, back: 0 };
    weekly[week].ship += shipped;
    weekly[week].back += backlog;

    totalShip += shipped;
    totalBack += backlog;
  });

  // hedef fallback'i için herhangi bir domain yeterli (değerler eşit)
  const domainForTarget =
    (filters.domain && filters.domain !== 'all') ? filters.domain : 'GMore';

  const data = {
    domain: 'MultiTarget',
    currentValue: totalShip > 0 ? (totalBack / totalShip) * 100 : 0,
    targetValue: resolveTargetValue(
      domainForTarget,
      filters,
      "Decrease the number of pending + rejected pallet backlog (Ops)"
    ),
    weeklyData: [],
    trend: []
  };

  Object.keys(weekly)
    .filter(_isWeek27Plus)
    .sort((a, b) => _weekNum(a) - _weekNum(b))
    .forEach(w => {
      const ww = weekly[w];
      if (ww.ship > 0) {
        const val = (ww.back / ww.ship) * 100;
        data.weeklyData.push({ week: w, value: val });
        data.trend.push(val);
      }
    });

  return (data.weeklyData.length || totalShip > 0) ? [data] : [];
}

function calculateSatisfactionScoreData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], score = _toNum(r[3]) ?? 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { sum: 0, cnt: 0 };
    res[dom].weekly[week].sum += score;
    res[dom].weekly[week].cnt += 1;
  });

  const out = [];
  ['G10', 'GMore'].forEach(dom => {
    const weeks = Object.keys(res[dom]?.weekly || {}).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let Tsum = 0, Tcnt = 0;
    const data = {
      domain: dom,
      currentValue: 0,
      targetValue: resolveTargetValue(dom, filters, "Increase Franchise Satisfaction Score"),
      weeklyData: [], trend: []
    };
    weeks.forEach(w => {
      const ww = res[dom].weekly[w];
      const val = ww.cnt ? (ww.sum / ww.cnt) : 0;
      Tsum += ww.sum; Tcnt += ww.cnt;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = Tcnt ? (Tsum / Tcnt) : 0;
    out.push(data);
  });

  return out;
}

/* --- G) ► NEW  F&V Complaint rate % (yalnızca GMore) ----------------------- */
function calculateFVCustomerComplaintData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2], order = +r[3] || 0, fb = +r[4] || 0;
    if (dom !== 'GMore') return;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { order: 0, fb: 0 };
    res[dom].weekly[week].order += order;
    res[dom].weekly[week].fb += fb;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totO = 0, totFb = 0;
    const data = {
      domain: "GMore",
      currentValue: 0,
      targetValue: resolveTargetValue("GMore", filters, "Decrease GMore customer quality complaint rate (problematic order) F&V"),
      weeklyData: [], trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const val = ww.order > 0 ? (ww.fb / ww.order) * 100 : 0;
      totO += ww.order; totFb += ww.fb;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = totO > 0 ? (totFb / totO) * 100 : 0;
    out.push(data);
  });
  return out;
}


/* --- H) ► NEW  GMore Fresh Order Penetration % (yalnızca GMore) ------------- */
function calculateGMoreFreshOrderData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2];
    const other = +r[3] || 0, fv = +r[4] || 0;

    if (dom !== 'GMore') return;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { other: 0, fv: 0 };
    res[dom].weekly[week].other += other;
    res[dom].weekly[week].fv += fv;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totO = 0, totF = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Increase GMore fresh order penetration"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const total = ww.other + ww.fv;
      const val = total > 0 ? (ww.fv / total) * 100 : 0;
      totO += ww.other; totF += ww.fv;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    const totalAll = totO + totF;
    data.currentValue = totalAll > 0 ? (totF / totalAll) * 100 : 0;
    out.push(data);
  });
  return out;
}


function calculateNonAgreedShipmentData(sheetData, filters) {
  const res = { G10: { weekly: {} }, GMore: { weekly: {} } };

  sheetData.forEach(r => {
    const week = r[0];                               // A: Week (sayı ya da "W27")
    const wh = String(r[1] || '').trim();          // B: Warehouse Name
    const dom = _normDomain(r[2]);                  // C: Domain  (Getir10 / Getir More)
    const prob = _toNum(r[3]) || 0;                  // D: Mutabakatsızlık Tutar
    const total = _toNum(r[4]) || 0;                  // E: Merkez Depo Gönderilen Tutar

    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom].weekly[week]) res[dom].weekly[week] = { prob: 0, total: 0 };
    res[dom].weekly[week].prob += prob;
    res[dom].weekly[week].total += total;
  });

  const out = [];
  ['G10', 'GMore'].forEach(dom => {
    const weeks = Object.keys(res[dom].weekly)
      .filter(_isWeek27Plus)
      .sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totP = 0, totT = 0;
    const data = {
      domain: dom,
      currentValue: 0,
      targetValue: resolveTargetValue(dom, filters, "Decrease non-agreed / problematic shipment ratios"),
      weeklyData: [],
      trend: []
    };

    weeks.forEach(w => {
      const ww = res[dom].weekly[w];
      const val = ww.total > 0 ? (ww.prob / ww.total) * 100 : 0;
      totP += ww.prob; totT += ww.total;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });

    data.currentValue = totT > 0 ? (totP / totT) * 100 : 0;
    out.push(data);
  });

  return out;
}

function calculateLMDVariableCostData(sheetData, filters) {
  const res = {};
  sheetData.forEach(r => {
    const wh = r[0], dom = _normDomain(r[1]), week = r[2];
    const courier = +r[3] || 0, fleet = +r[4] || 0, orders = +r[5] || 0;
    if (!applyManagerFilters(wh, dom, filters)) return;
    if (dom !== 'G10' && dom !== 'GMore') return;
    if (!_isWeek27Plus(week)) return;

    if (!res[dom]) res[dom] = { domain: dom, weekly: {} };
    if (!res[dom].weekly[week]) res[dom].weekly[week] = { cost: 0, orders: 0 };
    res[dom].weekly[week].cost += (courier + fleet);
    res[dom].weekly[week].orders += orders;
  });

  const out = [];
  Object.values(res).forEach(d => {
    const weeks = Object.keys(d.weekly).filter(_isWeek27Plus).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return;

    let totC = 0, totO = 0;
    const data = {
      domain: d.domain,
      currentValue: 0,
      targetValue: resolveTargetValue(d.domain, filters, "Achieve LMD variable cost per order target for G10 and GMore (Courier + Fleet)"),
      weeklyData: [],
      trend: []
    };
    weeks.forEach(w => {
      const ww = d.weekly[w];
      const val = ww.orders > 0 ? (ww.cost / ww.orders) : 0;
      totC += ww.cost; totO += ww.orders;
      data.weeklyData.push({ week: w, value: val });
      data.trend.push(val);
    });
    data.currentValue = totO > 0 ? (totC / totO) : 0;
    out.push(data);
  });
  return out;
}

function calculateWasteAMRatioData(sheetData, filters) {
  const monthly = {}; // key -> { label, date, w, n, c }
  let totalWaste = 0, totalCOGS = 0;

  sheetData.forEach(r => {
    const m = _parseTRMonthYear(r[0]); if (!m) return;
    const wasteAM = +r[1] || 0, netAtik = +r[2] || 0, cogs = +r[3] || 0;

    if (!monthly[m.key]) monthly[m.key] = { label: m.label, date: m.date, w: 0, n: 0, c: 0 };
    monthly[m.key].w += wasteAM;
    monthly[m.key].n += netAtik;
    monthly[m.key].c += cogs;

    totalWaste += (wasteAM + netAtik);
    totalCOGS += cogs;
  });

  const keys = Object.keys(monthly).sort((a, b) => monthly[a].date - monthly[b].date);
  const domainForTarget =
    (filters.domain && filters.domain !== 'all') ? filters.domain : 'GMore';

  const data = {
    domain: 'MultiTarget',
    currentValue: totalCOGS > 0 ? (totalWaste / totalCOGS) * 100 : 0,
    targetValue: resolveTargetValue(
      domainForTarget,
      filters,
      "Achieve Waste + Waste A&M ratio target"
    ),
    weeklyData: [],
    trend: []
  };

  keys.forEach(k => {
    const mm = monthly[k];
    if (mm.c > 0) {
      const val = ((mm.w + mm.n) / mm.c) * 100;
      data.weeklyData.push({ week: mm.label, value: val });
      data.trend.push(val);
    }
  });

  return (data.weeklyData.length || totalCOGS > 0) ? [data] : [];
}

/* ==================== UNDERPERFORMING WAREHOUSES (EXPORT) ================== */
/**
 * Verilen metrik + domain için hedefi kaçıran depoları döner.
 * Sadece metrik altında kalanları (hedefe ulaşamayanları) getirir.
 * Dönüş: [{ warehouse, domain, value, targetValue?, targetText?, hit:false }, ...]
 */
function getUnderperformingWarehouses(metricName, domain, filters) {
  const metric = OKR_METRICS.find(m => m.name === metricName);
  if (!metric) return [];

  // Depo boyutu olmayan metriklerde listeleme yapma
  if (["Achieve Waste + Waste A&M ratio target", "Decrease the number of pending + rejected pallet backlog (Ops)"].includes(metricName)) {
    return [];
  }

  const unit = metric.unit || '';
  const isIncrease = (metricName || '').toLowerCase().includes('increase');

  // Bazı metrikler (GMore özel)
  // GMore'a ÖZEL metrikler (sadece bu isimlerde domain GMore'a zorlanır)
  const GMORE_ONLY_METRICS = [
    "Increase GMore fresh order penetration",
    "Decrease GMore customer quality complaint rate (problematic order) F&V"
  ];

  // Domain seçimini koru; sadece yukarıdaki beyaz listedekilerde GMore'a zorla
  const dom = GMORE_ONLY_METRICS.includes(metricName) ? 'GMore' : _normDomain(domain);

  const sd = getSheetData(metricName);
  if (!sd || !sd.length) return [];

  // Domain hedefi/kişi hedefi (özel vaka hariç)
  const resolvedTarget = resolveTargetValue(dom, filters, metricName);

  // ==> Depo bazında değerleri toplayalım
  const map = {}; // wh -> aggregator

  // yardımcı: wh kaydı al
  const _rec = (wh) => (map[wh] || (map[wh] = { warehouse: wh, domain: dom }));

  if (metricName === "Decrease number of WHs with availability below 85%") {
    // En son haftayı bul (ilgili domain için), sadece o haftadaki <85% depolar listelensin
    const weekSet = new Set();
    sd.forEach(r => {
      const wh = String(r[0] || '').trim();
      const d = _normDomain(r[1]);
      const w = String(r[2] || '').trim();
      if (d !== dom) return;
      if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      weekSet.add(w);
    });
    const weeks = Array.from(weekSet).sort((a, b) => _weekNum(a) - _weekNum(b));
    if (!weeks.length) return [];
    const latest = weeks[weeks.length - 1];

    const rows = [];
    sd.forEach(r => {
      const wh = String(r[0] || '').trim();
      const d = _normDomain(r[1]);
      const w = String(r[2] || '').trim();
      const avail = +r[3] || 0, total = +r[4] || 0;
      if (d !== dom) return;
      if (!applyManagerFilters(wh, d, filters)) return;
      if (w !== latest) return;
      if (total <= 0) return;
      const ratio = (avail / total) * 100;
      const under = ratio < 85;
      if (under) {
        rows.push({
          warehouse: wh, domain: d, value: +ratio.toFixed(2),
          targetText: "≥85%", hit: false
        });
      }
    });
    // Kötüden iyiye sırala
    rows.sort((a, b) => a.value - b.value);
    return rows;
  }

  // Genel şablon: haftalar arası birikimli (W27+ veya W31+ bazı metriklerde)
  const useW31 = (metricName === "Decrease missed order ratio" || metricName === "Decrease lateness percent");

  const add = (wh, obj) => {
    const rec = _rec(wh);
    Object.keys(obj).forEach(k => rec[k] = (rec[k] || 0) + obj[k]);
    map[wh] = rec;
  };

  if (metricName === "Decrease lateness percent") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const late = +r[3] || 0, ord = +r[4] || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (useW31 ? !_isWeek31Plus(w) : !_isWeek27Plus(w)) return;
      add(wh, { late, ord });
    });
  } else if (metricName === "Increase stock accuracy in dark stores") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const A = +r[3] || 0, B = +r[4] || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { A, B });
    });
  } else if (metricName === "Increase throughput for G10 and GMore") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const orders = +r[3] || 0, hrs = +r[4] || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { orders, hrs });
    });
  } else if (metricName === "Decrease missed order ratio") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const del = +r[3] || 0, miss = +r[4] || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (useW31 ? !_isWeek31Plus(w) : !_isWeek27Plus(w)) return;
      add(wh, { del, miss });
    });
  } else if (metricName === "Decrease the problematic order ratio (G10 and GMore)") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const ord = +r[3] || 0, fb = +r[4] || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { ord, fb });
    });
  } else if (metricName === "Decrease non-agreed / problematic shipment ratios") {
    sd.forEach(r => {
      const w = String(r[0] || '').trim(), wh = String(r[1] || '').trim(), d = _normDomain(r[2]);
      const prob = _toNum(r[3]) || 0, total = _toNum(r[4]) || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { prob, total });
    });
  } else if (metricName === "Achieve LMD variable cost per order target for G10 and GMore (Courier + Fleet)") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const courier = +r[3] || 0, fleet = +r[4] || 0, ord = +r[5] || 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { cost: (courier + fleet), ord });
    });
  } else if (metricName === "Increase Franchise Satisfaction Score") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const score = _toNum(r[3]) ?? 0;
      if (d !== dom) return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { sum: score, cnt: 1 });
    });
  } else if (metricName === "Increase GMore fresh order penetration") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const other = +r[3] || 0, fv = +r[4] || 0;
      if (d !== 'GMore') return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { other, fv });
    });
  } else if (metricName === "Decrease GMore customer quality complaint rate (problematic order) F&V") {
    sd.forEach(r => {
      const wh = String(r[0] || '').trim(), d = _normDomain(r[1]), w = String(r[2] || '').trim();
      const ord = +r[3] || 0, fb = +r[4] || 0;
      if (d !== 'GMore') return; if (!applyManagerFilters(wh, d, filters)) return;
      if (!_isWeek27Plus(w)) return;
      add(wh, { ord, fb });
    });
  }

  // Depo değerini hesapla
  const rows = [];
  Object.values(map).forEach(rec => {
    let value = null, hit = false, targetText = null, targetValue = resolvedTarget;

    if (metricName === "Decrease lateness percent") {
      value = rec.ord > 0 ? (rec.late / rec.ord) * 100 : 0;
      hit = value <= (resolvedTarget ?? 0);
    } else if (metricName === "Increase stock accuracy in dark stores") {
      value = rec.A > 0 ? (rec.B / rec.A) * 100 : 0;
      hit = resolvedTarget != null ? (value >= resolvedTarget) : false;
    } else if (metricName === "Increase throughput for G10 and GMore") {
      value = rec.hrs > 0 ? (rec.orders / rec.hrs) : 0;
      hit = resolvedTarget != null ? (value >= resolvedTarget) : false;
    } else if (metricName === "Decrease missed order ratio") {
      const total = (rec.del || 0) + (rec.miss || 0);
      value = total > 0 ? (rec.miss / total) * 100 : 0;
      hit = value <= (resolvedTarget ?? 0);
    } else if (metricName === "Decrease the problematic order ratio (G10 and GMore)") {
      value = rec.ord > 0 ? (rec.fb / rec.ord) * 100 : 0;
      hit = value <= (resolvedTarget ?? 0);
    } else if (metricName === "Decrease non-agreed / problematic shipment ratios") {
      value = rec.total > 0 ? (rec.prob / rec.total) * 100 : 0;
      hit = value <= (resolvedTarget ?? 0);
    } else if (metricName === "Achieve LMD variable cost per order target for G10 and GMore (Courier + Fleet)") {
      value = rec.ord > 0 ? (rec.cost / rec.ord) : 0;
      hit = resolvedTarget != null ? (value <= resolvedTarget) : false;
    } else if (metricName === "Increase Franchise Satisfaction Score") {
      value = rec.cnt > 0 ? (rec.sum / rec.cnt) : 0;
      hit = resolvedTarget != null ? (value >= resolvedTarget) : false;
    } else if (metricName === "Increase GMore fresh order penetration") {
      const tot = (rec.other || 0) + (rec.fv || 0);
      value = tot > 0 ? (rec.fv / tot) * 100 : 0;
      hit = resolvedTarget != null ? (value >= resolvedTarget) : false;
    } else if (metricName === "Decrease GMore customer quality complaint rate (problematic order) F&V") {
      value = rec.ord > 0 ? (rec.fb / rec.ord) * 100 : 0;
      hit = value <= (resolvedTarget ?? 0);
    }

    if (value == null) return;

    // Sadece hedefi kaçıran depolar
    if (!hit) {
      rows.push({
        warehouse: rec.warehouse,
        domain: rec.domain,
        value: +(+value).toFixed(2),
        targetValue: resolvedTarget,
        targetText,
        hit
      });
    }
  });

  // Sıralama: "decrease" metriklerde en kötü üstte, "increase" metriklerde en kötü altta olacak şekilde
  if (isIncrease) rows.sort((a, b) => a.value - b.value);
  else rows.sort((a, b) => b.value - a.value);

  return rows;
}

/* ==================== GOOGLE SHEETS EXPORT ================================= */
/**
 * Google Sheets'e formatlı export fonksiyonu - gelişmiş hata yönetimi
 */
function exportToGoogleSheets(metricName, domain, unit, rows) {
  try {
    // Debug log
    console.log('Export başlatıldı:', metricName, domain, rows?.length);
    
    // Input validation
    if (!metricName || !domain) {
      throw new Error('Metrik adı veya domain eksik');
    }
    
    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      throw new Error('Export edilecek veri bulunamadı');
    }
    
    // Quota kontrolü (basit)
    const existingFiles = DriveApp.getFilesByName(`Depo Detayları - ${metricName}`);
    let count = 0;
    while (existingFiles.hasNext()) {
      existingFiles.next();
      count++;
    }
    if (count > 50) {
      throw new Error('Çok fazla dosya oluşturulmuş. IT ile iletişime geçin.');
    }
    
    // 1. Yeni Spreadsheet oluştur
    const timestamp = Utilities.formatDate(new Date(), 'GMT+3', 'dd.MM.yyyy HH:mm');
    const sheetName = `${metricName} - ${domain}`;
    const fileName = `Depo Detayları - ${sheetName} - ${timestamp}`;
    
    console.log('Spreadsheet oluşturuluyor:', fileName);
    const ss = SpreadsheetApp.create(fileName);
    
    if (!ss) {
      throw new Error('Spreadsheet oluşturulamadı. Google Drive izinlerini kontrol edin.');
    }
    
    // 2. Locale'i US olarak ayarla
    ss.setSpreadsheetLocale('en_US');
    
    const sheet = ss.getActiveSheet();
    sheet.setName(sheetName.substring(0, 100));
    
    // 3. Başlık satırını hazırla
    const headers = ['Depo', 'Domain', 'Değer', 'Hedef', 'Durum'];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // 4. Veriyi ekle
    if (rows && rows.length > 0) {
      const dataRows = rows.map(r => [
        r.warehouse || '',
        r.domain || '',
        r.value != null ? r.value : '',
        r.targetText || (r.targetValue != null ? r.targetValue : ''),
        r.hit ? 'Hedefte' : 'Geride'
      ]);
      
      const dataRange = sheet.getRange(2, 1, dataRows.length, headers.length);
      dataRange.setValues(dataRows);
      
      // 5. Sayısal formatlamalar
      if (unit) {
        const valueRange = sheet.getRange(2, 3, dataRows.length, 1);
        if (unit === '%') {
          valueRange.setNumberFormat('0.00"%"');
        } else if (unit === '₺') {
          valueRange.setNumberFormat('"₺"0.00');
        } else if (unit === 'pp') {
          valueRange.setNumberFormat('0.00" pp"');
        } else {
          valueRange.setNumberFormat('0.00');
        }
        
        const targetRange = sheet.getRange(2, 4, dataRows.length, 1);
        const targetValues = dataRows.map(row => row[3]);
        const hasNumericTargets = targetValues.some(val => typeof val === 'number');
        if (hasNumericTargets) {
          if (unit === '%') {
            targetRange.setNumberFormat('0.00"%"');
          } else if (unit === '₺') {
            targetRange.setNumberFormat('"₺"0.00');
          } else if (unit === 'pp') {
            targetRange.setNumberFormat('0.00" pp"');
          } else {
            targetRange.setNumberFormat('0.00');
          }
        }
      }
      
      // 6. Koşullu formatlama
      const statusRange = sheet.getRange(2, 5, dataRows.length, 1);
      const statusRule1 = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Hedefte')
        .setBackground('#D4EDDA')
        .setFontColor('#155724')
        .setRanges([statusRange])
        .build();
      
      const statusRule2 = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Geride')
        .setBackground('#F8D7DA')
        .setFontColor('#721C24')
        .setRanges([statusRange])
        .build();
      
      sheet.setConditionalFormatRules([statusRule1, statusRule2]);
    }
    
    // 7. Genel formatlama
    const totalRows = Math.max(2, (rows?.length || 0) + 1);
    const allDataRange = sheet.getRange(1, 1, totalRows, headers.length);
    
    allDataRange.setFontFamily('Avenir');
    allDataRange.setFontSize(11);
    
    // 8. Başlık formatlaması
    headerRange.setBackground('#5D3EBC');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(12);
    
    // 9. Çerçeveler
    allDataRange.setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
    headerRange.setBorder(true, true, true, true, false, true, '#5D3EBC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // 10. Gridlines kapat
    sheet.setHiddenGridlines(true);
    
    // 11. Kolon genişlikleri
    sheet.autoResizeColumns(1, headers.length);
    const minWidths = [120, 80, 100, 100, 100];
    minWidths.forEach((width, i) => {
      const currentWidth = sheet.getColumnWidth(i + 1);
      if (currentWidth < width) {
        sheet.setColumnWidth(i + 1, width);
      }
    });
    
    // 12. Başlık dondur
    sheet.setFrozenRows(1);
    
    // 13. Zebra striping
    if (rows && rows.length > 1) {
      for (let i = 2; i <= totalRows; i++) {
        if (i % 2 === 0) {
          const rowRange = sheet.getRange(i, 1, 1, headers.length);
          rowRange.setBackground('#F8F9FA');
        }
      }
    }
    
    // 14. Bilgi satırı
    sheet.insertRowBefore(1);
    const infoRange = sheet.getRange(1, 1, 1, headers.length);
    infoRange.merge();
    infoRange.setValue(`${metricName} - ${domain} | Oluşturulma: ${timestamp}`);
    infoRange.setBackground('#FFD300');
    infoRange.setFontColor('#5D3EBC');
    infoRange.setFontWeight('bold');
    infoRange.setFontSize(13);
    infoRange.setHorizontalAlignment('center');
    infoRange.setBorder(true, true, true, true, false, false, '#FFD300', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    sheet.setFrozenRows(2);
    
    // 15. Paylaşım ayarları
    console.log('Paylaşım izinleri ayarlanıyor...');
    const file = DriveApp.getFileById(ss.getId());
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    
    const url = ss.getUrl();
    console.log('Export tamamlandı:', url);
    
    return url;
    
  } catch (error) {
    console.error('Google Sheets export error:', error);
    
    // Detaylı hata mesajı döndür
    let userMessage = 'Sheets export hatası: ';
    if (error.message.includes('quota')) {
      userMessage += 'Günlük kullanım limiti aşıldı.';
    } else if (error.message.includes('permission')) {
      userMessage += 'İzin hatası. Getir hesabınızla giriş yaptığınızdan emin olun.';
    } else if (error.message.includes('Drive')) {
      userMessage += 'Google Drive erişim sorunu.';
    } else {
      userMessage += error.message;
    }
    
    return userMessage; // Hata mesajını string olarak döndür
  }
}

/**
 * Test fonksiyonu - gerektiğinde kullan
 */
function testSheetsExport() {
  const testData = [
    { warehouse: 'Test Depo 1', domain: 'G10', value: 82.5, targetValue: 85, hit: false },
    { warehouse: 'Test Depo 2', domain: 'G10', value: 87.3, targetValue: 85, hit: true },
    { warehouse: 'Test Depo 3', domain: 'GMore', value: 79.1, targetText: '≥85%', hit: false }
  ];
  
  const url = exportToGoogleSheets(
    'Test Metric',
    'G10',
    '%',
    testData
  );
  
  console.log('Test sheet URL:', url);
  return url;
}

/* ==================== 4) DUMMY DATA ======================================= */
function generateDummyData(metric, filters) {
  try {
    const weeks = ['W1', 'W2', 'W3', 'W4'];
    const out = [];
    ['G10', 'GMore'].forEach(dom => {
      if (metric.target[dom] == null) return;

      // Domain filtresini kontrol et (yönetici filtreleri dummy data için uygulanmaz)
      if (filters.domain && filters.domain !== 'all' && filters.domain !== dom) return;

      const isInc = (metric.name || '').toLowerCase().includes('increase');
      const obj = { domain: dom, currentValue: 0, targetValue: metric.target[dom], weeklyData: [], trend: [] };
      let sum = 0, cnt = 0;
      weeks.forEach((w, i) => {
        if (filters.week && filters.week !== 'all' && filters.week !== w) return;
        let val;
        if (!isInc && metric.target[dom] === 0) {
          const start = 10 + (dom === 'GMore' ? 2 : 0);
          val = Math.max(0, Math.round(start - i * 1.5 + (Math.random() - 0.5)));
        } else if (isInc) {
          val = metric.target[dom] * (0.7 + Math.random() * 0.4) + i * 0.05 * metric.target[dom];
        } else {
          val = metric.target[dom] * (1.3 - Math.random() * 0.4) - i * 0.05 * metric.target[dom];
        }
        obj.weeklyData.push({ week: w, value: val });
        obj.trend.push(val);
        sum += val; cnt++;
      });
      if (cnt) obj.currentValue = sum / cnt;
      out.push(obj);
    });
    return out;
  } catch (e) { console.error('generateDummyData error', e); return []; }
}

/* ==================== 5) DATA TOPLAMA (getOKRData) ========================= */
function getOKRData(filters = {}) {
  const data = [];

  OKR_METRICS.forEach(m => {
    const md = { name: m.name, unit: m.unit, target: m.target, behavior: m.behavior || null, data: [], _isDummy: false };

    // [NEW] Throughput, yönetici filtresi varsa gösterme/hesaba katma
    if (m.name === "Increase throughput for G10 and GMore" && _hasManagerFilter(filters)) {
      md.data = [];                 // boş bırak → kart render edilmez, özet hesaplara dahil olmaz
      data.push(md);
      return;                       // sonraki metriğe geç
    }

    try {
      if (m.name === "Decrease number of WHs with availability below 85%") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateWHBelow85Data(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Decrease the number of pending + rejected pallet backlog (Ops)") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculatePalletBacklogOpsData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Achieve LMD variable cost per order target for G10 and GMore (Courier + Fleet)") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateLMDVariableCostData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Decrease the problematic order ratio (G10 and GMore)") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateProblematicOrderRatioData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Increase throughput for G10 and GMore") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateThroughputData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Achieve Waste + Waste A&M ratio target") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateWasteAMRatioData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Decrease missed order ratio") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateMissedOrderData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Decrease lateness percent") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateLatenessData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Decrease non-agreed / problematic shipment ratios") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateNonAgreedShipmentData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Increase stock accuracy in dark stores") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateStockAccuracyData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Increase Franchise Satisfaction Score") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateSatisfactionScoreData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Increase GMore fresh order penetration") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateGMoreFreshOrderData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else if (m.name === "Decrease GMore customer quality complaint rate (problematic order) F&V") {
        const sd = getSheetData(m.name);
        md.data = sd ? calculateFVCustomerComplaintData(sd, filters) : (md._isDummy = true, generateDummyData(m, filters));

      } else {
        md.data = generateDummyData(m, filters);
        md._isDummy = true;
      }
    } catch (e) {
      console.error('Metric error:', m.name, e);
      md.data = generateDummyData(m, filters);
      md._isDummy = true;
    }

    data.push(md);
  });

  return data;
}

/* ==================== 6) FİLTRE OPSİYONLARI ================================ */
function getFilterOptions() {
  const opts = {
    warehouses: ['all'],
    domains: ['all', 'G10', 'GMore'],
    weeks: ['all'],
    sahaYoneticileri: ['all'],
    operasyonMudurleri: ['all'],
    bolgeMudurleri: ['all'] // ← YENİ
  };

  const sheets = [
    "Decrease number of WHs with availability below 85%",
    "Decrease lateness percent",
    "Increase stock accuracy in dark stores",
    "Increase throughput for G10 and GMore",
    "Decrease missed order ratio",
    "Decrease the problematic order ratio (G10 and GMore)",
    "Decrease non-agreed / problematic shipment ratios",              // ← önemli
    "Decrease GMore customer quality complaint rate (problematic order) F&V",
    "Increase GMore fresh order penetration",
    "Decrease the number of pending + rejected pallet backlog (Ops)"
  ];

  const wSet = new Set(), wkSet = new Set();
  const START = new Date(2025, 6, 1);

  sheets.forEach(s => {
    const sd = getSheetData(s);
    if (!sd) return;

    sd.forEach(r => {
      if (s === "Decrease the number of pending + rejected pallet backlog (Ops)") {
        const dt = _parseTRDate(r[1]);
        const resp = String(r[2] || '').trim().toLowerCase();
        if (!dt || dt < START) return;
        if (resp !== 'operasyon') return;
        wkSet.add(_isoWeekLabel(dt));
        return;
      }

      if (s === "Decrease non-agreed / problematic shipment ratios") {
        const wkCell = r[0], whCell = r[1];
        if (whCell) wSet.add(String(whCell));
        if (wkCell) wkSet.add(String(wkCell));
        return;
      }

      if (r[0]) wSet.add(String(r[0])); // warehouse
      if (r[2]) wkSet.add(String(r[2])); // week
    });
  });

  const managers = getWarehouseManagers();
  const sahaSet = new Set(), opsSet = new Set(), bolgeSet = new Set(); // ← YENİ
  Object.values(managers).forEach(info => {
    if (info.sahaYoneticisi) sahaSet.add(info.sahaYoneticisi);
    if (info.operasyonMuduru) opsSet.add(info.operasyonMuduru);
    if (info.bolgeMuduru) bolgeSet.add(info.bolgeMuduru);          // ← YENİ
  });

  if (wSet.size) opts.warehouses = ['all', ...Array.from(wSet).sort()];
  if (wkSet.size) opts.weeks = ['all', ...Array.from(wkSet).sort((a, b) => (parseInt(a.replace(/\D/g, '')) || 0) - (parseInt(b.replace(/\D/g, '')) || 0))];
  if (sahaSet.size) opts.sahaYoneticileri = ['all', ...Array.from(sahaSet).sort()];
  if (opsSet.size) opts.operasyonMudurleri = ['all', ...Array.from(opsSet).sort()];
  if (bolgeSet.size) opts.bolgeMudurleri = ['all', ...Array.from(bolgeSet).sort()]; // ← YENİ

  return opts;
}

/* ==================== 7) PROGRESS ve DASHBOARD ÖZET ======================== */
function _computeProgressServer(isInc, unit, beh, curr, tgt) {
  const clamp = v => Math.max(0, Math.min(100, v));
  const c = +curr, t = +tgt;
  if (!isFinite(c)) return 0;
  if (beh === "decrease_to_zero_numeric") return clamp(100 / (1 + Math.max(0, c)));
  if (isInc) {
    return (isFinite(t) && t > 0) ? clamp((c / t) * 100) : 0;
  } else {
    if (isFinite(t) && t > 0) {
      return clamp(100 * Math.min(1, t / Math.max(c, 1e-9)));
    }
    if (t === 0) return unit === '%' ? clamp(100 - c) : clamp(100 / (1 + Math.max(0, c)));
    return 0;
  }
}

function getWarehouseSummary(filters = {}) {
  const managers = getWarehouseManagers();
  const set = getWarehousesByFilter(filters);
  let total = 0, g10 = 0, gmore = 0;

  set.forEach(wh => {
    total++;
    const dom = managers[wh]?.domain;
    if (dom === 'G10') g10++;
    else if (dom === 'GMore') gmore++;
  });

  // Availability <85% için son hafta özetini kullan
  let belowG10 = 0, belowGMore = 0, lastWeek = null;
  try {
    const sd = getSheetData("Decrease number of WHs with availability below 85%");
    if (sd) {
      const arr = calculateWHBelow85Data(sd, filters); // mevcut mantığı kullanır
      (arr || []).forEach(d => {
        if (!d) return;
        const weeks = (d.weeklyData || []).map(x => x.week);
        if (weeks.length) lastWeek = weeks[weeks.length - 1];
        if (d.domain === 'G10') belowG10 = +d.currentValue || 0;
        if (d.domain === 'GMore') belowGMore = +d.currentValue || 0;
      });
    }
  } catch (e) { }

  return {
    total,
    g10,
    gmore,
    below85: { G10: belowG10, GMore: belowGMore },
    lastWeek
  };
}

// Hedefe ulaştı mı? (server)
function _hitServer(isInc, unit, beh, curr, tgt) {
  const c = +curr, t = +tgt;
  if (!isFinite(c) || !isFinite(t)) return false;
  if (beh === 'decrease_to_zero_numeric') return c <= 0;
  return isInc ? (c >= t) : (c <= t);
}

function getDashboardSummary(filters = {}) {
  const data = getOKRData(filters);
  const sum = { totalOKRs: 0, onTrack: 0, atRisk: 0, offTrack: 0, avgProgress: 0 };

  let tot = 0, cnt = 0;

  data.forEach(m => {
    if (m._isDummy) return; // dummy metrikleri dahil etme
    const isInc = (m.name || '').toLowerCase().includes('increase');
    const unit = m.unit || '';
    const beh = m.behavior || null;

    (m.data || []).forEach(d => {
      if (d.targetValue == null) return;

      sum.totalOKRs++;

      // Ortalama ilerleme için yüzdelik progress'i koruyoruz
      const p = _computeProgressServer(isInc, unit, beh, d.currentValue, d.targetValue);
      tot += p; cnt++;

      // Sınıflandırma: sadece Hedefte / Geride
      const hit = _hitServer(isInc, unit, beh, d.currentValue, d.targetValue);
      if (hit) sum.onTrack++; else sum.offTrack++;
    });
  });

  if (cnt) sum.avgProgress = tot / cnt;
  sum.atRisk = 0; // risk kaldırıldı
  return sum;
}
