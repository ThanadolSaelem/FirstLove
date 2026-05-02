// ============================================================
// FL_Main.gs — Main Orchestrator + Web App
// ─────────────────────────────────────────────────────────────
// Architecture:
//   3 platform folders (SHOPEE / TIKTOK / LAZADA)
//   ↓ FL_scanAndImportAll()
//   ↓ FL_detectAndImport(file, platform)
//       → FL_detectFileType()  ← อ่านแค่ sheet names (เร็ว)
//       → route to correct parser
// ============================================================

// ─── Web App Entry Points ────────────────────────────────────

/**
 * Route URL ?page= to different HTML files
 * ?page=dashboard  → Dashboard_Executive.html
 * ?page=annual     → Dashboard_Annual.html
 * (default/import) → Index.html (import + config)
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'import';
  const pageMap = {
    'dashboard': 'Dashboard_Executive',
    'annual':    'Dashboard_Annual',
    'ad':        'Dashboard_Ad',
    'import':    'Index',
  };
  const file  = pageMap[page] || 'Index';
  const titleMap = {
    'Dashboard_Executive': 'First Love — Executive Dashboard',
    'Dashboard_Annual':    'First Love — Annual Dashboard',
    'Dashboard_Ad':        'First Love — Ad Performance',
    'Index':               'First Love — Data Import',
  };
  return HtmlService
    .createHtmlOutputFromFile(file)
    .setTitle(titleMap[file] || 'First Love')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle POST requests from external API (Vercel)
 */
function doPost(e) {
  try {
    const action = e.pathInfo ? e.pathInfo.replace(/^\//, '') : '';
    if (!action || typeof this[action] !== 'function') {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Function not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    let args = [];
    if (e.postData && e.postData.contents) {
      let parsed = JSON.parse(e.postData.contents);
      if (Array.isArray(parsed)) {
          args = parsed;
      } else {
          args = [parsed];
      }
    }
    
    const result = this[action].apply(this, args);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Called from Index.html via google.script.run
 * action: 'scan' | 'setup'
 */
function FL_handleAction(action) {
  try {
    if (action === 'setup') {
      FL_setupSheets();
      return { success: true, message: 'Setup สำเร็จ — สร้าง sheets ครบแล้ว' };
    }
    if (action === 'scan') {
      const logs = FL_scanAndImportAll();
      return { success: true, message: 'Import เสร็จ', logs };
    }
    return { success: false, message: `ไม่รู้จัก action: ${action}` };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─── Dashboard Data Services ──────────────────────────────────

/**
 * Executive Dashboard: current month summary across all platforms
 * Called from Dashboard_Executive.html via google.script.run
 * @param {string} monthKey  e.g. '2026-01' (optional, defaults to latest)
 */
function FL_getExecutiveDashboardData(monthKey, platform = 'all') {
  try {
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'exec_' + (monthKey || 'latest') + '_' + platform;
    const cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(_) {} }

    const cfg = FL_getConfig();
    if (!cfg.OUTPUT_SHEET_ID) return { error: 'OUTPUT_SHEET_ID ยังไม่ได้กรอก' };

    const ss      = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);
    const msSheet = ss.getSheetByName(FL_SHEETS.MONTHLY_SUMMARY);
    const skSheet = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
    if (!msSheet) return { error: 'ไม่พบ sheet monthly_summary' };

    const msData  = msSheet.getDataRange().getValues();
    const headers = msData[0];
    const rows    = msData.slice(1).filter(r => r[0]).map(r => {
      // Force month_key (col 0) to string — Sheets may return Date objects
      if (r[0] instanceof Date) {
        r[0] = Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM");
      } else {
        r[0] = String(r[0]);
      }
      return r;
    });

    // All available months sorted newest first
    const allMonths = [...new Set(rows.map(r => r[0]))].sort().reverse();
    if (!allMonths.length) return { error: 'ยังไม่มีข้อมูลใน monthly_summary' };

    const isAllMode   = (monthKey === 'all');
    const targetMonth = isAllMode ? 'all' : String(monthKey || allMonths[0]);
    const prevMonth   = isAllMode ? null : FL_getPrevMonthKey(targetMonth);

    const idx = (col) => headers.indexOf(col);
    // sumField: month='all' → sum across all months
    const sumField = (month, platform, field) =>
      rows
        .filter(r => (month === 'all' || r[0] === month) && (!platform || r[1] === platform))
        .reduce((s, r) => s + (parseFloat(r[idx(field)]) || 0), 0);

    // Metadata for Boxes & Commissions (Phase 3)
    const metaMap = typeof FL_getSkuMetadataMap === 'function' ? FL_getSkuMetadataMap() : {};
    const skuSheetRef = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
    const skuRows    = skuSheetRef ? skuSheetRef.getDataRange().getValues() : [];
    const skuHeaders = skuRows[0] || [];
    const skuIdx     = (col) => skuHeaders.indexOf(col);
    const iYearCol   = skuIdx('year');
    const tYearStr   = isAllMode ? null : targetMonth.substring(0, 4);

    // Pre-filter sku rows for the target scope
    const targetSkuRows = skuRows.slice(1).filter(r => {
      if (!isAllMode) {
        if (iYearCol >= 0 && r[iYearCol] && String(r[iYearCol]) !== tYearStr) return false;
      }
      let m = r[0];
      if (m instanceof Date) m = Utilities.formatDate(m, Session.getScriptTimeZone(), 'yyyy-MM');
      m = String(m).replace(/^'/, '');
      return isAllMode || m === targetMonth;
    });

    const boxCounts = {};
    targetSkuRows.forEach(row => {
      const skuRef = (row[skuIdx('sku_ref')] || '').toString().toUpperCase();
      const units  = parseFloat(row[skuIdx('units_sold')]) || 0;
      const meta   = metaMap[skuRef];
      if (meta) {
        const bSize = meta.boxSize;
        if (bSize && bSize !== 'nan') boxCounts[bSize] = (boxCounts[bSize] || 0) + units;
      }
    });

    // Totals for the selected scope
    const curGross = sumField(targetMonth, null, 'gross');
    const curFees  = sumField(targetMonth, null, 'platform_fees');
    const curNet   = sumField(targetMonth, null, 'transferred');

    // MoM% only meaningful for single-month mode
    const prevGross = (!isAllMode && prevMonth) ? sumField(prevMonth, null, 'gross') : 0;
    const prevNet   = (!isAllMode && prevMonth) ? sumField(prevMonth, null, 'transferred') : 0;
    const momGross  = prevGross ? Math.round((curGross - prevGross) / prevGross * 100) : null;
    const momNet    = prevNet   ? Math.round((curNet   - prevNet)   / prevNet   * 100) : null;

    // Per-platform breakdown
    const byPlatform = {};
    ['shopee', 'tiktok', 'lazada'].forEach(p => {
      const g = sumField(targetMonth, p, 'gross');
      const f = sumField(targetMonth, p, 'platform_fees');
      byPlatform[p] = {
        gross:         g,
        platform_fees: f,
        transferred:   sumField(targetMonth, p, 'transferred'),
        feeRate:       g > 0 ? Math.abs(f) / g * 100 : 0,
      };
    });

    // Platform fee trend — last 6 months (always based on real months, not 'all')
    const last6 = allMonths.slice(0, 6).reverse();
    const feeTrend = last6.map(m => ({
      month:  m,
      label:  FL_monthLabel(m),
      shopee: Math.abs(sumField(m, 'shopee', 'platform_fees')),
      tiktok: Math.abs(sumField(m, 'tiktok', 'platform_fees')),
      lazada: Math.abs(sumField(m, 'lazada', 'platform_fees')),
    }));

    // Waterfall
    const curDiscount = sumField(targetMonth, null, 'seller_discount');
    const waterfall = {
      gross:    curGross,
      discount: Math.abs(curDiscount),
      fees:     Math.abs(curFees),
      net:      curNet,
    };

    // Units by Category × Platform (grouped chart) — selected scope
    const iCatCol = skuIdx('category');
    const iSkuCol = skuIdx('sku_ref');
    const unitsByCategoryPlatform = {};
    targetSkuRows.forEach(row => {
      const cat  = (row[iCatCol] || '').toString().trim();
      const plat = (row[skuIdx('platform')] || '').toString().toLowerCase();
      if (!cat || cat === 'เซต' || cat === 'อื่นๆ') return;
      const u = parseFloat(row[skuIdx('units_sold')]) || 0;
      if (!unitsByCategoryPlatform[cat]) {
        unitsByCategoryPlatform[cat] = { shopee: 0, tiktok: 0, lazada: 0, total: 0 };
      }
      unitsByCategoryPlatform[cat].total += u;
      if (['shopee','tiktok','lazada'].includes(plat)) unitsByCategoryPlatform[cat][plat] += u;
    });

    // Ad Spend — sum across all months if isAllMode
    let adSpendAmount = 0;
    if (typeof FL_getAdSpend === 'function' || typeof FL_getAdSpendAnnual === 'function') {
      if (isAllMode) {
        // Sum all recorded ad spend
        if (typeof FL_getAdSpendAnnual === 'function') {
          const allYears = [...new Set(allMonths.map(m => m.substring(0, 4)))];
          allYears.forEach(y => {
            const map = FL_getAdSpendAnnual(y);
            Object.values(map).forEach(v => { adSpendAmount += (parseFloat(v) || 0); });
          });
        }
      } else if (typeof FL_getAdSpend === 'function') {
        adSpendAmount = FL_getAdSpend(targetMonth);
      }
    }
    const adSpend = {
      total: adSpendAmount,
      roas:  adSpendAmount > 0 ? Math.round(curNet / adSpendAmount * 100) / 100 : null,
    };

    // Stock status (REQ-05) — computed once per dashboard load
    const stockStatus = (typeof FL_getStockStatus === 'function') ? FL_getStockStatus() : [];

    // Collect ALL SKU records for the month (grouped by SKU + Platform)
    const skuPerformance = [];
    const skuPerformanceMap = {}; // key: "SKU|platform"
    const dnMap = typeof FL_getDisplayNameMap === 'function' ? FL_getDisplayNameMap() : {};

    targetSkuRows.forEach(row => {
      const sku  = (row[skuIdx('sku_ref')] || '').toString().toUpperCase();
      const plat = (row[skuIdx('platform')] || '').toString().toLowerCase();
      const cat  = row[skuIdx('category')] || 'อื่นๆ';
      const u    = parseFloat(row[skuIdx('units_sold')]) || 0;
      const rev  = parseFloat(row[skuIdx('revenue')])    || 0;

      const key = `${sku}|${plat}`;
      if (!skuPerformanceMap[key]) {
        const meta        = metaMap[sku] || {};
        const displayName = dnMap[sku] || meta.displayName || sku;
        skuPerformanceMap[key] = { sku, displayName, category: cat, platform: plat, units: 0, revenue: 0 };
        skuPerformance.push(skuPerformanceMap[key]);
      }
      skuPerformanceMap[key].units   += u;
      skuPerformanceMap[key].revenue += rev;
    });


    const result = {
      month:         targetMonth,
      prevMonth:     prevMonth,
      allMonths:     allMonths,
      monthlyTarget: parseFloat(cfg.MONTHLY_TARGET) || 2000000,
      stockThreshold: parseFloat(cfg.STOCK_THRESHOLD) || 50,
      summary: {
        gross:         curGross,
        platform_fees: curFees,
        transferred:   curNet,
        logistics:     boxCounts,
        momGross:      momGross,
        momNet:        momNet,
      },
      byPlatform:               byPlatform,
      feeTrend:                 feeTrend,
      waterfall:                waterfall,
      unitsByCategoryPlatform:  unitsByCategoryPlatform,
      skuPerformance:           skuPerformance,
      stockStatus:              stockStatus,
      adSpend:                  adSpend,
      isAllMode:                isAllMode,
    };
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(_) {}
    return result;
  } catch(e) {
    return { error: e.toString() };
  }
}

/**
 * Annual Dashboard: full year data grouped by month
 * Called from Dashboard_Annual.html via google.script.run
 * @param {string|number} year  e.g. '2026'
 */
function FL_getAnnualDashboardData(year) {
  try {
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'annual_' + (year || 'latest');
    const cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(_) {} }

    const cfg = FL_getConfig();
    if (!cfg.OUTPUT_SHEET_ID) return { error: 'OUTPUT_SHEET_ID ยังไม่ได้กรอก' };

    const ss      = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);
    const msSheet = ss.getSheetByName(FL_SHEETS.MONTHLY_SUMMARY);
    const skSheet = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
    if (!msSheet) return { error: 'ไม่พบ sheet monthly_summary' };

    const msData  = msSheet.getDataRange().getValues();
    const headers = msData[0];
    const allRows = msData.slice(1).filter(r => r[0]).map(r => {
      // Force month_key (col 0) to string — Sheets may return Date objects
      if (r[0] instanceof Date) {
        r[0] = Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM");
      } else {
        r[0] = String(r[0]);
      }
      return r;
    });

    // Detect available years
    const allYears   = [...new Set(allRows.map(r => r[0].substring(0, 4)))].sort().reverse();
    const targetYear = String(year || allYears[0]);
    // Default prevYear if no specific compareYear provided
    const prevYear   = String(parseInt(targetYear) - 1);

    const idx = (col) => headers.indexOf(col);
    const rowsForYear = (y) => allRows.filter(r => String(r[0]).startsWith(y));

    // Monthly series (all 12 months, zeros where no data)
    const months  = ['01','02','03','04','05','06','07','08','09','10','11','12'];
    const monthly = months.map(m => {
      const key  = `${targetYear}-${m}`;
      const mRows = allRows.filter(r => r[0] === key);
      return {
        month:  key,
        label:  FL_monthLabel(key),
        gross:  mRows.reduce((s,r) => s+(parseFloat(r[idx('gross')])||0), 0),
        fees:   Math.abs(mRows.reduce((s,r) => s+(parseFloat(r[idx('platform_fees')])||0), 0)),
        net:    mRows.reduce((s,r) => s+(parseFloat(r[idx('transferred')])||0), 0),
        shopee: mRows.filter(r=>r[1]==='shopee').reduce((s,r)=>s+(parseFloat(r[idx('transferred')])||0),0),
        tiktok: mRows.filter(r=>r[1]==='tiktok').reduce((s,r)=>s+(parseFloat(r[idx('transferred')])||0),0),
        lazada: mRows.filter(r=>r[1]==='lazada').reduce((s,r)=>s+(parseFloat(r[idx('transferred')])||0),0),
      };
    });

    // Annual KPIs
    const yearRows    = rowsForYear(targetYear);
    const totalNet    = yearRows.reduce((s,r)=>s+(parseFloat(r[idx('transferred')])||0), 0);
    const totalGross  = yearRows.reduce((s,r)=>s+(parseFloat(r[idx('gross')])||0), 0);
    const totalFees   = Math.abs(yearRows.reduce((s,r)=>s+(parseFloat(r[idx('platform_fees')])||0), 0));
    const filledMonths= monthly.filter(m => m.net > 0).length;
    const avgMonthly  = filledMonths ? Math.round(totalNet / filledMonths) : 0;

    // YoY
    const prevYearRows = rowsForYear(prevYear);
    const prevNet = prevYearRows.reduce((s,r)=>s+(parseFloat(r[idx('transferred')])||0), 0);
    const yoy = prevNet ? Math.round((totalNet - prevNet) / prevNet * 100) : null;

    // Total units sold (annual)
    const metaMap = typeof FL_getSkuMetadataMap === 'function' ? FL_getSkuMetadataMap() : {};
    let totalUnits = 0;
    let skuPerformance  = [];
    let categoryPie     = {};
    let quadrant        = [];
    let marginRanking   = [];
    if (skSheet) {
      const skData    = skSheet.getDataRange().getValues();
      const skHeaders = skData[0];
      const iYr       = skHeaders.indexOf('year');
      const iSku      = skHeaders.indexOf('sku_ref');
      const iCat      = skHeaders.indexOf('category');
      const iUnits    = skHeaders.indexOf('units_sold');
      const iRev      = skHeaders.indexOf('revenue');
      const iPlat     = skHeaders.indexOf('platform');
      const skRows    = skData.slice(1).filter(r =>
        iYr >= 0 ? String(r[iYr]) === targetYear : String(r[0]).startsWith(targetYear)
      );
      const prevSkRows = skData.slice(1).filter(r =>
        iYr >= 0 ? String(r[iYr]) === prevYear : String(r[0]).startsWith(prevYear)
      );
      const skuMap    = {};
      const prevSkuMap = {};
      const dnMap     = typeof FL_getDisplayNameMap === 'function' ? FL_getDisplayNameMap() : {};
      skRows.forEach(r => {
        const sku = r[iSku];
        const cat = r[iCat];
        const u   = parseFloat(r[iUnits]) || 0;
        const rev = parseFloat(r[iRev])    || 0;
        totalUnits += u;
        if (!skuMap[sku]) {
          const meta        = metaMap[(sku || '').toUpperCase()] || metaMap[sku] || {};
          const displayName = dnMap[(sku || '').toUpperCase()] || meta.displayName || sku;
          skuMap[sku] = { sku, displayName, category: cat, units: 0, revenue: 0 };
        }
        skuMap[sku].units   += u;
        skuMap[sku].revenue += rev;
        categoryPie[cat] = (categoryPie[cat] || 0) + u;
      });

      // Monthly units aggregation (for Units Sold bar chart)
      const unitsByMonth = {};
      skRows.forEach(r => {
        const mk = String(r[0]).replace(/^'/, '');
        const u  = parseFloat(r[iUnits]) || 0;
        const p  = (r[iPlat] || '').toString().toLowerCase();
        if (!unitsByMonth[mk]) unitsByMonth[mk] = { total: 0, shopee: 0, tiktok: 0, lazada: 0 };
        unitsByMonth[mk].total += u;
        if (['shopee','tiktok','lazada'].includes(p)) unitsByMonth[mk][p] += u;
      });
      monthly.forEach(m => {
        const um = unitsByMonth[m.month] || {};
        m.units       = um.total  || 0;
        m.shopeeUnits = um.shopee || 0;
        m.tiktokUnits = um.tiktok || 0;
        m.lazadaUnits = um.lazada || 0;
      });

      prevSkRows.forEach(r => {
        const sku = r[iSku];
        const u   = parseFloat(r[iUnits]) || 0;
        prevSkuMap[sku] = (prevSkuMap[sku] || 0) + u;
      });
      skuPerformance = Object.values(skuMap).sort((a,b) => b.units - a.units).slice(0, 10);

      // Product Lifecycle Quadrant: units (x) × YoY growth % (y)
      quadrant = Object.values(skuMap).map(s => {
        const prev = prevSkuMap[s.sku] || 0;
        const growth = prev ? Math.round((s.units - prev) / prev * 100) : (s.units > 0 ? 100 : 0);
        return {
          sku: s.sku, displayName: s.displayName, category: s.category,
          units: s.units, revenue: s.revenue, growth,
        };
      }).filter(s => s.units > 0);

      // Contribution Margin Ranking: revenue − commission per SKU (top 10)
      marginRanking = Object.values(skuMap).map(s => {
        const meta = metaMap[(s.sku || '').toUpperCase()] || metaMap[s.sku] || {};
        const commission = (meta.commission || 0) * s.units;
        const margin = s.revenue - commission;
        const marginPct = s.revenue ? Math.round((margin / s.revenue) * 100) : 0;
        return {
          sku: s.sku, displayName: s.displayName, category: s.category,
          revenue: s.revenue, commission, margin, marginPct, units: s.units,
        };
      }).sort((a,b) => b.margin - a.margin).slice(0, 10);
    }

    // Platform fee distribution + rate
    const feeByPlatform = {};
    ['shopee','tiktok','lazada'].forEach(p => {
      const pRows = yearRows.filter(r => r[1] === p);
      const pFees = Math.abs(pRows.reduce((s,r) => s+(parseFloat(r[idx('platform_fees')])||0), 0));
      const pGross = pRows.reduce((s,r) => s+(parseFloat(r[idx('gross')])||0), 0);
      feeByPlatform[p] = { amount: pFees, rate: pGross > 0 ? Math.round(pFees / pGross * 1000) / 10 : 0 };
    });

    // Ad Spend per month (annual)
    const adSpendMap = (typeof FL_getAdSpendAnnual === 'function') ? FL_getAdSpendAnnual(targetYear) : {};
    monthly.forEach(m => { m.adSpend = adSpendMap[m.month] || 0; });

    const result = {
      year:          targetYear,
      allYears:      allYears,
      kpis: {
        totalNet,
        totalGross,
        totalFees,
        totalUnits,
        avgMonthly,
        yoy,
      },
      monthly,
      skuPerformance,
      categoryPie,
      feeByPlatform,
      quadrant,
      marginRanking,
    };
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(_) {}
    return result;
  } catch(e) {
    return { error: e.toString() };
  }
}

// ─── Shared Helpers ───────────────────────────────────────────

/** Return deployed Web App URL for navigation — returns null if not deployed */
function FL_getWebAppUrl() {
  try {
    var svc = ScriptApp.getService();
    return svc ? svc.getUrl() : null;
  } catch (e) {
    return null;
  }
}

/** Thai month label: '2026-01' → 'ม.ค. 2026' */
function FL_monthLabel(key) {
  key = String(key || '');
  const thaiMonths = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                      'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const parts = key.split('-');
  if (parts.length < 2) return key;
  return `${thaiMonths[parseInt(parts[1])]} ${parts[0]}`;
}

/** Previous month key: '2026-01' → '2025-12' */
function FL_getPrevMonthKey(key) {
  key = String(key || '');
  const parts = key.split('-');
  if (parts.length < 2) return key;
  let y = parseInt(parts[0]);
  let m = parseInt(parts[1]) - 1;
  if (m === 0) { m = 12; y--; }
  return `${y}-${String(m).padStart(2, '0')}`;
}

/** Format number with commas */
function FL_fmt(n) {
  return (n || 0).toLocaleString();
}

// ─── Main Scanner ─────────────────────────────────────────────

// Helper: normalize trigger unit (same options as KC version)
function _FL_normalizeTriggerUnit(unit) {
  var u = String(unit || '').toLowerCase().trim();
  if (u === 'day' || u === 'days' || u === 'd' || u === 'วัน') return 'days';
  if (u === 'hour' || u === 'hours' || u === 'hr' || u === 'hrs' || u === 'h' || u === 'ชม' || u === 'ชม.') return 'hours';
  // 'mins', 'min', 'minutes', 'minute', 'm', 'นาที' → all map to mins
  return 'mins';
}

function _FL_triggerIntervalMs(interval, unit) {
  var iv = parseInt(interval, 10);
  if (isNaN(iv) || iv < 1) iv = 1;
  var u = _FL_normalizeTriggerUnit(unit);
  if (u === 'days')  return iv * 24 * 60 * 60000;
  if (u === 'hours') return iv * 60 * 60000;
  return iv * 60000; // mins — iv อาจเป็น total minutes (เช่น 61 = 1h1m)
}



/**
 * สแกน 3 platform folders หาไฟล์ xlsx ใหม่ แล้ว detect + import
 * @returns {Array.<{file, platform, file_type, status, status_type, message, detail}>}
 */

/**
 * Recursively collect all files of a given MIME type from a folder and all its subfolders.
 * Depth-first; no hardcoded subfolder names.
 * @param {DriveApp.Folder} folder
 * @param {string} mimeType
 * @returns {DriveApp.File[]}
 */
function FL_collectFilesRecursive(folder, mimeType) {
  const results = [];
  const files = folder.getFilesByType(mimeType);
  while (files.hasNext()) results.push(files.next());
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    FL_collectFilesRecursive(subs.next(), mimeType).forEach(f => results.push(f));
  }
  return results;
}

function FL_doScanAndImportAll() {
  const log = [];
  const cfg = FL_getConfig();

  const platformTasks = [
    { platform: 'shopee', folderId: cfg.FOLDERS.SHOPEE },
    { platform: 'tiktok', folderId: cfg.FOLDERS.TIKTOK },
    { platform: 'lazada', folderId: cfg.FOLDERS.LAZADA },
  ];

  platformTasks.forEach(task => {
    if (!task.folderId) {
      log.push({
        file: '-', platform: task.platform, file_type: '-',
        status: 'skipped', status_type: 'skip',
        message: 'skipped', detail: 'Folder ID ยังไม่ได้กรอก',
      });
      return;
    }

    try {
      const folder   = DriveApp.getFolderById(task.folderId);
      const allFiles = FL_collectFilesRecursive(folder, MimeType.MICROSOFT_EXCEL);

      allFiles.forEach(file => {
        const result = FL_detectAndImport(file, task.platform);
        log.push({
          file_name:   file.getName(),
          platform:    task.platform,
          file_type:   result.file_type || '-',
          status:      result.status,
          status_type: result.status_type || result.status,
          message:     result.message,
          detail:      result.detail || '',
        });
      });

      if (allFiles.length === 0) {
        log.push({
          file_name: '-', platform: task.platform, file_type: '-',
          status: 'skipped', status_type: 'skip',
          message: 'skipped', detail: 'ไม่พบไฟล์ xlsx ใน folder (รวม subfolders)',
        });
      }
    } catch (e) {
      log.push({
        file_name: '-', platform: task.platform, file_type: '-',
        status: 'error', status_type: 'error',
        message: 'error', detail: e.toString(),
      });
    }
  });

  return log;
}

// ─── Detect + Import Dispatcher ──────────────────────────────

function FL_detectAndImport(file, platform) {
  if (FL_isFileProcessed(file.getId())) {
    return {
      file_type: '-', status: 'skipped', status_type: 'skip',
      message: 'skipped', detail: 'import แล้ว (ข้ามซ้ำ)',
    };
  }

  let fileType;
  try {
    fileType = FL_detectFileType(file, platform);
  } catch (e) {
    FL_logError(file.getName(), platform, e);
    return {
      file_type: 'unknown', status: 'error', status_type: 'error',
      message: 'error', detail: 'ตรวจ file type ล้มเหลว: ' + e.message,
    };
  }

  if (fileType === 'unknown') {
    const sheetNames = (() => { try { return FL_getSheetNames(file).join(', '); } catch(e) { return '?'; } })();
    FL_logError(file.getName(), platform, new Error(`ไม่รู้จัก file type (sheets: ${sheetNames})`));
    return {
      file_type: 'unknown', status: 'error', status_type: 'error',
      message: 'error',
      detail: `ไม่สามารถระบุ file type ได้ (sheets: ${sheetNames})`,
    };
  }

  try {
    return FL_runParser(file, platform, fileType);
  } catch (e) {
    FL_logError(file.getName(), `${platform}_${fileType}`, e);
    return {
      file_type: fileType, status: 'error', status_type: 'error',
      message: 'error', detail: e.message,
    };
  }
}

function FL_runParser(file, platform, fileType) {
  const name = file.getName();

  if (platform === 'shopee' && fileType === 'income') {
    const data = FL_parseShopeeIncome(file);
    FL_writeMonthlySummary(data);
    FL_markFileProcessed(file.getId(), name, 'shopee', 'income', data.monthKey);
    return { file_type: 'income', status: 'ok', status_type: 'success', message: 'ok',
             detail: `${data.monthKey} | NET: ฿${FL_fmt(data.transferred)}` };
  }

  if (platform === 'shopee' && fileType === 'orders') {
    const data  = FL_parseShopeeOrder(file);
    FL_writeSkuMonthly(data);
    FL_markFileProcessed(file.getId(), name, 'shopee', 'orders', data.monthKey);
    const units = data.skus.reduce((s, x) => s + x.units, 0);
    return { file_type: 'orders', status: 'ok', status_type: 'success', message: 'ok',
             detail: `${data.monthKey} | ${units} units | ${data.skus.length} SKUs` };
  }

  if (platform === 'tiktok' && fileType === 'income') {
    const data = FL_parseTikTokIncome(file);
    FL_writeMonthlySummary(data);
    FL_markFileProcessed(file.getId(), name, 'tiktok', 'income', data.monthKey);
    return { file_type: 'income', status: 'ok', status_type: 'success', message: 'ok',
             detail: `${data.monthKey} | NET: ฿${FL_fmt(data.transferred)}` };
  }

  if (platform === 'tiktok' && fileType === 'orders') {
    const data  = FL_parseTikTokOrder(file);
    FL_writeSkuMonthly(data);
    FL_markFileProcessed(file.getId(), name, 'tiktok', 'orders', data.monthKey);
    const units = data.skus.reduce((s, x) => s + x.units, 0);
    return { file_type: 'orders', status: 'ok', status_type: 'success', message: 'ok',
             detail: `${data.monthKey} | ${units} units | ${data.skus.length} SKUs` };
  }

  // Phase 3 — detect only, not yet implemented
  if (platform === 'tiktok' && fileType === 'live') {
    return { file_type: 'live', status: 'skipped', status_type: 'skip', message: 'skipped',
             detail: 'Live Performance — รอ Phase 3 (Live Commission)' };
  }

  if (platform === 'lazada' && fileType === 'income') {
    const data = FL_parseLazadaIncome(file);
    FL_writeMonthlySummary(data);
    FL_markFileProcessed(file.getId(), name, 'lazada', 'income', data.monthKey);
    return { file_type: 'income', status: 'ok', status_type: 'success', message: 'ok',
             detail: `${data.monthKey} | NET: ฿${FL_fmt(data.transferred)} | ${data.order_count} orders` };
  }

  if (platform === 'lazada' && fileType === 'orders') {
    const data  = FL_parseLazadaOrder(file);
    FL_writeSkuMonthly(data);
    FL_markFileProcessed(file.getId(), name, 'lazada', 'orders', data.monthKey);
    const units = data.skus.reduce((s, x) => s + x.units, 0);
    const warn  = data.warning ? ` ⚠️ ${data.warning}` : '';
    return { file_type: 'orders', status: 'ok', status_type: 'success', message: 'ok',
             detail: `${data.monthKey} | ${units} units | ${data.skus.length} SKUs${warn}` };
  }

  throw new Error(`ไม่มี handler สำหรับ platform="${platform}" fileType="${fileType}"`);
}

// ─── Manual Trigger ──────────────────────────────────────────
function FL_runManual() {
  const results = FL_scanAndImportAll();
  results.forEach(r => {
    Logger.log(`[${r.platform}] [${r.file_type}] ${r.file_name} → ${r.status}: ${r.detail}`);
  });
}

// ─── Auto Trigger (Time-driven) ──────────────────────────────

/**
 * บันทึก interval ลง config แล้ว install trigger ในขั้นตอนเดียว
 * เรียกจาก toggle switch ใน Index.html — ไม่ต้อง save config ก่อน
 * @param {number} interval
 * @param {string} unit  'days' | 'hours' | 'mins'
 */
function FL_installTriggerWithConfig(interval, unit) {
  const cfg = FL_getConfig();
  cfg.triggerInterval = Number(interval) || 1;
  cfg.triggerUnit     = String(unit || 'hours');
  PropertiesService.getScriptProperties()
    .setProperty('FL_CONFIG', JSON.stringify(cfg));
  return FL_installTrigger();  // อ่าน config ที่เพิ่งบันทึก
}

function FL_installTrigger() {
  FL_removeTrigger();
const cfg      = FL_getConfig();
var interval = Number(cfg.triggerInterval) || 5;   // fallback default
var unit     = _FL_normalizeTriggerUnit(cfg.triggerUnit);
var trig     = ScriptApp.newTrigger('FL_scanAndImportAll').timeBased();
var label;

if (unit === 'days') {
  interval = Math.min(7, Math.max(1, interval));
  trig.everyDays(interval);
  label = interval + ' วัน';
} else if (unit === 'hours') {
  interval = Math.min(12, Math.max(1, interval));
  trig.everyHours(interval);
  label = interval + ' ชม.';
} else {
  // mins — iv อาจเป็น total minutes รวม (เช่น 61 = 1h1m)
  // GAS everyMinutes รองรับแค่ 1,5,10,15,30 → ใช้ 1-min trigger + soft guard
  interval = Math.max(1, interval);
  trig.everyMinutes(1);
  var dispH = Math.floor(interval / 60);
  var dispM = interval % 60;
  var parts = [];
  if (dispH > 0) parts.push(dispH + ' ชม.');
  if (dispM > 0) parts.push(dispM + ' นาที');
  label = (parts.length ? parts.join(' ') : '1 นาที') + ' (soft)';
}
trig.create();
Logger.log('FL_installTrigger: ✅ ทุก ' + label);
return { success: true,
         message: 'ติดตั้ง Auto‑Scan (ทุก ' + label + ') เรียบร้อยแล้ว' };
}

function FL_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'FL_scanAndImportAll') ScriptApp.deleteTrigger(t);
  });
  return { success: true, message: 'ยกเลิก Auto-Scan ปิดระบบอัตโนมัติเรียบร้อย' };
}

function FL_getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const isActive = triggers.some(t => t.getHandlerFunction() === 'FL_scanAndImportAll');
  return { isActive };
}

// ─── Guarded wrapper – respects soft interval (ใช้กับ time-driven trigger เท่านั้น)
function FL_scanAndImportAll() {
  // Soft‑interval guard (prevents double run if trigger fires early)
  var cfg   = FL_getConfig();
  var ms    = _FL_triggerIntervalMs(cfg.triggerInterval, cfg.triggerUnit);
  var prop  = PropertiesService.getScriptProperties();
  var last  = parseInt(prop.getProperty('FL_LAST_RUN') || '0', 10);
  var now   = new Date().getTime();

  if (last > 0 && now - last < ms) {
    var remainMin = Math.ceil((ms - (now - last)) / 60000);
    Logger.log('FL_scanAndImportAll: ยังไม่ถึงเวลา (เหลือประมาณ ' + remainMin + ' นาที)');
    return []; // no work
  }
  prop.setProperty('FL_LAST_RUN', String(now));
  return FL_doScanAndImportAll();
}

/**
 * กดปุ่มมือ — ข้าม soft guard + อัปเดต SKU Master อัตโนมัติก่อนสแกน
 * เรียกจาก Index.html → startImport()
 */
function FL_scanNow() {
  PropertiesService.getScriptProperties()
    .setProperty('FL_LAST_RUN', String(new Date().getTime()));
  try { FL_importSkuMasterData(); } catch(e) { Logger.log('SKU master update skipped: ' + e); }
  let result;
  try {
    result = FL_doScanAndImportAll();
  } catch(e) {
    const msg = (e && e.message) ? e.message : String(e);
    Logger.log('FL_scanNow: FL_doScanAndImportAll threw: ' + msg);
    return [{
      file_name: 'System', platform: '-', file_type: '-',
      status: 'error', status_type: 'error',
      message: 'error', detail: 'เกิดข้อผิดพลาดระบบ: ' + msg,
    }];
  }
  try { FL_syncProductNames(); } catch(e) { Logger.log('FL_syncProductNames skipped: ' + e); }
  FL_clearDashboardCache();
  return result;
}

// ─── Cache + Warmup ───────────────────────────────────────────

/** Clear all dashboard caches — called automatically after import.
 *  FIX: dynamic year range (current year ± 3) แทน hardcode เพื่อรองรับปีถัดไป */
function FL_clearDashboardCache() {
  try {
    const cache = CacheService.getScriptCache();
    const keys  = [];
    const cy = new Date().getFullYear();
    const years = [];
    for (let y = cy - 3; y <= cy + 1; y++) years.push(String(y));
    const months = ['latest'];
    years.forEach(y => {
      for (let m = 1; m <= 12; m++) months.push(`${y}-${String(m).padStart(2, '0')}`);
    });
    months.forEach(m => {
      ['all','shopee','tiktok','lazada'].forEach(p => keys.push('exec_' + m + '_' + p));
    });
    [...years, 'latest'].forEach(y => keys.push('annual_' + y));
    cache.removeAll(keys);
  } catch(_) {}
}

/**
 * Keeps V8 runtime warm — run by time-driven trigger every 10 min.
 * Lightweight: just reads config, no Sheets open.
 */
function FL_warmup() {
  try { FL_getConfig(); } catch(_) {}
}

/**
 * Install warmup trigger — run ONCE from Apps Script editor.
 * Apps Script > Run > FL_setupWarmupTrigger
 */
function FL_setupWarmupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'FL_warmup')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('FL_warmup')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('✅ Warmup trigger installed — FL_warmup runs every 10 minutes.');
}

/** Remove warmup trigger */
function FL_removeWarmupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'FL_warmup')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Warmup trigger removed.');
}