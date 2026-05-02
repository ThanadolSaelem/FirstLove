// ============================================================
// FL_DataStore.gs — Write Parsed Data to Google Sheets
// ============================================================

/**
 * Setup output sheets (headers) ถ้ายังไม่มี
 */
function FL_setupSheets() {
  const ss = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);

  Object.entries(FL_SHEETS).forEach(([key, sheetName]) => {
    if (!ss.getSheetByName(sheetName)) {
      const sheet   = ss.insertSheet(sheetName);
      const headers = FL_HEADERS[key];
      if (headers) {
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#6B46C1')
          .setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }
    }
  });

  Logger.log('FL_setupSheets: เสร็จแล้ว');
}

// ─── Duplicate Check ─────────────────────────────────────────

function FL_isFileProcessed(fileId) {
  const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.PROCESSED_FILES);
  if (!sheet) return false;
  return sheet.getDataRange().getValues().some(row => row[0] === fileId);
}

/**
 * @param {string} fileId
 * @param {string} fileName
 * @param {string} platform   - 'shopee' | 'tiktok' | 'lazada'
 * @param {string} fileType   - 'income' | 'orders' | 'live'
 * @param {string} monthKey   - 'YYYY-MM'
 */
function FL_markFileProcessed(fileId, fileName, platform, fileType, monthKey) {
  const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.PROCESSED_FILES);
  if (!sheet) return;
  // Prefix "'" forces Sheets to store as text — prevents auto-conversion to Date
  sheet.appendRow([fileId, fileName, platform, fileType, "'" + monthKey, new Date()]);
}

// ─── Write Revenue Summary ───────────────────────────────────

function FL_writeMonthlySummary(data) {
  const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.MONTHLY_SUMMARY);
  if (!sheet) throw new Error('ไม่พบ sheet monthly_summary — รัน Setup ก่อน');

  const existing  = sheet.getDataRange().getValues();
  const duplicate = existing.some(row =>
    row[0] === data.monthKey && row[1] === data.platform
  );
  if (duplicate) {
    Logger.log(`FL_writeMonthlySummary: ข้าม (มีอยู่แล้ว) ${data.monthKey} / ${data.platform}`);
    return;
  }

  sheet.appendRow([
    "'" + data.monthKey,  // Prefix forces Sheets to store as text string
    data.platform,
    data.gross            || 0,
    data.seller_discount  || 0,
    data.net_revenue      || 0,
    data.platform_fees    || 0,
    data.transferred      || 0,
    new Date(),
    data.sourceFile       || '',
  ]);
  Logger.log(`FL_writeMonthlySummary: เขียน ${data.monthKey} / ${data.platform} สำเร็จ`);
}

// ─── Write SKU Summary ───────────────────────────────────────

function FL_writeSkuMonthly(data) {
  const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
  if (!sheet) throw new Error('ไม่พบ sheet sku_monthly — รัน Setup ก่อน');

  const NC   = FL_HEADERS.SKU_MONTHLY.length;  // 9 cols (includes year)
  const now  = new Date();
  const mk   = String(data.monthKey);
  const pl   = String(data.platform).toLowerCase();
  const year = mk.substring(0, 4);

  // ── 1. Read all data rows in ONE batch call ──────────────────
  const lr = sheet.getLastRow();
  const existing = lr > 1 ? sheet.getRange(2, 1, lr - 1, NC).getValues() : [];

  // ── 2. Dedup: drop matching month+platform; backfill year on old rows ──
  const kept = existing
    .filter(r => !(String(r[0]).replace(/^'/, '') === mk && String(r[1]).toLowerCase() === pl))
    .map(r => { if (!r[8]) r[8] = String(r[0]).replace(/^'/, '').substring(0, 4); return r; });

  // ── 3. Build new rows ────────────────────────────────────────
  const newRows = data.skus.map(sku => [
    mk, data.platform, sku.skuRef, sku.category,
    sku.units, sku.revenue, now, data.sourceFile || '', year,
  ]);

  // ── 4. Rewrite data area in ONE batch call ───────────────────
  if (lr > 1) sheet.getRange(2, 1, lr - 1, NC).clearContent();
  const total = kept.length + newRows.length;
  if (total > 0) {
    sheet.getRange(2, 1, total, 1).setNumberFormat('@');  // force month_key as text
    sheet.getRange(2, 1, total, NC).setValues([...kept, ...newRows]);
  }

  Logger.log(`FL_writeSkuMonthly: +${newRows.length} rows → ${mk}/${data.platform} (total ${total})`);
}

/**
 * ONE-TIME: เติม year column ให้ rows เก่าที่ยังไม่มี
 * รันครั้งเดียวจาก Apps Script editor
 */
function FL_backfillSkuMonthlyYear() {
  const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
  if (!sheet) { Logger.log('ไม่พบ sheet sku_monthly'); return; }

  // Ensure year header exists at col 9
  const hdr = sheet.getRange(1, 1, 1, 9).getValues()[0];
  if (hdr[8] !== 'year') {
    sheet.getRange(1, 9).setValue('year')
         .setFontWeight('bold').setBackground('#6B46C1').setFontColor('#FFFFFF');
  }

  const lr = sheet.getLastRow();
  if (lr < 2) { Logger.log('FL_backfillSkuMonthlyYear: ไม่มีข้อมูล'); return; }

  const rows = sheet.getRange(2, 1, lr - 1, 9).getValues();
  const updates = rows.map(r => [r[8] || String(r[0]).replace(/^'/, '').substring(0, 4)]);
  sheet.getRange(2, 9, lr - 1, 1).setNumberFormat('@').setValues(updates);
  Logger.log(`FL_backfillSkuMonthlyYear: อัปเดต ${lr - 1} rows เสร็จ`);
}

// ─── Stock Tracking (REQ-05) ─────────────────────────────────

/**
 * บันทึก stock เข้า — เรียกจาก Index.html → form "บันทึก Stock เข้า"
 * @param {Object} entry { date:string|Date, sku_ref:string, quantity:number, note:string }
 * @returns {Object} { success, message }
 */
function FL_addStockIn(entry) {
  try {
    const cfg = FL_getConfig();
    if (!cfg.OUTPUT_SHEET_ID) return { success: false, message: 'ยังไม่ได้ตั้งค่า Output Sheet ID' };
    const ss = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);
    let sheet = ss.getSheetByName(FL_SHEETS.STOCK_IN);
    if (!sheet) {
      sheet = ss.insertSheet(FL_SHEETS.STOCK_IN);
      sheet.appendRow(FL_HEADERS.STOCK_IN);
      sheet.getRange(1, 1, 1, FL_HEADERS.STOCK_IN.length)
        .setFontWeight('bold').setBackground('#6B46C1').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    const sku = (entry && entry.sku_ref || '').toString().trim();
    const qty = parseFloat(entry && entry.quantity);
    if (!sku) return { success: false, message: 'กรุณาเลือก SKU' };
    if (!qty || qty <= 0) return { success: false, message: 'จำนวนต้องมากกว่า 0' };
    const date = entry.date ? new Date(entry.date) : new Date();
    sheet.appendRow([date, sku, qty, (entry.note || '').toString(), new Date()]);
    FL_clearDashboardCache();
    return { success: true, message: `บันทึก stock เข้า ${qty} ชิ้น (${sku}) สำเร็จ` };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * คำนวณ stock คงเหลือต่อ SKU = SUM(stock_in.quantity)
 *   − SUM(sku_monthly.units_sold ตั้งแต่วันที่ stock_in แรกของ SKU นั้น)
 * @returns {Array<{sku, displayName, totalIn, sold, remain, status, threshold}>}
 */
function FL_getStockStatus() {
  try {
    const cfg = FL_getConfig();
    if (!cfg.OUTPUT_SHEET_ID) return [];
    const threshold = parseFloat(cfg.STOCK_THRESHOLD) || 50;
    const ss = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);

    // ── 1) Aggregate stock_in by SKU + earliest date ─────────────
    const stSheet = ss.getSheetByName(FL_SHEETS.STOCK_IN);
    const inMap = {};   // sku → { totalIn, firstDate }
    if (stSheet) {
      const stData = stSheet.getDataRange().getValues();
      for (let i = 1; i < stData.length; i++) {
        const r = stData[i];
        const date = r[0] instanceof Date ? r[0] : new Date(r[0]);
        const sku  = (r[1] || '').toString().trim();
        const qty  = parseFloat(r[2]) || 0;
        if (!sku || !qty) continue;
        if (!inMap[sku]) inMap[sku] = { totalIn: 0, firstDate: date };
        inMap[sku].totalIn  += qty;
        if (date < inMap[sku].firstDate) inMap[sku].firstDate = date;
      }
    }

    // ── 2) Aggregate sku_monthly.units_sold from firstDate onward ─
    const skSheet = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
    const soldMap = {}; // sku (UPPER) → units sold from firstDate
    if (skSheet) {
      const data = skSheet.getDataRange().getValues();
      const hdr  = data[0] || [];
      const iMonth = 0;
      const iSku   = hdr.indexOf('sku_ref');
      const iUnits = hdr.indexOf('units_sold');
      for (let i = 1; i < data.length; i++) {
        const r  = data[i];
        const sku = (r[iSku] || '').toString().trim();
        if (!sku) continue;
        const skuKey = Object.keys(inMap).find(k => k.toUpperCase() === sku.toUpperCase());
        if (!skuKey) continue;
        const firstDate = inMap[skuKey].firstDate;
        // month_key '2026-01' → check if month >= firstDate's month
        let mk = r[iMonth];
        if (mk instanceof Date) mk = Utilities.formatDate(mk, Session.getScriptTimeZone(), 'yyyy-MM');
        mk = String(mk).replace(/^'/, '');
        const [yy, mm] = mk.split('-').map(Number);
        if (!yy || !mm) continue;
        const rowMonthEnd = new Date(yy, mm, 0); // last day of that month
        // Count if rowMonth's last day >= firstDate
        if (rowMonthEnd < firstDate) continue;
        soldMap[skuKey] = (soldMap[skuKey] || 0) + (parseFloat(r[iUnits]) || 0);
      }
    }

    // ── 3) Build result ──────────────────────────────────────────
    const dnMap = typeof FL_getDisplayNameMap === 'function' ? FL_getDisplayNameMap() : {};
    const metaMap = typeof FL_getSkuMetadataMap === 'function' ? FL_getSkuMetadataMap() : {};
    const out = Object.keys(inMap).map(sku => {
      const totalIn = inMap[sku].totalIn;
      const sold    = soldMap[sku] || 0;
      const remain  = totalIn - sold;
      let status = 'ok';
      if (remain <= 0) status = 'out';
      else if (remain <= threshold) status = 'low';
      const meta = metaMap[sku.toUpperCase()] || {};
      const displayName = dnMap[sku.toUpperCase()] || meta.displayName || sku;
      return { sku, displayName, totalIn, sold, remain, status, threshold };
    }).sort((a, b) => a.remain - b.remain);
    return out;
  } catch (e) {
    Logger.log('FL_getStockStatus error: ' + e);
    return [];
  }
}

/**
 * Return SKU dropdown options for stock-in form (จาก sku_master).
 * @returns {Array<{sku, displayName}>}
 */
function FL_getSkuOptions() {
  try {
    const cfg = FL_getConfig();
    if (!cfg.OUTPUT_SHEET_ID) return [];
    const ss = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);
    const sheet = ss.getSheetByName(FL_SHEETS.SKU_MASTER);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const hdr  = data[0] || [];
    const iSku  = hdr.indexOf('sku_ref');
    const iName = hdr.indexOf('display_name');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const sku = (data[i][iSku] || '').toString().trim();
      if (!sku) continue;
      const name = iName >= 0 ? (data[i][iName] || '').toString().trim() : sku;
      out.push({ sku, displayName: name || sku });
    }
    return out.sort((a, b) => a.sku.localeCompare(b.sku));
  } catch (e) {
    Logger.log('FL_getSkuOptions error: ' + e);
    return [];
  }
}

// ─── One-shot Repair: TikTok Income month_key (REQ-06) ───────
/**
 * เครื่องมือซ่อมข้อมูลเก่าที่ import ก่อน REQ-06 fix
 * Re-parse ไฟล์ TikTok income ทุกไฟล์ใน processed_files แล้วเทียบกับแถวใน
 * monthly_summary — ถ้า month_key ไม่ตรง จะ:
 *   - ลบแถวเดิม (month_key เก่า)
 *   - merge เข้าแถว month_key ที่ถูกต้อง (ถ้ามี) หรือ append ใหม่
 *
 * รันครั้งเดียวจาก Apps Script editor: FL_repairTikTokIncomeMonthKeys
 * @returns {{ scanned, repaired, errors }}
 */
function FL_repairTikTokIncomeMonthKeys() {
  const cfg = FL_getConfig();
  const ss  = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);
  const pfSheet = ss.getSheetByName(FL_SHEETS.PROCESSED_FILES);
  const msSheet = ss.getSheetByName(FL_SHEETS.MONTHLY_SUMMARY);
  if (!pfSheet || !msSheet) return { scanned: 0, repaired: 0, errors: ['ไม่พบ sheet processed_files หรือ monthly_summary'] };

  const pfData = pfSheet.getDataRange().getValues();
  const msData = msSheet.getDataRange().getValues();
  const msHdr  = msData[0];
  const iMonth = 0, iPlat = 1, iSrc = msHdr.indexOf('source_file');

  // Index monthly_summary by (month_key, platform) → row index (1-based)
  const msIdx = {};
  for (let i = 1; i < msData.length; i++) {
    const mk  = String(msData[i][iMonth]).replace(/^'/, '');
    const plt = String(msData[i][iPlat]).toLowerCase();
    msIdx[`${mk}|${plt}`] = i + 1; // 1-based row in sheet
  }

  let scanned = 0, repaired = 0;
  const errors = [];
  const rowsToDelete = []; // sheet row numbers (descending)

  for (let i = 1; i < pfData.length; i++) {
    const fileId   = pfData[i][0];
    const fileName = pfData[i][1];
    const platform = String(pfData[i][2]).toLowerCase();
    const fileType = String(pfData[i][3]).toLowerCase();
    const oldMk    = String(pfData[i][4]).replace(/^'/, '');
    if (platform !== 'tiktok' || fileType !== 'income') continue;
    scanned++;

    let driveFile;
    try { driveFile = DriveApp.getFileById(fileId); }
    catch (e) { errors.push(`${fileName}: เปิดไฟล์ไม่ได้ — ${e}`); continue; }

    let parsed;
    try { parsed = FL_parseTikTokIncome(driveFile); }
    catch (e) { errors.push(`${fileName}: parse ไม่ได้ — ${e}`); continue; }

    const newMk = parsed.monthKey;
    if (newMk === oldMk) continue; // already correct

    // Found mismatch — repair
    const oldKey = `${oldMk}|tiktok`;
    const newKey = `${newMk}|tiktok`;
    const oldRow = msIdx[oldKey];
    if (!oldRow) continue;

    if (msIdx[newKey]) {
      // newMk row already exists → just delete oldMk (assume correct row preferred)
      rowsToDelete.push(oldRow);
    } else {
      // No row at newMk → rewrite oldMk's month_key in place
      msSheet.getRange(oldRow, 1).setNumberFormat('@').setValue("'" + newMk);
      msIdx[newKey] = oldRow;
      delete msIdx[oldKey];
    }

    // Update processed_files month_key
    pfSheet.getRange(i + 1, 5).setNumberFormat('@').setValue("'" + newMk);
    repaired++;
  }

  // Delete rows in descending order to avoid index shift
  rowsToDelete.sort((a, b) => b - a).forEach(r => msSheet.deleteRow(r));

  if (repaired > 0) FL_clearDashboardCache();
  Logger.log(`FL_repairTikTokIncomeMonthKeys: scanned=${scanned}, repaired=${repaired}, errors=${errors.length}`);
  return { scanned, repaired, errors };
}

// ─── Ad Spend (ค่าโฆษณา) ─────────────────────────────────────

function _FL_adSheet(create) {
  const cfg = FL_getConfig();
  if (!cfg.OUTPUT_SHEET_ID) return null;
  const ss = SpreadsheetApp.openById(cfg.OUTPUT_SHEET_ID);
  let sheet = ss.getSheetByName(FL_SHEETS.AD_SPEND);
  if (!sheet && create) {
    sheet = ss.insertSheet(FL_SHEETS.AD_SPEND);
    sheet.appendRow(FL_HEADERS.AD_SPEND);
    sheet.getRange(1, 1, 1, FL_HEADERS.AD_SPEND.length)
      .setFontWeight('bold').setBackground('#6B46C1').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * บันทึกข้อมูลโฆษณาแบบแยก platform/type รายเดือน
 * entries = [{platform, ad_type, ad_amount, sales_amount}, ...]
 */
function FL_saveAdSpendDetail(monthKey, entries) {
  try {
    const mk = String(monthKey || '').trim();
    if (!mk || !/^\d{4}-\d{2}$/.test(mk)) return { success: false, message: 'month_key ไม่ถูกต้อง' };
    const sheet = _FL_adSheet(true);
    if (!sheet) return { success: false, message: 'ยังไม่ได้ตั้งค่า Output Sheet ID' };

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).replace(/^'/, '') === mk) sheet.deleteRow(i + 1);
    }
    const now = new Date();
    (entries || []).forEach(e => {
      const adAmt = parseFloat(e.ad_amount) || 0;
      const saleAmt = parseFloat(e.sales_amount) || 0;
      if (adAmt === 0 && saleAmt === 0) return;
      sheet.appendRow(["'" + mk, e.platform || '', e.ad_type || '', adAmt, saleAmt, now]);
    });
    FL_clearDashboardCache();
    return { success: true, message: `บันทึกข้อมูลโฆษณา ${mk} สำเร็จ` };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/** ดึงข้อมูลโฆษณาแบบ detail รายเดือน → [{platform, ad_type, ad_amount, sales_amount}] */
function FL_getAdSpendDetail(monthKey) {
  try {
    const sheet = _FL_adSheet(false);
    if (!sheet) return [];
    const mk = String(monthKey || '').trim();
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).replace(/^'/, '') !== mk) continue;
      if (data[i].length >= 6) {
        result.push({ platform: data[i][1], ad_type: data[i][2],
                      ad_amount: parseFloat(data[i][3]) || 0, sales_amount: parseFloat(data[i][4]) || 0 });
      }
    }
    return result;
  } catch (e) { Logger.log('FL_getAdSpendDetail error: ' + e); return []; }
}

/** ดึงข้อมูลโฆษณาทั้งปี → { '2026-01': [{platform,ad_type,ad_amount,sales_amount},...], ... } */
function FL_getAdSpendDetailAnnual(year) {
  try {
    const sheet = _FL_adSheet(false);
    if (!sheet) return {};
    const y = String(year || '');
    const out = {};
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const mk = String(data[i][0]).replace(/^'/, '');
      if (!mk.startsWith(y) || data[i].length < 6) continue;
      if (!out[mk]) out[mk] = [];
      out[mk].push({ platform: data[i][1], ad_type: data[i][2],
                     ad_amount: parseFloat(data[i][3]) || 0, sales_amount: parseFloat(data[i][4]) || 0 });
    }
    return out;
  } catch (e) { Logger.log('FL_getAdSpendDetailAnnual error: ' + e); return {}; }
}

/** ดึงยอดรวม ad_amount ของเดือนนั้น (backward-compat กับ dashboard) */
function FL_getAdSpend(monthKey) {
  return FL_getAdSpendDetail(monthKey).reduce((s, e) => s + e.ad_amount, 0);
}

/** ดึงยอดรวม ad_amount ทั้งปีเป็น { month_key: total } (backward-compat) */
function FL_getAdSpendAnnual(year) {
  const annual = FL_getAdSpendDetailAnnual(year);
  const out = {};
  Object.keys(annual).forEach(mk => {
    out[mk] = annual[mk].reduce((s, e) => s + e.ad_amount, 0);
  });
  return out;
}

/** FL_saveAdSpend — legacy single-total form ยังใช้ได้ (เก็บใต้ platform='total') */
function FL_saveAdSpend(monthKey, amount) {
  return FL_saveAdSpendDetail(monthKey, [
    { platform: 'total', ad_type: '', ad_amount: parseFloat(amount) || 0, sales_amount: 0 }
  ]);
}

// ─── Error Log ───────────────────────────────────────────────

function FL_logError(fileName, platform, error) {
  try {
    const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
    const sheet = ss.getSheetByName(FL_SHEETS.ERROR_LOG);
    if (!sheet) return;
    sheet.appendRow([new Date(), fileName, platform, error.toString()]);
  } catch (e) {
    Logger.log('FL_logError failed: ' + e.toString());
  }
}