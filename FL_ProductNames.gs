// ============================================================
// FL_ProductNames.gs — Product Display Name Management
// ============================================================
// ระบบจัดการชื่อแสดงผลของสินค้า แยกออกจาก sku_master
// เพื่อให้ผู้ใช้แก้ชื่อได้จาก Dashboard UI โดยไม่ต้องแตะ Sheets

// ─── DEV / TEST ONLY ─────────────────────────────────────────
/**
 * รันครั้งเดียวจาก Apps Script editor เพื่อทดสอบระบบ
 * จะเพิ่ม 2 SKU ปลอมเข้า product_names ด้วย status: 'new'
 * หลังทดสอบเสร็จให้รัน FL_testCleanup() เพื่อลบออก
 */
function FL_testInjectNewSkus() {
  const sheet = FL_ensureProductNamesSheet();
  const now   = new Date();
  const testRows = [
    ['TEST_Bundle_Probiotic2+VitC3', FL_autoExtractDisplayName('Bundle_Probiotic2+VitC3'), 'new', now],
    ['TEST_GlutaX9-8',               FL_autoExtractDisplayName('GlutaX9-8'),               'new', now],
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, testRows.length, 4).setValues(testRows);
  Logger.log('FL_testInjectNewSkus: เพิ่ม ' + testRows.length + ' test SKU แล้ว — เปิด dashboard เพื่อดู badge');
}

/** ลบ test rows ออก */
function FL_testCleanup() {
  const sheet = FL_ensureProductNamesSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).startsWith('TEST_')) sheet.deleteRow(i + 1);
  }
  Logger.log('FL_testCleanup: ลบ TEST_ rows ออกแล้ว');
}

/**
 * ONE-TIME FIX — รันครั้งเดียวเพื่อแก้ข้อมูลที่ถูก mark เป็น 'new' ผิดๆ
 * - SKU ที่อยู่ใน FL_SKU_DISPLAY_MAP → เปลี่ยนเป็น 'confirmed'
 * - ลบ rows ซ้ำออก (เก็บ row แรก)
 * - ลบ TEST_ rows
 */
function FL_fixProductNamesStatus() {
  const sheet = FL_ensureProductNamesSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const hdr   = {};
  data[0].forEach((h, i) => { if (h) hdr[String(h).trim()] = i; });
  const iSku  = hdr['sku_ref']      ?? 0;
  const iStat = hdr['status']       ?? 2;
  const iUpd  = hdr['updated_at']   ?? 3;

  const seen  = new Set();
  const now   = new Date();

  // วน reverse เพื่อ deleteRow ได้ถูก index
  for (let i = data.length - 1; i >= 1; i--) {
    const sku   = String(data[i][iSku] || '').trim();
    const upper = sku.toUpperCase();

    // ลบ TEST_ rows
    if (sku.startsWith('TEST_')) { sheet.deleteRow(i + 1); continue; }

    // ลบ duplicate (เก็บ row แรกที่เจอ = index เล็กที่สุด = ลบ index ใหญ่กว่า)
    if (seen.has(upper)) { sheet.deleteRow(i + 1); continue; }
    seen.add(upper);

    // confirm ถ้าอยู่ใน display map แล้ว
    const isKnown = Object.keys(FL_SKU_DISPLAY_MAP).some(k => k.toUpperCase() === upper);
    if (isKnown && String(data[i][iStat]).trim() === 'new') {
      sheet.getRange(i + 1, iStat + 1).setValue('confirmed');
      sheet.getRange(i + 1, iUpd  + 1).setValue(now);
    }
  }
  Logger.log('FL_fixProductNamesStatus: เสร็จ — ลบซ้ำ, ลบ TEST_, confirm SKU ที่รู้จักแล้ว');
}
// ============================================================

const PRODUCT_NAMES_SHEET = 'product_names';

// ─── Sheet Bootstrap ─────────────────────────────────────────

/**
 * ตรวจว่า product_names sheet มีอยู่ ถ้าไม่มีให้สร้างใหม่
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function FL_ensureProductNamesSheet() {
  const ss      = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  let   sheet   = ss.getSheetByName(PRODUCT_NAMES_SHEET);
  const headers = ['sku_ref', 'display_name', 'status', 'updated_at'];

  if (!sheet) {
    sheet = ss.insertSheet(PRODUCT_NAMES_SHEET);
    Logger.log('FL_ensureProductNamesSheet: สร้าง sheet product_names ใหม่');
  }

  // ตรวจว่ามี header อยู่แล้วหรือยัง (row 1 ต้องมี 'sku_ref')
  const firstCell = sheet.getRange(1, 1).getValue().toString().trim();
  if (firstCell !== 'sku_ref') {
    // เพิ่ม header row ที่ตำแหน่ง 1 (ดัน data เดิมลงมา)
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('FL_ensureProductNamesSheet: เพิ่ม header row แล้ว');
  }

  // Style header
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold')
       .setBackground('#6B46C1')
       .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 180);

  return sheet;
}

// ─── Read ─────────────────────────────────────────────────────

/**
 * อ่านทุก entry ใน product_names sheet
 * @returns {Array<{sku:string, displayName:string, status:string, updatedAt:string}>}
 */
function FL_getProductNames() {
  try {
    const ss    = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
    const sheet = ss.getSheetByName(PRODUCT_NAMES_SHEET);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const hdr = {};
    data[0].forEach((h, i) => { if (h) hdr[String(h).trim()] = i; });

    const iSku  = hdr['sku_ref']      ?? 0;
    const iDisp = hdr['display_name'] ?? 1;
    const iStat = hdr['status']       ?? 2;
    const iUpd  = hdr['updated_at']   ?? 3;

    return data.slice(1)
      .filter(r => (r[iSku] || '').toString().trim())
      .map(r => ({
        sku:         String(r[iSku]  || '').trim(),
        displayName: String(r[iDisp] || '').trim(),
        status:      String(r[iStat] || 'confirmed').trim(),
        updatedAt:   r[iUpd] ? String(r[iUpd]) : '',
      }));
  } catch (e) {
    Logger.log('FL_getProductNames error: ' + e);
    return [];
  }
}

/**
 * คืน entries ที่ยังไม่ได้ confirm (status: 'new')
 * เรียกจาก Dashboard UI เพื่อแสดง badge + modal
 * @returns {Array<{sku:string, displayName:string, status:string}>}
 */
function FL_getNewProducts() {
  return FL_getProductNames().filter(p => p.status === 'new');
}

/**
 * สร้าง lookup map  sku.toUpperCase() → displayName
 * Priority: product_names sheet > FL_SKU_DISPLAY_MAP hardcoded > raw sku
 * @returns {Object.<string,string>}
 */
function FL_getDisplayNameMap() {
  const map = {};
  // ชั้นต่ำสุด: hardcoded map ใน FL_Config.gs
  Object.entries(FL_SKU_DISPLAY_MAP).forEach(([k, v]) => {
    map[k.toUpperCase()] = v;
  });
  // ชั้นสูงกว่า: ที่ user ตั้งใน product_names sheet (ทั้ง confirmed + new)
  FL_getProductNames()
    .filter(p => p.displayName)
    .forEach(p => { map[p.sku.toUpperCase()] = p.displayName; });
  return map;
}

// ─── Auto-Extract ────────────────────────────────────────────

/**
 * สกัดชื่อไทยเบื้องต้นจาก SKU ref string
 * ใช้ FL_SKU_DISPLAY_MAP ก่อน ถ้าไม่เจอค่อย parse pattern
 * @param {string} skuRef
 * @returns {string}
 */
function FL_autoExtractDisplayName(skuRef) {
  if (!skuRef) return '';

  // 1. ตรวจ display map โดยตรง
  if (FL_SKU_DISPLAY_MAP[skuRef]) return FL_SKU_DISPLAY_MAP[skuRef];
  const upper   = skuRef.toUpperCase();
  const mapKey  = Object.keys(FL_SKU_DISPLAY_MAP).find(k => k.toUpperCase() === upper);
  if (mapKey) return FL_SKU_DISPLAY_MAP[mapKey];

  // 2. Pattern tokenization
  const name = skuRef
    .replace(/^Bundle_?/i,              '')
    .replace(/GlutaX9-?(\d+)/gi,        'กลูต้า $1 กระปุก')
    .replace(/Gluta(\d+)/gi,            'กลูต้า $1')
    .replace(/VitC-?(\d+)/gi,           'วิตซี $1')
    .replace(/Probiotic-?(\d+)/gi,      'โปรไบโอติก $1')
    .replace(/Triple_Bottle/gi,         'กลูต้า 2 + วิตซี 1 + โปรไบโอติก + ขวด')
    .replace(/Triple/gi,                'กลูต้า 2 + วิตซี 1 + โปรไบโอติก')
    .replace(/Capsule_?(\d+in\d+)/gi,   'ตลับวิตามิน $1')
    .replace(/Capsule_?(\d+)/gi,        'ตลับวิตามิน $1')
    .replace(/[_+]+/g,                  ' + ')
    .replace(/\s{2,}/g,                 ' ')
    .trim();

  return name || skuRef;
}

// ─── Sync (auto-detect new SKUs) ─────────────────────────────

/**
 * เทียบ SKU ทั้งหมดใน sku_monthly กับ product_names
 * SKU ที่ไม่เคยรู้จัก → เพิ่มเข้า sheet พร้อม status: 'new'
 * เรียกอัตโนมัติหลัง import ทุกครั้ง
 * @returns {number} จำนวน SKU ใหม่ที่เพิ่ม
 */
function FL_syncProductNames() {
  try {
    const ss      = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
    const pnSheet = FL_ensureProductNamesSheet();

    // Build known set
    const known    = FL_getProductNames();
    const knownSet = new Set(known.map(p => p.sku.toUpperCase()));

    // อ่าน sku_monthly
    const skSheet = ss.getSheetByName(FL_SHEETS.SKU_MONTHLY);
    if (!skSheet) return 0;

    const skData = skSheet.getDataRange().getValues();
    if (skData.length < 2) return 0;

    const skHdr   = skData[0];
    const iSkuCol = skHdr.indexOf('sku_ref');
    if (iSkuCol < 0) return 0;

    const allSkus = [...new Set(
      skData.slice(1)
            .map(r => (r[iSkuCol] || '').toString().trim())
            .filter(Boolean)
    )];

    const newSkus = allSkus.filter(sku => !knownSet.has(sku.toUpperCase()));
    if (!newSkus.length) return 0;

    // SKU ที่อยู่ใน FL_SKU_DISPLAY_MAP อยู่แล้ว → confirmed ทันที
    // เฉพาะที่ไม่รู้จักจริงๆ → new (ต้องให้ user ตั้งชื่อ)
    const knownDisplayMap = FL_SKU_DISPLAY_MAP;
    const now  = new Date();
    const rows = newSkus.map(sku => {
      const upper     = sku.toUpperCase();
      const mapKey    = Object.keys(knownDisplayMap).find(k => k.toUpperCase() === upper);
      const isKnown   = !!mapKey;
      return [sku, FL_autoExtractDisplayName(sku), isKnown ? 'confirmed' : 'new', now];
    });
    pnSheet.getRange(pnSheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);

    Logger.log('FL_syncProductNames: เพิ่ม ' + newSkus.length + ' SKU ใหม่: ' + newSkus.join(', '));
    return newSkus.length;
  } catch (e) {
    Logger.log('FL_syncProductNames error: ' + e);
    return 0;
  }
}

// ─── Write (from UI) ─────────────────────────────────────────

/**
 * บันทึกชื่อเดียว — เรียกจาก Dashboard UI
 * อัปเดต row เดิมหรือ append ใหม่ แล้ว set status → 'confirmed'
 * @param {string} sku
 * @param {string} displayName
 * @returns {{ok:boolean, error?:string}}
 */
function FL_saveProductName(sku, displayName) {
  try {
    const sheet = FL_ensureProductNamesSheet();
    const data  = sheet.getDataRange().getValues();
    const hdr   = {};
    data[0].forEach((h, i) => { if (h) hdr[String(h).trim()] = i; });

    const iSku  = hdr['sku_ref']      ?? 0;
    const iDisp = hdr['display_name'] ?? 1;
    const iStat = hdr['status']       ?? 2;
    const iUpd  = hdr['updated_at']   ?? 3;

    const upper = sku.toString().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iSku] || '').toUpperCase() === upper) {
        sheet.getRange(i + 1, iDisp + 1).setValue(displayName.trim());
        sheet.getRange(i + 1, iStat + 1).setValue('confirmed');
        sheet.getRange(i + 1, iUpd  + 1).setValue(new Date());
        FL_clearDashboardCache();
        return { ok: true };
      }
    }
    // ไม่เจอ — append
    sheet.appendRow([sku, displayName.trim(), 'confirmed', new Date()]);
    FL_clearDashboardCache();
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * บันทึกหลายรายการพร้อมกัน — เรียกจาก Dashboard UI (กด "บันทึกทั้งหมด")
 * @param {Array<{sku:string, displayName:string}>} items
 * @returns {{ok:boolean, saved:number}}
 */
function FL_saveProductNames(items) {
  let saved = 0;
  (items || []).forEach(item => {
    if (item.sku && item.displayName) {
      const res = FL_saveProductName(item.sku, item.displayName);
      if (res.ok) saved++;
    }
  });
  return { ok: true, saved };
}
