// ============================================================
// FL_Metadata.gs — Management of SKU Master Data (Phase 3)
// ============================================================

/**
 * IMPORT SKU MAPPING DATA
 * รันใน Apps Script เพื่อนำเข้า/รีเซ็ตข้อมูล SKU master
 * ⚠️ commission_per_unit = placeholder — ต้องยืนยันกับลูกค้าก่อน Phase 3
 */
function FL_importSkuMasterData() {
  const ss = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.SKU_MASTER);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('ไม่พบ sheet sku_master กรุณารัน Setup ก่อน');
    return;
  }

  // Header: sku_ref, display_name, brand, category, box_size, commission_per_unit, updated_at
  // ลูกค้าสามารถแก้คอลัมน์ display_name ในชีตได้โดยตรง ไม่ต้องแตะโค้ด
  const data = [
    { sku: 'GLUTAX9-1',              displayName: 'กลูต้า 1 กระปุก',                         brand: 'FirstLove',   boxSize: 'AA', commission: 10.0 },
    { sku: 'GLUTAX9-2',              displayName: 'กลูต้า 2 กระปุก',                         brand: 'FirstLove',   boxSize: 'AA', commission: 20.0 },
    { sku: 'GLUTAX9-3',              displayName: 'กลูต้า 3 กระปุก',                         brand: 'FirstLove',   boxSize: 'B',  commission: 30.0 },
    { sku: 'GLUTAX9-4',              displayName: 'กลูต้า 4 กระปุก',                         brand: 'FirstLove',   boxSize: 'B',  commission: 40.0 },
    { sku: 'GLUTAX9-6',              displayName: 'กลูต้า 6 กระปุก',                         brand: 'FirstLove',   boxSize: 'C',  commission: 60.0 },
    { sku: 'Bundle_Gluta2+VitC1',    displayName: 'กลูต้า 2 + วิตซี 1',                     brand: 'FirstLove',   boxSize: 'AA', commission: 25.0 },
    { sku: 'Bundle_Gluta1+VitC1',    displayName: 'กลูต้า 1 + วิตซี 1',                     brand: 'FirstLove',   boxSize: 'AA', commission: 15.0 },
    { sku: 'Bundle_Gluta4+VitC2',    displayName: 'กลูต้า 4 + วิตซี 2',                     brand: 'FirstLove',   boxSize: 'B',  commission: 50.0 },
    { sku: 'Bundle_Triple',          displayName: 'กลูต้า 2 + วิตซี 1 + โปรไบโอติก',       brand: 'FirstLove',   boxSize: 'B',  commission: 32.0 },
    { sku: 'Bundle_Triple_Bottle',   displayName: 'กลูต้า 2 + วิตซี 1 + โปรไบโอติก + ขวด', brand: 'FirstLove',   boxSize: 'B',  commission: 35.0 },
    { sku: 'Bundle_Capsule_2in1',    displayName: 'ตลับวิตามิน 2in1',                        brand: 'FirstLove',   boxSize: 'AA', commission: 20.0 },
    { sku: 'Bundle_Capsule_3',       displayName: 'ตลับวิตามิน 3',                           brand: 'FirstLove',   boxSize: 'AA', commission: 25.0 },
    { sku: 'VITC-1',                 displayName: 'วิตซี 1',                                  brand: 'FirstLove',   boxSize: 'AA', commission: 5.0  },
    { sku: 'VITC-2',                 displayName: 'วิตซี 2',                                  brand: 'FirstLove',   boxSize: 'AA', commission: 10.0 },
    { sku: 'VITC-3',                 displayName: 'วิตซี 3',                                  brand: 'FirstLove',   boxSize: 'B',  commission: 15.0 },
    { sku: 'PROBIOTIC-1',            displayName: 'โปรไบโอติก 1',                            brand: 'FirstLove',   boxSize: 'AA', commission: 7.0  },
    { sku: 'PROBIOTIC-2',            displayName: 'โปรไบโอติก 2',                            brand: 'FirstLove',   boxSize: 'B',  commission: 14.0 },
    { sku: 'ERB-01',                 displayName: 'Erb ผลิตภัณฑ์',                           brand: 'Erb',         boxSize: 'AA', commission: 15.0 },
    { sku: 'SORDOR-01',              displayName: 'Sordorstyle ผลิตภัณฑ์',                   brand: 'Sordorstyle', boxSize: 'B',  commission: 20.0 },
  ];

  const now = new Date();
  const rows = data.map(item => [
    item.sku,
    item.displayName || item.sku,  // display_name (ลูกค้าแก้ได้ในชีต)
    item.brand,
    // FIX: normalize ก่อน getCategory เพื่อ match key ใน FL_CATEGORY_MAP
    FL_getCategory(FL_normalizeSKU(item.sku)),
    item.boxSize,
    item.commission,
    now,
  ]);

  // FIX: clearContent ด้วย numColumns ที่ถูกต้อง (เดิมใช้ lastRow-1 ผิด)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  Logger.log('FL_importSkuMasterData: สำเร็จ ' + rows.length + ' รายการ');
  return rows.length;
}

/**
 * Get SKU Metadata Map (ใช้ใน FL_Main.gs สำหรับ commission + logistics)
 * @returns {Object.<string, {boxSize:string, commission:number}>}
 */
function FL_getSkuMetadataMap() {
  const ss = SpreadsheetApp.openById(FL_getConfig().OUTPUT_SHEET_ID);
  const sheet = ss.getSheetByName(FL_SHEETS.SKU_MASTER);
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  // FIX: ใช้ header map แทน positional destructure — ป้องกัน column shift
  // Expected header: sku_ref, display_name, brand, category, box_size, commission_per_unit, updated_at
  const hdr = {};
  values[0].forEach((h, i) => { if (h) hdr[h.toString().trim()] = i; });

  // FIX: header order = sku_ref(0), display_name(1), brand(2), category(3),
  //                     box_size(4), commission_per_unit(5), updated_at(6)
  const iSku   = hdr['sku_ref']             ?? 0;
  const iDisp  = hdr['display_name']        ?? -1;  // อาจไม่มีใน sheet เก่า
  const iBox   = hdr['box_size']            ?? 4;
  const iComm  = hdr['commission_per_unit'] ?? 5;

  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sku = (row[iSku] || '').toString().trim();
    if (!sku) continue;
    map[sku.toUpperCase()] = {
      displayName: iDisp >= 0 ? (row[iDisp] || '').toString().trim() : '',
      boxSize:     (row[iBox]  || '').toString().trim(),
      commission:  parseFloat(row[iComm]) || 0,
    };
  }
  return map;
}