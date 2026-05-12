function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 Import Order')
    .addItem('📦 Shopee', 'importShopee')
    .addItem('🎵 TikTok', 'importTiktok')
    .addItem('🛍️ Lazada', 'importLazada')
    .addSeparator()
    .addItem('✅ ดำเนินการต่อ (หลัง Paste)', 'importContinue')
    .addSeparator()
    .addItem('🔧 Fix Verify_Income formulas', 'fixVerifyIncome')
    .addToUi();
}

function fixVerifyIncome() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Verify_Income');
  if (!sh) { SpreadsheetApp.getUi().alert('ไม่พบชีต Verify_Income'); return; }

  sh.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A1000="","",IF(B2:B1000="shopee",' +
      'IFERROR(SUMIF(shopee_income!$A:$A,"ค่าคอมมิชชั่น*",shopee_income!$B:$B)' +
      '+SUMIF(shopee_income!$A:$A,"ค่าบริการ*",shopee_income!$B:$B)' +
      '+SUMIF(shopee_income!$A:$A,"ค่าธรรมเนียม*",shopee_income!$B:$B)' +
      '+SUMIF(shopee_income!$A:$A,"ค่าธุรกรรม*",shopee_income!$B:$B),0),' +
    'IF(B2:B1000="tiktok",' +
      'IFERROR(VALUE(INDEX(tiktok_income!$B:$B,MATCH("*Total*Fee*",tiktok_income!$A:$A,0))),0),' +
    'IF(B2:B1000="lazada",' +
      'IFERROR(SUMIF(lazada_income!$B:$B,"*ธรรมเนียม*",lazada_income!$C:$C),0)' +
      '+IFERROR(SUMIF(lazada_income!$B:$B,"*Premium*",lazada_income!$C:$C),0),"")))))'
  );

  sh.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A1000="","",IF(F2:F1000=0,"⏳",' +
      'IF(ABS(F2:F1000-G2:G1000)<=1,"✅",' +
      '"❌ diff="&TEXT(F2:F1000-G2:G1000,"#,##0.00")))))'
  );

  SpreadsheetApp.getUi().alert('✅ แก้ F2 และ H2 เรียบร้อย');
}

// ===== Entry points =====
function importShopee() { startImport('shopee'); }
function importTiktok() { startImport('tiktok'); }
function importLazada() { startImport('lazada'); }

// ===== Column config per platform =====
const PLATFORM_CONFIG = {
  shopee: {
    sheet: 'shopee_order',
    headers: ['สถานะการสั่งซื้อ', 'เลขอ้างอิง SKU', 'ยอดชำระเงิน'],
    search: [
      ['สถานการสั่งซื้อ', 'สถานะการสั่งซื้อ', 'order status'],
      ['เลขอ้างอิง sku', 'sku reference', 'sku seller'],
      ['ยอดชำระเงิน', 'total amount', 'buyer paid'],
    ],
  },
  tiktok: {
    sheet: 'tiktok_order',
    headers: ['Order Status', 'Seller SKU', 'SKU Subtotal After Discount'],
    search: [
      ['order status'],
      ['seller sku'],
      ['sku subtotal after discount', 'subtotal after discount'],
    ],
  },
  lazada: {
    sheet: 'lazada_order',
    headers: ['status', 'sellerSku', 'paidPrice', 'shippingFee', 'sellerDiscount', 'refundAmount', 'net_rev [auto]'],
    search: [
      ['status'],
      ['sellersku', 'seller sku'],
      ['paidprice', 'paid price', 'item price'],
      ['shippingfee', 'shipping fee'],
      ['sellerdiscount', 'seller discount'],
      ['refundamount', 'refund amount'],
    ],
  },
};

// ===== Step 1: prepare staging sheet, then close dialog so user can paste =====
function startImport(platform) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  // Store which platform we're importing
  PropertiesService.getUserProperties().setProperty('importPlatform', platform);

  // Create / clear staging sheet
  let staging = ss.getSheetByName('__staging__');
  if (staging) {
    staging.clearContents();
  } else {
    staging = ss.insertSheet('__staging__');
  }
  ss.setActiveSheet(staging);
  staging.getRange('A1').activate();

  // Single blocking dialog — tells user to paste AFTER closing this dialog
  ui.alert(
    `เตรียม ${platform.toUpperCase()} พร้อมแล้ว`,
    `กด OK เพื่อปิด dialog นี้\n\n` +
    `จากนั้น:\n` +
    `1. เปิดไฟล์ order ของ ${platform} (Excel / Sheets)\n` +
    `2. กด Ctrl+A แล้ว Ctrl+C\n` +
    `3. คลิกชีต __staging__ แล้วกด Ctrl+V วางข้อมูล\n` +
    `4. วางเสร็จแล้ว คลิกเมนู '✅ ดำเนินการต่อ (หลัง Paste)'`,
    ui.ButtonSet.OK
  );
}

// ===== Step 2: read staging and write to target sheet =====
function importContinue() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const ui       = SpreadsheetApp.getUi();
  const props    = PropertiesService.getUserProperties();
  const platform = props.getProperty('importPlatform');

  if (!platform) {
    ui.alert('❌ ไม่พบข้อมูลการ import\nกรุณาเริ่มใหม่โดยคลิก Shopee / TikTok / Lazada ก่อน');
    return;
  }
  props.deleteProperty('importPlatform');

  const cfg     = PLATFORM_CONFIG[platform];
  const staging = ss.getSheetByName('__staging__');

  if (!staging) {
    ui.alert('❌ ไม่พบชีต __staging__\nกรุณาเริ่มใหม่');
    return;
  }

  const allData = staging.getDataRange().getValues();
  if (allData.length < 2) {
    ui.alert('❌ ไม่พบข้อมูลใน __staging__\nกรุณาเริ่มใหม่แล้ววางข้อมูลให้ครบ');
    ss.deleteSheet(staging);
    return;
  }

  const rawHeaders = allData[0].map(h => String(h).trim().toLowerCase());

  // Match columns by header name
  const colIdxList = cfg.search.map(candidates => {
    for (const candidate of candidates) {
      const idx = rawHeaders.findIndex(h => h.includes(candidate));
      if (idx >= 0) return idx;
    }
    return -1;
  });

  // Check for missing required columns (lazada col 7 = auto-calculated, skip)
  const required = platform === 'lazada' ? colIdxList.slice(0, 6) : colIdxList;
  const missing  = cfg.search.slice(0, required.length)
    .map((c, i) => ({ name: c[0], idx: colIdxList[i] }))
    .filter(x => x.idx < 0);

  if (missing.length > 0) {
    ui.alert(
      '❌ ไม่พบคอลัมน์',
      `หาไม่เจอ: ${missing.map(m => m.name).join(', ')}\n\n` +
      `หัวคอลัมน์ที่เจอในไฟล์:\n${rawHeaders.slice(0, 20).join(' | ')}`,
      ui.ButtonSet.OK
    );
    ss.deleteSheet(staging);
    return;
  }

  // Extract needed columns
  const extracted = allData.slice(1).map(row =>
    colIdxList.filter(i => i >= 0).map(i => row[i])
  );

  // Write to target sheet
  const target = ss.getSheetByName(cfg.sheet);
  if (!target) {
    ui.alert(`❌ ไม่พบชีต "${cfg.sheet}"`);
    ss.deleteSheet(staging);
    return;
  }

  target.clearContents();

  const usedHeaders = cfg.headers.slice(0, colIdxList.filter(i => i >= 0).length);
  target.getRange(1, 1, 1, usedHeaders.length).setValues([usedHeaders]);

  if (extracted.length > 0) {
    target.getRange(2, 1, extracted.length, extracted[0].length).setValues(extracted);
  }

  // Lazada: restore net_rev formula in column G
  if (platform === 'lazada' && extracted.length > 0) {
    target.getRange(1, 7).setValue('net_rev [auto]');
    target.getRange(2, 7).setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",IFERROR(VALUE(C2:C),0)-IFERROR(VALUE(D2:D),0)-IFERROR(VALUE(E2:E),0)-IFERROR(VALUE(F2:F),0)))'
    );
  }

  ss.deleteSheet(staging);
  ss.setActiveSheet(target);
  ui.alert(`✅ สำเร็จ! นำเข้า ${extracted.length} แถว ลงใน ${cfg.sheet}`);
}
