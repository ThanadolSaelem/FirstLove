function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 Import')
    .addItem('📦 Shopee — Order', 'importShopeeOrder')
    .addItem('🎵 TikTok — Order', 'importTiktokOrder')
    .addItem('🛍️ Lazada — Order', 'importLazadaOrder')
    .addSeparator()
    .addItem('💰 Shopee — Income', 'importShopeeIncome')
    .addItem('💰 TikTok — Income', 'importTiktokIncome')
    .addItem('💰 Lazada — Income', 'importLazadaIncome')
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
function importShopeeOrder()  { startJob('order',  'shopee'); }
function importTiktokOrder()  { startJob('order',  'tiktok'); }
function importLazadaOrder()  { startJob('order',  'lazada'); }
function importShopeeIncome() { startJob('income', 'shopee'); }
function importTiktokIncome() { startJob('income', 'tiktok'); }
function importLazadaIncome() { startJob('income', 'lazada'); }

// legacy aliases (ถ้าเคยผูก trigger ไว้ชื่อเดิม)
function importShopee() { importShopeeOrder(); }
function importTiktok() { importTiktokOrder(); }
function importLazada() { importLazadaOrder(); }

// ===== Order column config =====
const ORDER_CONFIG = {
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

// ===== Income column config =====
// shopee_income: col A = label (Thai), col B = amount  → SUMIF on A, sum B
// tiktok_income: col A = label (English), col B = amount → MATCH on A, INDEX B
// lazada_income: col B = transaction type, col C = amount → SUMIF on B, sum C
const INCOME_CONFIG = {
  shopee: {
    sheet: 'shopee_income',
    mode: 'two-col',
    outHeaders: ['รายการ', 'จำนวนเงิน'],
    search: [
      ['รายการ', 'รายละเอียด', 'description', 'ประเภท', 'หัวข้อ', 'item'],
      ['มูลค่า', 'จำนวนเงิน', 'amount', 'ยอดเงิน', 'total', 'value'],
    ],
    hint: 'ไฟล์ settlement Shopee — วางทั้งไฟล์\nสคริปต์จะหาคอลัมน์ "รายการ" และ "มูลค่า" เอง',
  },
  tiktok: {
    sheet: 'tiktok_income',
    mode: 'two-col',
    outHeaders: ['Description', 'Amount'],
    search: [
      ['description', 'item', 'type', 'รายการ', 'รายละเอียด'],
      ['amount', 'total', 'จำนวนเงิน', 'settlement amount', 'value'],
    ],
    hint: 'ไฟล์ settlement TikTok — วางทั้งไฟล์\nสคริปต์จะหาคอลัมน์ "Description" และ "Amount" เอง',
  },
  lazada: {
    sheet: 'lazada_income',
    mode: 'lazada',   // col B=type, col C=amount (col A = dummy)
    outHeaders: ['', 'Transaction Type', 'Amount'],
    search: [
      ['transaction type', 'ประเภทธุรกรรม', 'type', 'ประเภท'],
      ['amount', 'จำนวนเงิน', 'net amount', 'fee amount', 'total amount'],
    ],
    hint: 'ไฟล์ financial report Lazada — วางทั้งไฟล์\nสคริปต์จะหาคอลัมน์ "Transaction Type" และ "Amount" เอง',
  },
};

// ===== Step 1: create staging sheet and show instructions =====
function startJob(type, platform) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  PropertiesService.getUserProperties().setProperty('importJob', type + ':' + platform);

  let staging = ss.getSheetByName('__staging__');
  if (staging) { staging.clearContents(); }
  else { staging = ss.insertSheet('__staging__'); }
  ss.setActiveSheet(staging);
  staging.getRange('A1').activate();

  const isIncome = type === 'income';
  const cfg      = isIncome ? INCOME_CONFIG[platform] : null;
  const label    = isIncome ? 'Income/Settlement' : 'Order';
  const hint     = isIncome ? '\n\n💡 ' + cfg.hint : '';

  ui.alert(
    `เตรียม ${platform.toUpperCase()} ${label} พร้อมแล้ว`,
    `กด OK เพื่อปิด dialog นี้${hint}\n\n` +
    `จากนั้น:\n` +
    `1. เปิดไฟล์ ${label} ของ ${platform}\n` +
    `2. กด Ctrl+A แล้ว Ctrl+C\n` +
    `3. คลิกชีต __staging__ แล้วกด Ctrl+V\n` +
    `4. วางเสร็จแล้ว คลิกเมนู '✅ ดำเนินการต่อ (หลัง Paste)'`,
    ui.ButtonSet.OK
  );
}

// ===== Step 2: read staging and route to correct processor =====
function importContinue() {
  const ui   = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  const job  = props.getProperty('importJob');

  if (!job) {
    ui.alert('❌ ไม่พบข้อมูลการ import\nกรุณาเริ่มใหม่โดยเลือก platform ก่อน');
    return;
  }
  props.deleteProperty('importJob');

  const [type, platform] = job.split(':');
  if (type === 'income') {
    processIncomeImport(platform);
  } else {
    processOrderImport(platform);
  }
}

// ===== Order import processor =====
function processOrderImport(platform) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ui      = SpreadsheetApp.getUi();
  const cfg     = ORDER_CONFIG[platform];
  const staging = ss.getSheetByName('__staging__');

  if (!staging) { ui.alert('❌ ไม่พบชีต __staging__\nกรุณาเริ่มใหม่'); return; }

  const allData = staging.getDataRange().getValues();
  if (allData.length < 2) {
    ui.alert('❌ ไม่พบข้อมูลใน __staging__\nกรุณาเริ่มใหม่แล้ววางข้อมูลให้ครบ');
    ss.deleteSheet(staging);
    return;
  }

  const rawHeaders = allData[0].map(h => String(h).trim().toLowerCase());

  const colIdxList = cfg.search.map(candidates => {
    for (const candidate of candidates) {
      const idx = rawHeaders.findIndex(h => h.includes(candidate));
      if (idx >= 0) return idx;
    }
    return -1;
  });

  const required = platform === 'lazada' ? colIdxList.slice(0, 6) : colIdxList;
  const missing  = cfg.search.slice(0, required.length)
    .map((c, i) => ({ name: c[0], idx: colIdxList[i] }))
    .filter(x => x.idx < 0);

  if (missing.length > 0) {
    ui.alert(
      '❌ ไม่พบคอลัมน์',
      `หาไม่เจอ: ${missing.map(m => m.name).join(', ')}\n\n` +
      `หัวคอลัมน์ที่เจอ:\n${rawHeaders.slice(0, 20).join(' | ')}`,
      ui.ButtonSet.OK
    );
    ss.deleteSheet(staging);
    return;
  }

  const extracted = allData.slice(1).map(row =>
    colIdxList.filter(i => i >= 0).map(i => row[i])
  );

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

// ===== Income import processor =====
function processIncomeImport(platform) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ui      = SpreadsheetApp.getUi();
  const cfg     = INCOME_CONFIG[platform];
  const staging = ss.getSheetByName('__staging__');

  if (!staging) { ui.alert('❌ ไม่พบชีต __staging__\nกรุณาเริ่มใหม่'); return; }

  const allData = staging.getDataRange().getValues();
  if (allData.length < 1) {
    ui.alert('❌ ไม่พบข้อมูลใน __staging__');
    ss.deleteSheet(staging);
    return;
  }

  // Find the header row by searching the first 15 rows for any search keyword
  const allKeywords = cfg.search.flat();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(15, allData.length); r++) {
    const row = allData[r].map(c => String(c).trim().toLowerCase());
    if (allKeywords.some(kw => row.some(c => c.includes(kw)))) {
      headerRowIdx = r;
      break;
    }
  }

  let colIdxList;
  let dataRows;

  if (headerRowIdx >= 0) {
    const hdr = allData[headerRowIdx].map(c => String(c).trim().toLowerCase());
    colIdxList = cfg.search.map(candidates => {
      for (const candidate of candidates) {
        const idx = hdr.findIndex(h => h.includes(candidate));
        if (idx >= 0) return idx;
      }
      return -1;
    });
    dataRows = allData.slice(headerRowIdx + 1);

    const missing = cfg.search
      .map((c, i) => ({ name: c[0], idx: colIdxList[i] }))
      .filter(x => x.idx < 0);

    if (missing.length > 0) {
      ui.alert(
        '❌ ไม่พบคอลัมน์',
        `หาไม่เจอ: ${missing.map(m => m.name).join(', ')}\n\n` +
        `หัวคอลัมน์ที่เจอ:\n${allData[headerRowIdx].slice(0, 15).join(' | ')}\n\n` +
        `💡 ลองวางไฟล์ใหม่ตรวจสอบว่าแถวแรกมีหัวคอลัมน์`,
        ui.ButtonSet.OK
      );
      ss.deleteSheet(staging);
      return;
    }
  } else {
    // No header found — assume 2-col data starting at row 0 (label, amount)
    if (allData[0].length < 2) {
      ui.alert('❌ ไม่พบ header row และข้อมูลมีน้อยกว่า 2 คอลัมน์\nกรุณาตรวจสอบไฟล์');
      ss.deleteSheet(staging);
      return;
    }
    colIdxList = [0, 1];
    dataRows = allData;
  }

  // Build output rows (skip fully-empty rows)
  const [iA, iB] = colIdxList;
  const dataFiltered = dataRows.filter(row =>
    String(row[iA]).trim() !== '' || String(row[iB]).trim() !== ''
  );

  let outputRows;
  if (cfg.mode === 'lazada') {
    // lazada_income: col A = dummy, col B = type, col C = amount
    outputRows = dataFiltered.map(row => ['', row[iA], row[iB]]);
  } else {
    // shopee_income / tiktok_income: col A = label, col B = amount
    outputRows = dataFiltered.map(row => [row[iA], row[iB]]);
  }

  if (outputRows.length === 0) {
    ui.alert('❌ ไม่พบข้อมูล data rows หลัง header\nกรุณาตรวจสอบไฟล์');
    ss.deleteSheet(staging);
    return;
  }

  const target = ss.getSheetByName(cfg.sheet);
  if (!target) {
    ui.alert(
      `❌ ไม่พบชีต "${cfg.sheet}"`,
      `กรุณาสร้างชีตชื่อ "${cfg.sheet}" ก่อน`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    ss.deleteSheet(staging);
    return;
  }

  target.clearContents();
  target.getRange(1, 1, 1, cfg.outHeaders.length).setValues([cfg.outHeaders]);
  target.getRange(2, 1, outputRows.length, cfg.outHeaders.length).setValues(outputRows);

  ss.deleteSheet(staging);
  ss.setActiveSheet(target);
  ui.alert(`✅ สำเร็จ! นำเข้า ${outputRows.length} แถว ลงใน ${cfg.sheet}`);
}
