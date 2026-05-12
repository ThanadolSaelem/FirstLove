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

  // C2: gross formula — MAP per row, with month-key filtering for tiktok/lazada
  // shopee_income: per-import (no month col) → same value every shopee row (acceptable)
  // tiktok_income: col A=Type, col D=Subtotal-before-discounts, col F=Month Key
  // lazada_income: col B=type, col C=amount, col D=Month Key
  sh.getRange('C2').setFormula(
    '=MAP(A2:A1000,B2:B1000,LAMBDA(mk,plat,' +
      'IF(mk="","",IF(plat="shopee",' +
        'IFERROR(SUMIF(shopee_income!$B:$B,"สินค้าราคาปกติ",shopee_income!$C:$C),0),' +
      'IF(plat="tiktok",' +
        'IFERROR(SUMPRODUCT((tiktok_income!$A$2:$A$10000="Order")*(tiktok_income!$F$2:$F$10000=mk)*tiktok_income!$D$2:$D$10000),0),' +
      'IF(plat="lazada",' +
        'IFERROR(SUMPRODUCT((lazada_income!$D$2:$D$10000=mk)*(lazada_income!$C$2:$C$10000>0)*lazada_income!$C$2:$C$10000),0),' +
      '""))))))'
  );

  // E2: gross ✓
  sh.getRange('E2').setFormula(
    '=MAP(A2:A1000,C2:C1000,D2:D1000,LAMBDA(mk,f,g,' +
      'IF(mk="","",IF(f=0,"⏳",IF(ABS(f-g)<=1,"✅","❌ diff="&TEXT(f-g,"#,##0.00"))))))'
  );

  // F2: fee formula — MAP per row with month filtering
  // shopee_income: SUMIF section-total by parent label "ค่าธรรมเนียม"
  // tiktok_income: col E=Total Fees, col F=Month Key
  // lazada_income: col B=type, col C=amount, col D=Month Key → fee rows have negative amount
  sh.getRange('F2').setFormula(
    '=MAP(A2:A1000,B2:B1000,LAMBDA(mk,plat,' +
      'IF(mk="","",IF(plat="shopee",' +
        'IFERROR(SUMIF(shopee_income!$A:$A,"ค่าธรรมเนียม",shopee_income!$D:$D),0),' +
      'IF(plat="tiktok",' +
        'IFERROR(SUMPRODUCT((tiktok_income!$A$2:$A$10000="Order")*(tiktok_income!$F$2:$F$10000=mk)*tiktok_income!$E$2:$E$10000),0),' +
      'IF(plat="lazada",' +
        'IFERROR(SUMPRODUCT((lazada_income!$D$2:$D$10000=mk)*(lazada_income!$C$2:$C$10000<0)*lazada_income!$C$2:$C$10000),0),' +
      '""))))))'
  );

  // H2: fee ✓
  sh.getRange('H2').setFormula(
    '=MAP(A2:A1000,F2:F1000,G2:G1000,LAMBDA(mk,f,g,' +
      'IF(mk="","",IF(f=0,"⏳",IF(ABS(f-g)<=1,"✅","❌ diff="&TEXT(f-g,"#,##0.00"))))))'
  );

  // I2: net/transferred formula — MAP per row with month filtering
  // shopee_income: SUMIF parent-section "*โอนแล้ว*", col D=section total
  // tiktok_income: col B=Settlement Amount, col F=Month Key
  // lazada_income: col C=amount, col D=Month Key → SUM all (positive + negative = net)
  sh.getRange('I2').setFormula(
    '=MAP(A2:A1000,B2:B1000,LAMBDA(mk,plat,' +
      'IF(mk="","",IF(plat="shopee",' +
        'IFERROR(SUMIF(shopee_income!$A:$A,"*โอนแล้ว*",shopee_income!$D:$D),0),' +
      'IF(plat="tiktok",' +
        'IFERROR(SUMPRODUCT((tiktok_income!$A$2:$A$10000="Order")*(tiktok_income!$F$2:$F$10000=mk)*tiktok_income!$B$2:$B$10000),0),' +
      'IF(plat="lazada",' +
        'IFERROR(SUMPRODUCT((lazada_income!$D$2:$D$10000=mk)*lazada_income!$C$2:$C$10000),0),' +
      '""))))))'
  );

  // K2: net ✓
  sh.getRange('K2').setFormula(
    '=MAP(A2:A1000,I2:I1000,J2:J1000,LAMBDA(mk,f,g,' +
      'IF(mk="","",IF(f=0,"⏳",IF(ABS(f-g)<=1,"✅","❌ diff="&TEXT(f-g,"#,##0.00"))))))'
  );

  SpreadsheetApp.getUi().alert('✅ อัปเดต C2, E2, F2, H2, I2, K2 เรียบร้อยแล้ว');
}

// ===== Helpers =====
function _toMonthKey(raw) {
  if (!raw) return '';
  if (raw instanceof Date) {
    return raw.getFullYear() + '-' + String(raw.getMonth() + 1).padStart(2, '0');
  }
  const s = String(raw).trim();
  const m1 = s.match(/^(\d{4})[\/\-](\d{1,2})/);
  if (m1) return m1[1] + '-' + m1[2].padStart(2, '0');
  const m2 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m2) return m2[3] + '-' + m2[2].padStart(2, '0');
  return '';
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
// shopee_income: 4-col (A=parent section, B=sub-label, C=sub-amount, D=section total)
//   → mode 'shopee-income': paste raw, keep all 4 cols
// tiktok_income: 2-col (A=label, B=amount) → mode 'two-col'
// lazada_income: 3-col (A=dummy, B=type, C=amount) → mode 'lazada'
const INCOME_CONFIG = {
  shopee: {
    sheet: 'shopee_income',
    mode: 'shopee-income',
    outHeaders: ['รายการหลัก', 'รายการย่อย', 'จำนวนเงิน', 'ยอดรวมส่วน'],
    hint: 'ไฟล์ settlement Shopee (แท็บ Summary)\nวางทั้งเนื้อหา — ระบบเก็บ 4 คอลัมน์แรกอัตโนมัติ',
  },
  tiktok: {
    sheet: 'tiktok_income',
    mode: 'tiktok-income',
    outHeaders: ['Type', 'Settlement Amount', 'Total Revenue', 'Subtotal Before Discounts', 'Total Fees', 'Month Key'],
    search: [
      ['type'],
      ['total settlement amount'],
      ['total revenue'],
      ['subtotal before discounts'],
      ['total fees'],
      ['order settled time'],  // optional — used to derive Month Key
    ],
    optionalFrom: 5,
    hint: 'ไฟล์ income TikTok (Order details) — วางทั้งไฟล์\nสคริปต์จะหาคอลัมน์ Type, Settlement, Revenue, Subtotal before discounts, Total Fees เอง',
  },
  lazada: {
    sheet: 'lazada_income',
    mode: 'lazada',
    outHeaders: ['', 'Transaction Type', 'Amount', 'Month Key'],
    search: [
      ['transaction type', 'ประเภทธุรกรรม', 'ชื่อรายการธุรกรรม', 'type', 'ประเภท'],
      ['จำนวนเงิน(รวมภาษี)', 'amount', 'จำนวนเงิน', 'net amount', 'fee amount', 'total amount'],
      ['วันที่ปรับปรุงเข้ายอดของฉัน', 'วันที่ทำรายการ', 'settlement date', 'date', 'updated date'],  // optional
    ],
    optionalFrom: 2,
    hint: 'ไฟล์ financial report Lazada — วางทั้งไฟล์\nสคริปต์จะหาคอลัมน์ "ชื่อรายการธุรกรรม" และ "จำนวนเงิน" เอง',
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

  // === Shopee income: fixed 4-col structure (A=parent, B=sub-label, C=sub-amount, D=section-total) ===
  // No header searching needed — just take all 4 cols as-is
  if (cfg.mode === 'shopee-income') {
    const outputRows = allData
      .map(row => [
        row[0] !== undefined ? row[0] : '',
        row[1] !== undefined ? row[1] : '',
        row[2] !== undefined ? row[2] : '',
        row[3] !== undefined ? row[3] : '',
      ])
      .filter(row => row.some(c => String(c).trim() !== ''));

    if (outputRows.length === 0) {
      ui.alert('❌ ไม่พบข้อมูลใน __staging__\nกรุณาเริ่มใหม่แล้ววางข้อมูลให้ครบ');
      ss.deleteSheet(staging);
      return;
    }

    const target = ss.getSheetByName(cfg.sheet);
    if (!target) {
      ui.alert(
        `❌ ไม่พบชีต "${cfg.sheet}"`,
        `กรุณาสร้างชีตชื่อ "${cfg.sheet}" ก่อน`,
        ui.ButtonSet.OK
      );
      ss.deleteSheet(staging);
      return;
    }

    target.clearContents();
    target.getRange(1, 1, 1, cfg.outHeaders.length).setValues([cfg.outHeaders]);
    target.getRange(2, 1, outputRows.length, 4).setValues(outputRows);
    ss.deleteSheet(staging);
    ss.setActiveSheet(target);
    ui.alert(`✅ สำเร็จ! นำเข้า ${outputRows.length} แถว ลงใน ${cfg.sheet}`);
    return;
  }

  // === TikTok / Lazada income: search for header row then extract columns ===

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

    // Only required columns (before optionalFrom) must be found
    const requiredSearch = cfg.optionalFrom !== undefined ? cfg.search.slice(0, cfg.optionalFrom) : cfg.search;
    const missing = requiredSearch
      .map((c, i) => ({ name: c[0], idx: colIdxList[i] }))
      .filter(x => x.idx < 0);

    if (missing.length > 0) {
      ui.alert(
        '❌ ไม่พบคอลัมน์',
        `หาไม่เจอ: ${missing.map(m => m.name).join(', ')}\n\n` +
        `หัวคอลัมน์ที่เจอ:\n${allData[headerRowIdx].slice(0, 15).join(' | ')}`,
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
    // lazada_income: A=dummy, B=type, C=amount, D=Month Key (derived from settlement date)
    const iDate = colIdxList[2]; // optional date column
    outputRows = dataFiltered.map(row => ['', row[iA], row[iB], _toMonthKey(iDate >= 0 ? row[iDate] : null)]);
  } else if (cfg.mode === 'tiktok-income') {
    // tiktok_income: A=Type, B=Settlement, C=Revenue, D=Subtotal-before-discounts, E=Fees, F=Month Key
    const [i0, i1, i2, i3, i4, i5] = colIdxList;
    outputRows = dataFiltered.map(row => [row[i0], row[i1], row[i2], row[i3], row[i4], _toMonthKey(i5 >= 0 ? row[i5] : null)]);
  } else {
    // two-col: col A = label, col B = amount
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
      ui.ButtonSet.OK
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
