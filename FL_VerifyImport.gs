function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 Import Order')
    .addItem('📦 Shopee', 'importShopee')
    .addItem('🎵 TikTok', 'importTiktok')
    .addItem('🛍️ Lazada', 'importLazada')
    .addSeparator()
    .addItem('🔧 Fix Verify_Income formulas', 'fixVerifyIncome')
    .addToUi();
}

function fixVerifyIncome() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Verify_Income');
  if (!sh) { SpreadsheetApp.getUi().alert('ไม่พบชีต Verify_Income'); return; }

  // F2 — fee จาก income sheet (SUMIF ครบทุก row สำหรับ Shopee)
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

  // H2 — fee ✓ check (เทียบ F vs G)
  sh.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A1000="","",IF(F2:F1000=0,"⏳",' +
      'IF(ABS(F2:F1000-G2:G1000)<=1,"✅",' +
      '"❌ diff="&TEXT(F2:F1000-G2:G1000,"#,##0.00")))))'
  );

  SpreadsheetApp.getUi().alert('✅ แก้ F2 และ H2 เรียบร้อย');
}

// ===== Entry points =====
function importShopee()  { runImport('shopee');  }
function importTiktok()  { runImport('tiktok');  }
function importLazada()  { runImport('lazada');  }

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

// ===== Main logic =====
function runImport(platform) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const cfg = PLATFORM_CONFIG[platform];

  // Step 1: instruct user
  const step1 = ui.alert(
    `Import ${platform.toUpperCase()} Order`,
    `วิธีใช้:\n` +
    `1. เปิดไฟล์ order ของ ${platform} (ใน Excel หรือ Google Sheets)\n` +
    `2. กด Ctrl+A เพื่อ select ทั้งหมด\n` +
    `3. กด Ctrl+C เพื่อ Copy\n` +
    `4. กลับมาที่หน้านี้ แล้วกด OK`,
    ui.ButtonSet.OK_CANCEL
  );
  if (step1 !== ui.Button.OK) return;

  // Step 2: open staging sheet for paste
  let staging = ss.getSheetByName('__staging__');
  if (staging) {
    staging.clearContents();
  } else {
    staging = ss.insertSheet('__staging__');
  }
  ss.setActiveSheet(staging);
  staging.getRange('A1').activate();

  const step2 = ui.alert(
    'วาง (Paste) ข้อมูลได้เลย',
    'กด Ctrl+V เพื่อวางข้อมูลลงในชีตนี้\nแล้วกด OK เมื่อเสร็จแล้ว',
    ui.ButtonSet.OK_CANCEL
  );
  if (step2 !== ui.Button.OK) {
    ss.deleteSheet(staging);
    return;
  }

  // Step 3: read staging data
  const allData = staging.getDataRange().getValues();
  if (allData.length < 2) {
    ui.alert('❌ ไม่พบข้อมูล กรุณาลองใหม่');
    ss.deleteSheet(staging);
    return;
  }

  const rawHeaders = allData[0].map(h => String(h).trim().toLowerCase());

  // Step 4: match columns
  const colIdxList = cfg.search.map((candidates, i) => {
    for (const candidate of candidates) {
      const idx = rawHeaders.findIndex(h => h.includes(candidate));
      if (idx >= 0) return idx;
    }
    return -1; // not found
  });

  // Check which columns are missing (skip last col for lazada = auto-calculated)
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

  // Step 5: extract only needed columns
  const extracted = allData.slice(1).map(row => {
    const pickedCols = colIdxList.filter(i => i >= 0).map(i => row[i]);
    return pickedCols;
  });

  // Step 6: write to target sheet
  const target = ss.getSheetByName(cfg.sheet);
  if (!target) {
    ui.alert(`❌ ไม่พบชีต "${cfg.sheet}"`);
    ss.deleteSheet(staging);
    return;
  }

  target.clearContents();

  // Write header row
  const usedHeaders = cfg.headers.slice(0, colIdxList.filter(i => i >= 0).length);
  target.getRange(1, 1, 1, usedHeaders.length).setValues([usedHeaders]);

  // Write data rows
  if (extracted.length > 0) {
    target.getRange(2, 1, extracted.length, extracted[0].length).setValues(extracted);
  }

  // For lazada: restore net_rev formula in column G
  if (platform === 'lazada' && extracted.length > 0) {
    target.getRange(1, 7).setValue('net_rev [auto]');
    target.getRange(2, 7).setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",IFERROR(VALUE(C2:C),0)-IFERROR(VALUE(D2:D),0)-IFERROR(VALUE(E2:E),0)-IFERROR(VALUE(F2:F),0)))'
    );
  }

  // Cleanup & done
  ss.deleteSheet(staging);
  ss.setActiveSheet(target);
  ui.alert(`✅ สำเร็จ! นำเข้า ${extracted.length} แถว ลงใน ${cfg.sheet}`);
}
