// ============================================================
// FL_Parser_Lazada.gs — Lazada Order + Income Parsers
// ============================================================
// ⚠️ ทั้งสองไฟล์ใช้ inlineStr (ไม่มี sharedStrings.xml)
//    FL_readXlsxSheet จัดการให้อัตโนมัติแล้ว
// ============================================================

// ─────────────────────────────────────────────────────────────
// PARSER 1: Lazada Order File
// ชื่อไฟล์: {hash}.xlsx
// Sheet: sheet1
// Header: Row 1, Data: Row 2+
// 1 row = 1 unit (orderItem level)
//
// คอลัมน์หลัก:
//   F  (5):  sellerSku     → normalize → SKU ref (+ itemName fallback)
//   I  (8):  createTime    → "DD Mon YYYY HH:MM" → month_key
//   AU (46): paidPrice     → ราคาที่ลูกค้าจ่ายจริง (รวมค่าส่ง)
//   AV (47): unitPrice     → ราคาสินค้าก่อนลดราคา
//   AW (48): sellerDiscountTotal → ส่วนลดจากผู้ขาย
//   AX (49): shippingFee   → ค่าส่ง
//   AZ (51): itemName      → ชื่อสินค้าเต็ม (ใช้ fallback SKU mapping)
//   BN (65): status        → filter: confirmed / delivered / shipped
//   BW (74): refundAmount  → ยอดคืนเงิน
// ─────────────────────────────────────────────────────────────

/**
 * @param {DriveApp.File} driveFile
 * @returns {{ monthKey, platform:'lazada', skus:Array, revenue_summary:Object, sourceFile }}
 */
function FL_parseLazadaOrder(driveFile) {
  const filename = driveFile.getName();

  const sheetNames = FL_getSheetNames(driveFile);
  const possible   = ['sheet1', 'Sheet1', 'Lazada Order', 'OrderExport'];
  const actualName = possible.find(p => sheetNames.includes(p)) || sheetNames[0];

  const { rows } = FL_readXlsxSheet(driveFile, actualName);
  if (rows.length < 2) throw new Error(`Lazada Order: ไม่พบ data rows ในหน้าชีต "${actualName}"`);

  const hdr = FL_buildHeaderMap(rows[0]);

  const COL_SKU      = FL_findCol(hdr, ['sellerSku', 'Seller SKU', 'SKU']);
  const COL_DATE     = FL_findCol(hdr, ['createTime', 'Created Time', 'Date']);
  const COL_UNIT_PX  = FL_findCol(hdr, ['unitPrice', 'Unit Price']);
  const COL_PAID     = FL_findCol(hdr, ['paidPrice', 'Paid Price']);
  const COL_SHIPPING = FL_findCol(hdr, ['shippingFee', 'Shipping Fee']);
  const COL_DISCOUNT = FL_findCol(hdr, ['sellerDiscountTotal', 'Seller Discount']);
  const COL_ITEM     = FL_findCol(hdr, ['itemName', 'Item Name', 'Product Name']);
  const COL_STATUS   = FL_findCol(hdr, ['status', 'Status', 'สถานะ']);
  const COL_REFUND   = FL_findCol(hdr, ['refundAmount', 'Refund Amount']);

  if (COL_STATUS < 0) throw new Error('Lazada Order: ไม่พบคอลัมน์ Status');
  if (COL_SKU < 0)    throw new Error('Lazada Order: ไม่พบคอลัมน์ SKU');

  // Infer month_key จาก filename ก่อน ถ้าชื่อเป็น hash ให้ดู data row แรก
  let monthKey = FL_monthKeyFromFilename(filename);
  if (!monthKey && rows.length >= 2) {
    monthKey = FL_monthKeyFromDateStr(rows[1][COL_DATE]);
  }
  if (!monthKey) throw new Error(`Lazada Order: ไม่สามารถหา month_key (ไฟล์: ${filename})`);

  const skuMap = {};
  let totalRevenue = 0;
  let totalGross   = 0;
  let totalShipping = 0;
  let unknownCount = 0;

  for (let i = 1; i < rows.length; i++) {
    const row    = rows[i];

    // Multi-month export guard: skip orders from months other than monthKey.
    if (COL_DATE >= 0) {
      const rowMk = FL_monthKeyFromDateStr(rows[i][COL_DATE]);
      if (rowMk && rowMk !== monthKey) continue;
    }

    const status = (row[COL_STATUS] || '').toString().trim().toLowerCase();

    // ✅ นับ: confirmed, delivered, shipped
    // ❌ ไม่นับ: canceled, failed, returned
    if (!FL_LAZADA_VALID_STATUSES.has(status)) continue;

    const rawSku   = (row[COL_SKU]  || '').toString().trim();
    const itemName = (row[COL_ITEM] || '').toString().trim();

    // ใช้ FL_normalizeSKUWithFallback (SKU map → Lazada map → itemName keyword)
    const skuRef   = FL_normalizeSKUWithFallback(rawSku, itemName);
    const category = FL_getCategory(skuRef);

    const unitPx   = FL_toNum(row[COL_UNIT_PX]);
    const paidPx   = FL_toNum(row[COL_PAID]);
    const shipping = FL_toNum(row[COL_SHIPPING]);
    const discount = Math.abs(FL_toNum(row[COL_DISCOUNT]));
    const refund   = Math.abs(FL_toNum(row[COL_REFUND]));

    // Net revenue per unit = paidPrice - shippingFee - sellerDiscount - refund
    // (paidPrice รวมค่าส่งอยู่แล้ว ดังนั้นต้องหักค่าส่งออก)
    const revenue = paidPx - shipping - discount - refund;

    totalGross    += unitPx - refund;
    totalShipping += shipping;
    totalRevenue  += revenue;

    if (category === 'อื่นๆ') unknownCount++;

    if (!skuMap[skuRef]) {
      skuMap[skuRef] = { skuRef, category, units: 0, revenue: 0 };
    }
    skuMap[skuRef].units   += 1;   // 1 row = 1 unit
    skuMap[skuRef].revenue += revenue;
  }

  const result = {
    monthKey,
    platform:    'lazada',
    skus:        Object.values(skuMap),
    // revenue_summary ใช้สำหรับ monthly_summary sheet (ถ้าไม่มี Income file)
    revenue_summary: {
      gross:           totalGross,
      seller_discount: 0,          // ไม่มีข้อมูลรวม ใช้ 0
      net_revenue:     totalRevenue,
      platform_fees:   0,          // ไม่มีข้อมูลรวม ใช้ Income file แทน
      transferred:     totalRevenue,
    },
    sourceFile: filename,
  };

  if (unknownCount > 0) {
    result.warning = `พบ SKU ที่ไม่รู้จัก ${unknownCount} รายการ — ตรวจสอบ FL_LAZADA_SKU_MAP`;
  }

  return result;
}

// ─────────────────────────────────────────────────────────────
// PARSER 2: Lazada Income File
// ชื่อไฟล์: S{hash}.xlsx
// Sheet: Income Overview
// Header: Row 1, Data: Row 2+
// โครงสร้างพิเศษ: Transaction-level — 1 order มีหลาย rows
//
// คอลัมน์หลัก:
//   C (2):  วันที่ทำรายการ       → transaction date
//   D (3):  ชื่อรายการธุรกรรม    → transaction type
//   E (4):  จำนวนเงิน(รวมภาษี)  → amount (บวก/ลบ) ← SUM per orderNumber = Net
//   J (9):  วันที่สร้างคำสั่งซื้อ → ใช้ month_key
//   K (10): หมายเลขคำสั่งซื้อ    → grouping key
//   M (12): SKU ร้านค้า          → sellerSku
//   R (17): ชื่อสินค้า            → itemName fallback
// ─────────────────────────────────────────────────────────────

/**
 * @param {DriveApp.File} driveFile
 * @returns {{ monthKey, platform:'lazada', gross, net_revenue, transferred,
 *             fee_breakdown:Object, sourceFile }}
 */
function FL_parseLazadaIncome(driveFile) {
  const filename = driveFile.getName();

  const { rows } = FL_readXlsxSheet(driveFile, 'Income Overview');
  if (rows.length < 2) throw new Error('Lazada Income: ไม่พบ data rows');

  const hdr = FL_buildHeaderMap(rows[0]);

  // Column indices — header map ก่อน fallback fixed
  const COL_TXN_DATE  = FL_findCol(hdr, ['วันที่ทำรายการ', 'วันที่'])                     ?? 2;   // C
  const COL_TXN_TYPE  = FL_findCol(hdr, ['ชื่อรายการธุรกรรม'])                             ?? 3;   // D
  const COL_AMOUNT    = FL_findCol(hdr, ['จำนวนเงิน(รวมภาษี)', 'จำนวนเงิน'])              ?? 4;   // E
  const COL_ORDER_DATE= FL_findCol(hdr, ['วันที่สร้างคำสั่งซื้อ'])                          ?? 9;   // J
  const COL_ORDER_NUM = FL_findCol(hdr, ['หมายเลขคำสั่งซื้อ'])                              ?? 10;  // K
  const COL_SKU       = FL_findCol(hdr, ['SKU ร้านค้า', 'sellerSku'])                      ?? 12;  // M
  const COL_ITEM      = FL_findCol(hdr, ['ชื่อสินค้า'])                                     ?? 17;  // R

  // ─── Infer month_key ─────────────────────────────────────
  let monthKey = FL_monthKeyFromFilename(filename);
  if (!monthKey) {
    // หา month_key จาก order date ของ row แรกที่มีข้อมูล
    for (let i = 1; i < rows.length; i++) {
      const mk = FL_monthKeyFromDateStr(rows[i][COL_ORDER_DATE])
              || FL_monthKeyFromDateStr(rows[i][COL_TXN_DATE]);
      if (mk) { monthKey = mk; break; }
    }
  }
  if (!monthKey) throw new Error(`Lazada Income: ไม่สามารถหา month_key (ไฟล์: ${filename})`);

  // ─── Aggregate by transaction type ───────────────────────
  // Transaction types:
  //   ยอดรวมค่าสินค้า            → gross (บวก)
  //   ค่าธรรมเนียมการชำระเงิน    → payment_fee (ลบ)
  //   หักค่าธรรมเนียมการขายสินค้า → selling_fee (ลบ)
  //   Premium Package            → premium_fee (ลบ)

  let gross       = 0;
  let paymentFee  = 0;
  let sellingFee  = 0;
  let premiumFee  = 0;
  let otherFee    = 0;
  let totalNet    = 0;
  const otherTypes = {};

  // Per-order net (ใช้ตรวจสอบ)
  const orderMap = {};

  for (let i = 1; i < rows.length; i++) {
    const row     = rows[i];
    const txnType = (row[COL_TXN_TYPE] || '').toString().trim();
    const amount  = FL_toNum(row[COL_AMOUNT]);
    const orderNo = (row[COL_ORDER_NUM] || '').toString().trim();

    if (!orderNo) continue;

    // สะสม per-order
    if (!orderMap[orderNo]) orderMap[orderNo] = 0;
    orderMap[orderNo] += amount;
    totalNet          += amount;

    // จัดประเภท fee
    const lowerType = txnType.toLowerCase();
    if (txnType.includes('ยอดรวมค่าสินค้า') || lowerType.includes('item price')) {
      gross += amount;
    } else if (txnType.includes('ชำระเงิน') || lowerType.includes('payment fee')) {
      paymentFee += amount;
    } else if (txnType.includes('ขายสินค้า') || txnType.includes('commission') || 
               lowerType.includes('transaction fee')) {
      sellingFee += amount;
    } else if (lowerType.includes('premium')) {
      premiumFee += amount;
    } else {
      otherFee += amount;
      otherTypes[txnType] = (otherTypes[txnType] || 0) + 1;
    }
  }

  const totalFees = paymentFee + sellingFee + premiumFee + otherFee;

  const unknownTxnKeys = Object.keys(otherTypes);
  if (unknownTxnKeys.length > 0) {
    Logger.log('FL_parseLazadaIncome: ⚠️ unrecognized txn types: '
      + unknownTxnKeys.map(t => `"${t}" (${otherTypes[t]}x)`).join(', '));
  }

  return {
    monthKey,
    platform:        'lazada',
    gross:           gross,
    seller_discount: 0,         // Lazada Income ไม่แยก seller discount
    net_revenue:     gross + totalFees,  // gross - fees
    platform_fees:   totalFees,
    transferred:     totalNet,  // sum ทุก transaction = ยอดที่โอนจริง
    fee_breakdown: {
      payment_fee: paymentFee,
      selling_fee: sellingFee,
      premium_fee: premiumFee,
      other_fee:   otherFee,
    },
    order_count:  Object.keys(orderMap).length,
    sourceFile:   filename,
  };
}