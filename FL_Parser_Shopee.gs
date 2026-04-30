// ============================================================
// FL_Parser_Shopee.gs — Shopee Income + Order Parsers
// ============================================================

/**
 * Parse Shopee Income file → monthly revenue summary
 * ไฟล์: Income_โอนเงินสำเร็จ_th_YYYYMMDD_YYYYMMDD.xlsx
 * Sheet: Summary (row-based)
 *
 * @param {DriveApp.File} driveFile
 * @returns {Object} revenue summary
 */
function FL_parseShopeeIncome(driveFile) {
  const filename = driveFile.getName();
  const monthKey = FL_monthKeyFromFilename(filename);
  if (!monthKey) throw new Error(`ไม่สามารถหา month_key จากชื่อไฟล์: ${filename}`);

  const { rows } = FL_readXlsxSheet(driveFile, 'Summary');
  if (!rows || rows.length < 2) throw new Error('Shopee Income: ไม่พบข้อมูลใน Summary sheet');

  // Shopee Summary sheet = label-value overview (same pattern as TikTok Reports).
  // Each row: label in col A or B, numeric value in the adjacent column.
  // Guard: skip cells whose string starts with digits followed by Thai text (e.g. "3. จำนวนเงิน...")
  // to prevent parseFloat returning the leading digit instead of the actual value.
  function findValue(keywords) {
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const hasKeyword = row.some(col =>
        col && keywords.some(k => col.toString().toLowerCase().includes(k.toLowerCase()))
      );
      if (hasKeyword) {
        for (let j = 0; j < row.length; j++) {
          const val = row[j];
          if (val === '' || val === null || val === undefined) continue;
          if (!/^-?\d+(\.\d+)?$/.test(val.toString().replace(/,/g, '').trim())) continue;
          const num = FL_toNum(val);
          if (num !== 0) return num;
        }
      }
    }
    return 0;
  }

  const gross       = findValue(['สินค้าราคาปกติ',              'Selling Price', 'Original Price']);
  const discount    = findValue(['ส่วนลดสินค้าจากผู้ขาย',     'Seller-absorbed discount', 'Seller discount']);
  const transferred = findValue(['จำนวนเงินทั้งหมดที่โอนแล้ว', 'Amount to be Credited',   'Settlement Amount']);
  // Read total fees directly from the "ค่าธรรมเนียม" section total (col D = section sum).
  // Plug method was off by the Shopee-funded subsidy (ส่วนลดสินค้าที่ออกโดย Shopee).
  // Fallback to plug if the label is not found (older file versions).
  const shopee_subsidy = findValue(['ส่วนลดสินค้าที่ออกโดย Shopee', 'Shopee-funded discount']);
  const _feeFromSheet  = findValue(['ค่าธรรมเนียม', 'Service Fee', 'Total Fee']);
  const platform_fees  = _feeFromSheet !== 0
    ? _feeFromSheet
    : transferred - gross - discount - shopee_subsidy;

  return {
    monthKey,
    platform:        'shopee',
    gross,
    seller_discount: discount,
    net_revenue:     gross + discount,
    platform_fees,
    transferred,
    sourceFile:      filename,
  };
}

/**
 * Parse Shopee Order file → SKU sales summary
 * ไฟล์: Order_all_YYYYMMDD_YYYYMMDD.xlsx
 * Sheet: orders, Header Row 1
 *
 * @param {DriveApp.File} driveFile
 * @returns {Object} SKU summary
 */
function FL_parseShopeeOrder(driveFile) {
  const filename = driveFile.getName();
  const monthKey = FL_monthKeyFromFilename(filename);
  if (!monthKey) throw new Error(`ไม่สามารถหา month_key จากชื่อไฟล์: ${filename}`);

  const sheetNames = FL_getSheetNames(driveFile);
  const possible   = ['orders', 'Order Data', 'Sheet1', 'Order', 'Order Export'];
  const actualName = possible.find(p => sheetNames.includes(p)) || sheetNames[0];
  
  const { rows } = FL_readXlsxSheet(driveFile, actualName);
  if (rows.length < 2) throw new Error(`Shopee Order: ไม่พบ data rows ในหน้าชีต "${actualName}"`);

  const hdr = FL_buildHeaderMap(rows[0]);

  const COL_STATUS = FL_findCol(hdr, ['สถานะการสั่งซื้อ', 'Order Status', 'Status']);
  const COL_SKU    = FL_findCol(hdr, ['เลขอ้างอิง SKU (SKU Reference No.)', 'เลขอ้างอิง SKU', 'SKU Reference No.', 'SKU']);
  const COL_QTY    = FL_findCol(hdr, ['จำนวน', 'Quantity', 'Qty']);
  const COL_PRICE  = FL_findCol(hdr, ['ราคาขาย', 'Unit Price', 'Price']);

  if (COL_STATUS < 0) throw new Error('Shopee Order: ไม่พบคอลัมน์ สถานะการสั่งซื้อ / Order Status');
  if (COL_SKU    < 0) throw new Error('Shopee Order: ไม่พบคอลัมน์ SKU Reference');
  if (COL_QTY    < 0) throw new Error('Shopee Order: ไม่พบคอลัมน์ จำนวน / Quantity');

  const successStatuses = ['สำเร็จแล้ว', 'สำเร็จ', 'Completed', 'Delivered', 'Shipped', 'เสร็จสิ้น'];
  const skuMap = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const status = (row[COL_STATUS] || '').toString().trim();
    if (!successStatuses.includes(status)) continue;

    const skuRef = FL_normalizeSKU(row[COL_SKU] || '');
    const qty    = FL_toNum(row[COL_QTY]) || 1;
    const price  = COL_PRICE >= 0 ? FL_toNum(row[COL_PRICE]) : 0;

    if (!skuMap[skuRef]) {
      skuMap[skuRef] = { skuRef, category: FL_getCategory(skuRef), units: 0, revenue: 0 };
    }
    skuMap[skuRef].units   += qty;
    skuMap[skuRef].revenue += price * qty;
  }

  return {
    monthKey,
    platform:   'shopee',
    skus:       Object.values(skuMap),
    sourceFile: filename,
  };
}