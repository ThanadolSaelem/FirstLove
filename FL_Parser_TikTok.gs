// ============================================================
// FL_Parser_TikTok.gs — TikTok Income + Order Parsers
// ============================================================

/**
 * Parse TikTok Income file → monthly revenue summary
 * ไฟล์: income_YYYYMMDDHHMMSS_UTC+7.xlsx
 * Sheet: Reports (row-based)
 *
 * @param {DriveApp.File} driveFile
 * @returns {Object} revenue summary
 */
function FL_parseTikTokIncome(driveFile) {
  const filename = driveFile.getName();

  // ── Month key: prefer "Order settled time" in Order details sheet ──
  // Filename date = export date (e.g. 2026-04-07), NOT settle month (e.g. 2026-01).
  // Use filename only as fallback when sheet/column missing.
  let monthKey = null;
  try {
    const sheetNames = FL_getSheetNames(driveFile);
    const orderDetailsName = sheetNames.find(n =>
      /order\s*details/i.test(n) || /รายละเอียดคำสั่งซื้อ/.test(n)
    );
    if (orderDetailsName) {
      const { rows: odRows } = FL_readXlsxSheet(driveFile, orderDetailsName);
      if (odRows && odRows.length > 1) {
        const odHdr = FL_buildHeaderMap(odRows[0]);
        const colSettled = odHdr['Order settled time'] ?? FL_findCol(odHdr, [
          'Order settled time', 'Settled time', 'Settlement time',
          'เวลาชำระเงิน'
        ]);
        if (colSettled >= 0) {
          for (let i = 1; i < Math.min(odRows.length, 200); i++) {
            const mk = FL_monthKeyFromDateStr(odRows[i][colSettled]);
            if (mk) { monthKey = mk; break; }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('FL_parseTikTokIncome: month_key lookup from Order details failed: ' + e);
  }
  if (!monthKey) monthKey = FL_monthKeyFromFilename(filename);
  if (!monthKey) throw new Error(`ไม่สามารถหา month_key จากชื่อไฟล์: ${filename}`);

  const { rows } = FL_readXlsxSheet(driveFile, 'Reports');

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

  // Exact-match variant: prevents 'Seller discounts' from hitting 'Subtotal after seller discounts'.
  function findExact(keywords) {
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const hasKeyword = row.some(col =>
        col && keywords.some(k => col.toString().trim().toLowerCase() === k.toLowerCase())
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

  const gross           = findValue(['Subtotal before discounts', 'ยอดก่อนหักส่วนลด', 'ยอดขายสินค้าก่อนหักส่วนลด']);
  const seller_discount = findExact(['Seller discounts', 'ส่วนลดจากร้านค้า']);
  const net_revenue     = findValue(['Total Revenue', 'รายได้ทั้งหมด']);
  const transferred     = findValue(['Total settlement amount', 'ยอดที่ชำระให้ผู้ขาย', 'ยอดเงินที่โอน']);
  // Read Total Fees directly from the Reports sheet — matches the figure the customer tracks.
  // (Plug method gave a larger absolute value because it absorbed all adjustments.)
  const platform_fees   = findExact(['Total Fees', 'ค่าธรรมเนียมรวม']);

  return {
    monthKey,
    platform:             'tiktok',
    gross,
    seller_discount,
    net_revenue,
    platform_fees,
    transferred,
    commission:           findValue(['Commission', 'ค่าคอมมิชชั่น']),
    affiliate_commission: findValue(['Affiliate commission', 'ค่าคอมมิชชั่นพันธมิตร']),
    sourceFile:           filename,
  };
}

/**
 * Parse TikTok Order file → SKU sales summary
 * ไฟล์: ทั้งหมด_คำสั่งซื้อ-YYYY-MM-DD-HH_MM.xlsx
 * Sheet: OrderSKUList — inlineStr, Row 1=headers, Row 2=descriptions(SKIP), Row 3+=data
 *
 * @param {DriveApp.File} driveFile
 * @returns {Object} SKU summary
 */
function FL_parseTikTokOrder(driveFile) {
  const filename = driveFile.getName();

  const sheetNames = FL_getSheetNames(driveFile);
  const possible   = ['OrderSKUList', 'Order List', 'Sheet1', 'Orders'];
  const actualName = possible.find(p => sheetNames.includes(p)) || sheetNames[0];

  const { rows } = FL_readXlsxSheet(driveFile, actualName);
  if (rows.length < 2) throw new Error(`TikTok Order: ไม่พบ data rows ในหน้าชีต "${actualName}"`);

  // Row 1 = headers, Row 2 = description row (skip), Row 3+ = data
  const hdr = FL_buildHeaderMap(rows[0]);

  const COL_STATUS = hdr['Order Status']                ?? FL_findCol(hdr, ['Status', 'สถานะ']);
  const COL_SKU    = hdr['Seller SKU']                  ?? FL_findCol(hdr, ['SKU', 'Seller SKU']);
  const COL_QTY    = hdr['Quantity']                    ?? FL_findCol(hdr, ['Qty', 'จำนวน']);
  const COL_NET    = hdr['SKU Subtotal After Discount']  ?? FL_findCol(hdr, ['Revenue', 'ยอดขาย']);
  const COL_PRICE  = hdr['SKU Unit Original Price']     ?? FL_findCol(hdr, ['Price', 'ราคา']);
  const COL_DATE   = hdr['Created Time']                ?? FL_findCol(hdr, ['Paid Time', 'Date']);

  // Determine data start row:
  // Standard TikTok OrderSKUList has a description row 2 (skip) → start at index 2
  // If no known headers found, fall back to starting at index 1
  const dataStartRow = (COL_STATUS >= 0 || COL_SKU >= 0) ? 2 : 1;

  // ── Month key: prefer data (Created Time) over filename ──────────────
  // Filename date = export date (e.g. 2026-04-07), NOT order date (e.g. 2026-01)
  let monthKey = null;
  if (COL_DATE >= 0) {
    // Scan first few data rows to find Created Time
    for (let i = dataStartRow; i < Math.min(dataStartRow + 10, rows.length); i++) {
      const mk = FL_monthKeyFromDateStr(rows[i][COL_DATE]);
      if (mk) { monthKey = mk; break; }
    }
  }
  // Fallback to filename if data date not found
  if (!monthKey) monthKey = FL_monthKeyFromFilename(filename);
  if (!monthKey) throw new Error(`TikTok Order: ไม่สามารถหา month_key (ไฟล์: ${filename})`);

  const successStatuses = ['เสร็จสมบูรณ์', 'Completed', 'Delivered', 'Finished', 'สำเร็จ', 'Shipped'];
  const skuMap = {};

  for (let i = dataStartRow; i < rows.length; i++) {
    const row = rows[i];

    // Multi-month export guard: skip orders from months other than monthKey.
    // TikTok exports often span 60–90 days; without this filter all orders in the
    // file would be attributed to the first order's month.
    if (COL_DATE >= 0) {
      const rowMk = FL_monthKeyFromDateStr(rows[i][COL_DATE]);
      if (rowMk && rowMk !== monthKey) continue;
    }

    const status = (row[COL_STATUS] || '').toString().trim();
    if (status && !successStatuses.includes(status)) continue;

    const rawSku = (row[COL_SKU] || '').toString().trim();
    // Skip non-product rows (e.g. "ขวดน้ำ ฟรี" = free gift item)
    if (!rawSku || rawSku === 'ขวดน้ำ ฟรี') continue;

    const skuRef = FL_normalizeSKU(rawSku);
    const qty    = FL_toNum(row[COL_QTY]) || 1;
    const net    = FL_toNum(row[COL_NET]);
    const price  = FL_toNum(row[COL_PRICE]);

    if (!skuMap[skuRef]) {
      skuMap[skuRef] = { skuRef, category: FL_getCategory(skuRef), units: 0, revenue: 0 };
    }
    skuMap[skuRef].units   += qty;
    skuMap[skuRef].revenue += net !== 0 ? net : price * qty;
  }

  return {
    monthKey,
    platform:   'tiktok',
    skus:       Object.values(skuMap),
    sourceFile: filename,
  };
}