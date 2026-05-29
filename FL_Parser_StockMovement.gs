// ============================================================
// FL_Parser_StockMovement.gs — รายงานการซื้อขายสินค้า (accounting)
// ============================================================
// รองรับไฟล์ชื่อ Stock_Movement_report_export_as_of_*.xlsx
// ที่ดึงมาจากสำนักงานบัญชี — นำเข้าเฉพาะรายการที่เพิ่มสต็อก
// (ยอดยกมา, ซื้อ, ลดหนี้ฝั่งขาย) ข้ามรายการ ขาย
// ============================================================

const FL_STOCK_PRODUCT_MAP = {
  'P00001': 'กลูต้า',
  'P00002': 'โปรไบโอติก',
  'P00003': 'วิตามินซี',
};

function FL_parseStockMovement(driveFile) {
  const filename  = driveFile.getName();
  const sheetName = FL_getSheetNames(driveFile)[0];
  if (!sheetName) throw new Error('Stock Movement: ไม่พบ sheet ในไฟล์');

  const { rows } = FL_readXlsxSheet(driveFile, sheetName);
  if (rows.length < 9) throw new Error('Stock Movement: ไฟล์มีข้อมูลน้อยเกินไป');

  // ── 1. หา product code จากหัวรายงาน (scan 6 rows แรก เผื่อ xlsx parser ข้าม empty row) ──
  let productCode = null;
  for (let ri = 0; ri < Math.min(rows.length, 6); ri++) {
    const m = (rows[ri] || []).join(' ').match(/P\d{5}/);
    if (m) { productCode = m[0]; break; }
  }
  if (!productCode) throw new Error('Stock Movement: ไม่พบรหัสสินค้าในชื่อรายงาน');
  const category = FL_STOCK_PRODUCT_MAP[productCode];
  if (!category) throw new Error('Stock Movement: ไม่รู้จักรหัสสินค้า ' + productCode);

  // ── 2. หา data start row (หลัง header ที่มี "ลำดับ" หรือ "วันที่") ──
  let dataStart = 8;
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const r = rows[i];
    const a = (r[0] || '').toString();
    const b = (r[1] || '').toString();
    if (a.includes('ลำดับ') || b === 'วันที่') { dataStart = i + 1; break; }
  }

  // ── 3. Parse data rows — นำเข้าเฉพาะ qty > 0 (stock additions) ──
  const COL_DATE = 1, COL_TYPE = 2, COL_REF = 3, COL_QTY = 5;

  const entries = [];
  for (let i = dataStart; i < rows.length; i++) {
    const r       = rows[i];
    const txnType = (r[COL_TYPE] || '').toString().trim();
    if (!txnType) continue;

    const qty = FL_toNum(r[COL_QTY]);
    if (qty <= 0) continue;   // skip ขาย (negative) + zero rows

    // Parse date: Date object OR "DD/MM/YYYY" string
    let date;
    const raw = r[COL_DATE];
    if (raw instanceof Date) {
      date = raw;
    } else {
      const parts = String(raw).split('/');
      if (parts.length === 3) {
        date = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      } else {
        date = new Date(raw);
      }
    }
    if (!date || isNaN(date.getTime())) continue;

    const ref  = (r[COL_REF] || '').toString().trim();
    const note = txnType + (ref ? ' ' + ref : '') + ' | ' + filename;

    entries.push({ date, sku: category, qty, note });
  }

  return { category, productCode, entries, sourceFile: filename };
}
