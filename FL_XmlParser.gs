// ============================================================
// FL_XmlParser.gs — Core XLSX XML Reader + File Type Detector
// ============================================================

// ─────────────────────────────────────────────────────────────
// FILE TYPE DETECTION
// ─────────────────────────────────────────────────────────────

/**
 * อ่านชื่อ sheets ทั้งหมดจากไฟล์ xlsx
 * อ่านแค่ workbook.xml — เบาและเร็ว ไม่ต้อง parse worksheet
 *
 * @param {DriveApp.File} driveFile
 * @returns {string[]}  เช่น ['Summary', 'Income', 'Service Fee Details']
 */
function FL_getSheetNames(driveFile) {
  const blob = driveFile.getBlob().setContentType('application/zip');
  const zipEntries = Utilities.unzip(blob);
  const entryMap = {};
  zipEntries.forEach(e => { entryMap[e.getName()] = e; });

  if (!entryMap['xl/workbook.xml']) return [];

  const wbDoc = XmlService.parse(entryMap['xl/workbook.xml'].getDataAsString());
  const wbNS  = wbDoc.getRootElement().getNamespace();
  const sheetsEl = wbDoc.getRootElement().getChild('sheets', wbNS);
  if (!sheetsEl) return [];

  return sheetsEl.getChildren('sheet', wbNS)
    .map(s => s.getAttribute('name') ? s.getAttribute('name').getValue() : '');
}

/**
 * ตรวจจับประเภทไฟล์จาก sheet names (+ peek cells ถ้าจำเป็น)
 *
 * Detection rules:
 *
 * SHOPEE:
 *   income  → sheets มี "Summary" AND "Income"
 *   orders  → sheets มี "orders"
 *
 * TIKTOK:
 *   income  → sheets มี "Reports" AND "Order details"
 *   orders  → sheets มี "OrderSKUList"
 *   live    → sheets มี "Sheet1" → peek row 3 col A ≈ "ไลฟ์สตรีม"
 *
 * LAZADA:
 *   income  → sheets มี "Income Overview"
 *   orders  → sheets มี "sheet1" → peek row 1 col A = "orderItemId"
 *
 * @param {DriveApp.File} driveFile
 * @param {'shopee'|'tiktok'|'lazada'} platform
 * @returns {'income'|'orders'|'live'|'unknown'}
 */
function FL_detectFileType(driveFile, platform) {
  const sheetNames = FL_getSheetNames(driveFile);
  const has = name => sheetNames.includes(name);

  // ── Shopee ──────────────────────────────────────────────────
  if (platform === 'shopee') {
    if (has('Summary') && has('Income'))  return 'income';
    const orderSheets = ['orders', 'Order Data', 'Sheet1', 'Order', 'Order Export'];
    if (orderSheets.some(s => has(s))) return 'orders';
    return 'unknown';
  }

  // ── TikTok ──────────────────────────────────────────────────
  if (platform === 'tiktok') {
    if (has('Reports') && has('Order details')) return 'income';
    if (has('OrderSKUList') || has('Order List')) return 'orders';
    if (has('Sheet1')) {
      // Peek row 3 col A เพื่อยืนยัน live performance header
      try {
        const { byRowNum } = FL_readXlsxSheet(driveFile, 'Sheet1');
        const row3ColA = ((byRowNum[3] || [])[0] || '').toString();
        if (row3ColA.includes('ไลฟ์สตรีม') || row3ColA.includes('Live')) return 'live';
        // ถ้า Sheet1 แต่มี header คำสั่งซื้อ/SKU -> ให้เป็น orders
        const headerA = ((byRowNum[1] || [])[0] || '').toString();
        if (headerA.includes('Order ID') || headerA.includes('เลขที่คำสั่งซื้อ')) return 'orders';
      } catch (e) { }
      return 'live'; 
    }
    return 'unknown';
  }

  // ── Lazada ──────────────────────────────────────────────────
  if (platform === 'lazada') {
    if (has('Income Overview')) return 'income';
    if (has('sheet1') || has('Sheet1')) {
      const sName = has('sheet1') ? 'sheet1' : 'Sheet1';
      try {
        const { rows } = FL_readXlsxSheet(driveFile, sName);
        const headerA = ((rows[0] || [])[0] || '').toString().trim();
        if (headerA === 'orderItemId') return 'orders';
      } catch (e) { }
      return 'orders';
    }
    return 'unknown';
  }

  return 'unknown';
}

// ─────────────────────────────────────────────────────────────
// CORE XLSX READER
// ─────────────────────────────────────────────────────────────

/**
 * อ่านไฟล์ xlsx และ return rows พร้อม Excel row number
 * รองรับทั้ง sharedStrings และ inlineStr (TikTok / Lazada)
 *
 * @param {DriveApp.File} driveFile
 * @param {string|null}   sheetName  - ชื่อ sheet (null = sheet แรก)
 * @returns {{ rows: string[][], byRowNum: Object.<number, string[]> }}
 */
function FL_readXlsxSheet(driveFile, sheetName) {
  const blob = driveFile.getBlob().setContentType('application/zip');
  const zipEntries = Utilities.unzip(blob);
  const entryMap = {};
  zipEntries.forEach(e => { entryMap[e.getName()] = e; });

  const NS = XmlService.getNamespace(
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  );

  // ─── 1. Load sharedStrings ───────────────────────────────
  const shared = [];
  if (entryMap['xl/sharedStrings.xml']) {
    try {
      const ssDoc = XmlService.parse(
        entryMap['xl/sharedStrings.xml'].getDataAsString()
      );
      ssDoc.getRootElement().getChildren('si', NS).forEach(si => {
        shared.push(si.getValue());
      });
    } catch (e) {}
  }

  // ─── 2. Resolve sheet path ──────────────────────────────
  let targetWsPath = null;
  if (sheetName && entryMap['xl/workbook.xml']) {
    try {
      const wbDoc = XmlService.parse(entryMap['xl/workbook.xml'].getDataAsString());
      const wbNS = wbDoc.getRootElement().getNamespace();
      const relNS = XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
      
      const sheetsEl = wbDoc.getRootElement().getChild('sheets', wbNS);
      let rId = null;
      if (sheetsEl) {
        const els = sheetsEl.getChildren('sheet', wbNS);
        for (let i = 0; i < els.length; i++) {
          const nameAttr = els[i].getAttribute('name');
          if (nameAttr && nameAttr.getValue() === sheetName) {
            const idAttr = els[i].getAttribute('id', relNS);
            if (idAttr) rId = idAttr.getValue();
            break;
          }
        }
      }

      // Parse _rels if rId found
      if (rId && entryMap['xl/_rels/workbook.xml.rels']) {
        const relsDoc = XmlService.parse(entryMap['xl/_rels/workbook.xml.rels'].getDataAsString());
        const relsNS = relsDoc.getRootElement().getNamespace();
        const relEls = relsDoc.getRootElement().getChildren('Relationship', relsNS);
        for (let i = 0; i < relEls.length; i++) {
          const idAttr = relEls[i].getAttribute('Id');
          if (idAttr && idAttr.getValue() === rId) {
            const targetAttr = relEls[i].getAttribute('Target');
            if (targetAttr) {
               let target = targetAttr.getValue();
               if (target.startsWith('/xl/')) targetWsPath = target.substring(1);
               else if (target.startsWith('worksheets/')) targetWsPath = 'xl/' + target;
               else targetWsPath = 'xl/worksheets/' + target;
            }
            break;
          }
        }
      }
    } catch (e) {
      Logger.log('Error resolving rels: ' + e);
    }
  }

  // Fallback if relation not found
  if (!targetWsPath) {
    const allWsKeys = Object.keys(entryMap).filter(k => k.toLowerCase().startsWith('xl/worksheets/') && k.endsWith('.xml'));
    if (allWsKeys.length === 1) {
      targetWsPath = allWsKeys[0];
    } else {
      // Very crude fallback
      targetWsPath = 'xl/worksheets/sheet1.xml';
    }
  }

  // ─── 3. Parse worksheet XML ──────────────────────────────
  const wsKey = targetWsPath;
  if (!entryMap[wsKey]) {
    throw new Error(`FL_readXlsxSheet: ไม่พบ sheet "${sheetName || 'sheet1'}" (mapped to ${wsKey})`);
  }

  const wsDoc = XmlService.parse(entryMap[wsKey].getDataAsString());
  const wsNS  = wsDoc.getRootElement().getNamespace();
  const sheetData = wsDoc.getRootElement().getChild('sheetData', wsNS);
  if (!sheetData) return { rows: [], byRowNum: {} };

  // ─── 3b. Accumulate cells per rowNum ────────────────────
  // Some xlsx files (e.g. TikTok Order export) store each cell in its own
  // <row r="N"> element instead of grouping all cells of a row together.
  // The old approach overwrote byRowNum[N] on every <row> element, leaving
  // only the last cell. Fix: accumulate into a map first, then build arrays.
  const rowAccumulator = {}; // rowNum → { colIdx: value }

  sheetData.getChildren('row', wsNS).forEach(rowEl => {
    const rowAttr = rowEl.getAttribute('r');
    const rowNum  = parseInt(rowAttr ? rowAttr.getValue() : '0');
    if (!rowAccumulator[rowNum]) rowAccumulator[rowNum] = {};

    rowEl.getChildren('c', wsNS).forEach(cEl => {
      const refAttr    = cEl.getAttribute('r');
      const ref        = refAttr ? refAttr.getValue() : '';
      const colLetters = ref.replace(/\d+$/, '');
      const colIdx     = FL_colToIdx(colLetters);
      const typeAttr   = cEl.getAttribute('t');
      const type       = typeAttr ? typeAttr.getValue() : 'n';

      let val = '';
      try {
        if (type === 'inlineStr') {
          const isEl = cEl.getChild('is', wsNS);
          if (isEl) val = isEl.getValue();
        } else if (type === 's') {
          const vEl = cEl.getChild('v', wsNS);
          if (vEl) val = shared[parseInt(vEl.getText())] || '';
        } else {
          const vEl = cEl.getChild('v', wsNS);
          if (vEl) val = vEl.getText() || '';
        }
      } catch (e) { val = ''; }

      rowAccumulator[rowNum][colIdx] = val;
    });
  });

  // Build rows array and byRowNum from accumulator (sorted by row number)
  const rows     = [];
  const byRowNum = {};
  Object.keys(rowAccumulator)
    .map(Number)
    .sort((a, b) => a - b)
    .forEach(rowNum => {
      const rowData = rowAccumulator[rowNum];
      const keys    = Object.keys(rowData);
      if (keys.length === 0) return;
      const maxIdx  = Math.max(...keys.map(Number));
      const row     = Array.from({ length: maxIdx + 1 }, (_, i) =>
        rowData[i] !== undefined ? rowData[i] : ''
      );
      rows.push(row);
      byRowNum[rowNum] = row;
    });

  return { rows, byRowNum };
}

// ─────────────────────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────────────────────

/** Convert column letters to 0-based index: A→0, Z→25, AA→26 */
function FL_colToIdx(col) {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.toUpperCase().charCodeAt(i) - 64);
  }
  return result - 1;
}

/** Build header→colIdx map จาก header row */
function FL_buildHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    if (h) map[h.toString().trim()] = i;
  });
  return map;
}

/** Normalize SKU: FL_SKU_MAP → FL_LAZADA_SKU_MAP → as-is */
function FL_normalizeSKU(sku) {
  if (!sku) return 'UNKNOWN';
  const s = sku.toString().toUpperCase().trim()
    .replace(/\s+/g, '').replace(/\+\s*/g, '+');
  if (FL_SKU_MAP[s]) return FL_SKU_MAP[s];
  // Fallback: try original trimmed value (preserves spaces for Thai keys like 'ตลับวิตามิน 2-1')
  const trimmed = sku.toString().trim();
  if (FL_SKU_MAP[trimmed]) return FL_SKU_MAP[trimmed];
  if (FL_LAZADA_SKU_MAP[trimmed]) return FL_LAZADA_SKU_MAP[trimmed];
  return trimmed;
}

/** Normalize SKU with itemName fallback (ใช้กับ Lazada) */
function FL_normalizeSKUWithFallback(sku, itemName) {
  if (sku) {
    const s = sku.toString().toUpperCase().trim().replace(/\s+/g,'').replace(/\+\s*/g,'+');
    if (FL_SKU_MAP[s]) return FL_SKU_MAP[s];
  }
  if (sku && FL_LAZADA_SKU_MAP[sku.toString().trim()]) {
    return FL_LAZADA_SKU_MAP[sku.toString().trim()];
  }
  if (itemName) {
    const lower = itemName.toString().toLowerCase();
    for (const entry of FL_ITEM_NAME_FALLBACK) {
      if (lower.includes(entry.keyword.toLowerCase())) return entry.skuRef;
    }
  }
  return sku ? sku.toString().trim() : 'UNKNOWN';
}

/** Get product category */
function FL_getCategory(skuRef) {
  return FL_CATEGORY_MAP[skuRef] || 'อื่นๆ';
}

/** Parse float, return 0 if NaN */
function FL_toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(v.toString().replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

/** Extract "YYYY-MM" from filename */
function FL_monthKeyFromFilename(filename) {
  // Support 20260101, 2026-01-01, 2026.01.01, 2026 01 01
  const m1 = filename.match(/(\d{4})[._\-\s]?(\d{2})[._\-\s]?\d{2}/);
  if (m1) return `${m1[1]}-${m1[2]}`;
  
  // Support month names optionally in future, but YYYY-MM is standard for these reports
  return null;
}

/** Extract "YYYY-MM" from date string — supports DD/MM/YYYY (TikTok), DD Mon YYYY (Lazada EN),
 *  Thai month names (Lazada TH), ISO YYYY-MM-DD */
function FL_monthKeyFromDateStr(dateStr) {
  if (!dateStr) return null;
  const s = dateStr.toString().trim();

  // DD/MM/YYYY or DD/MM/YYYY HH:MM:SS  ← TikTok Order "Created Time"
  const dmy = s.match(/^(\d{1,2})\/(\d{2})\/(\d{4})/);
  if (dmy) return `${dmy[3]}-${dmy[2].padStart(2, '0')}`;

  // YYYY/MM/DD  ← TikTok Income "Order settled time"
  const ymdSlash = s.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  if (ymdSlash) return `${ymdSlash[1]}-${ymdSlash[2]}`;

  // English Lazada: 12 Jan 2026
  const lazadaEn = s.match(/\d{1,2}\s+(\w{3})\s+(\d{4})/);
  if (lazadaEn) {
    const MON = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',
                 Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
    return `${lazadaEn[2]}-${MON[lazadaEn[1]] || '01'}`;
  }

  // Thai Lazada: 12 เม.ย. 2026
  const lazadaTh = s.match(/\d{1,2}\s+([\u0E00-\u0E7F.]+)\s+(\d{4})/);
  if (lazadaTh) {
    const MON_TH = {
      'ม.ค.':'01','ก.พ.':'02','มี.ค.':'03','เม.ย.':'04','พ.ค.':'05','มิ.ย.':'06',
      'ก.ค.':'07','ส.ค.':'08','ก.ย.':'09','ต.ค.':'10','พ.ย.':'11','ธ.ค.':'12',
      'มกราคม':'01','เมษายน':'04',
    };
    let monthPart = lazadaTh[1];
    if (!monthPart.endsWith('.')) monthPart += '.';
    const mm = MON_TH[monthPart] || MON_TH[lazadaTh[1]];
    if (mm) return `${lazadaTh[2]}-${mm}`;
  }

  // YYYY-MM-DD or ISO
  const iso = s.match(/(\d{4})-(\d{2})-\d{2}/);
  if (iso) return `${iso[1]}-${iso[2]}`;

  const d = new Date(s);
  if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  return null;
}

/** ค้นหา column index จาก header map โดยลอง candidates หลายชื่อ */
function FL_findCol(hdr, candidates) {
  for (const c of candidates) {
    if (hdr[c] !== undefined) return hdr[c];
    for (const [k, v] of Object.entries(hdr)) {
      if (k.includes(c) || c.includes(k)) return v;
    }
  }
  return -1;
}