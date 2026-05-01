// ============================================================
// FL_Config.gs — First Love Sales Automation
// ============================================================

const FL_DEFAULT_CONFIG = {
  // ─── Google Drive Folder IDs (แค่ 3 folder แยกตาม platform) ──
  // วางไฟล์ทุกประเภท (income/orders/live) ไว้ใน folder ของ platform นั้น
  // ระบบจะตรวจจับ file type เองจาก sheet names
  FOLDERS: {
    SHOPEE: '1-h66Z56evcoIOZ3L0dFxtxlxZC_bycQL',   // Income + Orders ของ Shopee รวมกัน
    TIKTOK: '1g6gdFpVjnPCYrsjhy2Oe1c-RhqTjvczQ',   // Income + Orders + Live Performance ของ TikTok รวมกัน
    LAZADA: '1FpJWRbVHbDoiKuj9kwu270nQ5FBxi6cC',   // Income + Orders ของ Lazada รวมกัน
  },
  OUTPUT_SHEET_ID: '1iG41vcWgMV3E5gB8z9VczXD_IoKODRMMSG3LooAeZoQ',
  MC_LOG_SHEET_ID: '1u8AgUZ-oG07rJCgTBpm4mOthU4PFsxcTwEfL2Fs7-0w',
  // Auto‑scan interval configuration – default 1 hour
  triggerInterval: 1,
  triggerUnit: 'hours',
  // เป้าหมายยอดโอนสุทธิรายเดือน (บาท) — แก้ได้จาก Index.html → Config
  MONTHLY_TARGET: 2000000,
  // Stock low-stock threshold (REQ-05) — แก้ได้จาก Index.html → Config
  STOCK_THRESHOLD: 50,
};

/**
 * อ่านค่า config จาก Script Properties
 * ถ้ายังไม่มีค่า ให้ใช้ FL_DEFAULT_CONFIG
 */
function FL_getConfig() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('FL_CONFIG');
  if (stored) {
    try { return JSON.parse(stored); } catch (e) {}
  }
  return FL_DEFAULT_CONFIG;
}

/**
 * บันทึก config ลง Script Properties
 * เรียกจาก Index.html → updateConfig()
 */
function updateConfig(newCfg) {
  PropertiesService.getScriptProperties()
    .setProperty('FL_CONFIG', JSON.stringify(newCfg));
  // ถ้ามีการระบุ Output Sheet ให้ Auto-Setup โครงสร้าง Tab อัตโนมัติเผื่อเป็นไฟล์ใหม่
  if (newCfg.OUTPUT_SHEET_ID) {
    try {
      FL_setupSheets();
    } catch(e) {
      Logger.log('Auto setup sheets failed (No access to ID or invalid ID): ' + e);
    }
  }
  return { ok: true, config: newCfg };
}

/**
 * คืน config ปัจจุบัน — เรียกจาก Index.html → getCurrentConfig()
 */
function getCurrentConfig() {
  return FL_getConfig();
}

// ─── Output sheet names ──────────────────────────────────────
const FL_SHEETS = {
  MONTHLY_SUMMARY: 'monthly_summary',
  SKU_MONTHLY:     'sku_monthly',
  SKU_MASTER:      'sku_master',
  PRODUCT_NAMES:   'product_names',
  PROCESSED_FILES: 'processed_files',
  ERROR_LOG:       'error_log',
  STOCK_IN:        'stock_in',
  AD_SPEND:        'ad_spend',
};

// Column headers for output sheets
const FL_HEADERS = {
  MONTHLY_SUMMARY: [
    'month_key','platform','gross','seller_discount',
    'net_revenue','platform_fees','transferred','imported_at','source_file'
  ],
  SKU_MASTER: [
    'sku_ref', 'display_name', 'brand', 'category', 'box_size', 'commission_per_unit', 'updated_at'
  ],
  SKU_MONTHLY: [
    'month_key','platform','sku_ref','category',
    'units_sold','revenue','imported_at','source_file','year'
  ],
  PRODUCT_NAMES:   ['sku_ref', 'display_name', 'status', 'updated_at'],
  PROCESSED_FILES: ['file_id','file_name','platform','file_type','month_key','imported_at'],
  ERROR_LOG:       ['timestamp','file_name','platform','error'],
  STOCK_IN:        ['date','sku_ref','quantity','note','recorded_at'],
  AD_SPEND:        ['month_key','amount','recorded_at'],
};

// ─── Lazada valid order statuses ─────────────────────────────
const FL_LAZADA_VALID_STATUSES = new Set([
  'confirmed', 'delivered', 'shipped', 'topay', 'packed', 
  'ยืนยันแล้ว', 'จัดส่งแล้ว', 'ส่งมอบแล้ว', 'เตรียมจัดส่ง'
]);

// ─── SKU Normalization Map ────────────────────────────────────
// Updated 2026-04-11: added GlutaX9-3, GlutaX9-6, Bundle_Gluta1+VitC1,
//   Bundle_Capsule_2in1, Bundle_Capsule_3 found in FirstLove 2025.xlsx
const FL_SKU_MAP = {
  'GLUTAX9-1':                         'GlutaX9-1',
  'GLUTAX9-2':                         'GlutaX9-2',
  'GLUTAX9-3':                         'GlutaX9-3',
  'GLUTAX9-4':                         'GlutaX9-4',
  'GLUTAX9-6':                         'GlutaX9-6',
  'GLUTAX9-2+VITC-1':                  'Bundle_Gluta2+VitC1',
  'GLUTAX9-1+VITC-1':                  'Bundle_Gluta1+VitC1',
  'GLUTAX9-2+VITC-1+PROBIOTIC-1':      'Bundle_Triple',
  'GLUTAX9-2+VITC-1+PROBIOTIC-1+ขวด': 'Bundle_Triple_Bottle',
  'GLUTAX9-4+VITC-2':                  'Bundle_Gluta4+VitC2',
  'GLUTAX9-2+PROBIOTIC-1':             'Bundle_Gluta2+Probiotic1',
  'VITC-1':      'VitC-1',
  'VITC-2':      'VitC-2',
  'VITC-3':      'VitC-3',
  'PROBIOTIC-1': 'Probiotic-1',
  'PROBIOTIC-2': 'Probiotic-2',
  // ตลับวิตามิน = vitamin capsule bundle sets (found in 2025 TikTok data)
  'ตลับวิตามิน 2-1': 'Bundle_Capsule_2in1',
  'ตลับวิตามิน 3':   'Bundle_Capsule_3',
};

// ─── Lazada-specific SKU map ──────────────────────────────────
const FL_LAZADA_SKU_MAP = {
  'BDS49520254': 'Bundle_Gluta2+VitC1',   // Set 1 เดือน กลูต้า2+วิตซี1
  'BDS49520252': 'GlutaX9-6',             // เซต 3 เดือน กลูต้า6 (found 2025)
};

// ─── itemName fallback (keyword → canonical ref) ─────────────
const FL_ITEM_NAME_FALLBACK = [
  { keyword: 'เซต 3 เดือน',     skuRef: 'GlutaX9-6'           },
  { keyword: 'set 3 เดือน',     skuRef: 'GlutaX9-6'           },
  { keyword: 'set 1 เดือน',     skuRef: 'Bundle_Gluta2+VitC1' },
  { keyword: 'set 2 เดือน',     skuRef: 'Bundle_Gluta4+VitC2' },
  { keyword: 'กลูต้า2+vitc',    skuRef: 'Bundle_Gluta2+VitC1' },
  { keyword: 'กลูต้า4+vitc',    skuRef: 'Bundle_Gluta4+VitC2' },
  { keyword: 'probiotic-2',     skuRef: 'Probiotic-2'          },
  { keyword: 'probiotic-1',     skuRef: 'Probiotic-1'          },
  { keyword: 'probiotic',       skuRef: 'Bundle_Triple'        },
  { keyword: 'กลูต้า 6 กระปุก', skuRef: 'GlutaX9-6'           },
  { keyword: 'กลูต้า 4 กระปุก', skuRef: 'GlutaX9-4'           },
  { keyword: 'กลูต้า 3 กระปุก', skuRef: 'GlutaX9-3'           },
  { keyword: 'กลูต้า 2 กระปุก', skuRef: 'GlutaX9-2'           },
  { keyword: 'กลูต้า',          skuRef: 'GlutaX9-1'           },
  { keyword: 'vitamin c',       skuRef: 'VitC-1'               },
  { keyword: 'วิตามินซี',       skuRef: 'VitC-1'               },
];

// ─── SKU Display Name Map (ชื่อแสดงภาษาไทย/อ่านง่าย) ────────
// ใช้ใน Dashboard เพื่อแสดงชื่อสินค้าที่เข้าใจง่ายแทนรหัส SKU
const FL_SKU_DISPLAY_MAP = {
  'GlutaX9-1':            'กลูต้า 1 กระปุก',
  'GlutaX9-2':            'กลูต้า 2 กระปุก',
  'GlutaX9-3':            'กลูต้า 3 กระปุก',
  'GlutaX9-4':            'กลูต้า 4 กระปุก',
  'GlutaX9-6':            'กลูต้า 6 กระปุก',
  'Bundle_Gluta2+VitC1':  'กลูต้า 2 + วิตซี 1',
  'Bundle_Gluta1+VitC1':  'กลูต้า 1 + วิตซี 1',
  'Bundle_Gluta4+VitC2':  'กลูต้า 4 + วิตซี 2',
  'Bundle_Triple':        'กลูต้า 2 + วิตซี 1 + โปรไบโอติก',
  'Bundle_Triple_Bottle':       'กลูต้า 2 + วิตซี 1 + โปรไบโอติก + ขวด',
  'Bundle_Gluta2+Probiotic1':  'กลูต้า 2 + โปรไบโอติก 1',
  'Bundle_Capsule_2in1':  'ตลับวิตามิน 2in1',
  'Bundle_Capsule_3':     'ตลับวิตามิน 3',
  'VitC-1':               'วิตซี 1',
  'VitC-2':               'วิตซี 2',
  'VitC-3':               'วิตซี 3',
  'Probiotic-1':          'โปรไบโอติก 1',
  'Probiotic-2':          'โปรไบโอติก 2',
};

// ─── Category map ─────────────────────────────────────────────
const FL_CATEGORY_MAP = {
  'GlutaX9-1':            'กลูต้า',
  'GlutaX9-2':            'กลูต้า',
  'GlutaX9-3':            'กลูต้า',
  'GlutaX9-4':            'กลูต้า',
  'GlutaX9-6':            'กลูต้า',
  'Bundle_Gluta2+VitC1':  'เซต',
  'Bundle_Gluta1+VitC1':  'เซต',
  'Bundle_Triple':        'เซต',
  'Bundle_Triple_Bottle':      'เซต',
  'Bundle_Gluta2+Probiotic1': 'เซต',
  'Bundle_Gluta4+VitC2':  'เซต',
  'Bundle_Capsule_2in1':  'เซต',
  'Bundle_Capsule_3':     'เซต',
  'VitC-1':               'วิตามินซี',
  'VitC-2':               'วิตามินซี',
  'VitC-3':               'วิตามินซี',
  'Probiotic-1':          'โปรไบโอติก',
  'Probiotic-2':          'โปรไบโอติก',
};