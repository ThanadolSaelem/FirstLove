# First Love Project — Full Context (Single File)
> สร้าง 2026-05-01 | ใช้เป็น context ใน session ใหม่ได้ทันที

---

## 1. PROJECT SUMMARY

"First Love" คือธุรกิจอาหารเสริมที่ขายบน 3 platform: **Shopee, TikTok, Lazada**
ระบบคือ Google Apps Script (GAS) รัน import + dashboard บน Google Sheets/Drive

- **GitHub:** https://github.com/ThanadolSaelem/FirstLove.git (uploaded 2026-05-01, user: ThanadolSaelem)
- **Working Dir:** `c:\Users\Mynew\Downloads\firstlove`
- **Output Sheet:** ตั้งค่าใน `FL_Config.gs` → `OUTPUT_SHEET_ID`
- **Web App URL:** `?page=dashboard` (Executive), `?page=annual` (Annual), `?page=import` (Index/Admin)

---

## 2. FILES หลัก

| ไฟล์ | หน้าที่ |
|---|---|
| `FL_Main.gs` | doGet/doPost, `FL_getExecutiveDashboardData()`, `FL_getAnnualDashboardData()` |
| `FL_DataStore.gs` | read/write sheets: monthly_summary, sku_monthly, stock_in, ad_spend |
| `FL_Config.gs` | FL_SHEETS, FL_HEADERS constants; `FL_getConfig()` |
| `FL_Parser_Shopee.gs` | parse Shopee Income + Order xlsx |
| `FL_Parser_TikTok.gs` | parse TikTok Income + Order xlsx |
| `FL_Parser_Lazada.gs` | parse Lazada Income + Order xlsx |
| `Dashboard_Executive.html` | Executive dashboard |
| `Dashboard_Annual.html` | Annual dashboard |
| `Index.html` | Import + Admin (stock-in form, ad spend form, setup/scan) |

---

## 3. OUTPUT SHEETS

| Sheet | Columns |
|---|---|
| `monthly_summary` | month_key, platform, gross, seller_discount, net_revenue, platform_fees, transferred |
| `sku_monthly` | month_key, platform, sku_ref, category, units_sold, revenue, year |
| `stock_in` | date, sku_ref, quantity, note |
| `ad_spend` | month_key, amount, recorded_at |
| `processed_files` | fileId, fileName, platform, fileType, month_key |
| `sku_master` | sku_ref, display_name, box_size, commission |
| `error_log` | date, fileName, platform, error |

---

## 4. DATA FLOW

```
Google Drive folders (Shopee / TikTok / Lazada)
  ↓ FL_scanAndImportAll()
  ↓ FL_detectAndImport(file, platform)
  ↓ FL_parseXxxIncome() + FL_parseXxxOrder()
  ↓ FL_writeMonthlySummary() + FL_writeSkuMonthly()
  → monthly_summary + sku_monthly sheets
  ↓ FL_getExecutiveDashboardData(monthKey, platform)
  → Dashboard_Executive.html renders
```

---

## 5. PARSER — การดึงค่าสำคัญ (ถูกต้อง ณ ปัจจุบัน)

### TikTok (FL_Parser_TikTok.gs)
อ่านจาก sheet `Reports`:

| field | keyword / วิธี |
|---|---|
| `gross` | `findValue(['Subtotal before discounts'])` |
| `seller_discount` | `findExact(['Seller discounts'])` ← exact match ป้องกันชนกับ "Subtotal after seller discounts" |
| `platform_fees` | `findExact(['Total Fees', 'ค่าธรรมเนียมรวม'])` ← อ่านตรง ห้ามใช้ plug |
| `transferred` | `findValue(['Total settlement amount'])` |
| `month_key` | อ่านจาก "Order settled time" ใน sheet **Order details** — ไม่ใช้ filename เพราะ filename = export date ≠ settle month |

### Shopee (FL_Parser_Shopee.gs)
อ่านจาก sheet `Summary`:

| field | keyword / วิธี |
|---|---|
| `gross` | `findValue(['สินค้าราคาปกติ'])` |
| `seller_discount` | `findValue(['ส่วนลดสินค้าจากผู้ขาย'])` |
| `transferred` | `findValue(['จำนวนเงินทั้งหมดที่โอนแล้ว'])` |
| `platform_fees` | `findValue(['ค่าธรรมเนียม'])` → col D = section total ← อ่านตรง ห้ามใช้ plug |
| fallback fees | `transferred - gross - discount - shopee_subsidy` (ใช้เมื่อ label หาไม่เจอ) |
| `month_key` | จากชื่อไฟล์ (Shopee filename มีวันที่ช่วงที่ถูกต้อง) |

Shopee `findValue` มี guard พิเศษ: `!/^-?\d+(\.\d+)?$/.test(val)` เพราะ cell "3. จำนวน..." จะทำให้ parseFloat = 3

### Lazada Income (FL_Parser_Lazada.gs)
อ่านจาก sheet `Income Overview` ระดับ transaction:

| field | วิธี |
|---|---|
| `gross` | rows type `"ยอดรวมค่าสินค้า"` |
| `platform_fees` | `paymentFee + sellingFee + premiumFee` (ทุกตัวติดลบ) |
| `transferred` | sum ทุก transaction per order |
| `month_key` | จาก filename หรือ order date column |

Lazada ถูกต้องตั้งแต่แรก ไม่มี plug method

---

## 6. DASHBOARD EXECUTIVE — Layout & JS

### Filters
- **เดือน:** `<select id="selMonth">` option แรก = "ทั้งหมด" (value="all"), เดือนล่าสุดก่อน
- **Platform:** `<select id="selPlatform">` — all / tiktok / shopee / lazada

### Layout (ROW 1–4)
```
ROW 1 [2col]
  LEFT  card "ภาพรวมรายเดือน"
    #kFees (ค่าธรรมเนียมรวม), #kNet (ยอดขายรวม), #bNet (badge)
    #chartWaterfall (height:180px)
  RIGHT card "รายรับจริง"
    #revMonthLbl, #revBig, #revMomBadge
    hr.rev-divider
    #tgtVal (เป้าหมาย), #tgtBadge, #tgtNote
    border-top → #chartCatPlatform (height:200px)

ROW 2 [2col]
  LEFT  card "สินค้าขายดี"
    table tbody#topProdBody
  RIGHT card "สัดส่วนยอดขาย & ค่า Platform Fee"
    flex row:
      LEFT (42%): #chartDonut (height:170px) + #platValRow
      RIGHT (58%): #chartFeeBar (height:150px) + #feeRateRow (HTML % rate)

ROW 3 [2col]: MC | Stock
ROW 4 [2col]: Insights | Ad Spend
```

### JS Variables
```javascript
let DATA = null;       // response จาก FL_getExecutiveDashboardData()
let MONTH_KEY = '';    // 'YYYY-MM' หรือ 'all'
let PLATFORM = 'all'; // 'all' | 'shopee' | 'tiktok' | 'lazada'
```

### JS Functions หลัก
| Function | หน้าที่ |
|---|---|
| `loadData(monthKey)` | เรียก backend FL_getExecutiveDashboardData |
| `buildMonthDropdown(allMonths, activeKey)` | build `<select id="selMonth">` |
| `selectMonth(monthKey)` | set MONTH_KEY + loadData() |
| `renderAll()` | render ทุก section จาก DATA |
| `renderFeeBar(byPlatform, platforms)` | grouped bar 3 แท่ง + feeRateRow HTML |
| `renderWaterfall(wf)` | Waterfall chart |
| `renderDonut(platforms, byP)` | Pie chart ยอดขายสัดส่วน |
| `renderCatPlatformChart(catPlat)` | grouped bar หมวด × platform |

### isAllMode Logic
```javascript
// backend
const isAllMode = (monthKey === 'all');
const sumField = (month, platform, field) =>
  rows.filter(r => (month === 'all' || r[0] === month) && ...)
      .reduce((s, r) => s + (parseFloat(r[idx(field)]) || 0), 0);

// frontend
const isAllMode = DATA.isAllMode || MONTH_KEY === 'all';
// ถ้า isAllMode: ซ่อน MoM badge, ซ่อน tgtNote, revMonthLbl = 'ยอดขายสะสม'
```

### chartFeeBar — ล่าสุด
- **grouped bar** 3 แท่ง (TikTok/Shopee/Lazada) ไม่ stack
- annotation บนแท่ง = ยอดบาท
- `#feeRateRow` HTML ใต้กราฟ = % rate font-size:13px font-weight:700 สีตาม platform
  (Google Charts ไม่รองรับ 2-line annotation จึงใช้ HTML overlay แทน)

---

## 7. DASHBOARD ANNUAL — Layout

```
Card 1: ภาพรวมรายได้รายเดือน (chartMonthBar = ComboChart bar+line Ad Spend)
        + Contribution Margin Ranking [side-by-side .chart-pair-equal]
Card 2: รายได้ตาม Platform รายเดือน Stacked (chartPlatformBar)
        + สัดส่วนค่าธรรมเนียม (chartFeePie + platLegend) [side-by-side]
Card 3: Heatmap
Card 4: Product Lifecycle Quadrant (กรอง เซต/Bundle ออก)
```

**feeByPlatform shape:** `{ amount: number, rate: number }` (ไม่ใช่ number เดี่ยว)
**monthly[m].adSpend:** ใส่ใน monthly array สำหรับ ComboChart line series

---

## 8. AD SPEND FEATURE

- Sheet `ad_spend`: month_key, amount, recorded_at — upsert by month_key
- บันทึก: Index.html → `FL_saveAdSpend(monthKey, amount)`
- อ่าน: `FL_getAdSpend(monthKey)` / `FL_getAdSpendAnnual(year)`
- Executive ROW 4 right: แสดง ROAS = transferred / adSpend
- Annual: ComboChart line series overlay บน chartMonthBar

---

## 9. สิ่งที่ถูกลบออกแล้ว (ห้าม re-add)

| สิ่งที่ลบ | เหตุผล |
|---|---|
| Commission KPI card (Executive) | ลบตาม Phase A 2026-04-30 |
| chartRpu / จำนวนชิ้นที่ขายได้รายเดือน (Executive) | ซ้ำกับ chartCatPlatform |
| chartSkuBar (Annual) | Phase A |
| bundleBody card (Annual) | Phase A |
| Area chart / chartLine (Annual) | Phase A |
| Month pills buttons | เปลี่ยนเป็น dropdown แล้ว |
| feeTrend chart (6 เดือน stacked) | เปลี่ยนเป็น grouped bar เดือนที่เลือก |

---

## 10. VERIFIED NUMBERS จาก First love 2026 (mike).xlsx

**TikTok ม.ค. 2026:**
- transferred: 97,411.86
- `Total Fees` (ถูก): **-22,580.40**
- plug method (ผิด): -32,862.54 (ต่าง 10,282)

**Shopee (3 เดือน verified):**
- ม.ค.: ถูก = **-18,107** | ผิด = -17,231 (ต่าง 876 = subsidy)
- ก.พ.: ถูก = **-14,503** | ผิด = -13,580 (ต่าง 923 = subsidy)
- มี.ค.: ถูก = **-14,865** | ผิด = -17,935 (ต่าง 3,070 = refund ถูกดูดเข้า plug)

**Lazada:** ถูกต้องตั้งแต่แรก ไม่มีบั๊ก

---

## 11. CALCULATION BUGS & FIXES — ประวัติบั๊ก 9 ตัว

### BUG 1 — TikTok fees ผิด (Plug Method)
**อาการ:** fees สูงเกิน 10,282 บาท
**สาเหตุ:** `platform_fees = transferred - gross - seller_discount` ดูด Affiliate commission, EAMS fee, adjustments เข้ามาด้วย
**แก้:** `findExact(['Total Fees'])` อ่านตรงจาก Reports sheet
**อุทาหรณ์:** ห้ามใช้ plug method กับ TikTok แม้จะดูสมเหตุสมผล

### BUG 2 — Shopee fees ผิด (Plug Method + Subsidy + Refund)
**อาการ:** fees ผิดทุกเดือน ทั้งสูงเกินและต่ำเกิน
**สาเหตุ:** plug ไม่รู้จัก Shopee-funded subsidy (บวก) และ refund (ลบ) ทำให้คำนวณ fees ผิดทิศทาง
**แก้:** อ่าน "ค่าธรรมเนียม" section total จาก col D ของ Summary sheet ตรงๆ; fallback ถึงจะใช้ plug+subsidy
**อุทาหรณ์:** เดือนที่มี refund → plug fees สูงเกิน, เดือนที่มี subsidy → plug fees สูงเกินเช่นกัน

### BUG 3 — TikTok month_key ผิด (Filename = Export Date)
**อาการ:** ข้อมูลม.ค. import เป็นเม.ย. เพราะไฟล์ export วันที่ 2026-04-07
**สาเหตุ:** TikTok filename date = วันที่ export ≠ วันที่ settle
**แก้:** อ่าน "Order settled time" จาก sheet Order details แทน; fallback ถึงใช้ filename
**repair:** รัน `FL_repairTikTokIncomeMonthKeys()` สำหรับข้อมูลเก่า
**อุทาหรณ์:** Shopee filename ถูก (มีวันที่ช่วง), Lazada filename เป็น hash (ดูจาก data row), TikTok ห้ามใช้ filename เป็น primary

### BUG 4 — Shopee findValue ดึงตัวเลขผิด Row
**อาการ:** `findValue(['ค่าธรรมเนียม'])` คืนค่า 3 แทนค่าจริง
**สาเหตุ:** Shopee Summary sheet มี cell "3. จำนวนเงิน..." → `parseFloat("3. จำนวน...") = 3`
**แก้:** เพิ่ม regex guard: `!/^-?\d+(\.\d+)?$/.test(val.toString().replace(/,/g,'').trim())`
**อุทาหรณ์:** TikTok Reports sheet ไม่มีปัญหานี้ แต่ Shopee มีเสมอ

### BUG 5 — Duplicate Check ทำให้ข้อมูลเก่าไม่ถูก Overwrite
**อาการ:** แก้ parser แล้วตัวเลข dashboard ยังเหมือนเดิม
**สาเหตุ:** `FL_writeMonthlySummary()` skip ถ้า (month_key, platform) มีอยู่แล้ว
**แก้:** ลบแถวเก่าจาก monthly_summary sheet ก่อน re-scan
**อุทาหรณ์:** ทุกครั้งที่แก้ parser ต้องเตือนลูกค้าเรื่องนี้ ก่อนบอกว่า "แก้แล้ว"

### BUG 6 — CacheService ทำให้ Dashboard ไม่อัปเดต
**อาการ:** แก้ sheet แล้ว reload ยังเห็นเลขเก่า
**สาเหตุ:** cache TTL 300 วินาที (5 นาที); key = `exec_{monthKey}_{platform}`
**แก้:** รอ 5 นาที หรือรัน `FL_clearDashboardCache()` จาก Apps Script editor
**อุทาหรณ์:** debug ตัวเลขผิดก่อนอื่นต้องถาม "clear cache แล้วหรือยัง"

### BUG 7 — feeByPlatform Shape เปลี่ยน แต่ Consumer ไม่อัปเดต
**อาการ:** Annual Fee Pie chart render ผิดหลังเพิ่ม Fee Rate %
**สาเหตุ:** backend เปลี่ยน shape จาก `number` → `{ amount, rate }` แต่ `renderFeePie()` ยังอ่านเป็น number
**แก้:** `renderFeePie()` รองรับ 2 shape ด้วย `typeof feeByPlatform[p] === 'object'`
**อุทาหรณ์:** เปลี่ยน response shape ต้องตรวจทุก consumer (Executive + Annual)

### BUG 8 — DOM null Error (element ใน Hidden Div)
**อาการ:** `TypeError: Cannot set properties of null (setting 'innerText') at renderAll`
**สาเหตุ:** ระหว่างแก้ layout `kFees/kNet/bNet/chartWaterfall` หลุดเข้าไปใน `<div style="display:none">`
**แก้:** ย้าย elements ออกมาอยู่ใน visible DOM ของ ROW 1 LEFT card
**อุทาหรณ์:** หลังแก้ HTML layout ให้ grep element ID ทุกตัวที่ JS เขียนถึง แล้วตรวจว่าไม่อยู่ใน hidden container

### BUG 9 — Google Charts ไม่รองรับ 2-Line Annotation
**อาการ:** `"43,846\n7.0%"` บน bar annotation แสดงเป็นบรรทัดเดียว; `annotationText` role เป็นแค่ tooltip hover
**สาเหตุ:** Google Charts ColumnChart annotation ไม่รองรับ multiline text
**แก้:** ใช้ HTML `<div id="feeRateRow">` ใต้กราฟแสดง % rate แทน annotation
**อุทาหรณ์:** ต้องการ custom text บน Google Charts → ใช้ HTML overlay ดีกว่าพยายาม hack annotation

---

## 12. IMPORTANT CAVEATS

- **Duplicate check:** `FL_writeMonthlySummary()` skip ถ้ามีแถว (month_key, platform) อยู่แล้ว → ต้องลบแถวก่อน re-import เสมอ
- **Cache:** CacheService 5 นาที — ต้อง clear หรือรอก่อน debug ตัวเลข
- **`FL_repairTikTokIncomeMonthKeys()`:** รันครั้งเดียวจาก Apps Script editor เพื่อซ่อม month_key เก่าของ TikTok
- **`FL_backfillSkuMonthlyYear()`:** รันครั้งเดียวเพื่อเติม year column ใน sku_monthly rows เก่า

---

## 13. PRINCIPLES สรุป

1. **ห้ามใช้ plug method สำหรับ fees** ทุก platform มีค่าพิเศษที่ plug ดูดเข้ามาผิด
2. **อ่าน fees ตรงจาก summary label** — TikTok: "Total Fees", Shopee: "ค่าธรรมเนียม" col D
3. **TikTok filename = export date ≠ settle month** — อ่านจาก Order details เสมอ
4. **แก้ parser → ต้อง re-import** — duplicate check ป้องกัน overwrite อัตโนมัติ
5. **Cache 5 นาที** — clear ก่อน debug ตัวเลขเสมอ
6. **เปลี่ยน response shape → ตรวจทุก consumer** — Executive + Annual dashboard
7. **HTML elements ต้องไม่อยู่ใน hidden container** — grep ก่อน/หลังแก้ layout
8. **Google Charts annotation ไม่รองรับ multiline** — HTML overlay แทน
