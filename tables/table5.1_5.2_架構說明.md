# 表5.1 & 表5.2 腳本架構說明

## 目錄位置

```
D:\OpenCodeLib\supervision_plan\tables\
├── table5.1\
│   ├── 表5.1.docx          ← Word 模板（僅提供標題列樣式）
│   ├── convert_5.1.py      ← 材料設備送審管制總表轉換腳本
│   ├── convert_5.1_v2.py   ← v2 版（不預填預設值）
│   └── SCOPE_5.1.md
├── table5.2\
│   ├── 表5.2.docx          ← Word 模板
│   ├── convert_5.2.py      ← 材料設備檢(試)驗管制總表轉換腳本
│   └── SCOPE_5.2.md
└── common\
    └── docx_table.py       ← 共用表格元件（add_cell）
```

---

## 三層架構

```
Excel 母本 (詳細價目表.xlsx / Table 1)   ← 資料源
     ↓ 腳本讀取、篩選、計算
Python 腳本 (convert_5.1.py / convert_5.2.py)   ← 控制中樞
     ↓ 格式、分頁、填入
DOCX 模板 (表5.1.docx / 表5.2.docx)   ← 僅提供標題列外觀樣式
```

模板角色非常輕，貢獻 **標題列（前 2 列）XML** 後 deepcopy 貼上，其餘全部由腳本用 `etree.SubElement` 從零重建。

---

## 腳本共同流程

### Step 1 — 讀取 Excel 母本 (`load_price_sheet`)
- 開啟 `詳細價目表.xlsx` / `Table 1` 工作表
- 起點：項次 `壹.三.1` 之後
- 只取 `壹.` 開頭項目
- 排除指定單位（5.1：`式` `工`；5.2：`工`）
- 排除名稱含 `小計` `合計` `總價` 的列
- ⑤ → 窗（修正 PDF 轉檔亂碼）
- 回傳 `[(項次, 材料名稱, 數量+單位), ...]`

### Step 2 — 計算列高 (`calc_name_lines` + `calc_row_height`)
- 用 **Pillow** 量測項次/材料名稱的文字寬度
- 推算需要幾行：`ceil(文字寬度_twip / 欄位寬度_twip)`
- 行高 = `max(226, 行數 × 240 + 0)` twip（atLeast）

### Step 3 — 分頁規劃 (`page_plan`)
- 依每組（奇數列+偶數列）實際高度累加
- 超過可用高度則換頁：
  - 5.1：`11170 − 2600 ≈ 8570 twip`
  - 5.2：`10600 − 1100 ≈ 9500 twip`
- 算出每頁放幾組，記錄在 `page_plan` 陣列

### Step 4 — 清空模板 body，逐頁重建
- 從模板取出前 2 列（標題列）XML 備份
- 備份 `sectPr`（版面設定）
- 刪除 body 所有子元素

### Step 5 — 逐頁輸出
1. 第 2 頁起插入**分頁符號**
2. 寫入**三行表頭**：
   - `表5-N 材料設備...總表-{頁碼}`（14pt、粗體、置中）
   - `工程名稱：臺南市政府社會局...`
   - `(監造單位使用) 第N頁共Σ頁 表單編號：E5N-N`
3. 用 `etree.SubElement` 建立新表格 + `tblGrid`（設定 `COL_W` 欄寬）
4. 貼上標題列（deepcopy）
5. 逐組填入資料列：
   - **奇數列**：流水號、項次、數量單位 + 預設值（施作前14日、否、V...）
   - **偶數列**：材料名稱
   - 格式由 `add_cell()` / `make_cell()` 控制：
     - 黑實線框（single, sz=8, color=000000）
     - 段落 exact 240 twip
     - 儲存格邊距 tcMar=0
     - 垂直置中
     - 標楷體 11pt（5.2 部分用 Times New Roman）
   - 垂直合併由 `DAT_MERGE` 控制

### Step 6 — 存檔
- 補回 `sectPr`，輸出 `.docx`

---

## 5.1 vs 5.2 關鍵差異

| 項目 | 表5.1 | 表5.2 |
|------|-------|-------|
| 欄位數 | 15 欄 | 10 欄 |
| 欄寬 COL_W | 15 個值 | 10 個值 |
| 排除單位 | `式` `工` | `工` |
| 垂直合併 DAT_MERGE | `{0,2,3,6,7,8,9,10,11,12,14}` | `{0,3,5,7,8}` |
| 預設填入 | C4=施作前14日, C5=否, C7=V, C8=V | 無預設值 |
| C5 抽樣頻率 | 無此欄 | 每頁第一組=「進場時至少一次」，其餘=「至少一次」 |
| 列高計算 | 只算 C1（材料名稱） | C1 + C5（抽樣頻率）取最大 |
| 表格寬度 | 5289 pct | 4771 pct |
| 共用元件 | 使用 `common.docx_table.add_cell` | 內建 `make_cell`（邏輯相同） |

---

## 腳本參數

```bash
# 表5.1
python -X utf8 tables/table5.1/convert_5.1.py --exclude-units 式 工

# 表5.2
python -X utf8 tables/table5.2/convert_5.2.py --exclude-units 工

# 共同參數
-p, --price     詳細價目表路徑
-t, --template  模板路徑
-o, --output    輸出路徑
--test-num      測試流水號（插入檔名 _test_N）
--max-pairs     每頁最多組數（預設 20）
--max-pages     最多頁數（預設 0=不限）
```