# 材料設備送審管制總表轉換 — 計畫說明

## 目的

將 **詳細價目表（Excel）** 作為母本，填入 **材料設備送審管制總表（表5）Word 模板**，產出完整的多頁 Word 文件。

## 輸入

| 檔案 | 說明 |
|------|------|
| `02_成德-詳細價目表.xlsx` | 詳細價目表母本（Table 1 工作表） |
| `表5.1.docx` | 表5 模板（位於 `supervision_plan/table5.1_converter/`，15 欄、2 列標題 + 1 組空白資料列） |

## 處理規則

### 資料篩選
- 起點：項次 `壹.三.1` 之後
- 項次須以 `壹.` 開頭
- 排除單位：`式`、`工`（可自訂）
- 排除名稱含 `小計`、`合計`、`總價` 之項目
- ⑤ → 窗（修正 PDF 轉檔錯誤）

### 分頁
- 動態分頁，依各資料列實際高度（Pillow 測量字寬 + 行高校準）決定每頁可容納組數
- 每頁總高度上限：11500 twip（標題列 1900 + 資料列 9600）
- 每頁之間插入分頁符號

### 標題列合併（R1 + R2）
- **合併**：C0（項次）、C2（契約數量）、C3（是否取樣試驗）、C6（預定試驗單位）、C14（備註）
- **不併**：C1（項次/名稱）、C4（預定送審日期/實際送審日期）、C5（是否驗廠/驗廠日期）、C7~C12（送審資料各欄）、C13（審查日期/審查結果）

### 資料列合併（奇數列 + 偶數列）
- **合併**：C0（流水號）、C2（數量+單位）、C3、C6、C7~C12（送審資料）、C14（備註）
- **不併**：C1（奇=項次、偶=名稱）、C4、C5、C13

## 輸出

| 產出 | 說明 |
|------|------|
| `表5_完成3.docx` | 完整多頁材料設備送審管制總表 |

## 使用方式

```bash
# 基本用法（使用預設路徑）
python -X utf8 supervision_plan/supervision_plan/table5.1_converter/convert_5.1.py

# 指定檔案
python -X utf8 supervision_plan/supervision_plan/table5.1_converter/convert_5.1.py \
    -p 02_成德-詳細價目表.xlsx \
    -t 表5.1.docx \
    -o output.docx

# 自訂每頁組數與排除單位
python -X utf8 supervision_plan/supervision_plan/table5.1_converter/convert_5.1.py --pairs-per-page 8 --exclude-units 式 工 日
```

## 檔案結構

```
supervision_plan/table5.1_converter/
├── convert_5.1.py   # 主程式
└── SCOPE_5.1.md     # 本計畫說明文件
```
