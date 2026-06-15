# 監造計畫書 — 表格產製及各章節轉換工具

產製公共工程監造計畫書所需之各章節 Word 文件，從詳細價目表（xlsx）自動轉換，
並可合併為完整監造計畫書。

## 目錄結構

```
supervision_plan/
├── data/
│   ├── 02_成德-詳細價目表.xlsx    # 輸入母本
│   ├── project_info.json          # 工程基本資料
│   └── 監造計畫各級距內容對照表.md  # 級距章節對照（含官方比對）
├── tables/                        # 各轉換模組
│   ├── foreword/                  # 前言
│   ├── ch1/                       # 第一章 監造範圍（從價目表讀取主要施工項目）
│   ├── ch2/                       # 第二章 監造組織及權責分工（含SVG組織圖）
│   ├── ch3/                       # 第三章 品質計畫審查作業程序（含SVG流程圖＋表3.1）
│   ├── ch4/                       # 第四章 施工計畫審查作業程序（含SVG流程圖＋表4.1/4.2）
│   ├── ch6/                       # 第六章 設備功能運轉測試程序（含表6.1/6.2）（機電條件）
│   ├── ch8/                       # 第八章 品質稽核（含SVG流程圖＋稽核查對表）
│   ├── ch9/                       # 第九章 文件紀錄管理系統
│   ├── table5.1/                  # 表5.1 材料送審管制總表
│   ├── table5.2/                  # 表5.2 檢(試)驗管制總表
│   ├── table5.3/                  # 表5.3 材料設備品質抽驗紀錄表
│   ├── table5.4/                  # 表5.4 抽驗結果通知單
│   ├── table5.5/                  # 表5.5 不合格改善追蹤表
│   └── table7.1/                  # 圖7.1 輕隔間施工抽查流程圖
├── common/                        # 共用元件
│   ├── __init__.py
│   ├── docx_utils.py              # docx 共用工具（列高、分頁、合併）
│   └── docx_table.py              # 通用表格元件（黑實線框、tcMar=0）
├── tools/
│   ├── input_project.py           # 工程基本資料輸入
│   ├── calibrate_row_height.py    # 列高校準
│   ├── check_pages.py             # docx 頁數估算
│   ├── backup.py                  # 備份工具
│   └── merge_docx.py              # [新增] 合併各章為完整監造計畫書
├── output/                        # 輸出 docx（已 gitignore）
├── AGENTS.md                      # AI 行為指南
├── requirements.txt
└── README.md
```

## 安裝相依套件

```bash
pip install -r requirements.txt
```

## 使用方法

### 工程基本資料輸入

```bash
python -X utf8 tools/input_project.py           # GUI 模式
python -X utf8 tools/input_project.py --cli     # CLI 模式
```

### 各章節腳本

```bash
# 前言
python -X utf8 tables/foreword/convert_foreword.py

# 第一章 監造範圍（--level 控制價目表階層深度）
python -X utf8 tables/ch1/convert_1.py --level 4

# 第二章 監造組織及權責分工
python -X utf8 tables/ch2/convert_2.py

# 第三章 品質計畫審查作業程序
python -X utf8 tables/ch3/convert_3.py

# 第四章 施工計畫審查作業程序
python -X utf8 tables/ch4/convert_4.py

# 第六章 設備功能運轉測試抽驗程序及標準（機電條件）
python -X utf8 tables/ch6/convert_6.py

# 第八章 品質稽核
python -X utf8 tables/ch8/convert_8.py

# 第九章 文件紀錄管理系統
python -X utf8 tables/ch9/convert_9.py
```

### 表格轉換

```bash
# 表5.1 材料送審管制總表
python -X utf8 tables/table5.1/convert_5.1.py --exclude-units 式 工

# 表5.2 檢(試)驗管制總表
python -X utf8 tables/table5.2/convert_5.2.py -o output/表5.2_完成.docx

# 表5.3 材料設備品質抽驗紀錄表
python -X utf8 tables/table5.3/convert_5.3.py --exclude-units 式 工

# 表5.4 抽驗結果通知單
python -X utf8 tables/table5.4/convert_5.4.py --exclude-units 式 工

# 表5.5 不合格改善追蹤表
python -X utf8 tables/table5.5/convert_5.5.py --exclude-units 式 工

# 圖7.1 輕隔間施工抽查流程圖
python -X utf8 tables/table7.1/convert_7.1.py
```

### 合併為完整監造計畫書

```bash
python -X utf8 tools/merge_docx.py
```

### 檢查 docx 頁數

```bash
python -X utf8 tools/check_pages.py output/表5.2_完成.docx
```

## 監造計畫三級距對照

依「公共工程施工品質管理作業要點第八點」，監造計畫依工程規模分三級距：

| 級距 | 金額範圍 | 章節數 |
|------|---------|--------|
| 一 | 公告金額（150萬）↑～<1000萬 | 4 + Ⓜ |
| 二 | 1000萬↑～<5000萬 | 6 + Ⓜ |
| 三 | 5000萬↑（查核金額以上） | 8 + Ⓜ |

> Ⓜ = 工程具機電設備者始須增訂第六章

詳細對照表請見 `data/監造計畫各級距內容對照表.md`。

## 行高機制（適用表格類轉換）

| 參數 | 說明 | 值 |
|------|------|-----|
| LINE_H_TWIP | 每行文字高 | 240 twip（10pt 標楷體） |
| CELL_TOP_TWIP | 儲存格頂部餘裕 | 0（tcMar=0） |
| MIN_ROW_H_TWIP | 列高下限 | 226 twip |
| trHeight hRule | 列高規則 | `atLeast` |
| 段落行距 | 文字高度規則 | `exact` at LINE_H_TWIP |

## 技術棧

- **Python** — 主程式語言
- **pandas / openpyxl** — 讀取詳細價目表
- **python-docx / lxml** — 產出 Word 文件
- **Pillow** — 文字寬度測量（行數計算）
