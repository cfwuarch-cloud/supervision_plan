# 監造計畫書表格輔助工具

產製公共工程監造計畫書所需之表格，目前支援：

- **表5.1** — 材料送審管制總表
- **表5.2** — 檢(試)驗管制總表

從詳細價目表（xlsx）自動轉換為排版完成的 Word 表格文件，減少手動繕打錯誤與重複工時。

## 目錄結構

```
supervision_plan/
├── data/                        # 輸入資料（詳細價目表母本）
│   └── 02_成德-詳細價目表.xlsx
├── tables/                      # 各表格轉換模組（可擴充）
│   ├── table5.1/                #   表5.1 材料送審管制總表
│   │   ├── convert_5.1.py       #     主程式
│   │   ├── 表5.1.docx           #     模板
│   │   └── 表5.1_完成13.docx    #     輸出範例
│   └── table5.2/                #   表5.2 檢(試)驗管制總表
│       ├── convert_5.2.py       #     主程式
│       ├── 表5.2.docx           #     模板
│       └── 表5.2_完成_test_4.docx #  輸出範例
├── common/                      # 共用元件（docx helper、樣式等）
├── output/                      # 輸出檔案（自動建立，已 gitignore）
├── tools/                       # 輔助工具
│   └── check_pages.py           #   估算 docx 頁數
├── requirements.txt             # Python 套件相依
├── README.md                    # 本檔
└── AGENTS.md                    # AI 行為指南
```

## 安裝相依套件

```bash
pip install -r requirements.txt
```

## 使用方法

### 表5.1 材料送審管制總表

```bash
python -X utf8 tables/table5.1/convert_5.1.py --exclude-units 式 工
```

### 表5.2 檢(試)驗管制總表

測試模式（僅輸出前 N 筆）：

```bash
python -X utf8 tables/table5.2/convert_5.2.py --test-num 1 --exclude-units 工
```

正式輸出：

```bash
python -X utf8 tables/table5.2/convert_5.2.py -o output/表5.2_完成.docx
```

### 檢查 docx 頁數

```bash
python -X utf8 tools/check_pages.py output/表5.2_完成.docx
```

## 技術棧

- **Python** — 主程式語言
- **pandas / openpyxl** — 讀取詳細價目表 xlsx
- **python-docx / lxml** — 產生與操作 Word 文件
- **Pillow** — 表格圖片處理（蓋章用）
