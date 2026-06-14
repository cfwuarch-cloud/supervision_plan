# AGENTS.md — Supervision Plan 專案開發指南

> 基於 Andrej Karpathy LLM 編碼行為觀察（https://github.com/multica-ai/andrej-karpathy-skills）

---

## 核心準則（Karpathy 四原則）

### 1. 先思考再編碼
- **明確陳述假設** — 不確定的時候問，不要猜
- **呈現多種解釋** — 遇到歧義時提出選項，不要默默選一個
- **適時提出異議** — 有更簡單的做法時要說出來
- **卡住時停下來** — 指出困惑點並提問

### 2. 保持簡單
- 不做未要求的額外功能
- 不為一次性用途建立抽象層
- 不寫不可能發生的錯誤處理
- 如果 200 行能寫成 50 行，就重寫
- **自問**：資深工程師會說這太複雜嗎？如果是，簡化

### 3. 精準修改
- 只動必須動的地方，不改鄰近的程式碼、註解或格式
- 不重構沒壞的東西
- 匹配既有風格
- 發現不相關的死程式時，**提及但不要刪除**
- 若你的改動產生了孤兒（未使用的 import/變數），清除它們
- **檢驗**：每一行改動都應直接對應到使用者的需求

### 4. 目標驅動執行
- 將指令轉換為可驗證的目標：
  - 「加驗證」→「寫無效輸入測試，讓它們通過」
  - 「修 bug」→「寫重現 bug 的測試，讓它通過」
  - 「重構 X」→「確保前後測試都通過」
- 多步驟任務列出簡短計畫與檢查點
- **關鍵洞見**：LLM 非常擅長在明確目標驅動下反覆迭代直到達成

---

## 專案資訊

### 目錄結構

```
supervision_plan/
├── data/                          # 輸入資料（詳細價目表母本）
│   └── 02_成德-詳細價目表.xlsx
├── tables/                        # 各表格轉換模組（可擴充）
│   ├── table5.1/                  #   表5.1 材料送審管制總表
│   │   ├── convert_5.1.py         #     主程式
│   │   ├── SCOPE_5.1.md           #     計畫說明
│   │   ├── 表5.1.docx             #     模板
│   │   └── 表5.1_完成13.docx      #     輸出範例
│   └── table5.2/                  #   表5.2 檢(試)驗管制總表
│       ├── convert_5.2.py         #     主程式
│       ├── SCOPE_5.2.md           #     計畫說明
│       ├── 表5.2.docx             #     模板
│       └── 表5.2_完成_test_4.docx #     輸出範例
├── common/                        # 共用元件（docx helper、樣式等）
├── output/                        # 輸出檔案（自動建立，已 gitignore）
├── tools/                         # 輔助工具
│   └── check_pages.py             #   估算 docx 頁數
├── AGENTS.md                      # 本檔 — AI 行為指南
├── README.md                      # 專案說明
└── requirements.txt               # 依賴套件
```

---

## 技術規範

### 語言與註解
- 所有回覆使用**繁體中文**
- 檔頭註解至少包含：修正歷程、作者（OpenCode Assistant / cfwuarch）、版本、最後更新、相依套件、使用方法、功能說明、參數說明

### 套件版本管理
- `requirements.txt` 需加入版本上限防止破壞性更新
- 格式：`套件名稱>=最低版本,<上限版本`
- 範例：`openpyxl>=3.0.0,<4.0.0`

### PowerShell 中文顯示
- 使用**單引號**包覆 PowerShell 指令
```bash
# 正確
powershell -Command '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Write-Host "測試"'
# 錯誤：雙引號會導致 $ 被 bash 解譯
powershell -Command "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Write-Host '測試'"
```
- 執行 Python 腳本時加 `-X utf8` 參數

### 備份作業
- 同步 GitHub 遠端與本機檔案備份
- 本機備份：排除 `.git`，壓縮 zip，格式 `專案名_YYYYMMDDHHMMSS.zip`，存放桌面
```bash
python -X utf8 tools/backup.py -p .
```

---

## 常用指令

```bash
# 表5.1 材料送審管制總表
python -X utf8 tables/table5.1/convert_5.1.py --exclude-units 式 工

# 表5.2 檢(試)驗管制總表（測試模式）
python -X utf8 tables/table5.2/convert_5.2.py --test-num 1 --exclude-units 工

# 表5.2 正式輸出
python -X utf8 tables/table5.2/convert_5.2.py -o output/表5.2_完成.docx

# 檢查 docx 頁數
python -X utf8 tools/check_pages.py output/表5.2_完成.docx
```
