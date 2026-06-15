# Ch2 監造組織及權責分工轉換 — 計畫說明

## 目的

產出監造計畫第二章「監造組織及權責分工」docx 文件。

## 輸入

| 檔案 | 說明 |
|------|------|
| `tables/ch2/ch2_內容.md` | 章節文字內容（待提供） |

## 處理規則

### 內容結構
1. 監造組織架構（SVG 組織圖 + 說明）
2. 工作職掌（權責分工表）
3. 開工前協調會議

### SVG 組織圖
- 用 python 生成 SVG XML
- 輸出至 `output/Ch2_組織圖.svg`
- docx 中以 VML 嵌入

### 權責分工表
- 以 `common/add_cell` 建立表格
- 黑實線框、標楷體 11pt

## 使用方式

```bash
python -X utf8 tables/ch2/convert_2.py
```
