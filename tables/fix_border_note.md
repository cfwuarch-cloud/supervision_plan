# 遍歷 Word 表格所有儲存格的正確方式

## 問題

`row.cells` **不會**遍歷到垂直合併延續格（vMerge continue），那些格的邊框/格式會被遺漏。

```python
# ❌ 會漏掉 vMerge continue 的格子
for row in table.rows:
    for cell in row.cells:
        tc = cell._tc  # 某些 tc 拿不到
```

## 原因

python-docx 的 `row.cells` 對垂直合併的處理不完整，延續列的格子可能不在回傳清單中。

## 解法：直接走 XML

```python
from docx.oxml.ns import qn
from lxml import etree

for table in doc.tables:
    for tr in table._tbl.findall(qn('w:tr')):       # 所有表格列
        for tc in tr.findall(qn('w:tc')):            # 該列所有儲存格（含合併格）
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('w:tcPr'))
            # 對 tcPr 做你想做的事...
```

## 完整範例：統一所有儲存格邊框為 1/4pt 細實線

```python
import docx
from docx.oxml.ns import qn
from lxml import etree

src = '你的檔案.docx'
doc = docx.Document(src)

for table in doc.tables:
    for tr in table._tbl.findall(qn('w:tr')):
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('w:tcPr'))
            old = tcPr.find(qn('w:tcBorders'))
            if old is not None:
                tcPr.remove(old)
            tcBorders = etree.SubElement(tcPr, qn('w:tcBorders'))
            for edge in ('top', 'left', 'bottom', 'right'):
                be = etree.SubElement(tcBorders, qn('w:' + edge))
                be.set(qn('w:val'), 'single')
                be.set(qn('w:sz'), '2')        # 2 = 1/4pt, 8 = 1pt
                be.set(qn('w:space'), '0')
                be.set(qn('w:color'), '000000')

doc.save(src)
```

## 邊框粗細對照（sz 值）

| sz | 實際粗細 |
|----|---------|
| 2  | 1/4 pt  |
| 4  | 1/2 pt  |
| 8  | 1 pt    |
| 12 | 1.5 pt  |
| 24 | 3 pt    |

單位：sz = 點數 × 8（1pt = 8）