# -*- coding: utf-8 -*-
"""
監造計畫書整本產製入口腳本

修正歷程：
  v1.0  2026/06/14  初始版本：依序執行各表格轉換

作者：OpenCode Assistant / cfwuarch
版本：v1.0
最後更新：2026/06/14

相依套件：
  - 各 tables/*/convert.py 之相依套件

使用方法：
  python -X utf8 build.py

功能說明：
  依序呼叫各表格轉換程式，產出整份監造計畫書所需之 Word 文件。

參數說明：
  無（未來可擴充）
"""
import subprocess
import sys
import os


BASE = os.path.dirname(os.path.abspath(__file__))


def run_convert(rel_path, description):
    """執行指定轉換腳本"""
    script = os.path.join(BASE, rel_path)
    if not os.path.isfile(script):
        print(f'⚠ 找不到 {script}，跳過')
        return
    print(f'\n===== {description} =====')
    result = subprocess.run([sys.executable, '-X', 'utf8', script],
                            cwd=BASE)
    if result.returncode != 0:
        print(f'⚠ {description} 失敗（returncode={result.returncode}）')
    else:
        print(f'✓ {description} 完成')


def main():
    run_convert('tables/table5.1/convert_5.1.py', '表5.1 材料送審管制總表')
    run_convert('tables/table5.2/convert_5.2.py', '表5.2 檢(試)驗管制總表')
    run_convert('tables/table5.3/convert_5.3.py', '表5.3 材料設備品質抽驗紀錄表')
    run_convert('tables/table5.4/convert_5.4.py', '表5.4 抽驗結果通知單')
    run_convert('tables/table5.5/convert_5.5.py', '表5.5 不合格改善追蹤表')
    run_convert('tables/table7.1/convert_7.1.py', '圖7.1 輕隔間施工抽查流程圖')


if __name__ == '__main__':
    main()
