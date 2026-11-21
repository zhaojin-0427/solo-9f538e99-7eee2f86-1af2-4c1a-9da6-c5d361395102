"""手工端到端验证脚本（隔离模式运行）：构造含各种受支持特性的工作簿。"""
import os
import sys

sys.path.insert(0, "/workspace")
for _p in ("/home/node/.local/lib/python3.11/site-packages",):
    if os.path.isdir(_p):
        sys.path.append(_p)

from openpyxl import Workbook, load_workbook

path = "/tmp/demo_input.xlsx"
wb = Workbook()
ws = wb.active
ws.title = "基础表"
ws2 = wb.create_sheet("调整-表 2")

# 基础表：一块带空格的数据
data = [[10, 20], [None, 40], [30, None]]
for i, row in enumerate(data, start=1):
    for j, v in enumerate(row, start=1):
        ws.cell(row=i, column=j, value=v)
ws["E1"] = 5
ws["E2"] = "=SUM($A$1:$B$3)+E1"          # SUM 区域内含空格 + 绝对区域
ws["E3"] = "=(E2-E1)*2/3"
ws["F9"] = "=E3+调整-表 2'!D4"            # 含特殊字符表名的跨表引用（下面修正引号）

ws2["D4"] = 7
ws2["D5"] = "=D4*D4"
wb.save(path)

# 修正跨表公式（openpyxl 需要带引号的表名）
wb = load_workbook(path)
wb["基础表"]["F9"] = "='调整-表 2'!D4+E3"
wb.save(path)

# 不支持内容的对照文件
bad_path = "/tmp/demo_bad.xlsx"
wb2 = Workbook()
b = wb2.active
b.title = "S"
b["A1"] = 1
b["A2"] = '=IF(A1>0,1,2)'     # 不支持的函数
b["B1"] = "=A2+1"             # 指标引用链 A? -> A2 不支持
wb2.save(bad_path)

print("created", path, bad_path)
