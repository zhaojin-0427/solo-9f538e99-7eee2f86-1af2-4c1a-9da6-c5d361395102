"""对受支持公式做独立求值，用于结果核对（不依赖 Excel/LibreOffice）。"""
from __future__ import annotations

from .formula import BinOp, Number, Range, Ref, SumCall, UnaryMinus, iter_range_cells


class EvalError(Exception):
    """求值失败（除零、文本参与运算等）。"""


def _to_number(value, where: str) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    raise EvalError(f"{where} 的内容 {value!r} 是文本，不能参与数值运算")


def evaluate(node, get_value, current_sheet: str):
    """get_value(sheet, col, row) -> 单元格原值（公式格应先算出数值）。"""
    if isinstance(node, Number):
        return node.value
    if isinstance(node, Ref):
        sheet = node.sheet or current_sheet
        return _to_number(get_value(sheet, node.col, node.row),
                          f"{sheet}!{node.col},{node.row}")
    if isinstance(node, UnaryMinus):
        return -evaluate(node.operand, get_value, current_sheet)
    if isinstance(node, BinOp):
        a = evaluate(node.left, get_value, current_sheet)
        b = evaluate(node.right, get_value, current_sheet)
        if node.op == "+":
            return a + b
        if node.op == "-":
            return a - b
        if node.op == "*":
            return a * b
        if b == 0:
            raise EvalError("公式中存在除以零")
        return a / b
    if isinstance(node, SumCall):
        total = 0.0
        if isinstance(node.arg, Ref):
            sheet = node.arg.sheet or current_sheet
            total = _to_number(get_value(sheet, node.arg.col, node.arg.row),
                               f"{sheet}!{node.arg.col},{node.arg.row}")
        else:
            rng: Range = node.arg
            sheet = (rng.sheet or current_sheet)
            for _s, c, r in iter_range_cells(rng):
                v = get_value(sheet, c, r)
                # 与 Excel 一致：SUM 忽略区域内的文本与逻辑值，空格按 0
                if v is None or v == "":
                    continue
                if isinstance(v, bool):
                    continue
                if isinstance(v, (int, float)):
                    total += float(v)
        return total
    raise EvalError("内部错误：未知 AST 节点")
