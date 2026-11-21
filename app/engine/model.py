"""openpyxl 工作簿的只读封装：维度、合并单元格重定向、单元格分类。"""
from __future__ import annotations

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .formula import MAX_COL, MAX_ROW, FormulaError, parse_formula


class WorkbookLoadError(Exception):
    pass


class CellKind:
    EMPTY = "empty"
    NUMBER = "number"
    FORMULA = "formula"
    TEXT = "text"
    UNSUPPORTED = "unsupported"


class WorkbookModel:
    def __init__(self, path: str):
        try:
            self.wb = load_workbook(path, data_only=False)
        except Exception as exc:  # openpyxl 各种读取异常
            raise WorkbookLoadError(f"无法读取 XLSX 文件：{exc}") from exc
        if len(self.wb.worksheets) > 6:
            raise WorkbookLoadError(f"工作簿含 {len(self.wb.worksheets)} 张表，超过最多 6 张的限制")
        self.sheet_names = [ws.title for ws in self.wb.worksheets]
        self._ws = {ws.title: ws for ws in self.wb.worksheets}
        for ws in self.wb.worksheets:
            if ws.max_row > MAX_ROW or ws.max_column > MAX_COL:
                raise WorkbookLoadError(
                    f"工作表 {ws.title} 为 {ws.max_row} 行 × {ws.max_column} 列，"
                    f"超过单表 {MAX_ROW} 行 × {MAX_COL} 列的限制"
                )
        # 合并区域：非锚点格 -> 锚点格
        self.anchor_of: dict[tuple, tuple] = {}
        self.merged_ranges: dict[str, list[tuple]] = {}
        for name, ws in self._ws.items():
            self.merged_ranges[name] = []
            for mr in ws.merged_cells.ranges:
                bounds = (mr.min_col, mr.min_row, mr.max_col, mr.max_row)
                self.merged_ranges[name].append(bounds)
                for r in range(mr.min_row, mr.max_row + 1):
                    for c in range(mr.min_col, mr.max_col + 1):
                        if (c, r) != (mr.min_col, mr.min_row):
                            self.anchor_of[(name, c, r)] = (name, mr.min_col, mr.min_row)

    # ---- 基本访问 ----
    def dimensions(self, sheet: str) -> tuple[int, int]:
        ws = self._ws[sheet]
        return ws.max_row or 0, ws.max_column or 0

    def address(self, key: tuple) -> str:
        sheet, col, row = key
        return f"{sheet}!{get_column_letter(col)}{row}"

    def canonical(self, key: tuple) -> tuple:
        """合并单元格的非锚点格重定向到锚点。"""
        return self.anchor_of.get(key, key)

    def in_dimensions(self, key: tuple) -> bool:
        sheet, col, row = key
        rows, cols = self.dimensions(sheet)
        return 1 <= row <= rows and 1 <= col <= cols

    def raw_value(self, key: tuple):
        sheet, col, row = key
        return self._ws[sheet].cell(row=row, column=col).value

    def classify(self, key: tuple) -> tuple[str, object]:
        """返回 (类别, 原始值/解析后的 ParsedFormula)。"""
        value = self.raw_value(key)
        if value is None:
            return CellKind.EMPTY, None
        if isinstance(value, bool):
            return CellKind.NUMBER, value
        if isinstance(value, (int, float)):
            return CellKind.NUMBER, value
        if isinstance(value, str):
            if value.startswith("="):
                try:
                    return CellKind.FORMULA, parse_formula(value, set(self.sheet_names))
                except FormulaError as exc:
                    return CellKind.UNSUPPORTED, str(exc)
            return CellKind.TEXT, value
        # ArrayFormula、DataTable 等 openpyxl 对象
        return CellKind.UNSUPPORTED, f"不支持的单元格内容类型：{type(value).__name__}"
