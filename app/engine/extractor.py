"""按追溯结果生成新工作簿：删去无用行列、改写引用，并独立复算核对。"""
from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .evaluator import EvalError, evaluate
from .formula import parse_formula, unparse
from .model import CellKind, WorkbookModel
from .tracer import TraceResult, addr


@dataclass
class ExtractResult:
    output_path: str
    col_map: dict
    row_map: dict
    indicator_mapping: list          # [(原地址, 新地址)]
    sheet_layout: dict               # 表 -> (保留原行号列表, 保留原列号列表)
    verify: list                     # [(原地址, 原结果, 新结果, 是否一致)]
    verify_errors: list = field(default_factory=list)
    dropped_sheets: list = field(default_factory=list)


def _build_maps(model: WorkbookModel, result: TraceResult):
    col_map: dict[tuple, int] = {}
    row_map: dict[tuple, int] = {}
    layout: dict[str, tuple[list, list]] = {}
    by_sheet: dict[str, set] = {}
    for sheet, col, row in result.required:
        by_sheet.setdefault(sheet, set()).add((col, row))
    for name in model.sheet_names:
        cells = by_sheet.get(name)
        if not cells:
            layout[name] = ([], [])
            continue
        used_cols = sorted({c for c, _ in cells})
        used_rows = sorted({r for _, r in cells})
        for new_c, old_c in enumerate(used_cols, start=1):
            col_map[(name, old_c)] = new_c
        for new_r, old_r in enumerate(used_rows, start=1):
            row_map[(name, old_r)] = new_r
        layout[name] = (used_rows, used_cols)
    return col_map, row_map, layout


def build_extract(model: WorkbookModel, result: TraceResult, output_path: str) -> ExtractResult:
    col_map, row_map, layout = _build_maps(model, result)

    wb_new = Workbook()
    wb_new.remove(wb_new.active)
    wb_new.calculation.fullCalcOnLoad = True
    wb_new.calculation.calcMode = "auto"

    dropped = []
    for name in model.sheet_names:
        used_rows, used_cols = layout[name]
        if not used_rows:
            dropped.append(name)
            continue
        ws_old = model._ws[name]
        ws_new = wb_new.create_sheet(title=name)
        for old_r in used_rows:
            for old_c in used_cols:
                key = (name, old_c, old_r)
                kind, payload = model.classify(key)
                new_r = row_map[(name, old_r)]
                new_c = col_map[(name, old_c)]
                target = ws_new.cell(row=new_r, column=new_c)
                src = ws_old.cell(row=old_r, column=old_c)
                if kind == CellKind.FORMULA:
                    target.value = "=" + unparse(payload.ast, col_map, row_map, name)
                elif kind == CellKind.EMPTY:
                    target.value = None
                else:
                    target.value = payload
                if src.has_style:
                    target._style = src._style
                target.number_format = src.number_format
        # 合并区域按新坐标重建
        for c1, r1, c2, r2 in result.merged_to_keep.get(name, []):
            ws_new.merge_cells(
                start_row=row_map[(name, r1)], start_column=col_map[(name, c1)],
                end_row=row_map[(name, r2)], end_column=col_map[(name, c2)],
            )
        # 保留行列的行高/列宽
        for old_r in used_rows:
            dim = ws_old.row_dimensions.get(old_r)
            if dim and dim.height:
                ws_new.row_dimensions[row_map[(name, old_r)]].height = dim.height
        for old_c in used_cols:
            letter = get_column_letter(old_c)
            dim = ws_old.column_dimensions.get(letter)
            if dim and dim.width:
                ws_new.column_dimensions[get_column_letter(col_map[(name, old_c)])].width = \
                    dim.width

    wb_new.save(output_path)

    indicator_mapping = []
    for ind in result.indicators:
        name, col, row = ind
        new_addr = f"{name}!{get_column_letter(col_map[(name, col)])}{row_map[(name, row)]}"
        indicator_mapping.append((addr(ind), new_addr))

    verify, verify_errors = _verify(model, result, output_path, col_map, row_map)

    return ExtractResult(
        output_path=output_path, col_map=col_map, row_map=row_map,
        indicator_mapping=indicator_mapping, sheet_layout=layout,
        verify=verify, verify_errors=verify_errors, dropped_sheets=dropped,
    )


def _compute_old_values(model: WorkbookModel, keys: set[tuple]):
    """在原工作簿上对 required 集合中的公式格独立求值。"""
    cache: dict[tuple, object] = {}

    def get_value(sheet, col, row):
        key = (sheet, col, row)
        if key in cache:
            return cache[key]
        kind, payload = model.classify(key)
        if kind == CellKind.FORMULA:
            val = evaluate(payload.ast, get_value, sheet)
        elif kind == CellKind.EMPTY:
            val = None
        else:
            val = payload
        cache[key] = val
        return val

    values = {}
    for key in keys:
        values[key] = get_value(*key)
    return values


def _verify(model: WorkbookModel, result: TraceResult, output_path: str,
            col_map: dict, row_map: dict):
    verify_errors: list[str] = []
    verify_rows: list = []

    try:
        old_values = _compute_old_values(model, result.required)
    except EvalError as exc:
        verify_errors.append(f"原工作簿复算失败：{exc}")
        return verify_rows, verify_errors

    wb2 = load_workbook(output_path, data_only=False)

    def new_get_value(sheet, col, row):
        v = wb2[sheet].cell(row=row, column=col).value
        if isinstance(v, str) and v.startswith("="):
            parsed = parse_formula(v, set(wb2.sheetnames))
            return evaluate(parsed.ast, new_get_value, sheet)
        return v

    for ind in result.indicators:
        name, col, row = ind
        new_key = (name, col_map[(name, col)], row_map[(name, row)])
        try:
            new_val = new_get_value(name, new_key[1], new_key[2])
        except EvalError as exc:
            verify_errors.append(f"新工作簿指标 {addr(new_key)} 复算失败：{exc}")
            verify_rows.append((addr(ind), old_values.get(ind), None, False))
            continue
        old_val = old_values.get(ind)
        same = _num_equal(old_val, new_val)
        verify_rows.append((addr(ind), old_val, new_val, same))
        if not same:
            verify_errors.append(
                f"指标 {addr(ind)} 复算不一致：原值 {old_val!r}，新值 {new_val!r}"
            )
    return verify_rows, verify_errors


def _num_equal(a, b, tol=1e-9):
    if a is None and b is None:
        return True
    try:
        fa, fb = float(a), float(b)
    except (TypeError, ValueError):
        return a == b
    scale = max(abs(fa), abs(fb), 1.0)
    return abs(fa - fb) <= tol * scale
