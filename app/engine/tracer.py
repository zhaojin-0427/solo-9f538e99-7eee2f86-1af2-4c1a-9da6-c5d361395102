"""从指标格出发逐层追溯复算所需单元格，处理禁发、循环、不支持内容。

追溯采用迭代 DFS：每个栈帧记录“当前格 + 尚未处理的引用项（Ref/Range）+ 引用链”。
循环通过 visiting 标记发现；禁发格、越界、不支持内容都记录带完整引用链的错误。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string, get_column_letter

from .formula import MAX_COL, MAX_ROW, Range, Ref, iter_range_cells
from .model import CellKind, WorkbookModel


class InputError(Exception):
    """用户录入的指标格/禁发区域不合法。"""


_A1_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})")


def _check_bounds(col: int, row: int, text: str):
    if col > MAX_COL or row > MAX_ROW:
        raise InputError(
            f"地址 {text} 超出允许范围（最多 {MAX_ROW} 行 × {MAX_COL} 列）"
        )


@dataclass
class TraceError:
    indicator: tuple           # 指标原地址 (sheet, col, row)
    indicator_text: str
    kind: str                  # forbidden | cycle | unsupported | dimension | text
    message: str
    chain: list[tuple] = field(default_factory=list)

    @property
    def chain_text(self) -> str:
        return " → ".join(addr(k) for k in self.chain)


@dataclass
class TraceResult:
    required: set[tuple]
    merged_to_keep: dict[str, list]   # 表 -> [(min_col,min_row,max_col,max_row)]
    errors: list[TraceError]
    indicators: list[tuple]
    layers: dict[tuple, int]          # 各格距指标的最大层数（展示用）


def addr(key: tuple) -> str:
    sheet, col, row = key
    return f"{sheet}!{get_column_letter(col)}{row}"


def parse_a1(text: str, known_sheets: list[str], default_sheet: str | None = None) -> tuple:
    s = str(text).strip().replace(" ", "")
    if not s:
        raise InputError("地址为空")
    if "!" in s:
        sheet, cell = s.split("!", 1)
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
    else:
        if default_sheet is None:
            raise InputError(f"地址 {text} 缺少表名，应填写“表名!单元格”")
        sheet, cell = default_sheet, s
    if sheet not in known_sheets:
        raise InputError(f"工作表 {sheet} 不存在")
    m = _A1_RE.fullmatch(cell)
    if not m:
        raise InputError(f"单元格地址 {text} 不是合法的 A1 格式")
    col, row = column_index_from_string(m.group(1).upper()), int(m.group(2))
    _check_bounds(col, row, text)
    return sheet, col, row


def parse_rect(text: str, known_sheets: list[str]) -> tuple:
    """解析 表名!A1:B2（单格也可），返回 (sheet, min_col, min_row, max_col, max_row)。"""
    s = str(text).strip().replace(" ", "")
    if not s:
        raise InputError("区域为空")
    if "!" not in s:
        raise InputError(f"区域 {text} 缺少表名，应填写“表名!A1:B2”")
    sheet, rest = s.split("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    if sheet not in known_sheets:
        raise InputError(f"工作表 {sheet} 不存在")
    parts = rest.split(":")
    if len(parts) == 1:
        parts.append(parts[0])
    if len(parts) != 2:
        raise InputError(f"区域 {text} 格式不正确")
    coords = []
    for p in parts:
        m = _A1_RE.fullmatch(p)
        if not m:
            raise InputError(f"区域 {text} 中 {p} 不是合法 A1 地址")
        col, row = column_index_from_string(m.group(1).upper()), int(m.group(2))
        _check_bounds(col, row, text)
        coords.append((col, row))
    (c1, r1), (c2, r2) = coords
    if c1 > c2 or r1 > r2:
        raise InputError(f"区域 {text} 的起点必须在终点左上方")
    return sheet, c1, r1, c2, r2


def rect_cells(rect: tuple):
    sheet, c1, r1, c2, r2 = rect
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            yield sheet, c, r


def _formula_deps(parsed, current_sheet: str):
    out = []
    for item in parsed.references:
        if isinstance(item, Ref):
            item.sheet = item.sheet or current_sheet
            out.append(item)
        else:
            item.start.sheet = item.start.sheet or current_sheet
            item.end.sheet = item.end.sheet or item.start.sheet
            out.append(item)
    return out


def trace(model: WorkbookModel, indicator_keys: list[tuple],
          forbidden_rects: list[tuple]) -> TraceResult:
    forbidden: set[tuple] = set()
    for rect in forbidden_rects:
        forbidden.update(rect_cells(rect))

    required: set[tuple] = set()
    layers: dict[tuple, int] = {}
    errors: list[TraceError] = []
    seen_errors: set[tuple] = set()

    def add_error(ind, kind, message, chain):
        sig = (ind, kind, message, tuple(chain))
        if sig not in seen_errors:
            seen_errors.add(sig)
            errors.append(TraceError(ind, addr(ind), kind, message, list(chain)))

    # 合并区域索引
    merged_index: dict[str, dict[tuple, tuple]] = {}
    merged_to_keep: dict[str, list] = {}
    for sheet, ranges in model.merged_ranges.items():
        merged_to_keep[sheet] = []
        idx = {}
        for bounds in ranges:
            c1, r1, c2, r2 = bounds
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    idx[(c, r)] = bounds
        merged_index[sheet] = idx

    def keep_merged(key: tuple):
        sheet, col, row = key
        bounds = merged_index.get(sheet, {}).get((col, row))
        if bounds and bounds not in merged_to_keep[sheet]:
            merged_to_keep[sheet].append(bounds)
            c1, r1, c2, r2 = bounds
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    required.add((sheet, c, r))

    # ---- 指标格前置校验 ----
    indicators: list[tuple] = []
    for ind in indicator_keys:
        ind_text = addr(ind)
        canon = model.canonical(ind)
        keep_merged(canon)
        if not model.in_dimensions(canon):
            add_error(ind, "dimension", "指标格超出工作表数据区域", [ind])
            continue
        if ind in forbidden or canon in forbidden:
            add_error(ind, "forbidden", "指标格本身位于禁发区域", [ind])
            continue
        kind, payload = model.classify(canon)
        if kind == CellKind.UNSUPPORTED:
            add_error(ind, "unsupported", payload, [ind])
            continue
        if kind == CellKind.TEXT:
            add_error(ind, "text", "指标格是文本，没有可复算的数值", [ind])
            continue
        required.add(canon)
        indicators.append(canon)

    visiting: set[tuple] = set()
    done: set[tuple] = set()

    def visit_cell(key, ind, chain) -> bool:
        """检查一个被引用格；返回 True 表示它是公式格、需要继续深入。"""
        canon = model.canonical(key)
        keep_merged(canon)
        if key in forbidden or canon in forbidden:
            add_error(ind, "forbidden", f"引用到禁发格 {addr(canon)}", chain + [canon])
            return False
        # Excel 语义：引用数据区域外视为空白（按 0），不是错误；
        # 但仍需把该位置纳入 required，使公式改写后指向新文件里一个真实空格
        if not model.in_dimensions(canon):
            required.add(canon)
            return False
        required.add(canon)
        if canon in visiting:
            cyc = _cycle_path(canon, frames)
            add_error(ind, "cycle",
                      "检测到循环引用：" + " → ".join(addr(k) for k in cyc),
                      chain + [canon])
            return False
        if canon in done:
            return False
        kind, payload = model.classify(canon)
        if kind == CellKind.UNSUPPORTED:
            add_error(ind, "unsupported",
                      f"{addr(canon)} 含不支持的内容：{payload}", chain + [canon])
            done.add(canon)
            return False
        if kind == CellKind.TEXT:
            add_error(ind, "text",
                      f"{addr(canon)} 是文本，不能作为数值引用", chain + [canon])
            done.add(canon)
            return False
        if kind == CellKind.EMPTY:
            done.add(canon)
            return False
        visiting.add(canon)
        if kind == CellKind.FORMULA:
            frames.append([canon, _formula_deps(payload, canon[0]), 0, chain + [canon]])
            layers[canon] = max(layers.get(canon, 0), len(chain))
            return True
        done.add(canon)
        visiting.discard(canon)
        layers[canon] = max(layers.get(canon, 0), len(chain))
        return False

    frames: list[list] = []

    def _cycle_path(target, frame_stack):
        for i, fr in enumerate(frame_stack):
            if fr[0] == target:
                return [f[0] for f in frame_stack[i:]] + [target]
        return [target, target]

    for ind in indicators:
        if ind in done:
            continue
        kind0, payload0 = model.classify(ind)
        visiting.add(ind)
        layers.setdefault(ind, 0)
        frames = [[ind, _formula_deps(payload0, ind[0]) if kind0 == CellKind.FORMULA else [],
                   0, [ind]]]
        steps = 0
        while frames:
            steps += 1
            if steps > 500000:
                add_error(ind, "cycle", "依赖链超过 50 万步，按异常循环处理",
                          [f[0] for f in frames])
                frames.clear()
                break
            cell, entries, ei, chain = frames[-1]
            if ei >= len(entries):
                frames.pop()
                visiting.discard(cell)
                done.add(cell)
                continue
            entry = entries[ei]
            frames[-1][2] += 1
            if isinstance(entry, Ref):
                key = (entry.sheet, entry.col, entry.row)
                visit_cell(key, ind, chain)
            else:
                rng: Range = entry
                # SUM 区域：每个位置（含空格）都是引用位置
                for s, c, r in iter_range_cells(rng):
                    visit_cell((s, c, r), ind, chain)

    return TraceResult(required=required, merged_to_keep=merged_to_keep,
                       errors=errors, indicators=indicators, layers=layers)
