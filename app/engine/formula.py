"""受支持的 Excel 公式子集：四则运算、A1 单格引用（含绝对地址）、单矩形区域 SUM。

语法（Excel 风格）：
    expr      := term (('+'|'-') term)*
    term      := factor (('*'|'/') factor)*
    factor    := NUMBER | reference | range | '(' expr ')' | '-' factor
    reference := [SHEET '!'] (('$'? COL '$'? ROW))
    range     := reference ':' reference
    sumcall   := 'SUM' '(' range | reference ')'
其它函数、字符串、外部工作簿引用等一律报不支持。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from openpyxl.utils import column_index_from_string, get_column_letter

MAX_COL = 40
MAX_ROW = 200


class FormulaError(Exception):
    """公式不支持或语法错误。"""


# ---- 词法 ----
_TOKEN_RE = re.compile(
    r"""
    (?P<sheet>'(?:[^']|'')+'\s*!|[A-Za-z0-9_一-鿿.]+!)?
    (?P<rest>
        \$(?P<d1>[A-Za-z]{1,3})\$(?P<r1>[0-9]{1,7})      # $A$1
      | \$(?P<d2>[A-Za-z]{1,3})(?P<r2>[0-9]{1,7})         # $A1
      | (?P<d3>[A-Za-z]{1,3})\$(?P<r3>[0-9]{1,7})         # A$1
      | (?P<d4>[A-Za-z]{1,3})(?P<r4>[0-9]{1,7})           # A1
      | (?P<number>[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)
      | (?P<id>[A-Za-z_][A-Za-z0-9_.]*)
      | (?P<op>[+\-*/():,])
      | (?P<space>\s+)
    )
    """,
    re.VERBOSE,
)


@dataclass
class Token:
    kind: str            # number | ref | ident | op | sheet
    value: object
    pos: int
    text: str


def _decode_sheet(raw: str) -> str:
    s = raw.strip()
    assert s.endswith("!")
    s = s[:-1].strip()
    if s.startswith("'"):
        s = s[1:-1].replace("''", "'")
    return s


def tokenize(formula: str) -> list[Token]:
    tokens: list[Token] = []
    i = 0
    n = len(formula)
    while i < n:
        m = _TOKEN_RE.match(formula, i)
        if not m:
            raise FormulaError(f"公式存在无法识别的内容：{formula[i:i + 8]!r}")
        pos = i
        i = m.end()
        if m.group("sheet"):
            tokens.append(Token("sheet", _decode_sheet(m.group("sheet")), pos, m.group("sheet")))
        gd = m.groupdict()
        if gd.get("space"):
            continue
        if gd.get("number"):
            tokens.append(Token("number", float(m.group("number")), pos, m.group("number")))
            continue
        for k in ("d1", "d2", "d3", "d4"):
            if gd.get(k):
                col_letters = m.group(k)
                row_num = int(m.group("r" + k[1:]))
                rk = "r" + k[1:]
                raw = m.group("rest")
                col_abs = raw.startswith("$")
                row_abs = "$" in raw[1:]
                tokens.append(
                    Token("ref", {"col": column_index_from_string(col_letters.upper()),
                                  "row": row_num,
                                  "col_abs": col_abs,
                                  "row_abs": row_abs}, pos, raw)
                )
                break
        else:
            if gd.get("id"):
                tokens.append(Token("ident", m.group("id").upper(), pos, m.group("id")))
            elif gd.get("op"):
                tokens.append(Token("op", m.group("op"), pos, m.group("op")))
    return tokens


# ---- AST ----
@dataclass
class Number:
    value: float


@dataclass
class Ref:
    sheet: Optional[str]          # None 表示本表
    col: int
    row: int
    col_abs: bool
    row_abs: bool


@dataclass
class Range:
    start: Ref
    end: Ref

    @property
    def sheet(self) -> Optional[str]:
        return self.start.sheet


@dataclass
class BinOp:
    op: str
    left: object
    right: object


@dataclass
class UnaryMinus:
    operand: object


@dataclass
class SumCall:
    arg: object                    # Ref 或 Range


@dataclass
class ParsedFormula:
    ast: object
    references: list = field(default_factory=list)   # 扁平引用列表：Ref / Range
    raw: str = ""


class _Parser:
    def __init__(self, tokens: list[Token], formula: str, known_sheets: set[str]):
        self.toks = tokens
        self.i = 0
        self.formula = formula
        self.known_sheets = known_sheets
        self.references: list = []
        self.pending_sheet: Optional[str] = None

    def _peek(self) -> Optional[Token]:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def _next(self) -> Token:
        t = self.toks[self.i]
        self.i += 1
        return t

    def parse(self) -> object:
        if not self.toks:
            raise FormulaError("公式为空")
        node = self._expr()
        if self._peek() is not None:
            raise FormulaError(f"公式末尾存在多余内容：{self.formula}")
        return node

    def _expr(self) -> object:
        node = self._term()
        while self._peek() and self._peek().kind == "op" and self._peek().value in "+-":
            op = self._next().value
            node = BinOp(op, node, self._term())
        return node

    def _term(self) -> object:
        node = self._factor()
        while self._peek() and self._peek().kind == "op" and self._peek().value in "*/":
            op = self._next().value
            node = BinOp(op, node, self._factor())
        return node

    def _factor(self) -> object:
        t = self._peek()
        if t is None:
            raise FormulaError("公式不完整")
        if t.kind == "op" and t.value == "-":
            self._next()
            return UnaryMinus(self._factor())
        if t.kind == "op" and t.value == "+":
            self._next()
            return self._factor()
        if t.kind == "op" and t.value == "(":
            self._next()
            node = self._expr()
            self._expect_op(")")
            return node
        if t.kind == "number":
            self._next()
            return Number(t.value)
        if t.kind == "ident":
            return self._call()
        if t.kind == "sheet" or t.kind == "ref":
            return self._ref_or_range()
        raise FormulaError(f"公式中存在不支持的写法：{t.text!r}")

    def _expect_op(self, op: str) -> None:
        t = self._peek()
        if not t or t.kind != "op" or t.value != op:
            raise FormulaError(f"公式缺少 {op!r}：{self.formula}")
        self._next()

    def _call(self) -> object:
        name_tok = self._next()
        name = name_tok.value
        self._expect_op("(")
        if name != "SUM":
            raise FormulaError(f"不支持的函数 {name}，仅支持单区域 SUM")
        arg = self._ref_or_range()
        if isinstance(arg, Range) is False and not isinstance(arg, Ref):
            raise FormulaError("SUM 参数必须是单元格或矩形区域引用")
        # SUM 不允许多参数/表达式
        t = self._peek()
        if t and t.kind == "op" and t.value == ",":
            raise FormulaError("SUM 仅支持单个矩形区域，不支持多参数")
        self._expect_op(")")
        node = SumCall(arg)
        return node

    def _consume_sheet_prefix(self) -> Optional[str]:
        t = self._peek()
        if t and t.kind == "sheet":
            self._next()
            sheet = t.value
            if self.known_sheets is not None and sheet not in self.known_sheets:
                raise FormulaError(f"公式引用了不存在的工作表：{sheet}")
            return sheet
        return None

    def _read_ref(self) -> Ref:
        sheet = self._consume_sheet_prefix()
        t = self._peek()
        if not t or t.kind != "ref":
            raise FormulaError("工作表前缀后必须是单元格引用")
        self._next()
        d = t.value
        if not (1 <= d["col"] <= MAX_COL and 1 <= d["row"] <= MAX_ROW):
            raise FormulaError(
                f"引用 {get_column_letter(d['col'])}{d['row']} 超出允许范围（{MAX_ROW} 行 × {MAX_COL} 列）"
            )
        return Ref(sheet, d["col"], d["row"], d["col_abs"], d["row_abs"])

    def _ref_or_range(self) -> object:
        first = self._read_ref()
        t = self._peek()
        if t and t.kind == "op" and t.value == ":":
            self._next()
            second_sheet = self._consume_sheet_prefix()
            t2 = self._peek()
            if not t2 or t2.kind != "ref":
                raise FormulaError("区域运算符 ':' 后必须是单元格引用")
            self._next()
            d = t2.value
            if not (1 <= d["col"] <= MAX_COL and 1 <= d["row"] <= MAX_ROW):
                raise FormulaError("区域引用超出允许范围")
            second = Ref(second_sheet, d["col"], d["row"], d["col_abs"], d["row_abs"])
            if (second.sheet or first.sheet) and second.sheet != first.sheet:
                raise FormulaError("区域引用的两个端点必须在同一张表内")
            second.sheet = first.sheet
            if first.col > second.col or first.row > second.row:
                raise FormulaError("区域起点必须位于终点左上方")
            node = Range(first, second)
            self.references.append(node)
            return node
        self.references.append(first)
        return first


def parse_formula(formula: str, known_sheets: set[str]) -> ParsedFormula:
    """解析公式，返回 AST 与扁平引用清单。"""
    if not isinstance(formula, str) or not formula.strip():
        raise FormulaError("公式为空")
    body = formula.strip()
    if body.startswith("="):
        body = body[1:]
    tokens = tokenize(body)
    parser = _Parser(tokens, formula, set(known_sheets))
    ast = parser.parse()
    return ParsedFormula(ast=ast, references=parser.references, raw=formula)


# ---- 反解析（按新布局重写）----
def _ref_text(ref: Ref, col_map: dict, row_map: dict, current_sheet: str) -> str:
    sheet = ref.sheet or current_sheet
    new_col = col_map[(sheet, ref.col)]
    new_row = row_map[(sheet, ref.row)]
    return (("$" if ref.col_abs else "") + get_column_letter(new_col)
            + ("$" if ref.row_abs else "") + str(new_row))


def unparse(node, col_map: dict, row_map: dict, current_sheet: str) -> str:
    """按新行列映射重写公式；同一表内的引用不加表名前缀。"""
    if isinstance(node, Number):
        v = node.value
        return ("%g" % v) if abs(v - round(v)) > 1e-12 or abs(v) >= 1e16 else str(int(round(v)))
    if isinstance(node, Ref):
        sheet = node.sheet or current_sheet
        base = _ref_text(node, col_map, row_map, current_sheet)
        if node.sheet and node.sheet != current_sheet:
            return f"{_quote_sheet(node.sheet)}!{base}"
        return base
    if isinstance(node, Range):
        if node.start.sheet is None:
            node.start.sheet = current_sheet
            node.end.sheet = current_sheet
        s = unparse(node.start, col_map, row_map, current_sheet)
        e = unparse(node.end, col_map, row_map, current_sheet)
        return f"{s}:{e}"
    if isinstance(node, UnaryMinus):
        return "-" + unparse(node.operand, col_map, row_map, current_sheet)
    if isinstance(node, BinOp):
        return ("(" + unparse(node.left, col_map, row_map, current_sheet)
                + node.op + unparse(node.right, col_map, row_map, current_sheet) + ")")
    if isinstance(node, SumCall):
        return "SUM(" + unparse(node.arg, col_map, row_map, current_sheet) + ")"
    raise FormulaError("内部错误：未知 AST 节点")


def _quote_sheet(name: str) -> str:
    if re.fullmatch(r"[A-Za-z0-9_一-鿿.]+", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def iter_range_cells(rng: Range):
    for r in range(rng.start.row, rng.end.row + 1):
        for c in range(rng.start.col, rng.end.col + 1):
            yield rng.start.sheet, c, r
