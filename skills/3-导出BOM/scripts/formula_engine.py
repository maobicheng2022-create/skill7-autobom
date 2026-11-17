# -*- coding: utf-8 -*-
"""
formula_engine.py — AutoBOM FormulaEngine.cs 的 Python 移植
在 AutoBOM 里用户给某属性填公式(如 厚度 = ROUND({$体积}/({$表面积}/2),0))，
引用其它属性用 {属性名} 占位。本模块复刻其递归下降解析器。

支持: 四则 + - * / 括号 负号 比较 > < >= <= == !=(返回1/0) & 字符串拼接
函数 SUM MAX MIN ABS ROUND IF
数值参与计算; 非数值走字符串拼接; 空值数值上下文按0、文本按""
出错时返回原公式
"""
import re


def is_numeric_context(formula):
    body = re.sub(r"\{[^{}]+\}", "x", formula)
    return ("&" not in body and '"' not in body and "'" not in body and "IF(" not in body)


_PLAIN_CHARS = set("+-*/(),&><!=")


def is_plain_value(expr):
    if any(c in _PLAIN_CHARS for c in expr):
        return False
    for fn in ("IF(", "SUM(", "MAX(", "MIN(", "ABS(", "ROUND("):
        if fn in expr:
            return False
    return True


def to_display(val):
    if val is None:
        return ""
    if isinstance(val, bool):
        return "1" if val else "0"
    if isinstance(val, float):
        if abs(val - round(val)) < 1e-9:
            return str(int(round(val)))
        s = ("%.8f" % val).rstrip("0").rstrip(".")
        return s if s else "0"
    return str(val)


class _Parser:
    def __init__(self, src):
        self.s = src
        self.pos = 0

    def fail(self):
        raise ValueError("公式解析失败: " + self.s)

    def cur(self):
        return self.s[self.pos] if self.pos < len(self.s) else ""

    def skip_ws(self):
        while self.pos < len(self.s) and self.s[self.pos].isspace():
            self.pos += 1

    def parse(self):
        v = self.comparison()
        self.skip_ws()
        if self.pos < len(self.s):
            self.fail()
        return v

    def comparison(self):
        left = self.addsub()
        while True:
            self.skip_ws()
            op = ""
            rest = self.s[self.pos:]
            for o in (">=", "<=", "==", "!=", ">", "<"):
                if rest.startswith(o):
                    op = o
                    self.pos += len(o)
                    break
            if not op:
                break
            right = self.addsub()
            x, y = self.try_num(left), self.try_num(right)
            if x is not None and y is not None:
                r = 1.0 if self.compare_num(x, y, op) else 0.0
            else:
                cmp = (to_display(left) > to_display(right)) - (to_display(left) < to_display(right))
                r = 1.0 if self.compare_text(cmp, op) else 0.0
            left = r
        return left

    @staticmethod
    def compare_num(x, y, op):
        if op == ">": return x > y
        if op == "<": return x < y
        if op == ">=": return x >= y
        if op == "<=": return x <= y
        if op == "==": return abs(x - y) < 1e-9
        if op == "!=": return abs(x - y) >= 1e-9
        return False

    @staticmethod
    def compare_text(cmp, op):
        if op == ">": return cmp > 0
        if op == "<": return cmp < 0
        if op == ">=": return cmp >= 0
        if op == "<=": return cmp <= 0
        if op == "==": return cmp == 0
        if op == "!=": return cmp != 0
        return False

    def addsub(self):
        left = self.muldiv()
        while True:
            self.skip_ws()
            c = self.cur()
            if c == "+":
                self.pos += 1
                left = self.add_num(left, self.muldiv())
            elif c == "-":
                self.pos += 1
                left = self.sub_num(left, self.muldiv())
            elif c == "&":
                self.pos += 1
                left = self.concat_str(left, self.muldiv())
            else:
                break
        return left

    def muldiv(self):
        left = self.unary()
        while True:
            self.skip_ws()
            c = self.cur()
            if c == "*":
                self.pos += 1
                left = self.mul_num(left, self.unary())
            elif c == "/":
                self.pos += 1
                left = self.div_num(left, self.unary())
            else:
                break
        return left

    def unary(self):
        self.skip_ws()
        c = self.cur()
        if c == "-":
            self.pos += 1
            return -self.unary()
        if c == "+":
            self.pos += 1
            return self.unary()
        return self.primary()

    def primary(self):
        self.skip_ws()
        c = self.cur()
        if c == "(":
            self.pos += 1
            v = self.comparison()
            self.skip_ws()
            if self.cur() != ")":
                self.fail()
            self.pos += 1
            return v
        if c in ('"', "'"):
            return self.read_string()
        if c.isdigit() or c == ".":
            return self.read_number()
        if c.isalpha():
            return self.read_function()
        self.fail()

    def read_string(self):
        q = self.cur()
        self.pos += 1
        out = []
        while self.pos < len(self.s) and self.s[self.pos] != q:
            out.append(self.s[self.pos])
            self.pos += 1
        if self.pos >= len(self.s):
            self.fail()
        self.pos += 1
        return "".join(out)

    def read_number(self):
        start = self.pos
        while self.pos < len(self.s) and (self.s[self.pos].isdigit() or self.s[self.pos] == "."):
            self.pos += 1
        t = self.s[start:self.pos]
        try:
            return float(t)
        except ValueError:
            self.fail()

    def read_function(self):
        start = self.pos
        while self.pos < len(self.s) and self.s[self.pos].isalpha() and (self.pos - start) < 20:
            self.pos += 1
        name = self.s[start:self.pos].upper()
        self.skip_ws()
        if self.cur() != "(":
            self.fail()
        self.pos += 1
        args = []
        self.skip_ws()
        if self.cur() == ")":
            self.pos += 1
        else:
            while True:
                args.append(self.comparison())
                self.skip_ws()
                if self.cur() == ",":
                    self.pos += 1
                    continue
                if self.cur() == ")":
                    self.pos += 1
                    break
                self.fail()
        if name == "SUM":
            return sum(self.try_num(a) or 0 for a in args)
        if name == "MAX":
            nums = [self.try_num(a) for a in args]
            nums = [n for n in nums if n is not None]
            return max(nums) if nums else 0
        if name == "MIN":
            nums = [self.try_num(a) for a in args]
            nums = [n for n in nums if n is not None]
            return min(nums) if nums else 0
        if name == "ABS":
            x = self.try_num(args[0]) if args else None
            return abs(x) if x is not None else 0
        if name == "ROUND":
            x = self.try_num(args[0]) if args else None
            n = 0
            if x is not None:
                if len(args) > 1:
                    y = self.try_num(args[1])
                    n = int(y) if y is not None else 0
                return round(x, n)
            return 0
        if name == "IF":
            if len(args) < 3:
                return args[0] if args else ""
            cond = self.try_num(args[0])
            if cond is not None:
                truthy = abs(cond) > 1e-12
            else:
                truthy = not (args[0] is None or (isinstance(args[0], str) and args[0] == ""))
            return args[1] if truthy else args[2]
        self.fail()

    @staticmethod
    def try_num(o):
        if isinstance(o, (int, float)):
            return float(o)
        if isinstance(o, str):
            try:
                s = o.strip()
                if " " in s:
                    return None
                return float(s)
            except ValueError:
                return None
        return None

    def add_num(self, a, b):
        x, y = self.try_num(a), self.try_num(b)
        if x is not None and y is not None:
            return x + y
        return self.concat_str(a, b)

    def sub_num(self, a, b):
        x, y = self.try_num(a), self.try_num(b)
        return x - y if (x is not None and y is not None) else a

    def mul_num(self, a, b):
        x, y = self.try_num(a), self.try_num(b)
        return x * y if (x is not None and y is not None) else a

    def div_num(self, a, b):
        x, y = self.try_num(a), self.try_num(b)
        if x is not None and y is not None:
            if abs(y) < 1e-12:
                return a
            return x / y
        return a

    @staticmethod
    def concat_str(a, b):
        return to_display(a) + to_display(b)


def evaluate(formula, get_value):
    if not formula or not formula.strip():
        return ""
    try:
        numeric_ctx = is_numeric_context(formula)
        def _rep(m):
            name = m.group(1).strip()
            if not name:
                return "0" if numeric_ctx else '""'
            v = get_value(name)
            if not v:
                return "0" if numeric_ctx else '""'
            s = v.strip() if isinstance(v, str) else str(v).strip()
            if " " not in s:
                try:
                    return str(float(s))
                except ValueError:
                    pass
            return '"' + str(v).replace('"', "'") + '"'
        expr = re.sub(r"\{([^{}]+)\}", _rep, formula).strip()
        if not expr:
            return ""
        if is_plain_value(expr):
            return expr
        p = _Parser(expr)
        val = p.parse()
        return to_display(val)
    except Exception:
        return formula
