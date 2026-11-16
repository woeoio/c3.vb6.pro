"""C30-10 正则引擎「扩展档」的判据来源（Python re）。

为什么单开一份：VB 的 VBScript.RegExp 不认命名组/后顾/原子组/占有量词，
这些是 C30-10 加的超集语法（默认开启、不设开关）。它们的语义不再由
VBScript 提供，改由 PCRE2 家族的共同子集提供 —— 这里用 Python re 当代理，
因为 Python 3.11+ 已有原子组与占有量词，3.14 已有 \\z，与 PCRE2 在这些点上一致。

每条用例给两份模式：c3pat（本引擎的写法）与 pypat（等价的 Python 写法）。
两者不是简单字符串替换：VB 侧要覆盖 (?<n>) (?'n') (?P<n>) 三种命名写法与 \\k<>，
Python 只有 (?P<n>) 一种，所以不能拿同一份模式硬套。

用法: python ext_oracle.py   ->  stdout 每行 "id|mode|结果"
ReMain.bas 的 ExtCases() 写出的 re_c3_ext.txt 必须与它逐行相同。
"""
import re

# (id, c3 模式, Python 等价模式, 目标串, 模式 t=存在性 / m=第一次匹配+子组, 手工期望)
# 手工期望只给 Python 表达不出来的那两条，并在注释里写明理由。
CASES = [
    ("e01", r"(?<c>a)\k<c>",          r"(?P<c>a)(?P=c)",        "aa",     "t", None),
    ("e02", r"(?P<c>a)(?P=c)",        r"(?P<c>a)(?P=c)",        "aa",     "t", None),
    ("e03", r"(?'x'b)\k'x'",          r"(?P<x>b)(?P=x)",        "bb",     "t", None),
    ("e04", r"(?<n>\d+)-(\w+)",       r"(?P<n>\d+)-(\w+)",      "12-ab",  "m", None),
    ("e05", r"(?<x>a)",               r"(?P<x>a)",              "a",      "t", None),
    # 前向引用：PCRE 允许，Python 不允许；本引擎选择与 Python 一致（5017）
    ("e06", r"\k<x>(?<x>a)",          r"\k<x>(?<x>a)",          "a",      "t", None),
    # 后顾（定长）
    ("e07", r"(?<=foo)bar",           r"(?<=foo)bar",           "foobar", "m", None),
    ("e08", r"(?<=\d{2})x",           r"(?<=\d{2})x",           "a12x",   "m", None),
    ("e09", r"(?<!foo)bar",           r"(?<!foo)bar",           "foobar", "t", None),
    ("e10", r"(?<!foo)bar",           r"(?<!foo)bar",           "xbar",   "t", None),
    # 变长后顾：Python 报 "look-behind requires fixed-width"；本引擎 5017
    ("e11", r"(?<=a*)b",              r"(?<=a*)b",              "ab",     "t", None),
    # 原子组 / 占有量词（Python 3.11+ 有原生支持，正好当对照）
    ("e12", r"(?>a+)ab",              r"(?>a+)ab",              "aaab",   "t", None),
    ("e13", r"(?>a+)c",               r"(?>a+)c",               "aaac",   "t", None),
    ("e14", r"a++ab",                 r"a++ab",                 "aaab",   "t", None),
    ("e15", r"a*+a",                  r"a*+a",                  "aaa",    "t", None),
    ("e16", r"\d++",                  r"\d++",                  "123",    "t", None),
    # \A / \z / \Z
    ("e17", r"\Aabc",                 r"\Aabc",                 "abc",    "t", None),
    ("e18", r"\Aabc",                 r"\Aabc",                 "xabc",   "t", None),
    ("e19", r"abc\z",                 r"abc\Z",                 "abc",    "t", None),
    ("e20", r"abc\z",                 r"abc\Z",                 "abc\n",  "t", None),
    # PCRE 的 \Z = 末尾（允许结尾一个换行）；Python 没有对应物（它的 \Z 是绝对末尾），
    # 所以这一条是**手工期望**：abc\Z 在 "abc\n" 上应当命中。
    ("e21", r"abc\Z",                 None,                     "abc\n",  "t", "1"),
    ("e22", r"abc\Z",                 r"abc\Z",                 "abc",    "t", None),
    # \Q...\E：Python 没有引号区，用 re.escape 表达同一个字面量
    ("e23", r"\Qa.b\E",               r"a\.b",                  "a.b",    "t", None),
    ("e24", r"\Qa.b\E",               r"a\.b",                  "axb",    "t", None),
    ("e25", r"\Qa+b\E",               r"a\+b",                  "a+b",    "t", None),
]


def one(cid, pypat, subj, mode, manual):
    if manual is not None:
        return f"{cid}|{mode}|{manual}"
    try:
        rx = re.compile(pypat)
    except re.error:
        return f"{cid}|err|5017"
    if mode == "t":
        return f"{cid}|t|{1 if rx.search(subj) else 0}"
    m = rx.search(subj)
    if m is None:
        return f"{cid}|m|-"
    out = f"{cid}|m|{m.start()}:{m.end() - m.start()}"
    for g in range(1, rx.groups + 1):
        out += "|[" + (m.group(g) or "") + "]"
    return out


def main():
    for cid, _c3, pypat, subj, mode, manual in CASES:
        print(one(cid, pypat, subj, mode, manual))


if __name__ == "__main__":
    main()
