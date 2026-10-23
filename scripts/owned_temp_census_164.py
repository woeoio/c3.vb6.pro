# owned_temp_census_164.py — 账 #463 §B164: 「现造带权临时量、无人持有」在**语料**上的清点
#
# 问句（研究用，**零行为改动**）：
#   语料里有多少处「一条语句里现造的带权临时量，传进去之后没人持有」？
#   判据（权威全部来自 scripts/owned_temp_keeps_164.txt，那张表由
#   scripts/check_owned_temp_keeps.ps1 从 RTL **推导** ⇒ 哨兵双向对账 ⇒ 漂不了）：
#     ① mint 侧：出口在 [own] 节（新造且归我们所有）；不在 [borrow]（借来）/[alias]（别名）
#        —— 档号**只由出口的返回档定**（BSTR / VARIANT），**不是**消费者的 VB 形参档
#           （#461 的病根就是拿了形参档 ⇒ 一枚 double 被声明成 BSTR ⇒ cl C2440）
#     ② consumer 侧：keeps_pointer == '0'（不带走指针）· 不在 closed 族 ·
#        **不在消费轴上**（消费者自己放掉这一格 ⇒ 再放就是双放 ⇒ 堆损坏）
#     ③ 站点的形状 = 「整枚实参 = 一个 mint 调用」—— 站点取**发码文本**（--emit-c），
#        所以「消费者第几格」取的是**发码里那一格的位号**（不是 VB 实参下标）
#
# 追记（账 #466 §B164 清点，2026-10-22）：
#   ① **门 ⑩**（拼接站右操作数提名时，左操作数不许含调用 ⇒ 求值次序不许动）**已建模**
#      ⇒ 理由桶里多一项 concat-left-has-call。
#   ② **门 ⑨**（ElseIf 的条件：那一枚的声明落在 else 块内、收回落在块外 ⇒ 引用越界）
#      **没法从发码文本建模**（那条结构在发码里已经看不出是 ElseIf）⇒ 那一项仍缺 ⇒
#      本工具的**提名面仍是上界**（门 ⑨ 只会**减少**它）。
#   ③ 扫法修了一处**真缺陷**：老写法在找到一枚调用后，外层接着扫**它里面**的文本，
#      而递归进左操作数时又把同一段文本再扫一遍 ⇒ 深链（Charts 2020 那种几十层拼接）
#      上是**指数级** ⇒ 上一版在真语料上**跑不完**（>30 分钟被杀）。现在：整枚调用**只数一次**
#      （外层扫到它的右括号之后），左操作数整支归 'left' 族单独记 ⇒ 496 份语料 5.7 分钟跑完。
#
#   **这一版量到的三笔数**（496 份语料 / 本机 C3.exe；账 #467 §B164 左操作数那一刀之后）：
#     判据面（criteria） = 8,259 枚（BSTR 7,519 / VARIANT 740）
#     左操作数族（下一刀的承面，不进提名面）= 9,429 枚（BSTR 9,273 / VARIANT 156）
#       —— 比 #466 那版（5,949）大：这一版把**拼接站的左操作数本身**也算进这一族了
#          （那一版只算了「左操作数**里面**」的那些）
#       够得着（自己的文本不含 vb6_BSTR_Concat 子串）= 6,390
#         其中左操作数**整枚就是**它 = 2,454（在拼接站内部次序不变）
#         其中左操作数里还有别的部分 = 3,936（次序要另算）
#       够不着（含那个子串 ⇒ 拼接站那条判据会翻面）= 3,039
#     没进提名面的理由：concat-left-has-call 2,611 / others-have-call 2,543 / keeps 1,749 / closed 11
#     （borrow / alias / consumes 三条机械证明仍**全为 0**）
#   ⚠ **判据面 ≠ 实现面**：实现在真语料上只提名了 **1,757** 枚（BSTR 1,749 / VARIANT 8）⇒
#     判据的 8,400 里有一大半**够不着** —— 那些站点是发码层用**裸字符串**拼出来的
#     （`"vb6_CStr(" + expr + ")"` / `"vb6_VariantToString(" + expr + ")"` 那一族，
#     见 src/backend/expr/cgen_expr_binary_util.cpp）⇒ **压根不走实参表那条钩子**。
#     那不是判据的错，是**覆盖面**的账：谁想收那一族，得先把那些站点改成走实参表。
#
# 跑法:
#   python scripts/owned_temp_census_164.py --dump     # 只把语料发码落到缓存
#   python scripts/owned_temp_census_164.py            # 清点 + 机械证明
#   退出码: 0 = 证明都成立；1 = 有证明不成立（那是判据或表出了事）
import argparse
import collections
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

TABLE = os.path.join(ROOT, 'scripts', 'owned_temp_keeps_164.txt')
C3 = os.path.join(ROOT, '.build', 'C3.exe')
CACHE = os.path.join(ROOT, 'output', '_tmp', 'census_c')


# ---------------------------------------------------------------- 证据表
def load_table():
    own, alias, borrowed, keeps, closed, consumes = {}, set(), set(), {}, set(), {}
    section = None
    for raw in open(TABLE, encoding='utf-8'):
        line = raw.rstrip('\n')
        if not line.strip() or line.startswith('#'):
            # 节头：# --- ... owned ... / alias ... / 消费轴 ... / 借用轴 ...
            if '返回侧所有权：owned' in line:
                section = 'own'
            elif 'alias = 交出去的是' in line:
                section = 'alias'
            elif '消费轴' in line:
                section = 'consumes'
            elif '借用轴' in line:
                section = 'borrow'
            continue
        if line.startswith('['):
            tag, rest = line.split(']', 1)
            rest = rest.strip()
            if tag == '[own':
                own[rest.split()[0]] = rest.split()[1]
            elif tag == '[alias':
                alias.add(rest.split()[0])
            elif tag == '[borrow':
                borrowed.add(rest.split()[0])
            section = None
            continue
        if section == 'consumes':
            f = line.split('\t')
            if len(f) >= 3:
                consumes[(f[0], int(f[0].split('|')[1]))] = (f[1], f[2], f[3] if len(f) > 3 else '')
            continue
        if section is not None:
            continue
        f = line.split('\t')
        if len(f) < 6:
            continue
        nm, slot = f[0].split('|')
        keeps[(nm, int(slot))] = f[1]
        if f[4] == 'closed':
            closed.add(nm)
    return own, alias, borrowed, keeps, closed, consumes


# ---------------------------------------------------------------- 发码解析
IDENT = re.compile(r'[A-Za-z_][A-Za-z0-9_]*')
CALL = re.compile(r'^([A-Za-z_][A-Za-z0-9_]*)\s*\(')


def strip_code(src):
    """去掉注释与字符串字面量（保留位置，换成空格）"""
    out = list(src)
    i, n = 0, len(src)
    while i < n:
        c = src[i]
        if c == '/' and i + 1 < n and src[i + 1] == '/':
            while i < n and src[i] != '\n':
                out[i] = ' '
                i += 1
        elif c == '/' and i + 1 < n and src[i + 1] == '*':
            out[i] = out[i + 1] = ' '
            i += 2
            while i < n and not (src[i] == '*' and i + 1 < n and src[i + 1] == '/'):
                if src[i] != '\n':
                    out[i] = ' '
                i += 1
            if i < n:
                out[i] = ' '
                i += 2
        elif c in '"\'':
            q = c
            out[i] = ' '
            i += 1
            while i < n and src[i] != q:
                if src[i] == '\\':
                    out[i] = ' '
                    i += 1
                if i < n:
                    out[i] = ' '
                    i += 1
            if i < n:
                out[i] = ' '
                i += 1
        else:
            i += 1
    return ''.join(out)


def match_paren(s, i):
    d = 0
    while i < len(s):
        if s[i] == '(':
            d += 1
        elif s[i] == ')':
            d -= 1
            if d == 0:
                return i
        i += 1
    return -1


def split_args(s, o, c):
    args, d, st = [], 0, o + 1
    for i in range(o + 1, c):
        ch = s[i]
        if ch in '([{':
            d += 1
        elif ch in ')]}':
            d -= 1
        elif ch == ',' and d == 0:
            args.append(s[st:i])
            st = i + 1
    args.append(s[st:c])
    return [a.strip() for a in args]


def corpus_inputs():
    ins = []
    for dirpath, dirnames, filenames in os.walk(os.path.join(ROOT, 'tests')):
        dirnames[:] = [d for d in dirnames if d not in ('bin', 'obj', 'output')]
        for fn in filenames:
            if fn.lower().endswith(('.bas', '.vbp')):
                ins.append(os.path.join(dirpath, fn))
    ins.sort()
    return ins


def dump_corpus(inputs):
    os.makedirs(CACHE, exist_ok=True)
    rc_bad = 0
    for path in inputs:
        rel = os.path.relpath(path, ROOT).replace('\\', '/')
        flat = re.sub(r'[\\/ \-]', '_', rel)
        dst = os.path.join(CACHE, flat + '.c')
        if os.path.exists(dst):
            continue
        try:
            p = subprocess.run([C3, path, '--emit-c'], capture_output=True, timeout=180)
        except Exception:
            rc_bad += 1
            continue
        if p.returncode != 0:
            rc_bad += 1
        with open(dst, 'wb') as f:
            f.write(p.stdout)
    return rc_bad


# ---------------------------------------------------------------- 主
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dump', action='store_true', help='只把语料发码落到缓存')
    ap.add_argument('--c3', default='', help='用哪一个 C3.exe 发码（默认 .build/C3.exe）')
    ap.add_argument('--cache', default='', help='发码缓存目录（默认 output/_tmp/census_c）')
    args = ap.parse_args()
    if args.c3:
        globals()['C3'] = os.path.abspath(args.c3)
    if args.cache:
        globals()['CACHE'] = os.path.abspath(args.cache)

    own, alias, borrowed, keeps, closed, consumes = load_table()
    print('evidence: own=%d alias=%d borrow=%d keeps=%d closed=%d consumes=%d'
          % (len(own), len(alias), len(borrowed), len(keeps), len(closed), len(consumes)))

    inputs = corpus_inputs()
    print('corpus inputs =', len(inputs))
    print('measured binary =', C3)
    print('cache =', CACHE)
    rc_bad = dump_corpus(inputs)
    print('emit rc!=0 =', rc_bad, '(那一族是 *_neg* 负例夹具，本来就红)')
    if args.dump:
        return 0

    def mint_kind(callee):
        if callee in borrowed or callee in alias:
            return 0
        if callee in own:
            return 'V' if own[callee] == 'VARIANT' else 'B'
        return 0

    def verdict(consumer, slot):
        if consumer in closed:
            return 0
        if keeps.get((consumer, slot), '0') != '0':
            return 0
        if (consumer, slot) in consumes:
            return 0
        return 1

    # ---- 拼接站左操作数那一族: 左操作数的**发码文本**被拼接站那条
    #      「含 vb6_BSTR_Concat 子串」的判据读过 ⇒ 左操作数里的一切不许提名
    #      （提了名那条判据会翻面 ⇒ 行为变了）。发码这一侧用 concatLeftDepth_ 护住。
    CONCAT_STATIONS = ('vb6_BSTR_Concat', 'vb6_BSTR_ConcatFree')

    def arg_has_call(text):
        for i in range(len(text) - 1):
            if text[i] != '(':
                continue
            j = i
            while j > 0 and (text[j - 1].isalnum() or text[j - 1] == '_'):
                j -= 1
            if j != i:
                return True
        return False

    nom = []
    shapes = collections.Counter()
    left_family = []          # 拼接站**左操作数**里那一族（下一刀的承面，不进提名面）
    left_ok = []              # 那一族里**够得着**的（自己的文本不含拼接子串）
    left_reasons = collections.Counter()
    left_whole = 0            # 够得着 ∧ 左操作数**整枚就是**它（次序上最干净）
    left_parts = 0            # 够得着 ∧ 左操作数里还有**别的**部分（次序要另算）
    reasons = collections.Counter()   # 每一格**为什么没进**提名面

    def classify(consumer, slot, mint, kind, args, protected, bucket, argtext, lefttext):
        """把一格「整枚实参 = mint」归到提名面 / 某一族理由 / 左操作数族。

           lefttext = 所在那一格的**外层左操作数文本**（左操作数族专用；其余为空）。
        """
        nonlocal left_whole, left_parts
        # 「在拼接站的左操作数里」有两种形态：
        #   ① 这一格**就是**某个拼接站的左操作数（consumer ∈ 拼接站 ∧ slot = 0）
        #   ② 这一格在**某个**拼接站的左操作数**里面**（protected = 真）
        in_left = protected or (consumer in CONCAT_STATIONS and slot == 0)
        if in_left:
            # **够不够得着**（下一刀的那条本地判据）：那一枚**自己的文本**里不许含
            # 「vb6_BSTR_Concat」子串 —— 含了 ⇒ 拼接站那条判据会翻面（ConcatFree ⇒ Concat）
            # ⇒ 不够得着；不含 ⇒ ① 判据不翻面 ② 那一枚不可能是任何 ConcatFree 的左操作数
            # （ConcatFree 的左必是拼接，其文本必含子串）⇒ 不会双放。
            has_concat = 'vb6_BSTR_Concat' in argtext
            is_whole_left = (not protected) or (bool(lefttext) and lefttext.strip() == argtext.strip())
            left_family.append((consumer, slot, mint, kind))
            if has_concat:
                left_reasons['含拼接子串（判据会翻面）'] += 1
            else:
                left_ok.append((consumer, slot, mint, kind, is_whole_left))
                if is_whole_left:
                    left_whole += 1
                else:
                    left_parts += 1
            return
        if mint in borrowed:
            reasons['borrow'] += 1
            return
        if mint in alias:
            reasons['alias'] += 1
            return
        # 门 ⑩: 拼接站**右操作数**提名时，左操作数不许含调用（提了名 ⇒ 求值次序会动）
        if consumer in CONCAT_STATIONS and slot == 1 and arg_has_call(args[0]):
            reasons['concat-left-has-call'] += 1
            return
        # 别的实参有没有调用？（有 ⇒ 提名单会改动求值次序 ⇒ 不提名）
        others_pure = True
        for j in range(len(args)):
            if j == slot:
                continue
            if arg_has_call(args[j]):
                others_pure = False
        if not others_pure:
            reasons['others-have-call'] += 1
            return
        if consumer in closed:
            reasons['closed'] += 1
            return
        if keeps.get((consumer, slot), '0') != '0':
            reasons['keeps'] += 1
            return
        if (consumer, slot) in consumes:
            reasons['consumes'] += 1
            return
        nom.append((consumer, slot, mint, kind))
        shapes[(mint, kind, consumer, slot)] += 1

    def scan(code, protected, bucket, lefttext=''):
        """扫一段发码文本; protected = 处在拼接站左操作数里（那里不许提名）"""
        i = 0
        n = len(code)
        while i < n:
            m = IDENT.match(code, i)
            if not (m and m.end() < n and code[m.end()] == '('):
                i += 1
                continue
            consumer = code[i:m.end()]
            o = m.end()
            c = match_paren(code, o)
            if c < 0:
                i += 1
                continue
            args = split_args(code, o, c)
            is_concat = consumer in CONCAT_STATIONS
            for slot, a in enumerate(args):
                mm = CALL.match(a)
                if mm and a.count('(') == a.count(')'):
                    mint = mm.group(1)
                    kind = mint_kind(mint)
                    if kind:
                        classify(consumer, slot, mint, kind, args, protected, bucket, a, lefttext)
                # 递归进这一格: 拼接站的**左操作数**（slot 0）整支受保护 ⇒ 归 'left' 族;
                # 其余实参继承 protected。外层扫到 c+1 ⇒ 整枚调用**只数一次**。
                if is_concat and slot == 0:
                    scan(a, True, 'left', a)
                else:
                    scan(a, protected, bucket, lefttext)
            i = c + 1

    for path in sorted(os.listdir(CACHE)):
        if not path.endswith('.c'):
            continue
        code = strip_code(open(os.path.join(CACHE, path), encoding='latin-1').read())
        scan(code, False, 'nom')

    print('\nnomination set =', len(nom))
    proofs = [
        ('intersect closed', sum(1 for n in nom if n[0] in closed)),
        ('intersect borrow', sum(1 for n in nom if n[2] in borrowed)),
        ('intersect alias', sum(1 for n in nom if n[2] in alias)),
        ('intersect consumes', sum(1 for n in nom if (n[0], n[1]) in consumes)),
    ]
    bad = 0
    for name, v in proofs:
        print('  proof %-20s = %d' % (name, v))
        if v != 0:
            bad += 1
    b = sum(1 for n in nom if n[3] == 'B')
    print('  kinds: BSTR = %d  VARIANT = %d' % (b, len(nom) - b))
    print('  左操作数族（不进提名面，下一刀的承面）= %d' % len(left_family))
    lb = sum(1 for n in left_family if n[3] == 'B')
    print('    其中 BSTR = %d  VARIANT = %d' % (lb, len(left_family) - lb))
    print('    够得着（自己的文本不含 vb6_BSTR_Concat 子串）= %d' % len(left_ok))
    print('      其中左操作数**整枚就是**它 = %d（次序上最干净）' % left_whole)
    print('      其中左操作数里还有别的部分 = %d（次序要另算）' % left_parts)
    for r, v in left_reasons.most_common():
        print('      够不着: %-22s = %d' % (r, v))
    print('  没进提名面的理由:')
    for r, v in reasons.most_common():
        print('    %-22s = %d' % (r, v))
    print('\n--- top shapes ---')
    for (mint, kind, cons, slot), n in shapes.most_common(25):
        print('  mint=%-22s(%s) -> consumer=%-26s slot=%d n=%d' % (mint, kind, cons, slot, n))
    if bad:
        print('\nFAIL: %d proofs do not hold' % bad)
        return 1
    print('\nALL-PROOFS-HOLD')
    return 0


if __name__ == '__main__':
    sys.exit(main())
