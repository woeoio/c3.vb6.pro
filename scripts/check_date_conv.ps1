# ============================================================
# check_date_conv.ps1 - 账 #305 #417 (第九十四刀) 的结构性哨兵
#
# 为什么单开这一道: 「这串文本是哪一天」这一问在 RTL 里有三位各自抄写的主人 (vb6_DateValue、
# vb6_TimeValue、vb6rtl_conv.c 里那枚 CStr(日期形文本) 入口)，三处都自己调 VarDateFromStr、
# 三处都在解析不出时静默返 0；而 vb6_CDate 收文本那一支连日期解析都没问，转调的是**数值**解析
# (vb6_Val) —— 注释里却写着「字符串按日期解析」。实测 (本机 VB-family 运行期, 尺
# .build\k233\k234\k235\k236) CDate(2026/1/1) 答 2026 (那一年) 而不是那一天、CDate(abc) 答 0
# 且 Err 还是 0，而同一份程序里 DateValue(2026/1/1) 是对的。VB6 两句都是 13。
# 本刀把解析收成一条站点、把「认不出算几号」交给两位主人 (CDate 那一档还认纯数字串，
# DateValue/TimeValue 那一档不认 —— 这是实测出来的两档，不是随手兜底)。
# 发码 A/B 看不见这一族 (改的全在函数体，发出去的 C 文本一字不动 —— 本刀 BASE/NEW 两份
# --emit-c 逐字节相同)，语料普查也只说得出面有多大 (80 处文本形状的 CDate) => 行为判据住
# tests\b417 + Test-Vbp(两台)，「不许退回两遍」由这里钉形状。
#
# 规则:
#   R1  date.c 里解析那一句只许问 oleaut32 一次: VarDateFromStr 恰好 1 次 (住在唯一那条
#       vb6_TryDateFromStr 里，站点合计 4 = 定义 + 三位主人各一次转调 (DateParse / DateVal /
#       DateCanParse，第三枚是第九十五刀的 #418 给 IsDate 开的口))，交回 13 的那句恰好
#       1 次 (在 vb6_DateParse 里)，两位转换主人的定义各 1 次 (第三枚主人是判据，形状由
#       check_ispred_text 钉)，旧的「(int32_t) 砍成整天」不许回来。
#   R2  全仓 RTL 再不许有第二条日期解析路: src\rtl 整棵树里 VarDateFromStr 合计恰好 1；
#       conv.c 里日期那一族不许再出现 vb6_Val(v.bstrVal) 或 vb6_CStrDate(vb6_Val(...)。
#   R3  CDate 的文本分支问的是 vb6_DateVal；CStr(日期形文本) 入口问的是 vb6_DateParse。
#   R4  两位投影各归各的半边: DateValue = floor(解析)，TimeValue = 小数 (d - floor(d))。
#   R5  两位主人在 vb6rtl_builtin.h 各有声明 (发码那侧之外的人也得问得到它们)。
#   R6  发码那张表一字没动: cdate / datevalue / timevalue 三行各自还在 (本刀不动发码)。
#   R7  判据两头都在: run_tests.ps1 里 b417 两台各 1、本哨兵已登记。
#   R8  §11.183: 隐式转换矩阵里 String→Date 与 数值→Date 各恰好一行。
#   R9  §11.183: 赋值站的数组元素腿 —— 判据名/声明档问句/出口各恰好一处。
#   R10 §11.183: b208 夹具两台在跑。
#   R11 §11.184: 赋值站那一腿**不再按秩分叉** (`positional.size() == 1` 不许回来,
#       改成非空即认)。
#   R12 §11.184: 装箱站 (dateElemVariantBox) 同口径 —— `size() != 1` 不许回来。
#   R13 §11.184: b209 夹具两台在跑。
#   R14 §11.187: 赋值站那条腿认得 UDT 成员数组 (问 inferUdtFieldVb6Type, 且带 isArray 条件)。
#   R15 §11.187: b211 夹具两台在跑。
#   R16 §11.188: Variant 右值那一档走 vb6_CDate (vb6_VariantToDouble 不许留在赋值站)。
#   R17 §11.188: b212 夹具两台在跑。
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$D = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c')
$C = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_conv.c')
$H = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_builtin.h')
$G = Get-Content -Raw (Join-Path $root 'src\backend\detail\expr\cgen_expr_ident_builtin_table.inc')
$T = Get-Content -Raw (Join-Path $root 'tests\run_tests.ps1')
# §11.183 (第二百刀): 赋值站的 String→Date 隐含转换 —— 矩阵那一半与发码那一半各一份形状。
$S = Get-Content -Raw (Join-Path $root 'src\semantics\type_system.cpp')
$A = Get-Content -Raw (Join-Path $root 'src\backend\detail\stmt\cgen_assign_value_sem.inc')
# §11.184 (第二百零一刀): 同一问题的另一站 —— Date 元素装进 Variant 的那张标签。
$B = Get-Content -Raw (Join-Path $root 'src\backend\cgen_util_ctrl.cpp')
# 整棵 RTL 树的日期解析站点计数 (R2 的面): 递归读每个 .c/.h, 只数 VarDateFromStr
$rtlFiles = Get-ChildItem -Path (Join-Path $root 'src\rtl') -Recurse -Include *.c,*.h -File
$r2sites = 0
foreach ($f in $rtlFiles) {
    $r2sites += ([regex]::Matches((Get-Content -Raw $f.FullName), 'VarDateFromStr\(')).Count
}
$fail = @()
function N([string]$pat, [string]$txt) {
    return ([regex]::Matches($txt, $pat, 'Multiline')).Count
}
# R1
$r1parse = N 'VarDateFromStr\(' $D
$r1try = N 'vb6_TryDateFromStr\(' $D
$r1raise = N 'vb6_RaiseError\(13, vb6_BSTR_FromStr\(L"Type mismatch"\)\)' $D
$r1dp = N '^double vb6_DateParse\(BSTR s\) \{' $D
$r1dv = N '^double vb6_DateVal\(BSTR s\) \{' $D
$r1old = N '\(double\)\(int32_t\)result' $D
if ($r1parse -ne 1) { $fail += ('R1 VarDateFromStr in date.c = ' + $r1parse + ' (want 1)') }
# §11.191 第二百零八刀: 第四位消费方是 vb6_TryDateValQuiet (Variant 比较那一档的静默解析)。
# VarDateFromStr 仍然只 1 处 —— 新增的是消费方, 不是第二个解析站点。
if ($r1try -ne 5) { $fail += ('R1 vb6_TryDateFromStr sites = ' + $r1try + ' (want 5 = def + the four owners)') }
if ($r1raise -ne 1) { $fail += ('R1 raise-13 outlet in date.c = ' + $r1raise + ' (want 1, inside vb6_DateParse)') }
if ($r1dp -ne 1) { $fail += ('R1 vb6_DateParse definition = ' + $r1dp + ' (want 1)') }
if ($r1dv -ne 1) { $fail += ('R1 vb6_DateVal definition = ' + $r1dv + ' (want 1)') }
if ($r1old -ne 0) { $fail += ('R1 the retired (int32_t) truncation is back = ' + $r1old + ' (want 0)') }
# R2
if ($r2sites -ne 1) { $fail += ('R2 date-parse sites across src\rtl = ' + $r2sites + ' (want 1: one authority only)') }
$r2lenient = N 'return vb6_Val\(v\.bstrVal\);' $C
$r2csd = N 'vb6_CStrDate\(vb6_Val\(' $C
if ($r2lenient -ne 0) { $fail += ('R2 CDate text branch still asks vb6_Val = ' + $r2lenient + ' (want 0)') }
if ($r2csd -ne 0) { $fail += ('R2 CStr(date-text) fallback still asks vb6_Val = ' + $r2csd + ' (want 0)') }
# R3
$r3branch = N 'return vb6_DateVal\(v\.bstrVal\);' $C
$r3entry = N 'return vb6_CStrDate\(vb6_DateParse\(s\)\);' $C
if ($r3branch -ne 1) { $fail += ('R3 vb6_CDate vtBSTR -> vb6_DateVal = ' + $r3branch + ' (want 1)') }
if ($r3entry -ne 1) { $fail += ('R3 vb6_CStrDateFromBSTR -> vb6_DateParse = ' + $r3entry + ' (want 1)') }
# R4
$r4dv = N 'return floor\(vb6_DateParse\(dateStr\)\);' $D
$r4tv = N 'double d = vb6_DateParse\(timeStr\);' $D
$r4frac = N 'return d - floor\(d\);' $D
if ($r4dv -ne 1) { $fail += ('R4 DateValue projection = ' + $r4dv + ' (want 1)') }
if ($r4tv -ne 1) { $fail += ('R4 TimeValue parses through the authority = ' + $r4tv + ' (want 1)') }
if ($r4frac -ne 1) { $fail += ('R4 TimeValue fraction projection = ' + $r4frac + ' (want 1)') }
# R5
$r5dp = N 'double vb6_DateParse\(BSTR s\);' $H
$r5dv = N 'double vb6_DateVal\(BSTR s\);' $H
if ($r5dp -ne 1) { $fail += ('R5 vb6_DateParse declaration = ' + $r5dp + ' (want 1)') }
if ($r5dv -ne 1) { $fail += ('R5 vb6_DateVal declaration = ' + $r5dv + ' (want 1)') }
# R6
$r6cd = N '\{"cdate",\s*"vb6_CDate"\}' $G
$r6dvl = N '\{"datevalue",\s*"vb6_DateValue"\}' $G
$r6tv = N '\{"timevalue",\s*"vb6_TimeValue"\}' $G
if (($r6cd + $r6dvl + $r6tv) -ne 3) {
    $fail += ('R6 codegen table rows cdate/datevalue/timevalue = ' + ($r6cd + $r6dvl + $r6tv) + ' (want 3, this knife moves no codegen)')
}
# R7
$r7cells = N 'b417_date_conv' $T
$r7sent = N 'check_date_conv\.ps1' $T
if ($r7cells -ne 2) { $fail += ('R7 b417 run cells = ' + $r7cells + ' (want 2 = x64 + x86)') }
if ($r7sent -ne 1) { $fail += ('R7 this sentinel is not registered in run_tests.ps1 = ' + $r7sent + ' (want 1)') }

# R8 (§11.183 第二百刀): 隐式转换矩阵里 String→Date 与 数值→Date 各恰好一行。
# 旧形只有 Date↔Double 那两行 ⇒ 三条 String→Date 形态全报假 VB3003 (真机 P241 八行)。
$r8str = N 'if \(from == Vb6Type::String && to == Vb6Type::Date\) return true;' $S
$r8num = N 'if \(isNumeric\(from\) && to == Vb6Type::Date\) return true;' $S
if ($r8str -ne 1) { $fail += ('R8 matrix String->Date row = ' + $r8str + ' (want 1)') }
if ($r8num -ne 1) { $fail += ('R8 matrix numeric->Date row = ' + $r8num + ' (want 1)') }
# R9: 发码侧那条数组元素腿 —— 只能有一枚判据名、一处声明档问句、一处出口。
$r9decl = N 'bool dateTarget183 = knownDateVars_\.count\(tgtDt\) != 0;' $A
$r9leg = N 'it183->second == Vb6Type::Date' $A
$r9emit = N 'if \(dateTarget183 && inferExprType\(\*node\.value\) == Vb6Type::String\) \{' $A
if ($r9decl -ne 1) { $fail += ('R9 dateTarget183 declaration = ' + $r9decl + ' (want 1)') }
if ($r9leg -ne 1) { $fail += ('R9 arrayElemTypes_ Date leg = ' + $r9leg + ' (want 1)') }
if ($r9emit -ne 1) { $fail += ('R9 Date emit outlet = ' + $r9emit + ' (want 1)') }
# R10: 夹具两台在跑，且台账里登记。
$r10cells = N '\$Tests\\b208_string_to_date\\B208Res\.vbp' $T
if ($r10cells -ne 2) { $fail += ('R10 b208 run cells = ' + $r10cells + ' (want 2 = x64 + x86)') }
# R11 (§11.184 第二百零一刀): 赋值站那条腿不再按秩分叉。旧形 `positional.size() == 1`
# 把多维那一形挡在外面 ⇒ 裸 BSTR 进 double 槽 (真机 p242 的 S1/S3/S4 三条前全部 C2440)。
$r11old = N 'ic183\.positional\.size\(\) == 1' $A
$r11new = N '!ic183\.positional\.empty\(\)' $A
if ($r11old -ne 0) { $fail += ('R11 assignment leg still gates on rank = ' + $r11old + ' (want 0)') }
if ($r11new -ne 1) { $fail += ('R11 assignment leg rank-free gate = ' + $r11new + ' (want 1)') }
# R12: 装箱站同一口径。旧形对多维交空串 ⇒ _Generic 按 double 装成 VT_R8 (真机 p243 的 T4
# 改前答 5/Double, 而 T1/T2/T3 早就是 7/Date)。
$r12old = N 'ic198\.positional\.size\(\) != 1' $B
$r12new = N 'ic198\.positional\.empty\(\)' $B
if ($r12old -ne 0) { $fail += ('R12 boxing leg still gates on rank = ' + $r12old + ' (want 0)') }
if ($r12new -ne 1) { $fail += ('R12 boxing leg rank-free gate = ' + $r12new + ' (want 1)') }
# R13: b209 夹具两台在跑。
$r13cells = N '\$Tests\\b209_nd_date_elem\\B209Res\.vbp' $T
if ($r13cells -ne 2) { $fail += ('R13 b209 run cells = ' + $r13cells + ' (want 2 = x64 + x86)') }
# R14 (§11.187 第二百零四刀): 赋值站那条腿要认得 **UDT 成员数组** 那一形 ——
# 判据只问 inferUdtFieldVb6Type (与 Boolean/Currency 元素两问同一份), 不新开表;
# 且仍然要求 outIsArray (同一字段名也可能是标量 Date)。
$r14member = N 'inferUdtFieldVb6Type\(ic183\.callee\.get\(\), &isArr187\) == Vb6Type::Date' $A
$r14arr = N '&& isArr187\)' $A
if ($r14member -ne 1) { $fail += ('R14 UDT-member Date element leg = ' + $r14member + ' (want 1)') }
if ($r14arr -ne 1) { $fail += ('R14 member leg still asks isArray = ' + $r14arr + ' (want 1)') }
# R15: b211 夹具两台在跑。
$r15cells = N '\$Tests\\b211_udt_date_elem\\B211Res\.vbp' $T
if ($r15cells -ne 2) { $fail += ('R15 b211 run cells = ' + $r15cells + ' (want 2 = x64 + x86)') }
# R16 (§11.188 第二百零五刀): Variant 右值那一档走 vb6_CDate —— 旧形 vb6_VariantToDouble
# 对文本日期/"abc"/Null 三条都是错读数 (真机 p245 的 V1/V5/V9)。
# 只看两件事: 那一腿存在, 且它的出口就是 vb6_CDate。其余站的 vb6_VariantToDouble 是合法的
# (数值目标的 Variant 解包), 不在这里一刀切。
$r16gate = N 'if \(dateTarget183 && valueIsVariant188\) \{' $A
$r16out = N 'emitLine\(target \+ " = vb6_CDate' $A
if ($r16gate -ne 1) { $fail += ('R16 Variant leg gate = ' + $r16gate + ' (want 1)') }
if ($r16out -ne 1) { $fail += ('R16 Variant leg CDate outlet = ' + $r16out + ' (want 1)') }
# R17: b212 夹具两台在跑。
$r17cells = N '\$Tests\\b212_variant_to_date\\B212Res\.vbp' $T
if ($r17cells -ne 2) { $fail += ('R17 b212 run cells = ' + $r17cells + ' (want 2 = x64 + x86)') }
# R18 (§11.189 第二百零六刀): Variant 右值的第二形 —— 调用。两个要点:
#   ① 裸名被调/下标也算载体 (否则函数返回值永进不来, 真机 p246 的 W1/W4/W8 全错);
#   ② 但已知数组的元素不得当载体 (第一版缺这条, 把 d2 = q(3) 的 Date 元素包成了
#      vb6_CDate(VB6_SA_AT(double, q, 3)) —— 能编但形状错, 形状神谕当场报 b206 两行)。
$r18call = N 'node\.value->kind == ASTNodeKind::IndexOrCallExpr' $A
$r18arr = N 'ae188 != arrayElemTypes_\.end\(\)' $A
if ($r18call -lt 1) { $fail += ('R18 call-shaped Variant leg = ' + $r18call + ' (want >= 1)') }
if ($r18arr -ne 1) { $fail += ('R18 known-array exclusion = ' + $r18arr + ' (want 1)') }
# R19: b213 夹具两台在跑。
$r19cells = N '\$Tests\\b213_variant_call_to_date\\B213Res\.vbp' $T
if ($r19cells -ne 2) { $fail += ('R19 b213 run cells = ' + $r19cells + ' (want 2 = x64 + x86)') }
# R20 (§11.190 第二百零七刀): 比较站那一腿 —— Date 与 String 相比要折成日期序列。
# 旧形把 Date 侧 wrapToBSTR 后发 vb6_StrCmp(vb6_CStrDate(d), "…"), 比的是格式化文本
# (真机 p248 的 Y2/Y3 答 0、Y7 答 1)。三枚钉子: 腿在、元素档问句在、Select Case 走 double。
$B = Get-Content -Raw (Join-Path $root 'src\backend\expr\cgen_expr_binary.cpp')
$S = Get-Content -Raw (Join-Path $root 'src\backend\stmt\cgen_select.cpp')
$r20leg = ([regex]::Matches($B, '\(lDate190 && rStr190\) \|\| \(lStr190 && rDate190\)', 'Multiline')).Count
$r20elem = ([regex]::Matches($B, 'arrayElemTypes_\.find\(Symbol::toLower\(', 'Multiline')).Count
$r20sel = ([regex]::Matches($S, 'bool isDateSelect = \(testType == Vb6Type::Date\);', 'Multiline')).Count
if ($r20leg -ne 1) { $fail += ('R20 Date-vs-String comparison leg = ' + $r20leg + ' (want 1)') }
if ($r20elem -lt 1) { $fail += ('R20 Date element档 ask in comparison = ' + $r20elem + ' (want >= 1)') }
if ($r20sel -ne 1) { $fail += ('R20 Select Case Date carrier = ' + $r20sel + ' (want 1)') }
# R21: b214 夹具两台在跑。
$r21cells = N '\$Tests\\b214_date_string_compare\\B214Res\.vbp' $T
if ($r21cells -ne 2) { $fail += ('R21 b214 run cells = ' + $r21cells + ' (want 2 = x64 + x86)') }
# R22 (§11.191 第二百零八刀): Date 与 Variant 相比 —— 两处都在 RTL 里, 一处也不许消失:
#   ① Date 侧装箱必须是 VT_DATE (否则 RTL 按“数值恒小于串”比, 真机 Z1/Z2/Z8 全反);
#   ② VT_NULL 参与比较 = 无序 ⇒ 六个出口全假 (真机 Z5), 不许再按 0 解包。
$R = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl.c')
$D = Get-Content -Raw (Join-Path $root 'src\rtl\core\vb6rtl\vb6rtl_date.c')
$r22box = ([regex]::Matches($B, 'return "vb6_VariantDate\(" \+ e \+ "\)";', 'Multiline')).Count
$r22null = ([regex]::Matches($R, 'VT_NULL\) \|\| \(b && b->vt == \(vb6_vartype\)VT_NULL\)\)', 'Multiline')).Count
$r22quiet = ([regex]::Matches($D, 'int32_t vb6_TryDateValQuiet\(BSTR s, double\* out\)', 'Multiline')).Count
if ($r22box -ne 1) { $fail += ('R22 Date operand boxed as VT_DATE = ' + $r22box + ' (want 1)') }
if ($r22null -ne 1) { $fail += ('R22 Null-is-unordered rule = ' + $r22null + ' (want 1)') }
if ($r22quiet -ne 1) { $fail += ('R22 quiet date parse = ' + $r22quiet + ' (want 1)') }
# R23: b215 夹具两台在跑。
$r23cells = N '\$Tests\\b215_date_variant_compare\\B215Res\.vbp' $T
if ($r23cells -ne 2) { $fail += ('R23 b215 run cells = ' + $r23cells + ' (want 2 = x64 + x86)') }
# R24 (§11.192 第二百零九刀): Select Case 的三处口径 —— 主语 Date 时不走整枚槽那条链、
# Variant 标签按 CDate 折、Date 标签在 Variant 主语上改走 VarCmp（不再 VarSelCmp 抛 94）。
$sel = Get-Content -Raw (Join-Path $Root 'src\backend\stmt\cgen_select.cpp')
$r24skip = N 'if \(isVariantSelect && !isDateSelect\) \{' $sel
if ($r24skip -ne 1) { $fail += ('R24 date carrier beats variant carrier = ' + $r24skip + ' (want 1)') }
$r24cd = N 'lt190 == Vb6Type::Variant\) return "vb6_CDate\("' $sel
if ($r24cd -ne 1) { $fail += ('R24 Variant label -> vb6_CDate = ' + $r24cd + ' (want 1)') }
$r24dl = N 'dateLabelPresent' $sel
if ($r24dl -ne 5) { $fail += ('R24 dateLabelPresent reads = ' + $r24dl + ' (want 5 = decl + two sets + gate + comment)') }
$r24cp = N 'inferExprType\(\*lbl\) == Vb6Type::Date\) return "vb6_VarCmp"' $sel
if ($r24cp -ne 1) { $fail += ('R24 Date label -> vb6_VarCmp = ' + $r24cp + ' (want 1)') }
# R25: b216 夹具两台在跑。
$r25cells = N '\$Tests\\b216_select_variant_label_date\\B216Res\.vbp' $T
if ($r25cells -ne 2) { $fail += ('R25 b216 run cells = ' + $r25cells + ' (want 2 = x64 + x86)') }
# R26 (§11.193 第二百一十刀): 串主语的 Date 标签走 VB6 自己的短日期文本入口 (vb6_CStrDate),
# 三处形状 (Is / 范围 / 精确) 各 1 —— 少了任何一处就会退回 `vb6_StrCmp(BSTR, 44259.0)` 编不过。
$s26 = Get-Content -Raw (Join-Path $Root 'src\backend\stmt\cgen_select.cpp')
$r26wrap = N 'vb6_CStrDate\(" \+ txt \+ "\)' $s26
if ($r26wrap -ne 1) { $fail += ('R26 Date label -> vb6_CStrDate = ' + $r26wrap + ' (want 1: one owning lambda)') }
$r26use = N 'strLabel193\(' $s26
if ($r26use -ne 4) { $fail += ('R26 strLabel193 sites = ' + $r26use + ' (want 4 = def + Is + range + exact)') }
# R27: b217 夹具两台在跑。
$r27cells = N '\$Tests\\b217_string_select_date_label\\B217Res\.vbp' $T
if ($r27cells -ne 2) { $fail += ('R27 b217 run cells = ' + $r27cells + ' (want 2 = x64 + x86)') }
# R28 (§11.194 第二百一十一刀): Option Compare Text 与 Date 标签共存的夹具 —— 两台在跑，
# 且那条用的是文本入口 (vb6_StrCmpT) 而不是二进制入口。
$r28cells = N '\$Tests\\b218_octext_date_label\\B218Res\.vbp' $T
if ($r28cells -ne 2) { $fail += ('R28 b218 run cells = ' + $r28cells + ' (want 2 = x64 + x86)') }
$r28opt = N '(?m)^Option Compare Text\s*$' (Get-Content -Raw (Join-Path $Root 'tests\b218_octext_date_label\B218Main.bas'))
if ($r28opt -ne 1) { $fail += ('R28 b218 Option Compare Text = ' + $r28opt + ' (want 1: the whole point of the cell)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_date_conv: parse-sites=' + $r1parse + ' rtl-total=' + $r2sites +
    ' raise13=' + $r1raise + ' owners=2 projections=3 cgen-rows=3 cells=' + $r7cells +
    ' nd-cells=' + $r13cells + ' udt-cells=' + $r15cells + ' var-cells=' + $r17cells +
    ' call-cells=' + $r19cells + ' cmp-cells=' + $r21cells + ' vcmp-cells=' + $r23cells + ' sel-cells=' + $r24cd + ' b216-cells=' + $r25cells + ' s217-cells=' + $r27cells + ' oc218-cells=' + $r28cells)
exit 0
