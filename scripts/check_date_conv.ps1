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
# R29 (§11.195 第二百一十二刀): Select Case 的**主语**判据必须问三处权威 ——
# 裸名数组表 (arrayElemTypes_) 与 UDT 成员 (inferUdtFieldVb6Type) 各留读点，
# 而且 AllString 那条强制改走串承载的闸要带 !isDateSelect。
$r29elem = N 'arrayElemTypes_' $s26
$r29udt = N 'inferUdtFieldVb6Type' $s26
$r29gate = N '!isDateSelect && !dateBuiltinSubj196' $s26
if ($r29elem -lt 4) { $fail += ('R29 cgen_select arrayElemTypes_ reads = ' + $r29elem + ' (want >= 4: label side + subject side)') }
if ($r29udt -ne 2) { $fail += ('R29 cgen_select inferUdtFieldVb6Type reads = ' + $r29udt + ' (want 2: UDT field + UDT field array)') }
if ($r29gate -ne 1) { $fail += ('R29 AllString string-carrier gate = ' + $r29gate + ' (want 1: must not hijack a Date subject)') }
# R30: b219 夹具两台在跑。
$r30cells = N '\$Tests\\b219_date_subject_elem\\B219Res\.vbp' $T
if ($r30cells -ne 2) { $fail += ('R30 b219 run cells = ' + $r30cells + ' (want 2 = x64 + x86)') }
# R31 (§11.196 第二百一十三刀): 日期内建当主语 —— 那条 Variant 档的判据只许一处
# (名单表 kVariantDate196 恰 2 处读: 声明 + 循环), 降档与两条闸各必须显式引用它。
$r31sa = N '\$Tests\\b220_date_builtin_subject\\B220Res\.vbp' $T
if ($r31sa -ne 2) { $fail += ('R31 b220 run cells = ' + $r31sa + ' (want 2 = x64 + x86)') }
$r31name = N 'dateBuiltinSubj196' $s26
if ($r31name -ne 6) { $fail += ('R31 dateBuiltinSubj196 reads = ' + $r31name + ' (want 6 = decl + set + downgrade + AllString gate + carrier + isVariantSelect)') }
$r31list = N 'kVariantDate196' $s26
if ($r31list -ne 2) { $fail += ('R31 kVariantDate196 = ' + $r31list + ' (want 2 = the one name table + its loop)') }
$r31gate = N '!dateBuiltinSubj196' $s26
if ($r31gate -ne 1) { $fail += ('R31 AllString gate must exclude date builtins = ' + $r31gate + ' (want 1)') }
# R32 (§11.197 第二百一十四刀): CStr(零参日期内建) —— 两处必须同时在位：
#   ① 推断权威 (cgen_util_type.cpp) 对 now/date/time 答 Date 的那一问；
#   ② CStr 站点 (conv_cstr) 把 Date 档交给 vb6_CStrDate 那一臂。
# 缺①或②就回到 `vb6_CStr(vb6_VariantFromValue(vb6_Now()))` ⇒ 打裸序列号。
$r32cells = N '\$Tests\\b221_cstr_now_date\\B221Res\.vbp' $T
if ($r32cells -ne 2) { $fail += ('R32 b221 run cells = ' + $r32cells + ' (want 2 = x64 + x86)') }
$s221conv = Get-Content -Raw (Join-Path $Root 'src\backend\detail\expr\cgen_expr_call_conv_cstr.inc')
$s221type = Get-Content -Raw (Join-Path $Root 'src\backend\cgen_util_type.cpp')
$r32arm = N 'itNI2 == Vb6Type::Date\),?\s*callee = "vb6_CStrDate"' $s221conv
if ($r32arm -ne 1) { $fail += ('R32 CStr site Date arm = ' + $r32arm + ' (want 1: the only place that turns the Date answer into vb6_CStrDate)') }
$r32ask = N 'lb197 == "now"' $s221type
if ($r32ask -ne 1) { $fail += ('R32 inferExprType date-builtin ask = ' + $r32ask + ' (want 1: one authority for "is this Now/Date/Time a Date?")') }
# R33 (§11.198 第二百一十五刀): general date 的「日期部分为零只给时间」那一问 ——
# 判据必须走**序列号**（SYSTEMTIME 分不出 1899-12-30 是真日期还是零日），且只住一处。
$r33cells = N '\$Tests\\b222_time_general_date\\B222Res\.vbp' $T
if ($r33cells -ne 2) { $fail += ('R33 b222 run cells = ' + $r33cells + ' (want 2 = x64 + x86)') }
$s222fmt = Get-Content -Raw (Join-Path $Root 'src\rtl\core\vb6rtl\vb6rtl_format.c')
$r33rule = N 'if \(vb6_fmtDatePartIsZero\(serial\)\)' $s222fmt
if ($r33rule -ne 1) { $fail += ('R33 general-date zero-date rule = ' + $r33rule + ' (want 1: date-part-zero => time only, asked through the shared predicate)') }
# R34 (§11.199 第二百一十六刀): 「日期部分为零」的窗口是 (−1,1)（向零截断为 0），
# 不是 [0,1) —— 真机 P271: CStr(CDate(-0.5)) = "12:00:00"、CDate(-1.5) 仍带日期。
# Print # 那一站有自己的一份同规则窗口 (vb6rtl_file.c 的 vb6_FileDateText)，两处都要在。
$r34cells = N '\$Tests\\b223_negative_serial_text\\B223Res\.vbp' $T
if ($r34cells -ne 2) { $fail += ('R34 b223 run cells = ' + $r34cells + ' (want 2 = x64 + x86)') }
$s222file = Get-Content -Raw (Join-Path $Root 'src\rtl\core\vb6rtl\vb6rtl_file.c')
$r34file = N 'if \(vb6_fmtDatePartIsZero\(v\)\)' $s222file
if ($r34file -ne 1) { $fail += ('R34 Print # zero-date window = ' + $r34file + ' (want 1: the file station asks the same predicate)') }
# R35 (§11.200 第二百一十七刀): Write # 的 #...# 文本 —— 时间部分为零就整段不写。
# 真机 P273 p273w.txt: CDate(0) → "#1899-12-30#"、CDate(1) → "#1899-12-31#"、
# CDate(0.0000001) → "#1899-12-30#"; 0.5 仍是 "#1899-12-30 12:00:00#"。
$r35cells = N '\$Tests\\b224_write_date_time_zero\\B224Res\.vbp' $T
if ($r35cells -ne 2) { $fail += ('R35 b224 run cells = ' + $r35cells + ' (want 2 = x64 + x86)') }
$r35arm = N 'L"#%04d-%02d-%02d#"' $s222file
if ($r35arm -ne 1) { $fail += ('R35 Write # date-only arm = ' + $r35arm + ' (want 1: zero time omits the whole time segment)') }
# R36 (§11.201 第二百一十八刀): 零日判据跟着四入走 —— 四入到 24:00 就进位一天,
# 零日不再是零日。真机 P275: CStr(CDate(±0.9999999)) = "1899/12/31"、
# Print #f, CDate(0.9999999) = "1899/12/31 "（改前 "23:59:59 "）。
$r36cells = N '\$Tests\\b225_boundary_carry\\B225Res\.vbp' $T
if ($r36cells -ne 2) { $fail += ('R36 b225 run cells = ' + $r36cells + ' (want 2 = x64 + x86)') }
$r36def = N 'int vb6_fmtDatePartIsZero\(double serial\)' $s222fmt
if ($r36def -ne 1) { $fail += ('R36 carry-aware predicate = ' + $r36def + ' (want 1: one definition in format.c)') }
$r36decl = N 'int vb6_fmtDatePartIsZero\(double serial\);' (Get-Content -Raw (Join-Path $Root 'src\rtl\core\vb6rtl\vb6rtl_builtin.h'))
if ($r36decl -ne 1) { $fail += ('R36 predicate declared = ' + $r36decl + ' (want 1: file.c asks it too)') }
$r36carry = N '>= 86400\) ip \+= 1\.0;' $s222fmt
if ($r36carry -ne 1) { $fail += ('R36 carry step = ' + $r36carry + ' (want 1: rounding to 24:00 advances one day)') }
# R37 (§11.202 第二百一十九刀): DateAdd 的两半必须问同一句 —— 裸 (int32_t)date 配
# vb6_SecondsOfDay 会让进位落到时间那半而日期那半不进，午夜附近的运算差一整天。
# 真机 P278: DateAdd("s",1,CDate(44259.9999999)) = 2021/3/5 0:00:01（改前 3/4 0:00:01）。
$r37cells = N '\$Tests\\b226_dateadd_carry\\B226Res\.vbp' $T
if ($r37cells -ne 2) { $fail += ('R37 b226 run cells = ' + $r37cells + ' (want 2 = x64 + x86)') }
$s202date = Get-Content -Raw (Join-Path $Root 'src\rtl\core\vb6rtl\vb6rtl_date.c')
$r37split = N 'vb6_split_serial\(date, &datePart, &sod\);' $s202date
if ($r37split -ne 1) { $fail += ('R37 DateAdd asks splitter = ' + $r37split + ' (want 1)') }
$r37raw = N 'int32_t datePart = \(int32_t\)date;' $s202date
if ($r37raw -ne 0) { $fail += ('R37 raw day half = ' + $r37raw + ' (want 0: the raw truncation must not come back)') }
# R38 (§11.203 第二百二十刀): DateDiff / DatePart 的「哪一天」也要走同一句 splitter。
# 真机 P280: DateDiff("d", CDate(44259.9999999), CDate(44260)) = 0、
# DatePart("d", CDate(44259.9999999)) = 5（裸截断分别给 1 与 4）。
$r38cells = N '\$Tests\\b227_part_diff_carry\\B227Res\.vbp' $T
if ($r38cells -ne 2) { $fail += ('R38 b227 run cells = ' + $r38cells + ' (want 2 = x64 + x86)') }
$r38diff = N 'int32_t d1 = vb6_serial_day_half\(date1\), d2 = vb6_serial_day_half\(date2\);' $s202date
if ($r38diff -ne 1) { $fail += ('R38 DateDiff day half = ' + $r38diff + ' (want 1)') }
$r38part = N 'int32_t d = vb6_serial_day_half\(date\);' $s202date
if ($r38part -ne 2) { $fail += ('R38 DatePart + Weekday day half = ' + $r38part + ' (want 2: P282 Weekday(CDate(44259.9999999)) = 6)') }
$r38raw1 = N 'int32_t d1 = \(int32_t\)date1, d2 = \(int32_t\)date2;' $s202date
$r38raw2 = N 'int32_t d = \(int32_t\)date;' $s202date
if (($r38raw1 + $r38raw2) -ne 0) { $fail += ('R38 raw day halves = ' + ($r38raw1 + $r38raw2) + ' (want 0)') }
# R39 (§11.204 第二百二十一刀): Int/Fix 在 Date 实参上保持 Date 身份（VarType=7、
# CStr 走日期文本、IsDate=True）；数值实参仍是 Double。真机 P283/P284/P285。
$r39cells = N '\$Tests\\b228_int_fix_date\\B228Res\.vbp' $T
if ($r39cells -ne 2) { $fail += ('R39 b228 run cells = ' + $r39cells + ' (want 2 = x64 + x86)') }
$s204type = Get-Content -Raw (Join-Path $Root 'src\backend\cgen_util_type.cpp')
$r39rule = N 'if \(\(idn == "int" \|\| idn == "fix" \|\| idn == "abs" \|\| idn == "round"\)' $s204type
if ($r39rule -ne 1) { $fail += ('R39 Int/Fix (+Abs/Round) date identity = ' + $r39rule + ' (want 1)') }
# R40 (§11.205 第二百二十二刀): Abs / Round 也在 Date 实参上保持 Date；
# CInt/CLng/CCur/CSng/CDbl 不在此列。真机 P288/P289。
$r40cells = N '\$Tests\\b229_abs_round_date\\B229Res\.vbp' $T
if ($r40cells -ne 2) { $fail += ('R40 b229 run cells = ' + $r40cells + ' (want 2 = x64 + x86)') }
$r40set = N 'idn == "abs" \|\| idn == "round"' $s204type
if ($r40set -ne 1) { $fail += ('R40 Abs/Round in the date set = ' + $r40set + ' (want 1)') }

if ($fail.Count -gt 0) {
    foreach ($m in $fail) { Write-Host ('FAIL ' + $m) }
    exit 1
}
Write-Host ('PASS check_date_conv: parse-sites=' + $r1parse + ' rtl-total=' + $r2sites +
    ' raise13=' + $r1raise + ' owners=2 projections=3 cgen-rows=3 cells=' + $r7cells +
    ' nd-cells=' + $r13cells + ' udt-cells=' + $r15cells + ' var-cells=' + $r17cells +
    ' call-cells=' + $r19cells + ' cmp-cells=' + $r21cells + ' vcmp-cells=' + $r23cells + ' sel-cells=' + $r24cd + ' b216-cells=' + $r25cells + ' s217-cells=' + $r27cells + ' oc218-cells=' + $r28cells + ' subj219-cells=' + $r30cells + ' bld220-cells=' + $r31sa + ' now221-cells=' + $r32cells + ' gen222-cells=' + $r33cells + ' neg223-cells=' + $r34cells + ' wr224-cells=' + $r35cells + ' carry225-cells=' + $r36cells + ' dadd226-cells=' + $r37cells + ' pdiff227-cells=' + $r38cells + ' intfix228-cells=' + $r39cells + ' absr229-cells=' + $r40cells)
exit 0
