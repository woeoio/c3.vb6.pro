' Fix 161e: MsgBox 的首参是**数值**时, VB6 会隐式转成字符串再显示
' (MsgBox 7 弹窗里就是 "7"); C3 此前只包装了 COM 读 / Variant 两类形态,
' 纯数值 (Len(...) 返回值 / 数值变量 / 字面量 / 算术表达式 / 浮点 / 布尔)
' 一个都没包 ⇒ int32_t(double/bool) 直接当 BSTR 指针传给 vb6_MsgBox1
' ⇒ 弹窗是空的 (地址 7、123 之类根本不可读; 真读得到时显示的也是垃圾)。
' 修复: 对上面没认领的形态统一过 wrapToBSTR (cgen_expr_binary_util.cpp:81),
' 它按 inferExprType 分流, 且 String 分支原样返回 (不会把字符串实参再包一层)。
'
' ⚠ 本文件**只供 --emit-c 源码面断言 (Test-EmitcShape / Test-EmitcAbsent) 使用,
'   不登记为运行用例** —— MsgBox 是模态 MessageBoxW, 无头环境会阻塞线程
'   (run_tests.ps1:1190 那条注释: "运行中弹 MsgBox 的用例需获得前台焦点")。
Public Sub Main()
    Dim n As Long
    Dim s As String
    Dim d As Double
    Dim b As Boolean
    n = 7
    s = "hello"
    d = 3.5
    b = True

    MsgBox Len("你好hello")   ' ① 数值函数返回值 (int32_t) → vb6_CStrLong
    MsgBox n                  ' ② 数值变量 (Long)          → vb6_CStrLong
    MsgBox 123                ' ③ 数值字面量               → vb6_CStrLong
    MsgBox n + 1              ' ④ 算术表达式               → vb6_CStrLong
    MsgBox d                  ' ⑤ 浮点 (Double)            → vb6_CStrDbl
    MsgBox b                  ' ⑥ 布尔                     → vb6_CStrBool
    MsgBox s                  ' ⑦ 字符串 (对照)            → 原样, 不得被误包

    ' ---- 账 #490 §B178: MsgBox 的 VB6 全签名是五参 (prompt, buttons, title, helpfile, context),
    '      RTL 原型此前只有三参 ⇒ 四/五参形 (语料 10 处真调用) 是**静默多递** (本机 cl 不报),
    '      且第 4/5 位没有形参类型 ⇒ Variant 直接进 BSTR 槽。原型扩到五参之后, 缺位要补齐。
    Dim sty As Long
    Dim ttl As String
    Dim hlp As String
    Dim ctx As Long
    Dim v As Variant
    sty = 0
    ttl = "t"
    hlp = "h"
    ctx = 100
    v = "hv"

    MsgBox s, sty                   ' (8) 两参  -> 补 NULL, NULL, 0
    MsgBox s, sty, ttl              ' (9) 三参  -> 补 NULL, 0
    MsgBox s, sty, ttl, hlp         ' (10) 四参 -> 补 0
    MsgBox s, sty, ttl, hlp, ctx    ' (11) 五参 -> 原样 (原型收满)
    MsgBox s, sty, ttl, v, ctx      ' (12) helpfile 位是 Variant -> vb6_VariantToString
End Sub
