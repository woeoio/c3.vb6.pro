' <vbeclipse>: Variant **装箱表本身**的哨兵夹具。
' 为什么单独有这一条: 装箱档位以前有两份并行表 (wrapVariantValue 的 switch 与
' boxToVariant), 账 #123 只在其中一份修了 Byte、Fix 198 只在另一份修了 Boolean,
' 于是"某一档对不对"取决于表达式走哪条路 —— DT43 (dt1.CheckBox 的 VarType) 就是在
' 一个和装箱毫无关系的控件应用上偶然暴露的。这里把三条真实路线各钉一遍:
'   (1) 内置函数实参装箱   VarType(m_b)        —— DT43 断的那条
'   (2) Variant 赋值装箱   v = m_b             —— wrapVariantValue 那条
'   (3) 过程形参装箱       ReportV ..., m_b    —— packLetValueArg 那条
' 以及模块级/局部/函数返回三种来源 (账 #123 的原始形态就是"模块级 Byte 错、局部对")。
' Date/Single 两档本批已修进权威表也进了 needle (修前两者都读成 5); Currency 仍只打印
' 不钉 —— 它与 VB6 的 VT_CY=6 的差距没有真 VB6 读数可依据, 已在汇报里记成待办。
Option Explicit

Private m_b As Boolean
Private m_by As Byte
Private m_i As Integer
Private m_l As Long
Private m_s As String
Private m_sin As Single
Private m_dbl As Double
Private m_d As Date

Private Sub ReportV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag; VarType(v); "/"; TypeName(v)
End Sub

Private Function RetB() As Boolean
    RetB = True
End Function

Public Sub Main()
    Dim lb As Boolean
    Dim lby As Byte
    lb = True
    lby = 67
    m_b = True
    m_by = 67
    m_i = 7
    m_l = 9
    m_s = "x"
    m_sin = 1.5
    m_dbl = 2.5
    ' Date 用序列号给值: CDate("2026-01-01") 在本仓现在是 0 (日期串解析是另一格问题,
    ' 已单独记), 那会让本夹具的 clng/cdbl 两行读成 0, 分不清是装箱还是解析。
    m_d = 46023

    Debug.Print "VB-mod-bool="; VarType(m_b); "/"; TypeName(m_b)
    Debug.Print "VB-mod-byte="; VarType(m_by); "/"; TypeName(m_by)
    Debug.Print "VB-loc-bool="; VarType(lb); "/"; TypeName(lb)
    Debug.Print "VB-loc-byte="; VarType(lby); "/"; TypeName(lby)
    Debug.Print "VB-func-bool="; VarType(RetB()); "/"; TypeName(RetB())
    Debug.Print "VB-expr-bool="; VarType(lb And True)
    Debug.Print "VB-mod-long="; VarType(m_l)
    Debug.Print "VB-mod-int="; VarType(m_i)
    Debug.Print "VB-mod-str="; VarType(m_s)
    Debug.Print "VB-mod-sin="; VarType(m_sin)
    Debug.Print "VB-mod-dbl="; VarType(m_dbl)
    Debug.Print "VB-mod-date="; VarType(m_d)

    Dim v As Variant
    v = m_b
    Debug.Print "VB-asg-bool="; VarType(v); "/"; TypeName(v)
    v = m_by
    Debug.Print "VB-asg-byte="; VarType(v); "/"; TypeName(v)
    v = lb
    Debug.Print "VB-asg-loc-bool="; VarType(v)
    v = m_l
    Debug.Print "VB-asg-long="; VarType(v)
    v = m_d
    Debug.Print "VB-asg-date="; VarType(v)

    ReportV "VB-call-bool=", m_b
    ReportV "VB-call-byte=", m_by
    ReportV "VB-call-long=", m_l
    ReportV "VB-call-date=", m_d

    ' 新档位的**反向**也要钉: 装成 VT_DATE / VT_R4 之后, 数值提取端与 IsDate 还得答对
    ' (RTL 里 vb6_vtDate 早有 TypeName/Format 档, 缺的是 VariantToLong/ToDouble 那一侧)。
    Dim vd As Variant
    vd = m_d
    Dim vs As Variant
    vs = m_sin
    Debug.Print "VB-date-clng="; CLng(vd)
    Debug.Print "VB-date-cdbl="; CDbl(vd)
    Debug.Print "VB-date-isdate="; IsDate(vd)
    Debug.Print "VB-date-cdate="; (CDate(vd) = m_d)
    Debug.Print "VB-sin-cdbl="; CDbl(vs)
    Debug.Print "VB-sin-clng="; CLng(vs)
    Debug.Print "VB-DONE"
End Sub
