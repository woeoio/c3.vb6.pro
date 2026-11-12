VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbUdtField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 夹具: **UDT 字段声明为 Variant** 的读腿 (照 PhotoDemon pdDictionary.cls:
' dItem As Variant, m_DictEntries() As dictEntry)。读进 String 变量 / 函数返回值
' 时, 发码必须把整条读表达式包进 vb6_VariantToString —— 否则整只 vb6_VARIANT
' 结构体递给 BSTR 形参 = C2440。

Private Type VbEntry
    dKey As String
    dItem As Variant
End Type

Private m_Entries(0 To 3) As VbEntry

' R1: 局部 String 变量 ← Variant 字段
Public Function ReadToLocalString() As String
    Dim s As String
    s = m_Entries(0).dItem
    ReadToLocalString = s
End Function

' R2: 函数返回值 (BSTR) ← Variant 字段
Public Function ReadToReturnString() As String
    ReadToReturnString = m_Entries(0).dItem
End Function

' R3: 局部 Double 变量 ← Variant 字段
Public Function ReadToLocalDouble() As Double
    Dim d As Double
    d = m_Entries(0).dItem
    ReadToLocalDouble = d
End Function

' R4: 局部 Long 变量 ← Variant 字段
Public Function ReadToLocalLong() As Long
    Dim l As Long
    l = m_Entries(0).dItem
    ReadToLocalLong = l
End Function

' 反钉: 写腿 (具体型 → Variant 槽) 本来就对, 不许被这一刀带歪
Public Sub WriteBack(ByVal v As Variant)
    m_Entries(0).dItem = v
End Sub
