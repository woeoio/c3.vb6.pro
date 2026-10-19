VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TUHolder"
Option Explicit

' ai/032 §6.9 收口: 四档类型 (ULong/UInteger/ULongLong/SByte) 落在**类模块字段**位。
' 类字段不在 knownNarrowIntVars_ 那张局部表里 —— 本夹具钉住「CStr / 装箱 / 混算 / 整除」
' 四条路在字段位上与局部同形 (根因是 §6.6 的截断装箱, 装箱收口后自然消解)。
Public mU As ULong
Public mUI As UInteger
Public mU3 As ULongLong
Public mSB As SByte

Public Sub Show()
    Debug.Print "UF-u=" & CStr(mU)          ' 4294967295
    Debug.Print "UF-ui=" & CStr(mUI)        ' 65535
    Debug.Print "UF-u3=" & CStr(mU3)        ' 78187493530
    Debug.Print "UF-sb=" & CStr(mSB)        ' -128
    Debug.Print "UF-u-add=" & CStr(mU + 1)  ' 4294967296 (混算提升到 64)
    Debug.Print "UF-u-idiv=" & CStr(mU \ 2) ' 2147483647 (无符号整除档)
    Debug.Print "UF-u3-idiv=" & CStr(mU3 \ 2)' 39093746765

    Dim v As Variant
    v = mU
    Debug.Print "UF-u-box=" & CStr(v)           ' 4294967295
    Debug.Print "UF-u-box-tn=" & TypeName(v)    ' ULong
    v = mU3
    Debug.Print "UF-u3-box=" & CStr(v)          ' 78187493530
    v = mUI
    Debug.Print "UF-ui-box=" & CStr(v)          ' 65535
    ' ai/032 §6.9 收口: 这两格钉的是**类型身份**, 不只是值。
    ' 以前 SByte/UInteger 没有自己的出口, boxToVariant 落 default -> vb6_VariantFromValue,
    ' 由 RTL 的 _Generic 按 C 类型猜: int8_t 命中 `char:` 档、uint16_t 命中 `wchar_t:` 档,
    ' 两档都被装成 **VT_I4** ⇒ 值对(65535 / -128), 而 TypeName 答 "Long"、VarType 答 3。
    ' ULong/VT_UI4 早就有真出口(它的注释写着「答 "Long" 才是误导」), 这两档按同一口径补齐。
    Debug.Print "UF-ui-box-tn=" & TypeName(v)   ' UInteger (改前: Long)
    v = mSB
    Debug.Print "UF-sb-box=" & CStr(v)           ' -128
    Debug.Print "UF-sb-box-tn=" & TypeName(v)    ' SByte (改前: Long)
End Sub
