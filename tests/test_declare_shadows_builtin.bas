Option Explicit

' ============================================================
'  test_declare_shadows_builtin.bas - 交接 §7.8 缺陷 B 的最小回归夹具
'
'  `GetObject` 在 VB6 里**既是内置函数** (GetObject(, "Excel.Application"))
'  **也是 Win32 API** (gdi32!GetObjectA). 判定依据三件:
'    (1) 存在性 — 本模块确有同名的 Declare;
'    (2) 参数形状 — 声明的形参表是 SDK 边界 (3 参), 与内置的 2 参不同形;
'    (3) Lib/Alias — 声明由 `Declare ... Lib "gdi32.dll" Alias "GetObjectA"` 建出.
'  三件齐 ⇒ 调用点走 Declare; 没有 Declare 的同名站点仍走内置 (见
'  test_p24_04b/test_getobject.bas 那条对照针).
'
'  修前发码 (落进内置 vb6_GetObject(const wchar_t*, const wchar_t*) ⇒ C2197):
'    vb6_GetObject(vb6_VariantToLong((*SrcImg)), sizeof(tBmp), &tBmp)
'  修后发码 (走声明的别名转发桩):
'    vb6_lwa_GetObjectA(h, sizeof(bmp), &bmp)
' ============================================================

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" _
    (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Sub Main()
    Dim bmp As BITMAP
    Dim h As Long

    h = GetObject(0&, Len(bmp), bmp)
    Debug.Print CStr(h)
End Sub
