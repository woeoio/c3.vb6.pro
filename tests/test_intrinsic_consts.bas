' NOTE: deliberately NO Option Explicit here - that is the shape the defect needs.
' With Option Explicit the old compiler still read the right value, because the bare C
' name expanded the RTL #define. Without it the name became an implicit Variant local
' (Empty) and every read of it answered 0.

' 218 knife 3: every intrinsic VB constant must have exactly ONE answer source.
'
' Before this knife the answers lived in three places that never talked to each other:
'   * src/rtl/core/vb6rtl/vb6rtl_userctl.h  -> 14 names as C #defines (Picture/Async/Hit families)
'   * src/backend/detail/expr/cgen_expr_ident_dispatch.inc -> 26 names folded per name at codegen
'   * src/semantics/builtin/builtin_consts(.inc|_ext.inc)  -> the symbol table the parser knows
' The semantic layer knew none of the first two grids, so:
'   * a .bas reading vbPicTypeBitmap was treated as an undeclared identifier: an implicit Variant
'     local initialised to Empty, so the program ran and printed 0 instead of 1 (measured).
'   * vbUseSystem answered -1 through Const folding (table) and 0 through a direct read (fold)
'     inside the SAME program. The VB6 manual in this repo says 0
'     (docs/vb6-manual/09-.../Date ... md), so -1 was the wrong one and is fixed here.

Const C_VS As Long = vbUseSystem          ' folded by the symbol table
Const C_PIC As Long = vbPicTypeBitmap     ' used to be RTL-header only

Public Sub Main()
    Dim n As Long
    Dim p1 As Long, p2 As Long, p3 As Long

    ' The shape that was broken: the name on the RIGHT of an assignment. The old compiler
    ' registered it as an implicit Variant local (Empty), so it read 0 at run time - while the
    ' same name inside a concatenation still expanded the RTL #define and read the right value.
    p1 = vbPicTypeBitmap
    p2 = vbHitResultHit
    p3 = vbAsyncTypeByteArray
    Debug.Print "IC-ASSIGN=" & p1 & "," & p2 & "," & p3

    ' --- the 14 names the RTL header used to own (all zero on the old compiler) ---
    Debug.Print "IC-RTL-PT=" & vbPicTypeNone & "," & vbPicTypeBitmap & "," & vbPicTypeMetafile & "," & vbPicTypeIcon & "," & vbPicTypeEMetafile
    Debug.Print "IC-RTL-HR=" & vbHitResultOutside & "," & vbHitResultTransparent & "," & vbHitResultHit
    Debug.Print "IC-RTL-AT=" & vbAsyncTypePicture & "," & vbAsyncTypeFile & "," & vbAsyncTypeByteArray
    Debug.Print "IC-RTL-AR=" & vbAsyncReadSynchronous & "," & vbAsyncReadAsynchronous & "," & vbAsyncReadForceUpdate
    Debug.Print "IC-RTL-CONST=" & C_PIC

    ' --- the 26 names the codegen folds used to own (these were already right) ---
    Debug.Print "IC-FOLD-CALL=" & vbCallObject & "," & vbCallGetNext & "," & vbCallLet & "," & vbCallMethod & "," & vbCallSet
    Debug.Print "IC-FOLD-ROP=" & vbSrcCopy & "," & vbSrcAnd & "," & vbSrcPaint & "," & vbSrcInvert & "," & vbBlackness & "," & vbWhiteness
    Debug.Print "IC-FOLD-LOG=" & vbLogEventTypeConstant & "," & vbLogEventTypeInformation & "," & vbLogEventTypeWarning & "," & vbLogEventTypeError
    Debug.Print "IC-FOLD-DFM=" & vbGeneralDate & "," & vbLongDate & "," & vbShortDate & "," & vbLongTime & "," & vbShortTime
    Debug.Print "IC-FOLD-WK=" & vbFirstDayOfWeek & "," & vbFirstWeekOfYear & "," & vbUseSystem
    Debug.Print "IC-FOLD-MISC=" & vbSimple & "," & vbUseCompare & "," & vbUseDefault

    ' --- the -1 vs 0 conflict: three readings of one name, all three must agree ---
    n = vbUseSystem
    Debug.Print "IC-VS-CONST=" & C_VS
    Debug.Print "IC-VS-VAR=" & n
    Debug.Print "IC-VS-DIRECT=" & vbUseSystem

    Debug.Print "IC PASS"
End Sub
