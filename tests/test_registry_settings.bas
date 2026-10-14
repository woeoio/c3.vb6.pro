' Acct #305 para-B149 + para-B150 (2026-10-11 knife 40): registry round trip AND the
' late-bound two-dimensional element read, in one fixture. ASCII comments only (C3 reads
' sources as ANSI/GBK).
'
' Why a real round trip: vb6_SaveSetting writes under
'   HKCU\Software\VB and VBA Program Settings\<app>\<section>
' which is exactly where VB6 puts it (vb6_RegBuildKey), so the read side is measured
' against the registry itself instead of a fake. The app name is unique to this case
' and every key is deleted again at the end (and at the top, so a crashed run cannot
' poison the next one).
'
' What was wrong before this knife (readings .build/b1265_nd/r6.txt, both arches):
'   VarType(GetAllSettings(...)) = 12        (VB6: 8204 = vbArray + vbVariant)
'   IsNull(...)                  = False    (VB6 with no entries: Null)
'   UBound(v, 1) / UBound(v, 2)  = 0 / 0    (VB6: n-1 / 2)
'   v(0, 2)                      = ""        (VB6: the value)
'   five readings, rc=0, not one diagnostic.
' The shape pinned here is VB6's: first dimension 0..n-1 is the key index, the second
' dimension is 1..2 with column 1 = key name and column 2 = value. External witness: the
' real VB6 corpus file mdoceditor/E-Type/Main.bas:269 does
'   For i = 0 To UBound(varFiles, 1) : ... varFiles(i, 1)
'
' para-B150 rides along because it is the same slot: v(i, j) used to be emitted as
' vb6_VariantArrayGet(&v, i) -- the second index was dropped at the emission site --
' while UBound(v, 2) on that very slot already asked the runtime.
'
' VB6's DeleteSetting has three shapes (key / section / app). The two-argument shape
' (delete a whole section) is pinned by R6 below -- acct #305 para-B149 cell 3, knife 41.
' The one-argument shape (delete a whole app) is deliberately NOT padded: the RTL exit
' returns immediately when section is empty (vb6rtl_registry.c:97), so filling that slot
' with an empty string would turn "delete the app" into a silent no-op; it stays a loud
' C2198 instead. Three-argument calls below spell all three out.
Option Explicit

Private Const RS_APP As String = "C3ProbeB149"

Public Sub Main()
    Dim v As Variant
    Dim w As Variant

    DeleteSetting RS_APP, "Sec", "K1"
    DeleteSetting RS_APP, "Sec", "K2"

    ' R1: app/section not present => Null
    v = GetAllSettings(RS_APP, "Sec")
    Debug.Print "R1-miss-vt="; VarType(v)
    Debug.Print "R1-miss-isnull="; IsNull(v)

    ' R2: two keys => two-dimensional Variant array (0..1, 1..2)
    SaveSetting RS_APP, "Sec", "K1", "V1"
    SaveSetting RS_APP, "Sec", "K2", "V2"
    v = GetAllSettings(RS_APP, "Sec")
    Debug.Print "R2-vt="; VarType(v)
    Debug.Print "R2-isnull="; IsNull(v)
    Debug.Print "R2-lb1="; LBound(v, 1)
    Debug.Print "R2-ub1="; UBound(v, 1)
    Debug.Print "R2-lb2="; LBound(v, 2)
    Debug.Print "R2-ub2="; UBound(v, 2)
    Debug.Print "R2-cell01="; CStr(v(0, 1))
    Debug.Print "R2-cell02="; CStr(v(0, 2))
    Debug.Print "R2-cell12="; CStr(v(1, 2))

    ' R3: the returned array is a snapshot - writing it does not write the registry
    v(1, 2) = "W2"
    w = GetAllSettings(RS_APP, "Sec")
    Debug.Print "R3-slot-after-write="; CStr(v(1, 2))
    Debug.Print "R3-registry-unchanged="; CStr(w(1, 2))

    ' R5 (para-B150 second half): the ADDRESS form of that same element read. The codegen
    ' used to compare the emitted C text against the one-index name only, so a two-index
    ' element handed to VarPtr() did not match and fell to the compound-literal path
    ' ((intptr_t)&(void*){...} on a struct -> C2440). Now both spellings answer from one
    ' table and this emits (intptr_t)vb6_VariantArrayElemPtrND(&v, 2, 0, 1). Zero would mean
    ' the element address was not found, so the cell is read twice: by address, by value.
    Debug.Print "R5-elemptr-nonzero="; CStr(VarPtr(v(0, 1)) <> 0)
    Debug.Print "R5-cell01-still="; CStr(v(0, 1))

    ' R4: after deleting both keys the missing case is back
    DeleteSetting RS_APP, "Sec", "K1"
    DeleteSetting RS_APP, "Sec", "K2"
    v = GetAllSettings(RS_APP, "Sec")
    Debug.Print "R4-after-delete-vt="; VarType(v)
    Debug.Print "R4-after-delete-isnull="; IsNull(v)

    ' R6 (para-B149 cell 3): the two-argument shape = delete the whole section. Before this
    ' knife the emitted call carried two arguments against a three-parameter RTL exit, so a
    ' program using VB6's own spelling did not compile at all (error C2198: too few
    ' arguments) -- that is the negative control for this row (BASE compiler, same fixture).
    SaveSetting RS_APP, "Sec6", "K1", "V1"
    v = GetAllSettings(RS_APP, "Sec6")
    Debug.Print "R6-before-vt="; VarType(v)
    DeleteSetting RS_APP, "Sec6"
    v = GetAllSettings(RS_APP, "Sec6")
    Debug.Print "R6-two-arg-delete-vt="; VarType(v)
    Debug.Print "R6-two-arg-delete-isnull="; IsNull(v)

    Debug.Print "RS-DONE"
End Sub
