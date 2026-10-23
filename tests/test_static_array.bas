Attribute VB_Name = "StaticArrayStorage"
Option Explicit

' Acct #305 (甲) knife 160: a Static LOCAL ARRAY never got C's static storage. The scalar path asked
' (var.isStatic || inStaticProc_) and emitted "static <T> x = <const>;", but all three array branches
' of the declaration emitter ignored storageClass entirely, so every call minted a fresh descriptor:
'   - Static f(0 To 3) As Long      -> second call read back 0 (the VB6 init-once idiom silently
'                                     defeats itself: the static guard flag says "already filled"
'                                     while the array that was filled is gone)
'   - Static d() As Long + ReDim    -> second call read a NULL descriptor = Error 9
' Both readings were identical on x64 and x86 with the pre-fix compiler.
' The fix asks ONE question in ONE place (wantsStaticLocalStorage) and emits ONE shape (emitArrayCarrierDecl):
' static storage, and a first-call guard keyed on the pointer itself where the initializer is a runtime
' call (C forbids a non-constant initializer for a static local - the Fix 084aa rule). The dynamic form's
' initializer IS a constant NULL, so it gets the prefix and no guard: "not ReDim'd yet = Error 9" stays
' VB6's own answer.
'
'   fix / dyn / nd / copy / ptr = the five shapes that must now survive the round trip.
'   fresh / freshnd = the same two shapes declared with plain Dim in the same procedure - those must
'                     KEEP resetting (they are the red power of this fixture: widen the gate and both
'                     read 42/9, and the exit-destroy then frees a carrier the next call still reads).
'   erase = a Static fixed array Erase'd on the first call: the elements reset but the carrier is the
'           one this frame created once - UBound must still answer, not Error 9.
'   ptr   = the gate-visible use (VBFlexGridDemo UserEditingForm:412 hands VarPtr(CustomColors(0)) to
'           the OS CHOOSECOLOR struct): the buffer the OS writes into has to be THE SAME buffer on the
'           next call, or the custom-colour panel resets behind the user's back.

Private gPtr As LongPtr

' 1) fixed 1D + the init-once idiom, and a plain Dim arm in the same procedure
Public Sub FixArm(ByVal which As Long)
    Dim f(0 To 3) As Long
    Static sf(0 To 3) As Long
    If which = 1 Then
        sf(2) = 42
        f(2) = 42
    Else
        Debug.Print "K160-fix="; sf(2); " fresh="; f(2)
    End If
End Sub

' 2) dynamic + multi-dim, each with its plain Dim brother
Public Sub DynArm(ByVal which As Long)
    Dim nd(1 To 2, 1 To 2) As Long
    Static d() As Long
    Static sn(1 To 2, 1 To 2) As Long
    If which = 1 Then
        ReDim d(0 To 2)
        d(0) = 7
        sn(2, 1) = 9
        nd(2, 1) = 9
    Else
        Debug.Print "K160-dyn="; d(0); " ub="; UBound(d); " nd="; sn(2, 1); " freshnd="; nd(2, 1)
    End If
End Sub

' 3) whole-array assignment INTO a static slot (knife 158/159's copy, with a Static left value):
'    the copy itself must land and the clone must survive to the next call.
Public Sub CopyArm(ByVal which As Long)
    Dim src(0 To 3) As Long
    Static cs(0 To 3) As Long
    src(1) = 123456
    If which = 1 Then
        cs = src
    Else
        Debug.Print "K160-copy="; cs(1)
    End If
End Sub

' 4) Erase on a static fixed array: elements reset, carrier stays (no Error 9 next call)
Public Sub EraseArm(ByVal which As Long)
    Static es(0 To 3) As Long
    If which = 1 Then
        es(0) = 5
        Erase es
    Else
        Debug.Print "K160-erase="; es(0); " ub="; UBound(es)
    End If
End Sub

' 5) the inStaticProc_ arm cannot be pinned through the front end: this parser always hands
'     parseSubDecl / parseFunctionDecl an isStatic of false (parser_decl.cpp:35/40/95/100), so a
'     "Static Sub" header is not accepted at all and inStaticProc_ is never true. That is a separate
'     account (the real corpus has 10 such headers); here it only documents which arm is unreachable.

' 6) the buffer identity the OS struct keeps a pointer to
Public Sub PtrArm(ByVal which As Long)
    Static pc(0 To 15) As Long
    Dim now As LongPtr
    If which = 1 Then
        gPtr = VarPtr(pc(0))
    Else
        now = VarPtr(pc(0))
        If now = gPtr Then
            Debug.Print "K160-ptr=Y"
        Else
            Debug.Print "K160-ptr=N"
        End If
    End If
End Sub

Public Sub Main()
    FixArm 1
    FixArm 2
    DynArm 1
    DynArm 2
    CopyArm 1
    CopyArm 2
    EraseArm 1
    EraseArm 2
    PtrArm 1
    PtrArm 2
    Debug.Print "K160-DONE=1"
End Sub
