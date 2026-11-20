Attribute VB_Name = "LvMod"
Option Explicit

' Enum members defined in a standard module and referenced from a .frm go through
' the driver cross-module pre-registration, which runs *before* Pass 1: module-level
' Consts are not in the symbol table yet. If the evaluator does not scan the
' declarations first, members silently fall back to their ordinal (0,1,2) and the
' referencing side gets the wrong bit.
Public Const LSX2_FULL As Long = &H20
Public Const LSX2_TRACK As Long = &H8

Public Enum LvFlag2
    Flag2Full = LSX2_FULL
    Flag2Track = LSX2_TRACK
    Flag2Both = Flag2Full Or Flag2Track
End Enum
