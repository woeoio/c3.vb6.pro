Attribute VB_Name = "DllBlocks"
Option Explicit

' CoClass blocks so this fixture also carries the explicit-identity form of the four
' classes above (same shape as tests/b188_term; the boxing leg does not read them).

CoClass Good
    [Implementation("B187Good")]
    [ComCreatable(True)]
End CoClass

CoClass Driver
    [Implementation("B187Driver")]
    [ComCreatable(True)]
End CoClass

CoClass Hazard
    [Implementation("B187Hazard")]
    [ComCreatable(True)]
End CoClass

CoClass Ident
    [Implementation("B187Ident")]
    [ComCreatable(True)]
End CoClass

CoClass Arg
    [Implementation("B187Arg")]
    [ComCreatable(True)]
End CoClass

CoClass Ref
    [Implementation("B187Ref")]
    [ComCreatable(True)]
End CoClass

CoClass Name
    [Implementation("B187Name")]
    [ComCreatable(True)]
End CoClass

CoClass Nil
    [Implementation("B187Nil")]
    [ComCreatable(True)]
End CoClass

CoClass Decl
    [Implementation("B187Decl")]
    [ComCreatable(True)]
End CoClass
