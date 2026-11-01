Attribute VB_Name = "DllBlocks"
Option Explicit

' Knife 231 measure: puts ByRefCls on the IDispatch surface (probe arm BYREFARG=Bump:10).

CoClass ByRefProbe
    [Implementation("ByRefCls")]
    [ComCreatable(True)]
End CoClass
