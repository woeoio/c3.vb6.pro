Public Sub Main()
    ' ledger 245 / B94 behavior pin.
    ' IXMLDOMAttribute.value is a property Get whose type-library return type is VARIANT.
    ' mapType has no typed reader for that bucket, so the only legal reading is the Variant
    ' fallback vb6_VariantFromComResult(vb6_ComGetProp(...)). Trusting the right-hand context
    ' instead (a String target below) would emit GetStringProp and hand a BSTR* to a
    ' VARIANT-sized slot - that is the uninitialized-read family of ledger 245.
    ' Three spellings are pinned apart (each writes to its own variable so each emitted line
    ' is unique): plain read into a Variant, plain read into a String, and the call-shaped
    ' zero-arg read at.value(). Each takes a different consumer of the signature table.
    ' Witness against "always fall back": nodeName is a real BSTR property Get in the same
    ' type library and must still read through GetStringProp.
    Dim doc As New DOMDocument
    Dim el As IXMLDOMElement
    Dim at As IXMLDOMAttribute
    Dim v As Variant
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim n As Long
    doc.loadXML "<a b='1'/>"
    Set el = doc.documentElement
    Set at = el.getAttributeNode("b")
    v = at.value
    s1 = at.value
    s2 = at.value()
    n = Len(at.value)
    s3 = at.nodeName
    Debug.Print "CVP:done"
End Sub
