Attribute VB_Name = "Mod1"
' C2224 Printer/Printers data properties: .DeviceName/.Port/.Copies/.PrintQuality
' are data properties, not methods. Before the fix the generic paths emitted
' vb6_Printer_Object().DeviceName (field read on void*) and
' vb6_Printers_Collection(i).DeviceName (arg passed to a 0-arg function + field read on void*).
Sub ProbePrinterProps()
    Dim s As String
    Dim i As Long
    s = Printer.DeviceName
    s = Printer.Port
    Printer.Copies = 2
    Printer.PrintQuality = -1
    i = Printers.Count
    s = Printers(0).DeviceName
    s = Printers(i).Port
End Sub
