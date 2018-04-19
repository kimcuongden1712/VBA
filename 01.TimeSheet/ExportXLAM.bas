Attribute VB_Name = "ExportXLAM"
Sub Macro1()
    Dim x
    x = Application.GetOpenFilename("xlamFiles,*.xlam")
    If VarType(x) = vbBoolean Then Exit Sub
    On Error GoTo errHandler
    With Workbooks.Open(x)
        .IsAddin = False
        .SaveAs Filename:=Replace$(x, "xlam", "xlsm"), _
                FileFormat:=xlOpenXMLWorkbookMacroEnabled
        .Close
    End With
    'Kill x
    Exit Sub
errHandler:
    MsgBox Err.Number & "::" & Err.Description
End Sub

