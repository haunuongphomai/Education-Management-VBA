Attribute VB_Name = "AddRow"
Sub AddRowWithFormatAndFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newRow As Long
    Dim lastSerialNumber As Long

    On Error Resume Next
    ActiveSheet.Unprotect ""

    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastSerialNumber = ws.Cells(lastRow, 1).Value

    ws.Rows(lastRow + 1).Insert Shift:=xlDown
    ws.Rows(lastRow).Copy
    ws.Rows(lastRow + 1).PasteSpecial Paste:=xlPasteFormats
    ws.Rows(lastRow + 1).PasteSpecial Paste:=xlPasteFormulas
    ws.Rows(lastRow + 1).SpecialCells(xlCellTypeConstants).ClearContents
    newRow = lastRow + 1

    ws.Cells(newRow, 1).Value = lastSerialNumber + 1
    Application.CutCopyMode = False

    ActiveSheet.Protect ""
End Sub
