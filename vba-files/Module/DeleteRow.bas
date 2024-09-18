Attribute VB_Name = "DeleteRow"
Sub DeleteLastRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim response As VbMsgBoxResult
    Dim message As String
    Dim title As String

    ActiveSheet.Unprotect ""

    Set ws = ThisWorkbook.Sheets("BangDiemDanh")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    message = UniConvert("Bajn cos chawsc muoosn xosa dofng cuoosi cufng?", "Telex")
    title = UniConvert("Xasc Nhaajn", "Telex")
    response = Application.Assistant.DoAlert(title, message, msoAlertButtonYesNo, msoAlertIconCritical, 0, 0, 0)

    If response = 6 Then
        ws.Rows(lastRow).Delete
    End If
    ActiveSheet.Protect ""

End Sub
