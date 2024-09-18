Attribute VB_Name = "SetDate"
Public Sub SetDateInMonths()
    Dim ws As Worksheet
    Dim monthValue As Integer
    Dim yearValue As Integer
    Dim dayCounter As Integer
    Dim lastDayOfMonth As Integer
    Dim col As Integer

    Set ws = ThisWorkbook.Sheets("BangDiemDanh")

    monthValue = ws.Range("C3").Value
    yearValue = ws.Range("CD4").Value

    lastDayOfMonth = Day(DateSerial(yearValue, monthValue + 1, 0))

    dayCounter = 1

    For col = 3 To 33
        If dayCounter <= lastDayOfMonth Then
            ws.Cells(8, col).Value = DateSerial(yearValue, monthValue, dayCounter)
            dayCounter = dayCounter + 1
        Else
            ws.Cells(8, col).Value = ""
        End If
    Next col
End Sub


