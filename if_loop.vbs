Sub stock_check()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim i As Integer
    Dim stockValue As Double
    Dim threshold As Double
    threshold = 10 ' Set your threshold value
    middle = 15
    For i = 1 To 20
        stockValue = ws.Cells(i, 1).Value ' Assuming stock values are in column C
        If stockValue > middle Then
            ws.Cells(i, 2).Value = "Sufficient Stock" ' Write "Low Stock" in column D
        ElseIf stockValue >=threshold AND stockValue < middle Then
            ws.Cells(i, 2).Value = "Normal Stock" ' Write "Sufficient Stock" in column D
        Else
            ws.Cells(i, 2).Value = "Low Stock" ' Write "Sufficient Stock" in column D
        End If
    Next i
End Sub