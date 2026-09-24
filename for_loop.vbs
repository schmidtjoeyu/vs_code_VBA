Sub ForAddCheck()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim i As Integer
    Dim sum As Double
    sum = 0
    i = 1

    For i = 1 To 9
        ws.Cells(i, 2).Value = "Checked" ' Write "Checked" in column B
    Next i
End Sub