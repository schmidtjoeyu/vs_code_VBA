Sub WhileLoop()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim i As Integer
    i = 1

    Do While NOT IsEmpty(ws.Cells(i, 1)) ' Loop while the cell is not empty
        ws.Cells(i, 2).Value = "Checked" ' Write "Checked" in column B
        i = i + 1 ' Increment the counter
    Loop
End Sub

Sub SumNumber()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim i As Integer
    Dim sum As Double
    sum = 0
    i = 1

    Do Until sum >= 5000 ' Loop while the sum is less than 5000
        sum = sum + ws.Cells(i, 4).Value ' Add the value in column A to the sum
        Cells(1, 5).Value = sum ' Write the sum in column E, row 1
        i = i + 1 ' Increment the counter
    Loop
End Sub