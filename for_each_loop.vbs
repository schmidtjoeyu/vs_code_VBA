Sub UpdateInventory()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim i As Integer
    Dim stockValue As Double
    Dim threshold As Double
    threshold = 10 ' Set your threshold value
    middle = 15
    For Each cell In ws.Range("B2:B101") ' Assuming stock values are in column A
        If cell.Value < 50 Then
            cell.Interior.Color = RGB(255, 0, 0) ' Highlight low stock in red
        End If
    Next cell
End Sub
