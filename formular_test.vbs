Sub CalculateSalesAndDiscount()
    ' Calculate total sales
    Range("B2").Formula = "=SUM(A2:A20)"
    ' Apply 10% discount to each sale
    Dim lastRow As Integer

    ' 这句 VBA 用于获取 A 列最后一个有数据的单元格所在行号
    ' Rows.Count, 返回工作表总行数。现代 Excel 通常是 1048576。
    ' Range("A" & Rows.Count) 定位到 A 列最后一个单元格，即：Range("A1048576")
    ' .End(xlUp).Row 从最后一个单元格向上查找，直到找到第一个非空单元格，返回该单元格的行号。
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("B3:B" & lastRow).Formula = "=A3*0.9"
End Sub

' 相对引用
Sub RelativeReferenceExample()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row and apply a formula using relative references
    Dim i As Long
    For i = 1 To lastRow
        ws.Cells(i, 2).FormulaR1C1 = "=RC[-1]*2" ' Apply 10% discount to column A values in column B
    Next i
End Sub
