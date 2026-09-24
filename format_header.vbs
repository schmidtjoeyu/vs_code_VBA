Sub FormatHeader()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim headerCell As Range
    Set headerCell = ws.Range("A1:D1") ' Change to your header range

    With headerCell
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255) ' White font color
        .Interior.Color = RGB(0, 102, 204) ' Blue background color
        .HorizontalAlignment = xlCenter ' xlCenter: Excel 内置常量(xl前缀), 表示水平居中对齐
        .VerticalAlignment = xlCenter ' xlCenter: 同一常量, 此处表示垂直居中对齐
    End With
End Sub