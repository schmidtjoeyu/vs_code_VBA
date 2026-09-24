' 本脚本用于演示 For Each 循环的用法：遍历指定区域中的每个单元格
' cell 声明为 Range 类型是因为第 6 行的 For Each cell In ws.Range("A1:D6") 要遍历一个区域（Range 对象）中的每一个单元格。（Excel 对象模型里没有单独的"Cell"类型，单个单元格就是一个只包含 1 行 1 列的 Range 对象）。所以循环变量 cell 每次会被赋值成当前这个单元格对应的 Range 对象，才能用 cell.Value、cell.fillColor 等 Range 的属性/方法操作它。

Sub FillEmptyCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim cell As Range
    For Each cell In ws.Range("A1:D6") ' Change to your desired range
        If IsEmpty(cell.Value) Then
            cell.Interior.Color = RGB(255, 220, 0) ' Fill empty cells with yellow color
        End If
    Next cell
End Sub