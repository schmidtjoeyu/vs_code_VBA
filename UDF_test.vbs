
' originalPrice(i, 1) 是 Variant 元素，但 CalculatePrice 要求 Double。由于参数默认为 ByRef，两边类型必须匹配，所以编译报错。推荐解决方法：参数改为 ByVal
' 如果不写返回类型，VBA 默认会把返回值当作 Variant，函数仍可能运行，但题目说“there is no need to define a return type for the function to work correctly in Excel”这种说法通常应判断为 False，因为规范定义 UDF 时应指定合适的返回类型。
Function DiscountedPrice(ByVal price As Double, discountRate As Double) As Double  
    If discountRate > 0 And discountRate <= 1 Then
        DiscountedPrice = price * (1 - discountRate)
    Else
        DiscountedPrice = price
    End If
End Function

Sub ApplyDiscount()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim i As Integer

    ' 定义成Variant类型，以便存储从单元格范围获取的任意类型的数据，避免类型不匹配错误。
    Dim originalPrice As Variant
    Dim discountRate As Double
    discountRate = 0.1 ' 10% discount

    originalPrice = ws.Range("A1:A20").Value ' Assuming original

    For i = LBound(originalPrice, 1) To UBound(originalPrice, 1)
        ws.Cells(i, 2).Value = DiscountedPrice(originalPrice(i, 1), discountRate) ' Write discounted price in column B
    Next i
End Sub

