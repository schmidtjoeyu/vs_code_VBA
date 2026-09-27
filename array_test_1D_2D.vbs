Sub IterateArray()
    Dim arr(4) As Integer
    Dim i As Integer

    ' Initialize the array with values
    For i = 0 To 4
        arr(i) = (i + 1) * 10
    Next i

    ' Iterate through the array and display values
    ' LBound(arr)：返回数组 arr 的最小下标
    ' UBound(arr)：返回数组 arr 的最大下标
    For i = LBound(arr) To UBound(arr)
        MsgBox "Value at index " & i + 1 & " is: " & arr(i)
    Next i
End Sub

Sub Populate2DArray()
    Dim grid(2, 3) As String
    Dim i As Integer, j As Integer
    For i = 0 To 2
        For j = 0 To 3
            grid(i, j) = "R" & i + 1 & "C" & j + 1
        Next j
    Next i

    MsgBox "2D Array populated. Value at (2,3): " & grid(1, 2) ' Display value at row 2, column 3
End Sub

' Read data from a worksheet into an array, manipulate it, and write it back.
Sub ReadWriteArray()
    Dim dataArray() As Variant
    Dim i As Long
    ' Read data from Range into Array
    dataArray = Range("A1:A10").Value
    ' Modify Array
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        dataArray(i, 1) = dataArray(i, 1) * 2
    Next i
    ' Write Array back to Range
    Range("B1:B10").Value = dataArray
End Sub

Sub Iterate1DArray()
    Dim dataArray() As Variant
    Dim i As Long
    ' 得到的仍然是一个二维数组，结构为 10 行 × 1 列：
    dataArray = Range("A1:A10").Value

    ' Iterate through 1D Array and display values
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        MsgBox "Value at A" & i + 1 & " is: " & dataArray(i, 1)
    Next i
End Sub