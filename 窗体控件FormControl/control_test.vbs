Private Sub ComboBox1_Change()

    ComboBox1.AddItem "Option 1"
    ComboBox1.AddItem "Option 2"
    ComboBox1.AddItem "Option 3"
    
End Sub

' 当用户单击 ListBox1 时触发
Private Sub ListBox1_Click()

    ' 用于保存所有选中项的内容
    Dim selectedValue As String

    ' ListBox 项目的循环索引
    Dim i As Integer

    ' 遍历 ListBox1 中的所有项目
    For i = 0 To ListBox1.ListCount - 1

        ' 判断当前项目是否被选中
        If ListBox1.Selected(i) Then

            ' 将选中项目追加到字符串中，每项占一行
            selectedValue = selectedValue & ListBox1.List(i) & vbNewLine
        End If
    Next i

    ' 弹窗显示所有选中的项目
    MsgBox "You selected: " & selectedValue

End Sub

Private Sub TextBox1_Change()
    Dim value As String
    value = TextBox1.value
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Button Clicked. You entered the value: " & TextBox1.value
End Sub

Private Sub CommandButton1_Click()
    Dim Value As VbMsgBoxResult
    value = MsgBox("Are you sure you want to proceed?", vbYesNo + vbQuestion, "Confirmation")
        ' 将 MsgBox 的返回值与 vbYes 比较，判断用户是否确认。
    If value = vbYes Then
        ' 省略返回值接收时，可以不写括号。
        ' vbInformation 显示信息图标。
        MsgBox "Confirmed"
    Else
        ' vbExclamation 显示警告图标。
        MsgBox "Order canceled."
    End If
End Sub
