' 处理订购流程：获取订购数量，让用户确认，并显示处理结果。
Sub ProcessOrder()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    Dim cell As Range
    Set cell = ws.Range("A3")
    Dim t as Date
    t = Date
    ' 先用 String 保存 InputBox 的原始返回值，便于转换前进行校验。
    Dim inputValue As String

    ' Integer 用于保存用户输入的订购数量。
    Dim itemCount As Integer

    ' VbMsgBoxResult 用于保存 MsgBox 返回的按钮选择结果，
    ' 例如 vbYes、vbNo、vbOK 或 vbCancel。
    Dim response As VbMsgBoxResult

    ' InputBox(Prompt, Title, Default) 显示一个文本输入对话框：
    '   Prompt  ：对话框中的提示文字；
    '   Title   ：标题栏文字；
    '   Default ：输入框中预先显示的默认值。
    ' InputBox 返回 String；先保存原始输入，不立即进行数值转换。
    ' 用户单击“取消”时也会返回空字符串。
    inputValue = InputBox("Enter the number of items to order:", "Item Count", 1)

    ' Trim 去除输入前后的空格；如果没有内容，则提示错误并退出过程。
    If Len(Trim(inputValue)) = 0 Then
        MsgBox "Item count cannot be empty.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' IsNumeric 判断字符串是否可以解释为数字。
    If Not IsNumeric(inputValue) Then
        MsgBox "Item count must be a number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' 通过空值和数字校验后，再显式转换成 Integer。
    itemCount = Abs(CInt(inputValue))

    ' 只有正整数才进入确认步骤。
    If itemCount > 0 Then
        ' MsgBox(Prompt, Buttons, Title) 显示消息对话框并返回用户单击的按钮：
        '   Prompt  ：消息正文；
        '   Buttons ：按钮、图标等样式常量，可以用“+”组合；
        '   Title   ：标题栏文字。
        ' vbYesNo 显示“是/否”按钮，vbQuestion 显示问号图标。
        response = MsgBox("You are about to order " & itemCount & " items. Confirm?", vbYesNo + vbQuestion, "Order Confirmation")

        ' 将 MsgBox 的返回值与 vbYes 比较，判断用户是否确认。
        If response = vbYes Then
            ' 省略返回值接收时，可以不写括号。
            ' vbInformation 显示信息图标。
            MsgBox "Order of " & itemCount & " items confirmed!", vbInformation, "Confirmed"
        Else
            ' vbExclamation 显示警告图标。
            MsgBox "Order canceled.", vbExclamation, "Canceled"
        End If
    Else
        ' vbCritical 显示严重错误图标。
        MsgBox "Invalid item count.", vbCritical, "Error"
    End If

    ' 将最终的订购数量写入当前工作表的 A3 单元格。
    cell.Value = itemCount

    ' 这里的 _ 是 VBA 的“续行符”，表示当前语句还没有结束，下一行仍属于同一条语句。
    ThisWorkbook.SaveAs _
        Filename:="C:\Users\schmi\Downloads\Report_" & _
              Format(Date, "yyyy_mm_dd") & ".xlsm", _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled
End Sub
