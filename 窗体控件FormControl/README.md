Excel 里的 **Form Control（窗体控件）**，是 Excel 自带的一组比较简单的交互控件，用来让用户在工作表上点击、选择、勾选，然后触发宏或者改变单元格值。

常见的 Form Controls 有：按钮 Button、复选框 Check Box、单选按钮 Option Button、下拉框 Combo Box、列表框 List Box、滚动条 Scroll Bar、微调按钮 Spin Button 等。

比如你插入一个 **Button（Form Control）**，然后可以把它直接绑定到一个普通 VBA 宏：

```vba
Sub Test()
    MsgBox "Hello"
End Sub
```

用户点击按钮时，就执行 `Test()`。

它和 ActiveX Control 最大的区别是：**Form Control 更简单、更稳定、安全风险更低，但可编程能力也弱一些。**

| 对比 | Form Control | ActiveX Control |
|---|---|---|
| 使用难度 | 简单 | 较复杂 |
| 绑定 VBA | 直接指定 Macro | 通过事件代码 |
| 事件支持 | 少 | 很丰富 |
| 属性设置 | 较少 | 很多 |
| 稳定性 | 较高 | 相对容易出现兼容问题 |
| 安全限制 | 较少 | 更严格 |
| Mac 支持 | 相对更好 | 基本不支持 |
| 适合场景 | 普通 Excel 自动化 | 复杂 GUI |

例如 Form Control Button 通常是：

```text
按钮
 ↓
Assign Macro
 ↓
Sub MyMacro()
```

而 ActiveX Button 是：

```vba
Private Sub CommandButton1_Click()

End Sub
```

也就是说，ActiveX 有真正的 `Click`、`Change`、`MouseMove` 等事件机制，而 Form Control 更像是一个“触发宏的按钮”。

如果你只是做 Excel + VBA 自动化，例如：

```text
点击按钮
→ 导入数据
→ 处理数据
→ 生成报表
```

我通常更建议优先使用 **Form Control Button**，没有必要为了一个简单按钮用 ActiveX。

你可以在 Excel：

**Developer → Insert**

里面看到两组：

```text
Form Controls
----------------
Button
Check Box
Option Button
Combo Box
...

ActiveX Controls
----------------
Command Button
Check Box
Combo Box
...
```

所以你刚才问的“为什么 Excel 默认禁用 ActiveX”，这里就能看出区别：**Form Control 本质更轻量，而 ActiveX 是 Windows COM/ActiveX 技术的一部分，权限和复杂度都更高。**