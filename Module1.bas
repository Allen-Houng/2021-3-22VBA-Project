Attribute VB_Name = "Module1"
Sub MsgboxDemo()
Dim rst As Integer
rst = MsgBox("要結束程式嗎?", vbYesNo, "結束程式")
If rst = vbYes Then Application.Quit

End Sub
Sub IntegerDemo()
Dim i As Integer
i = 1000
MsgBox i
End Sub
Sub StringDemo()
Dim i As String
i = "VBA"
MsgBox i
End Sub
Sub SingleDemo()
Dim i As Single
i = 121.233
MsgBox i
End Sub
Sub DoubleDemo()
Dim i As Double
i = 121.233366963366
MsgBox i
End Sub
Sub DateDemo()
Dim i As Date
i = Now
MsgBox i
End Sub
Sub BooleanDemo()
Dim i As Boolean
i = True
MsgBox i
End Sub
Sub IntputBoxDemo()
Dim userString As String '宣告String變數
'InputBox為使用者輸入框的談條視窗，第一個引數為視窗主文字
'變數值為輸入框內容-變數=輸入框函數，即可將輸入框的內容存到變數
userString = InputBox("Please enter your VBA score")
'彈跳視窗呈現，顯示輸入的資料(字串相加用&)，本行程式 = "Your score" + 使用者Input
MsgBox "Your score is :" & userString
End Sub
