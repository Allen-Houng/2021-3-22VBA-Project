Attribute VB_Name = "Module1"
Sub MsgboxDemo()
Dim rst As Integer
rst = MsgBox("�n�����{����?", vbYesNo, "�����{��")
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
Dim userString As String '�ŧiString�ܼ�
'InputBox���ϥΪ̿�J�ت��ͱ������A�Ĥ@�Ӥ޼Ƭ������D��r
'�ܼƭȬ���J�ؤ��e-�ܼ�=��J�ب�ơA�Y�i�N��J�ت����e�s���ܼ�
userString = InputBox("Please enter your VBA score")
'�u�������e�{�A��ܿ�J�����(�r��ۥ[��&)�A����{�� = "Your score" + �ϥΪ�Input
MsgBox "Your score is :" & userString
End Sub
