Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim a, b, result

a = Parameter("aA")
b = Parameter("aB")
result = a + b

MyMsgBox.Show "Result: " + CStr(a) + " + " + CStr(b) + " = " + CStr(result)
