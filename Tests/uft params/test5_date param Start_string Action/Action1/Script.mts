Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim param, result

param = CDate(Parameter("aDate"))

result = FormatDateTime(param, vbLongDate)

MyMsgBox.Show "Long Date Format: " & CStr(result), "Date"
