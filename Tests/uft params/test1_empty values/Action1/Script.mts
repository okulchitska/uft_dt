Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim param1, param2

param1 = Parameter("aParam1")
param2 = Parameter("aParam2")

MyMsgBox.Show "Parameter 1: " + param1
MyMsgBox.Show "Parameter 2: " + param2
