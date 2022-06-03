Attribute VB_Name = "FunctionBoxInaktíváló"
Option Explicit

Sub BoxInaktíváló()

AppWindow.TextBox1.Locked = False
AppWindow.TextBox10.Locked = False
AppWindow.ComboBox1.Locked = False
AppWindow.ComboBox2.Locked = False
AppWindow.TextBox7.Locked = False
AppWindow.ComboBox8.Locked = False
AppWindow.TextBox5.Locked = False

AppWindow.TextBox1.BackColor = &H80000005
AppWindow.TextBox10.BackColor = &H80000005
AppWindow.ComboBox1.BackColor = &H80000005
AppWindow.ComboBox2.BackColor = &H80000005
AppWindow.TextBox7.BackColor = &H80000005
AppWindow.ComboBox8.BackColor = &H80000005
AppWindow.TextBox5.BackColor = &H80000005

End Sub
