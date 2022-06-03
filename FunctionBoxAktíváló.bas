Attribute VB_Name = "FunctionBoxAktíváló"
Option Explicit

Sub BoxAktíváló()

AppWindow.TextBox1.Locked = True
AppWindow.TextBox10.Locked = True
AppWindow.ComboBox1.Locked = True
AppWindow.ComboBox2.Locked = True
AppWindow.TextBox7.Locked = True
AppWindow.ComboBox8.Locked = True
AppWindow.TextBox5.Locked = True

AppWindow.TextBox1.BackColor = &H8000000F
AppWindow.TextBox10.BackColor = &H8000000F
AppWindow.ComboBox1.BackColor = &H8000000F
AppWindow.ComboBox2.BackColor = &H8000000F
AppWindow.TextBox7.BackColor = &H8000000F
AppWindow.ComboBox8.BackColor = &H8000000F
AppWindow.TextBox5.BackColor = &H8000000F

End Sub
