Attribute VB_Name = "FunctionTartalomEllenõrzés"
Option Explicit

Sub TartalomEllenõrzés()

If AppWindow.TextBox11 = "" Then
MsgBox "Bárcaszám hiányzik!"
ElseIf AppWindow.TextBox1 = "" Then
MsgBox "Munkaszám hiányzik!"
ElseIf AppWindow.TextBox10 = "" Then
MsgBox "RÁBAszám hiányzik!"
ElseIf AppWindow.TextBox9 = "" Then
MsgBox "Terület hiányzik!"
ElseIf AppWindow.TextBox8 = "" Then
MsgBox "Csapat hiányzik!"
ElseIf AppWindow.TextBox7 = "" Then
MsgBox "Kezdõ idõpont (-tól) hiányzik!"
ElseIf AppWindow.TextBox6 = "" Then
MsgBox "Záró idõpont (-ig) hiányzik!"
ElseIf AppWindow.TextBox5 = "" Then
MsgBox "Probléma leírás hiányzik!"
ElseIf AppWindow.TextBox4 = "" Then
MsgBox "Megoldás leírása hiányzik!"
ElseIf AppWindow.TextBox3 = "" Then
MsgBox "Javítás státusza hiányzik!"
ElseIf AppWindow.TextBox2 = "" Then
MsgBox "Mérés hiányzik!"
End If

End Sub
