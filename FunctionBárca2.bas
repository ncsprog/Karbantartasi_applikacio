Attribute VB_Name = "FunctionBárca2"
Option Explicit

Sub Bárca2()
'JelszóRejtés2

If AppWindow.TextBox54 = "" Then
MsgBox "Bárcaszám megadása kötelezõ!"
End If

Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
