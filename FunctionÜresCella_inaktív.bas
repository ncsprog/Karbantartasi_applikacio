Attribute VB_Name = "Function�resCella_inakt�v"
Option Explicit

Sub �resCella()
'Jelsz�Rejt�s2

Dim dummy As String
dummy = "dummy"

If AppWindow.TextBox11 = "" Then
AppWindow.TextBox11.Value = dummy
End If


Sheets("Start").Select
Range("b2").Select
'Jelsz�Rejt�s
End Sub
