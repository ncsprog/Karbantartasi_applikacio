Attribute VB_Name = "Function�resCella_inakt�v"
Option Explicit

Sub �resCella()

Dim dummy As String
dummy = "dummy"

If AppWindow.TextBox11 = "" Then
AppWindow.TextBox11.Value = dummy
End If


Sheets("Start").Select
Range("b2").Select
End Sub
