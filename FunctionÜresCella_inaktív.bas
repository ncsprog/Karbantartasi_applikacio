Attribute VB_Name = "FunctionÜresCella_inaktív"
Option Explicit

Sub ÜresCella()
'JelszóRejtés2

Dim dummy As String
dummy = "dummy"

If AppWindow.TextBox11 = "" Then
AppWindow.TextBox11.Value = dummy
End If


Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
