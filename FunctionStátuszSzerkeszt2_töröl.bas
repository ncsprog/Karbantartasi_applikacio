Attribute VB_Name = "FunctionSt�tuszSzerkeszt2_t�r�l"
Option Explicit

Sub St�tuszSzerkeszt2()
' - ez megadja a kijel�lt sor �rt�k�t - '

Munka12.Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "b"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - hol keressen - '

Dim hol As String
hol = "b" & AppWindow.ListBox29.Value + 1
Range(hol).Select
' - t�rl�s - '
Selection.Delete Shift:=xlUp


End Sub
