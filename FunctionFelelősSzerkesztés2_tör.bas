Attribute VB_Name = "FunctionFelel�sSzerkeszt�s2_t�r"
Option Explicit

Sub Felel�sSzerkeszt�s2()

Munka12.Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - hol keressen - '

Dim hol As String
hol = "d" & AppWindow.ListBox30.Value + 1
Range(hol).Select
' - t�rl�s - '
Selection.Delete Shift:=xlUp



End Sub
