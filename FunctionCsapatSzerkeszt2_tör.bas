Attribute VB_Name = "FunctionCsapatSzerkeszt2_t�r"
Option Explicit

Sub CsapatSzerkeszt2()

Munka12.Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "m"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - hol keressen - '

Dim hol As String
hol = "m" & AppWindow.ListBox39.Value + 1
Range(hol).Select
' - t�rl�s - '
Selection.Delete Shift:=xlUp

End Sub
