Attribute VB_Name = "FunctionTer�letSzerkeszt2_t�r"
Option Explicit

Sub Ter�letSzerkeszt2()


Munka12.Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "p"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - hol keressen - '

Dim hol As String
hol = "p" & AppWindow.ListBox38.Value + 1
Range(hol).Select
' - t�rl�s - '
Selection.Delete Shift:=xlUp

End Sub
