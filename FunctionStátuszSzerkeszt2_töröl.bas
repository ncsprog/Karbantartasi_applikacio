Attribute VB_Name = "FunctionStátuszSzerkeszt2_töröl"
Option Explicit

Sub StátuszSzerkeszt2()
' - ez megadja a kijelölt sor értékét - '

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
' - törlés - '
Selection.Delete Shift:=xlUp


End Sub
