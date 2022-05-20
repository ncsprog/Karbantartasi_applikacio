Attribute VB_Name = "FunctionKategóriaSzerkeszt2_tör"
Option Explicit

Sub KategóriaSzerkeszt2()


Munka12.Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "j"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

' - hol keressen - '

Dim hol As String
hol = "j" & AppWindow.ListBox37.Value + 1
Range(hol).Select
' - törlés - '
Selection.Delete Shift:=xlUp

End Sub
