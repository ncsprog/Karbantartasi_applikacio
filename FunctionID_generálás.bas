Attribute VB_Name = "FunctionID_generálás"
Option Explicit

Sub ID_generálás()

Dim most As Date
most = Now()

Sheets("adatok").Select
Columns("w:w").Select
Selection.End(xlDown).Select
Dim ID_nr As Long
ID_nr = ActiveCell + 1
Dim ID_rw As Long
ID_rw = ActiveCell.Row + 1
Dim ID_oszlop As String
ID_oszlop = "w"
Dim ID_koord As String
ID_koord = ID_oszlop & ID_rw
Range(ID_koord) = most


Sheets("Start").Select
Range("b2").Select

End Sub
