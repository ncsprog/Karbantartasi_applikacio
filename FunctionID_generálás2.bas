Attribute VB_Name = "FunctionID_generálás2"
Option Explicit

Sub ID_generálás2()
'JelszóRejtés2

Dim most As Date
most = Now()

Sheets("Létszám").Select
Columns("L:L").Select
Selection.End(xlDown).Select
Dim ID_nr As Long
ID_nr = ActiveCell + 1
Dim ID_rw As Long
ID_rw = ActiveCell.row + 1
Dim ID_oszlop As String
ID_oszlop = "l"
Dim ID_koord As String
ID_koord = ID_oszlop & ID_rw
Range(ID_koord) = most


Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
