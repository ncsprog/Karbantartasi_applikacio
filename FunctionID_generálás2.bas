Attribute VB_Name = "FunctionID_generálás2"
Option Explicit

Sub ID_generálás2()

Dim most As Date
most = Now()

Sheets("Létszám").Select
Columns("ag:ag").Select
Selection.End(xlDown).Select
Dim ID_nr As Long
ID_nr = ActiveCell + 1
Dim ID_rw As Long
ID_rw = ActiveCell.row + 1
Dim ID_oszlop As String
ID_oszlop = "ag"
Dim ID_koord As String
ID_koord = ID_oszlop & ID_rw
Range(ID_koord) = most

End Sub
