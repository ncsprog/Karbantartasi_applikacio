Attribute VB_Name = "FunctionID_generálás4"
Option Explicit

Sub ID_generálás4()

' - Létszámkeret - '

Dim most As Date
most = Date

Sheets("alapadatok").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim ID_nr As Long
ID_nr = ActiveCell + 1
Dim ID_rw As Long
ID_rw = ActiveCell.row + 1
Dim ID_oszlop As String
ID_oszlop = "g"
Dim ID_koord As String
ID_koord = ID_oszlop & ID_rw
Range(ID_koord) = most


Sheets("Start").Select
Range("b2").Select

End Sub
