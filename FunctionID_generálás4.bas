Attribute VB_Name = "FunctionID_gener�l�s4"
Option Explicit

Sub ID_gener�l�s4()

' - L�tsz�mkeret - '

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
