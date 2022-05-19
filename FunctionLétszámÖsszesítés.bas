Attribute VB_Name = "FunctionLétszámÖsszesítés"
Option Explicit

Sub LétszámÖsszesítés()
'JelszóRejtés2

Sheets("létszám").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim LÖ_rw As Long
LÖ_rw = ActiveCell.row
Dim LÖ_Koszlop As String
LÖ_Koszlop = "c"
Dim LÖ_Voszlop As String
LÖ_Voszlop = "af"
Dim kezd As String
kezd = LÖ_Koszlop & LÖ_rw
Dim vég As String
vég = LÖ_Voszlop & LÖ_rw


Dim létsz As Integer
létsz = Application.WorksheetFunction.Sum(Range(kezd, vég))
AppWindow.TextBox80.Value = létsz

Sheets("létszám").Select
Columns("ah:ah").Select
Selection.End(xlDown).Select
Dim szumRw As Long
szumRw = ActiveCell.row + 1
Dim szumC As String
szumC = "ah"
Dim szumRng As String
szumRng = szumC & szumRw

Range(szumRng).Value = létsz

Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
