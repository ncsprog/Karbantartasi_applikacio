Attribute VB_Name = "FunctionLétszámÖsszesítés"
Option Explicit

Sub LétszámÖsszesítés()
'JelszóRejtés2

Sheets("létszám").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim LÖ_rw As Long
LÖ_rw = ActiveCell.Row
Dim LÖ_Koszlop As String
LÖ_Koszlop = "c"
Dim LÖ_Voszlop As String
LÖ_Voszlop = "k"
Dim kezd As String
kezd = LÖ_Koszlop & LÖ_rw
Dim vég As String
vég = LÖ_Voszlop & LÖ_rw



Dim létsz As Integer
létsz = Application.WorksheetFunction.Sum(Range(kezd, vég))

Sheets("létszám").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim szumRw As Long
szumRw = ActiveCell.Row + 1
Dim szumC As String
szumC = "m"
Dim szumRng As String
szumRng = szumC & szumRw

Range(szumRng).Value = létsz

Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
