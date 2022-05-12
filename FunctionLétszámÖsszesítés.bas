Attribute VB_Name = "FunctionLétszámÖsszesítés"
Option Explicit

Sub LétszámÖsszesítés()


Sheets("Létszám").Select
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


Dim létszám As Integer
létszám = Application.WorksheetFunction.Sum(Range(kezd, vég))
AppWindow.TextBox80.Value = létszám

End Sub
