Attribute VB_Name = "FunctionG�p�ll�sokSz�ma"
Option Explicit

Sub G�p�ll�sokSz�ma()

Sheets("transfer_kulcsg�p").Select

Columns("r:r").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "r"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Dim �ll�sok As Long
�ll�sok = Application.WorksheetFunction.Count(Range("r1", ALkoord))
AppWindow.TextBox97.Value = "�ll�sok sz�ma: " & �ll�sok & " db"

End Sub
