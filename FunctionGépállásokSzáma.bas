Attribute VB_Name = "FunctionGépállásokSzáma"
Option Explicit

Sub GépállásokSzáma()
'JelszóRejtés2

Sheets("transfer_kulcsgép").Select

Columns("r:r").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
Dim ALoszlop As String
ALoszlop = "r"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Dim Állások As Long
Állások = Application.WorksheetFunction.Count(Range("r1", ALkoord))
AppWindow.TextBox97.Value = "Állások száma: " & Állások & " db"


'JelszóRejtés
End Sub
