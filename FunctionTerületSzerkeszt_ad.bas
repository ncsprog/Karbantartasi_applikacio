Attribute VB_Name = "FunctionTer�letSzerkeszt_ad"
Option Explicit

Sub Ter�letSzerkeszt()

'ez vissza adja a kijel�lt sor ID-t.

If AppWindow.TextBox109 = "" Then
MsgBox "Nincs megadva �j felel�s."
Else
Sheets("alapadatok").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row + 1
Dim ALoszlop As String
ALoszlop = "p"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw
Munka12.Range(ALkoord) = AppWindow.TextBox109.Value
AppWindow.TextBox109 = ""
End If

Sheets("Start").Select
Range("b2").Select


End Sub
