Attribute VB_Name = "FunctionFelel�sSzerkeszt�s_ad"
Option Explicit

Sub Felel�sSzerkeszt�s()

'ez vissza adja a kijel�lt sor ID-t.

If AppWindow.TextBox102 = "" Then
MsgBox "Nincs megadva �j felel�s."
Else
Sheets("alapadatok").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row + 1
Dim ALoszlop As String
ALoszlop = "d"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw
Munka12.Range(ALkoord) = AppWindow.TextBox102.Value
AppWindow.TextBox102 = ""
End If

Sheets("Start").Select
Range("b2").Select

'Jelsz�Rejt�s
End Sub
