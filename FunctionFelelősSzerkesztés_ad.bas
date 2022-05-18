Attribute VB_Name = "FunctionFelelõsSzerkesztés_ad"
Option Explicit

Sub FelelõsSzerkesztés()

'ez vissza adja a kijelölt sor ID-t.

If AppWindow.TextBox102 = "" Then
MsgBox "Nincs megadva új felelõs."
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

'JelszóRejtés
End Sub
