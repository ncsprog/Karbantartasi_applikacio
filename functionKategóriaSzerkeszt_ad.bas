Attribute VB_Name = "functionKateg�riaSzerkeszt_ad"
Option Explicit

Sub Kateg�riaSzerkeszt()

'ez vissza adja a kijel�lt sor ID-t.

If AppWindow.TextBox108 = "" Then
MsgBox "Nincs megadva �j felel�s."
Else
Sheets("alapadatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row + 1
Dim ALoszlop As String
ALoszlop = "j"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw
Munka12.Range(ALkoord) = AppWindow.TextBox108.Value
AppWindow.TextBox108 = ""
End If

Sheets("Start").Select
Range("b2").Select


End Sub
