Attribute VB_Name = "functionKategóriaSzerkeszt_ad"
Option Explicit

Sub KategóriaSzerkeszt()

'ez vissza adja a kijelölt sor ID-t.

If AppWindow.TextBox108 = "" Then
MsgBox "Nincs megadva új felelõs."
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
