Attribute VB_Name = "FunctionCsapatSzerkeszt_ad"
Option Explicit

Sub CsapatSzerkeszt()

'ez vissza adja a kijelölt sor ID-t.

If AppWindow.TextBox110 = "" Then
MsgBox "Nincs megadva új felelõs."
Else
Sheets("alapadatok").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row + 1
Dim ALoszlop As String
ALoszlop = "m"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw
Munka12.Range(ALkoord) = AppWindow.TextBox110.Value
AppWindow.TextBox110 = ""
End If

Sheets("Start").Select
Range("b2").Select

End Sub
