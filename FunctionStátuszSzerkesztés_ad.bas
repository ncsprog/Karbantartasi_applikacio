Attribute VB_Name = "FunctionStátuszSzerkesztés_ad"
Option Explicit

Sub StátuszSzerkesztés()

'ez vissza adja a kijelölt sor ID-t.

Sheets("alapadatok").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row + 1
Dim ALoszlop As String
ALoszlop = "b"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Munka12.Range(ALkoord) = AppWindow.TextBox101.Value

AppWindow.TextBox101 = ""
Sheets("Start").Select
Range("b2").Select

'JelszóRejtés
End Sub
