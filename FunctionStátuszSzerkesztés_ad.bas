Attribute VB_Name = "FunctionSt�tuszSzerkeszt�s_ad"
Option Explicit

Sub St�tuszSzerkeszt�s()

'ez vissza adja a kijel�lt sor ID-t.

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

'Jelsz�Rejt�s
End Sub
