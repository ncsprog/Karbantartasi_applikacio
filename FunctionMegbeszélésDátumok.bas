Attribute VB_Name = "FunctionMegbeszélésDátumok"
Option Explicit

Sub MegbeszélésDátumok()

Munka12.Range("x2:x3652") = ""

Dim Napok As Date, D0 As Date, Dx As Date
Munka12.Range("x2").Value = Date
Munka12.Range("x3").Value = Munka12.Range("x2") - AppWindow.TextBox134.Value
Napok = AppWindow.TextBox134.Value
D0 = Date
Dx = D0 - Napok

For Napok = Dx To D0 Step 1
'MsgBox Napok

Munka12.Select
Columns("x:x").Select
Selection.End(xlDown).Select
Dim Dtrw As Integer
Dtrw = ActiveCell.Row + 1
Dim koord As String
koord = "x" & Dtrw
Range(koord).Value = Napok

Next


End Sub
