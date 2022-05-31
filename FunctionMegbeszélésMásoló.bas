Attribute VB_Name = "FunctionMegbeszélésMásoló"
Option Explicit

Sub MegbeszélésMásoló()

Munka16.Select
Range("a:ax").Select
Selection = ""


Munka1.Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim Xrw As Long
Xrw = ActiveCell.Row + 1
Dim Xcl As String
Xcl = "x"
Dim X As String
X = Xcl & Xrw
Range("a1", X).Copy

Munka16.Select
Range("a1").PasteSpecial xlPasteValues

End Sub
