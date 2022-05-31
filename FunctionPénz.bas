Attribute VB_Name = "FunctionPénz"
Option Explicit

Sub Pénz()
'JelszóRejtés2

'kimutatást csesz a kimutatásra pénzügyileg :D

Munka10.Range("a1").AutoFilter 14, 10, xlTop10Items
'anyag ktsg összegmásolás
Columns("n:n").Select
Selection.End(xlDown).Select
Dim igA_ktsg As Integer
igA_ktsg = ActiveCell.Row
Dim koordA_ktsg As String
koordA_ktsg = "n" & igA_ktsg


'illesssze be a kombobokszba a találati értékeket




Range("n2", koordA_ktsg).Copy
Munka9.Range("e11").PasteSpecial xlPasteValues
'anyag ktsg megnevezés másolás
Columns("l:l").Select
Selection.End(xlDown).Select
Dim igA_név As Integer
igA_név = ActiveCell.Row
Dim koordA_név As String
koordA_név = "l" & igA_ktsg
Range("l2", koordA_név).Copy
Munka9.Range("d11").PasteSpecial xlPasteValues

ActiveSheet.AutoFilterMode = False              'törli a szûrõt

'bér ktsg igazítás, jelenleg top 10-es
Munka15.Range("a1").AutoFilter 15, 10, xlTop10Items
'bér ktsg összegmásolás

Columns("o:o").Select
Selection.End(xlDown).Select
Dim igB_ktsg As Integer
igB_ktsg = ActiveCell.Row
Dim koordB_ktsg As String
koordB_ktsg = "o" & igB_ktsg
Range("o2", koordB_ktsg).Copy
Munka9.Range("h11").PasteSpecial xlPasteValues
'bér ktsg megnevezés másolás
Columns("l:l").Select
Selection.End(xlDown).Select
Dim igB_név As Integer
igB_név = ActiveCell.Row
Dim koordB_név As String
koordB_név = "l" & igB_ktsg
Range("l2", koordB_név).Copy
Munka9.Range("g11").PasteSpecial xlPasteValues

ActiveSheet.AutoFilterMode = False              'törli a szûrõt

'külsõ ktsg igazítás, jelenleg top 10-es
Munka15.Range("a1").AutoFilter 15, 10, xlTop10Items
'külsõ ktsg összegmásolás

Columns("p:p").Select
Selection.End(xlDown).Select
Dim igK_ktsg As Integer
igK_ktsg = ActiveCell.Row
Dim koordK_ktsg As String
koordK_ktsg = "p" & igB_ktsg
Range("p2", koordB_ktsg).Copy
Munka9.Range("k11").PasteSpecial xlPasteValues
'külsõ ktsg megnevezés másolás
Columns("l:l").Select
Selection.End(xlDown).Select
Dim igK_név As Integer
igK_név = ActiveCell.Row
Dim koordK_név As String
koordK_név = "l" & igK_ktsg
Range("l2", koordK_név).Copy
Munka9.Range("j11").PasteSpecial xlPasteValues

ActiveSheet.AutoFilterMode = False              'törli a szûrõt
'Munka9.Range("d10").Select
'JelszóRejtés
End Sub
