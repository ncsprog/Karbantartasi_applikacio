Attribute VB_Name = "FunctionP�nz"
Option Explicit

Sub P�nz()
'Jelsz�Rejt�s2

'kimutat�st csesz a kimutat�sra p�nz�gyileg :D

Munka10.Range("a1").AutoFilter 14, 10, xlTop10Items
'anyag ktsg �sszegm�sol�s
Columns("n:n").Select
Selection.End(xlDown).Select
Dim igA_ktsg As Integer
igA_ktsg = ActiveCell.Row
Dim koordA_ktsg As String
koordA_ktsg = "n" & igA_ktsg


'illesssze be a kombobokszba a tal�lati �rt�keket




Range("n2", koordA_ktsg).Copy
Munka9.Range("e11").PasteSpecial xlPasteValues
'anyag ktsg megnevez�s m�sol�s
Columns("l:l").Select
Selection.End(xlDown).Select
Dim igA_n�v As Integer
igA_n�v = ActiveCell.Row
Dim koordA_n�v As String
koordA_n�v = "l" & igA_ktsg
Range("l2", koordA_n�v).Copy
Munka9.Range("d11").PasteSpecial xlPasteValues

ActiveSheet.AutoFilterMode = False              't�rli a sz�r�t

'b�r ktsg igaz�t�s, jelenleg top 10-es
Munka15.Range("a1").AutoFilter 15, 10, xlTop10Items
'b�r ktsg �sszegm�sol�s

Columns("o:o").Select
Selection.End(xlDown).Select
Dim igB_ktsg As Integer
igB_ktsg = ActiveCell.Row
Dim koordB_ktsg As String
koordB_ktsg = "o" & igB_ktsg
Range("o2", koordB_ktsg).Copy
Munka9.Range("h11").PasteSpecial xlPasteValues
'b�r ktsg megnevez�s m�sol�s
Columns("l:l").Select
Selection.End(xlDown).Select
Dim igB_n�v As Integer
igB_n�v = ActiveCell.Row
Dim koordB_n�v As String
koordB_n�v = "l" & igB_ktsg
Range("l2", koordB_n�v).Copy
Munka9.Range("g11").PasteSpecial xlPasteValues

ActiveSheet.AutoFilterMode = False              't�rli a sz�r�t

'k�ls� ktsg igaz�t�s, jelenleg top 10-es
Munka15.Range("a1").AutoFilter 15, 10, xlTop10Items
'k�ls� ktsg �sszegm�sol�s

Columns("p:p").Select
Selection.End(xlDown).Select
Dim igK_ktsg As Integer
igK_ktsg = ActiveCell.Row
Dim koordK_ktsg As String
koordK_ktsg = "p" & igB_ktsg
Range("p2", koordB_ktsg).Copy
Munka9.Range("k11").PasteSpecial xlPasteValues
'k�ls� ktsg megnevez�s m�sol�s
Columns("l:l").Select
Selection.End(xlDown).Select
Dim igK_n�v As Integer
igK_n�v = ActiveCell.Row
Dim koordK_n�v As String
koordK_n�v = "l" & igK_ktsg
Range("l2", koordK_n�v).Copy
Munka9.Range("j11").PasteSpecial xlPasteValues

ActiveSheet.AutoFilterMode = False              't�rli a sz�r�t
'Munka9.Range("d10").Select
'Jelsz�Rejt�s
End Sub
