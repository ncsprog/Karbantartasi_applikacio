Attribute VB_Name = "FunctionIdegenXL2_�ll�sid�"
Option Explicit

Sub IdegenXL2()

    'ki t�rli a kor�bbi forr�sadatokat
    Munka11.Range("a1:x10000") = ""
    Range("a1").Select
         
    'megynyitja m�sik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forr�sadatok\�ll�sid� adott id�szakban.xlsx"
    Windows("�ll�sid� adott id�szakban.xlsx").Activate
    'mett�l
    Sheets("FNDWRR").Select
    Dim t�lXL As Range
    Set t�lXL = Range("a1")
    'meddig
     Sheets("FNDWRR").Select
     Columns("a:a").Select
    Selection.End(xlDown).Select
    Dim sorXL As Long
    sorXL = ActiveCell.Row
    Dim igXL As String
    igXL = "v" & sorXL
   
   Range(t�lXL, igXL).Select
    
    Selection.Copy
    Windows("Karbantart�si applik�ci�.xlsm").Activate
    Munka11.Range("a1").PasteSpecial xlPasteValues
    Windows("�ll�sid� adott id�szakban.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

End Sub
