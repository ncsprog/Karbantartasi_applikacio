Attribute VB_Name = "FunctionIdegenXL_gazdas�gi"
Option Explicit

Sub IdegenXL()

    'ki t�rli a kor�bbi forr�sadatokat
    Munka10.Range("a1:p10000") = ""
    Range("a1").Select
         
    'megynyitja m�sik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forr�sadatok\gazdas�gi lek�rdezett adatok.xlsx"
    Windows("gazdas�gi lek�rdezett adatok.xlsx").Activate
    'kijel�li a k�v�nt adatokat
    'mett�l
    'Munka1.Select
    Dim t�lXL As Range
    Set t�lXL = Range("a1")
    'meddig
    Columns("a:a").Select
    Selection.End(xlDown).Select
    Dim sorXL As Long
    sorXL = ActiveCell.Row
    Dim igXL As String
    igXL = "p" & sorXL
   
   Range(t�lXL, igXL).Select
    
   'a kijel�ltet �tm�solja
    Selection.Copy
    Windows("Karbantart�si applik�ci�.xlsm").Activate
    Munka10.Range("a1").PasteSpecial xlPasteValues
    Windows("gazdas�gi lek�rdezett adatok.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

End Sub
