Attribute VB_Name = "FunctionIdegenXL"
Option Explicit

Sub IdegenXL()

    'ki t�rli a kor�bbi forr�sadatokat
    Munka5.Range("a1:x10000").clear
    Range("a1").Select
         
    'megynyitja m�sik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\gazdas�gi lek�rdezett adatok.xlsx"
    Windows("gazdas�gi lek�rdezett adatok.xlsx").Activate
    'abban m�kol:   kijel�li a k�v�nt adatokat
    'mett�l
    Sheets("FNDWRR").Select
    Dim t�lXL As Range
    Set t�lXL = Range("a1")
    'meddig
     Sheets("FNDWRR").Select
     Columns("a:a").Select
    Selection.End(xlDown).Select
    Dim sorXL As Long
    sorXL = ActiveCell.row
    Dim igXL As String
    igXL = "p" & sorXL
   
   Range(t�lXL, igXL).Select
    
   'm�kol �tt:  a kijel�ltet �tm�solja az emez f�zetbe
    Selection.Copy
    Windows("Karbantart�si applik�ci�.xlsm").Activate
    Munka5.Range("a1").PasteSpecial xlPasteValues
    Windows("gazdas�gi lek�rdezett adatok.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

'P�nz

End Sub
