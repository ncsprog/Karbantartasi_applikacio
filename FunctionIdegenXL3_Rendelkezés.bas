Attribute VB_Name = "FunctionIdegenXL3_Rendelkez�s"
Option Explicit

Sub IdegenXL3()
'Jelsz�Rejt�s2
' - Rendelkez�s - '


'ki t�rli a kor�bbi forr�sadatokat
    Munka11.Range("a1:x10000") = ""
    Range("a1").Select
         
    'megynyitja m�sik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forr�sadatok\�ll�sid� adott id�szakban.xlsx"
    Windows("�ll�sid� adott id�szakban.xlsx").Activate
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
    igXL = "v" & sorXL
   
   Range(t�lXL, igXL).Select
    
   'm�kol �tt:  a kijel�ltet �tm�solja az emez f�zetbe
    Selection.Copy
    Windows("Karbantart�si applik�ci�.xlsm").Activate
    Munka11.Range("a1").PasteSpecial xlPasteValues
    Windows("�ll�sid� adott id�szakban.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

'�ll�sid�

'Jelsz�Rejt�s
End Sub
