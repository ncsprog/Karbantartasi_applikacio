Attribute VB_Name = "FunctionIdegenXL3_Rendelkezés"
Option Explicit

Sub IdegenXL3()
'JelszóRejtés2
' - Rendelkezés - '


'ki törli a korábbi forrásadatokat
    Munka11.Range("a1:x10000") = ""
    Range("a1").Select
         
    'megynyitja másik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forrásadatok\Állásidõ adott idõszakban.xlsx"
    Windows("Állásidõ adott idõszakban.xlsx").Activate
    'abban mókol:   kijelöli a kívánt adatokat
    'mettõl
    Sheets("FNDWRR").Select
    Dim tólXL As Range
    Set tólXL = Range("a1")
    'meddig
     Sheets("FNDWRR").Select
     Columns("a:a").Select
    Selection.End(xlDown).Select
    Dim sorXL As Long
    sorXL = ActiveCell.row
    Dim igXL As String
    igXL = "v" & sorXL
   
   Range(tólXL, igXL).Select
    
   'mókol átt:  a kijelöltet átmásolja az emez füzetbe
    Selection.Copy
    Windows("Karbantartási applikáció.xlsm").Activate
    Munka11.Range("a1").PasteSpecial xlPasteValues
    Windows("Állásidõ adott idõszakban.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

'Állásidõ

'JelszóRejtés
End Sub
