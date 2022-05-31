Attribute VB_Name = "FunctionIdegenXL_gazdasági"
Option Explicit

Sub IdegenXL()
'JelszóRejtés2

    'ki törli a korábbi forrásadatokat
    Munka10.Range("a1:p10000") = ""
    Range("a1").Select
         
    'megynyitja másik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forrásadatok\gazdasági lekérdezett adatok.xlsx"
    Windows("gazdasági lekérdezett adatok.xlsx").Activate
    'abban mókol:   kijelöli a kívánt adatokat
    'mettõl
    'Munka1.Select
    Dim tólXL As Range
    Set tólXL = Range("a1")
    'meddig
     'Sheets("FNDWRR").Select
    Columns("a:a").Select
    Selection.End(xlDown).Select
    Dim sorXL As Long
    sorXL = ActiveCell.Row
    Dim igXL As String
    igXL = "p" & sorXL
   
   Range(tólXL, igXL).Select
    
   'mókol átt:  a kijelöltet átmásolja az emez füzetbe
    Selection.Copy
    Windows("Karbantartási applikáció.xlsm").Activate
    Munka10.Range("a1").PasteSpecial xlPasteValues
    Windows("gazdasági lekérdezett adatok.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

'Pénz
'JelszóRejtés
End Sub
