Attribute VB_Name = "FunctionIdegenXL3_Rendelkezés"
Option Explicit

Sub IdegenXL3()
'JelszóRejtés2
' - Rendelkezés - '


'ki törli a korábbi forrásadatokat
    Munka14.Range("a1:v10000") = ""
    Range("a1").Select
         
    'megynyitja másik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forrásadatok\Rendelkezésre állás és összállásidõ idõszakra.xlsx"
    Windows("Rendelkezésre állás és összállásidõ idõszakra.xlsx").Activate
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
    sorXL = ActiveCell.Row
    Dim igXL As String
    igXL = "v" & sorXL
   
   Range(tólXL, igXL).Select
    
   'mókol átt:  a kijelöltet átmásolja az emez füzetbe
    Selection.Copy
    Windows("Karbantartási applikáció.xlsm").Activate
    Munka14.Range("a1").PasteSpecial xlPasteValues
    Windows("Rendelkezésre állás és összállásidõ idõszakra.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

'Állásidõ

'JelszóRejtés
End Sub
