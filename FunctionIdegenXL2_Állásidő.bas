Attribute VB_Name = "FunctionIdegenXL2_Állásidõ"
Option Explicit

Sub IdegenXL2()

    'ki törli a korábbi forrásadatokat
    Munka11.Range("a1:x10000") = ""
    Range("a1").Select
         
    'megynyitja másik file-t
    Workbooks.Open "\\rabart\frs$\sajat$\09049\Ncsp\programok\Forrásadatok\Állásidõ adott idõszakban.xlsx"
    Windows("Állásidõ adott idõszakban.xlsx").Activate
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
    
    Selection.Copy
    Windows("Karbantartási applikáció.xlsm").Activate
    Munka11.Range("a1").PasteSpecial xlPasteValues
    Windows("Állásidõ adott idõszakban.xlsx").Activate
    ActiveWindow.Close
    Range("A1").Select

End Sub
