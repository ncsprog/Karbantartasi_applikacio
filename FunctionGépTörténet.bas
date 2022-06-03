Attribute VB_Name = "FunctionGépTörténet"
Option Explicit

Sub GépTörténet()

Sheets("szûrõ_transfer").Select
Columns("a:u").Select
Selection = ""

Dim kritérium As String
kritérium = AppWindow.TextBox73.Value

Munka1.Range("a1").AutoFilter 5, kritérium

' - Lista koordináta - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "u"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Dim terület As Range
Set terület = Munka1.Range("a1", ALkoord)
terület.Copy
Munka6.Range("a1").PasteSpecial xlPasteValues

Sheets("adatok").Select
ActiveSheet.AutoFilterMode = False              'törli a szûrõt
Application.CutCopyMode = False                 'a másolás kijelölélst törli

Sheets("szûrõ_transfer").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim Trw As Long
Trw = ActiveCell.Row
Dim Toszlop As String
Toszlop = "u"
Dim Tkoord As String
Tkoord = Toszlop & Trw


Dim rngList As Range
Set rngList = Munka6.Range("a1", Tkoord)
AppWindow.ListBox22.List = rngList.Value

Sheets("adatok").Select
Range("a1").Select

Sheets("Start").Select
Range("b2").Select

End Sub
