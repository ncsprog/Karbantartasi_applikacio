Attribute VB_Name = "FunctionAdatfelvételLista2"
Option Explicit

Sub AdatfelvételLista2()
'JelszóRejtés2

Sheets("szûrõ_transfer").Select
Columns("a:v").Select
Selection = ""

Munka1.Range("a1").AutoFilter 16, AppWindow.ComboBox5.Value

' - Lista koordináta - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
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
Trw = ActiveCell.row
Dim Toszlop As String
Toszlop = "u"
Dim Tkoord As String
Tkoord = Toszlop & Trw


Dim rngList As Range
Set rngList = Munka6.Range("a1", Tkoord)
AppWindow.ListBox20.List = rngList.Value

Sheets("adatok").Select
Range("a1").Select

Sheets("Start").Select
Range("b2").Select

'JelszóRejtés
End Sub
