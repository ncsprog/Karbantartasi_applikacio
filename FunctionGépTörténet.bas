Attribute VB_Name = "FunctionG�pT�rt�net"
Option Explicit

Sub G�pT�rt�net()

Sheets("sz�r�_transfer").Select
Columns("a:u").Select
Selection = ""

Dim krit�rium As String
krit�rium = AppWindow.TextBox73.Value

Munka1.Range("a1").AutoFilter 5, krit�rium

' - Lista koordin�ta - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "u"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Dim ter�let As Range
Set ter�let = Munka1.Range("a1", ALkoord)
ter�let.Copy
Munka6.Range("a1").PasteSpecial xlPasteValues

Sheets("adatok").Select
ActiveSheet.AutoFilterMode = False              't�rli a sz�r�t
Application.CutCopyMode = False                 'a m�sol�s kijel�l�lst t�rli

Sheets("sz�r�_transfer").Select
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
