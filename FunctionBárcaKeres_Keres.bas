Attribute VB_Name = "FunctionB�rcaKeres_Keres"
Option Explicit

Sub B�rcaKeres_Keres()

Munka15.Select
Range("a1:i1000").Clear

' - m�sol - '

Munka3.Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim BKKrw As Integer
BKKrw = ActiveCell.Row
Dim BKKig As String
BKKig = "d"
Dim BKKkoord As String
BKKkoord = BKKig & BKKrw
Range("a1", BKKkoord).Copy

' - odailleszt - '

Munka15.Select
Range("a1").PasteSpecial xlPasteValues
Range("a1").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim BKKrw2 As Integer
BKKrw2 = ActiveCell.Row
Dim BKKig2 As String
BKKig2 = "a"
Dim BKKkoord2 As String
BKKkoord2 = BKKig2 & BKKrw2

' - sz�r - '

Selection.AutoFilter
ActiveSheet.Range("a1", BKKkoord2).AutoFilter Field:=1, Criteria1:="*" & AppWindow.TextBox106.Value & "*"

Munka15.Select
Range("a1").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim BKKrw3 As Long
BKKrw3 = ActiveCell.Row
Dim kil�p� As Long
kil�p� = BKKrw3
If kil�p� > 10000 Then
MsgBox "Nincs tal�lat."
AppWindow.TextBox106 = ""
Exit Sub
Else
Dim BKKig3 As String
BKKig3 = "d"
Dim BKKkoord3 As String
BKKkoord3 = BKKig3 & BKKrw3
Range("a1", BKKkoord).Select
Selection.Copy
Range("f1").PasteSpecial xlPasteValues

Selection.AutoFilter
End If
' - sz�rt adatot vissza ad - '
Columns("i:i").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "i"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Dim rngList As Range
Set rngList = Munka15.Range("f2", ALkoord)
AppWindow.ListBox35.List = rngList.Value

AppWindow.TextBox106 = ""
Munka15.Select
Range("a1:i1000").Clear

End Sub
