Attribute VB_Name = "FunctionB�rcaKeres_Keres"
Option Explicit

Sub B�rcaKeres_Keres()

'Filter a tb106 �rt�k�t a n�v oszlopba, az eredm�nyt a lb35-be l�ki vissza �ltal
Munka15.Select
Range("a1:i1000").Clear

' - m�sol - '

Munka3.Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim BKKrw As Integer
BKKrw = ActiveCell.row
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
BKKrw2 = ActiveCell.row
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
Dim BKKrw3 As Integer
BKKrw3 = ActiveCell.row
Dim BKKig3 As String
BKKig3 = "d"
Dim BKKkoord3 As String
BKKkoord3 = BKKig3 & BKKrw3
Range("a1", BKKkoord).Select
Selection.Copy
'Application.CutCopyMode = False
'Range("a1:xx10000") = ""
Range("f1").PasteSpecial xlPasteValues

' - sz�rt adatot vissza ad - '
Columns("i:i").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.row
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
