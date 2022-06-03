Attribute VB_Name = "FunctionBárcaKeres_Keres"
Option Explicit

Sub BárcaKeres_Keres()

Munka15.Select
Range("a1:i1000").Clear

' - másol - '

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

' - szûr - '

Selection.AutoFilter
ActiveSheet.Range("a1", BKKkoord2).AutoFilter Field:=1, Criteria1:="*" & AppWindow.TextBox106.Value & "*"

Munka15.Select
Range("a1").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim BKKrw3 As Long
BKKrw3 = ActiveCell.Row
Dim kilépõ As Long
kilépõ = BKKrw3
If kilépõ > 10000 Then
MsgBox "Nincs találat."
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
' - szûrt adatot vissza ad - '
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
