Attribute VB_Name = "FunctionGépKeres2_Keres"
Option Explicit

Sub GépKeres2()
' - Keres - '

'Filter a tb107 értékét a név oszlopba, az eredményt a lb36-be löki vissza által
Munka15.Select
Range("a1:i3000").Clear

' - másol - '

Munka4.Select
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
' - meddig keressen - '
Range("a1").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim BKKrw2 As Integer
BKKrw2 = ActiveCell.Row
Dim BKKig2 As String
BKKig2 = "d"
Dim BKKkoord2 As String
BKKkoord2 = BKKig2 & BKKrw2

' - szûr - '

Selection.AutoFilter
ActiveSheet.Range("a1", BKKkoord2).AutoFilter Field:=4, Criteria1:="*" & AppWindow.TextBox107.Value & "*"
' - visszadás - '
Munka15.Select
Range("a1").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim BKKrw3 As Long
BKKrw3 = ActiveCell.Row
Dim kilépõ As Long
kilépõ = BKKrw3
If kilépõ > 10000 Then
MsgBox "Nincs találat."
AppWindow.TextBox107 = ""
Exit Sub
Else
Dim BKKig3 As String
BKKig3 = "c"
Dim BKKkoord3 As String
BKKkoord3 = BKKig3 & BKKrw3
Range("a1", BKKkoord).Select
Selection.Copy
'Application.CutCopyMode = False
'Range("a1:xx10000") = ""
Range("f1").PasteSpecial xlPasteValues

Selection.AutoFilter
End If
' - szûrt adatot vissza ad - '
Columns("h:h").Select
Selection.End(xlDown).Select
Dim ALrw As Long
ALrw = ActiveCell.Row
Dim ALoszlop As String
ALoszlop = "h"
Dim ALkoord As String
ALkoord = ALoszlop & ALrw

Dim rngList As Range
Set rngList = Munka15.Range("f2", ALkoord)
AppWindow.ListBox36.List = rngList.Value

AppWindow.TextBox107 = ""
Munka15.Select
Range("a1:i3000").Clear


End Sub
