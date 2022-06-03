Attribute VB_Name = "FunctionIdõszámítás"
Option Explicit

Sub Idõszámítás()

' - -tól koordináta - '
Sheets("adatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Tólrw As Long
Tólrw = ActiveCell.Row
Dim Tóloszlop As String
Tóloszlop = "j"
Dim Tólkoord As String
Tólkoord = Tóloszlop & Tólrw
Dim tól As String
tól = Range(Tólkoord).Value

' - -ig koordináta - '

Sheets("adatok").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Igrw As Long
Igrw = ActiveCell.Row
Dim Igoszlop As String
Igoszlop = "k"
Dim Igkoord As String
Igkoord = Igoszlop & Igrw
Dim ig As String
ig = Range(Igkoord).Value

' - hova koordináta - '

Sheets("adatok").Select
Columns("l:l").Select
Selection.End(xlDown).Select
Dim Lrw As Long
Lrw = ActiveCell.Row + 1
Dim Loszlop As String
Loszlop = "l"
Dim Lkoord As String
Lkoord = Loszlop & Lrw
'Range(Lkoord) = ig -tól

If tól > ig Then
Range(Lkoord) = ig - tól + Munka1.Range("x1")
Else
Range(Lkoord) = ig - tól
End If



Sheets("Start").Select
Range("b2").Select

End Sub
