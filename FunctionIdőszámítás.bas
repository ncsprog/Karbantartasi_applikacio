Attribute VB_Name = "FunctionId�sz�m�t�s"
Option Explicit

Sub Id�sz�m�t�s()

' - -t�l koordin�ta - '
Sheets("adatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim T�lrw As Long
T�lrw = ActiveCell.Row
Dim T�loszlop As String
T�loszlop = "j"
Dim T�lkoord As String
T�lkoord = T�loszlop & T�lrw
Dim t�l As String
t�l = Range(T�lkoord).Value

' - -ig koordin�ta - '

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

' - hova koordin�ta - '

Sheets("adatok").Select
Columns("l:l").Select
Selection.End(xlDown).Select
Dim Lrw As Long
Lrw = ActiveCell.Row + 1
Dim Loszlop As String
Loszlop = "l"
Dim Lkoord As String
Lkoord = Loszlop & Lrw
'Range(Lkoord) = ig -t�l

If t�l > ig Then
Range(Lkoord) = ig - t�l + Munka1.Range("x1")
Else
Range(Lkoord) = ig - t�l
End If



Sheets("Start").Select
Range("b2").Select

End Sub
