Attribute VB_Name = "FunctionAdatokMásolása"
Option Explicit

Sub AdatokMásolása()

Dim T1rw As Integer, T1koord As String, _
T2rw As Integer, T2koord As String, T_tõl As Integer, M_tõl As Integer, T_ig As Integer, _
M_ig As Integer, Dh As Integer, Dm As String, T3rw As Integer, T3koord As String, H As Integer, _
M As Integer, H_24 As Integer, M_60 As Integer, Csek1 As Integer, Csek2 As Integer, _
D As Integer, Mûszak As String, De As String, Du As String, Éj As String

' - Bárcaszám   "B:B" - '

Sheets("adatok").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.Row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = AppWindow.TextBox11.Value

' - Dátum   "C:C" - '

Sheets("adatok").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim Crw As Long
Crw = ActiveCell.Row + 1
Dim Coszlop As String
Coszlop = "c"
Dim Ckoord As String
Ckoord = Coszlop & Brw
Range(Ckoord).Value = Date


' - Munkaszám   "D:D" - '

Sheets("adatok").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim Drw As Long
Drw = ActiveCell.Row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Brw
Range(Dkoord).Value = AppWindow.TextBox1.Value

' - RÁBAszám   "E:E" - '

Sheets("adatok").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.Row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Brw
Range(Ekoord).Value = AppWindow.TextBox10.Value
 '- Gép   "F:F" - '

Sheets("adatok").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.Row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Brw

' - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - '

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
ActiveSheet.Range("a1", BKKkoord2).AutoFilter Field:=1, Criteria1:=AppWindow.TextBox10.Value
' - visszadás - '

Munka15.Select
Columns("c:c").Select
Selection.End(xlDown).Copy
Munka1.Select
Range(Fkoord).PasteSpecial xlPasteValues


' - Kulcs "G:G" - kulcs keresés - kulcs keresés - kulcs keresés - kulcs keresés - kulcs keresés - kulcs keresés - kulcs keresés - '

Sheets("adatok").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.Row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Brw

Munka15.Select
Columns("b:b").Select
Selection.End(xlDown).Copy
Munka1.Select
Range(Gkoord).PasteSpecial xlPasteValues

Munka15.Select
Selection.AutoFilter
Range("a1:i3000").Clear

' - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - gépkeresése - '

' - Terület   "H:H" - '

Sheets("adatok").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.Row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Brw
Range(Hkoord).Value = AppWindow.ComboBox1.Value
' - Csapat   "I:I" - '

Sheets("adatok").Select
Columns("i:i").Select
Selection.End(xlDown).Select
Dim Irw As Long
Irw = ActiveCell.Row + 1
Dim Ioszlop As String
Ioszlop = "i"
Dim Ikoord As String
Ikoord = Ioszlop & Brw
Range(Ikoord).Value = AppWindow.ComboBox2.Value

' - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - '

'idõ_tól

Munka1.Select
Columns("j:j").Select
Selection.End(xlDown).Select
'T1rw = ActiveCell.Row + 1
T1koord = "j" & Brw
Range(T1koord) = AppWindow.TextBox7

'idõ_ig

Munka1.Select
Columns("k:k").Select
Selection.End(xlDown).Select
'T2rw = ActiveCell.Row + 1
T2koord = "k" & Brw
Range(T2koord) = AppWindow.TextBox6

'Csekkolás

Csek1 = Len(AppWindow.TextBox7)
Csek2 = Len(AppWindow.TextBox6)

If Csek1 <> 5 Then
MsgBox "Kezdõ idõpont formátuma nem megfelelõ! óó:pp"
Exit Sub
End If
If Csek2 <> 5 Then
MsgBox "Befejezõ idõpont formátuma nem megfelelõ! óó:pp"
Exit Sub
End If


'Tól-óra, perc

T_tõl = Left(AppWindow.TextBox7, 2)
M_tõl = Right(AppWindow.TextBox7, 2)

'Ig-óra, perc

T_ig = Left(AppWindow.TextBox6, 2)
M_ig = Right(AppWindow.TextBox6, 2)

'Delta-óra, perc

'D = M_ig - M_tõl

'Óra, perc

H = 24
M = 60

'DeltaT

Munka1.Select
Columns("l:l").Select
Selection.End(xlDown).Select
'T3rw = ActiveCell.Row + 1
T3koord = "l" & Brw

'Számítás_óra

If T_tõl = T_ig Then
Dh = T_ig - T_tõl
ElseIf T_tõl > T_ig Then
    If M_tõl = 0 Then
    Dh = H - T_tõl + T_ig
    Else
    Dh = H - T_tõl + T_ig - 1
    End If
ElseIf T_ig > T_tõl Then
Dh = T_ig - T_tõl
End If

'Számítás_perc

If M_tõl = M_ig Then
D = M_ig - M_tõl
ElseIf M_tõl > M_ig Then
D = M - M_tõl + M_ig
ElseIf M_ig > M_tõl Then
D = M_ig - M_tõl
End If

If D < 10 Then
Dm = "0" & D
Else
Dm = D
End If

Range(T3koord) = Dh & ":" & Dm & " óra"

' - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - idõértékek - '

' - Probléma "N:N" - '

Sheets("adatok").Select
Columns("n:n").Select
Selection.End(xlDown).Select
Dim Nrw As Long
Nrw = ActiveCell.Row + 1
Dim Noszlop As String
Noszlop = "n"
Dim Nkoord As String
Nkoord = Noszlop & Brw
Range(Nkoord).Value = AppWindow.TextBox5.Value

' - Megoldás   "O:O" - '

Sheets("adatok").Select
Columns("o:o").Select
Selection.End(xlDown).Select
Dim Orw As Long
Orw = ActiveCell.Row + 1
Dim Ooszlop As String
Ooszlop = "o"
Dim Okoord As String
Okoord = Ooszlop & Brw
Range(Okoord).Value = AppWindow.TextBox4.Value

' - Státusz   "P:P" - '

Sheets("adatok").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim Prw As Long
Prw = ActiveCell.Row + 1
Dim Poszlop As String
Poszlop = "p"
Dim Pkoord As String
Pkoord = Poszlop & Brw
Range(Pkoord).Value = AppWindow.ComboBox4.Value
' - Mérés   "Q:Q" - '

Sheets("adatok").Select
Columns("q:q").Select
Selection.End(xlDown).Select
Dim Qrw As Long
Qrw = ActiveCell.Row + 1
Dim Qoszlop As String
Qoszlop = "q"
Dim Qkoord As String
Qkoord = Qoszlop & Brw
Range(Qkoord).Value = AppWindow.ComboBox3.Value

' - Megjegyzés   "V:V" - '

Sheets("adatok").Select
Columns("v:v").Select
Selection.End(xlDown).Select
Dim Vrw As Long
Vrw = ActiveCell.Row + 1
Dim Voszlop As String
Voszlop = "v"
Dim Vkoord As String
Vkoord = Voszlop & Brw
Range(Vkoord).Value = AppWindow.TextBox78.Value

' - Megjegyzés   "X:X" - '

Sheets("adatok").Select
Columns("x:x").Select
Selection.End(xlDown).Select
Dim Xrw As Long
Xrw = ActiveCell.Row + 1
Dim Xoszlop As String
Xoszlop = "x"
Dim Xkoord As String
Xkoord = Xoszlop & Brw
Range(Xkoord).Value = AppWindow.ComboBox8.Value
    
    
Sheets("Start").Select
Range("b2").Select

End Sub
