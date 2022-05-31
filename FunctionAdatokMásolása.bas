Attribute VB_Name = "FunctionAdatokMásolása"
Option Explicit

Sub AdatokMásolása()
'JelszóRejtés2

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
Ckoord = Coszlop & Crw
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
Dkoord = Doszlop & Drw
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
Ekoord = Eoszlop & Erw
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
Fkoord = Foszlop & Frw
Range(Fkoord).Value = "GÉP"

' - Kulcs "G:G" - '

Sheets("adatok").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.Row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Grw
Range(Gkoord).Value = "KULCS"

' - Terület   "H:H" - '

Sheets("adatok").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.Row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw
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
Ikoord = Ioszlop & Irw
Range(Ikoord).Value = AppWindow.ComboBox2.Value

' - -tól   "J:J" - '

Sheets("adatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Jrw As Long
Jrw = ActiveCell.Row + 1
Dim Joszlop As String
Joszlop = "j"
Dim Jkoord As String
Jkoord = Joszlop & Jrw
Range(Jkoord).Value = AppWindow.TextBox7.Value

' - -ig "K:K" - '

Sheets("adatok").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Krw As Long
Krw = ActiveCell.Row + 1
Dim Koszlop As String
Koszlop = "k"
Dim Kkoord As String
Kkoord = Koszlop & Krw
Range(Kkoord).Value = AppWindow.TextBox6.Value

' - Idõ "L:L" - '

'Sheets("adatok").Select
'Columns("l:l").Select
'Selection.End(xlDown).Select
'Dim Lrw As Long
'Lrw = ActiveCell.row + 1
'Dim Loszlop As String
'Loszlop = "l"
'Dim Lkoord As String
'Lkoord = Loszlop & Lrw
'Range(Lkoord).Value = idõ

' - Mûszak   "M:M" - '

Sheets("adatok").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim Mrw As Long
Mrw = ActiveCell.Row + 1
Dim Moszlop As String
Moszlop = "m"
Dim Mkoord As String
Mkoord = Moszlop & Mrw
Range(Mkoord).Value = "MÛSZAK"

' - Probléma "N:N" - '

Sheets("adatok").Select
Columns("n:n").Select
Selection.End(xlDown).Select
Dim Nrw As Long
Nrw = ActiveCell.Row + 1
Dim Noszlop As String
Noszlop = "n"
Dim Nkoord As String
Nkoord = Noszlop & Nrw
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
Okoord = Ooszlop & Orw
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
Pkoord = Poszlop & Prw
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
Qkoord = Qoszlop & Qrw
Range(Qkoord).Value = AppWindow.ComboBox3.Value

' - Felelõs   "R:R" - '

Sheets("adatok").Select
Columns("r:r").Select
Selection.End(xlDown).Select
Dim Rrw As Long
Rrw = ActiveCell.Row + 1
Dim Roszlop As String
Roszlop = "r"
Dim Rkoord As String
Rkoord = Roszlop & Rrw
Range(Rkoord).Value = "vatta"

'- Becsültdátum   "S:S" - '

Sheets("adatok").Select
Columns("s:s").Select
Selection.End(xlDown).Select
Dim Srw As Long
Srw = ActiveCell.Row + 1
Dim Soszlop As String
Soszlop = "s"
Dim Skoord As String
Skoord = Soszlop & Srw
Range(Skoord).Value = "vatta"

' - Visszaigazoltdátum   "T:T" - '

Sheets("adatok").Select
Columns("t:t").Select
Selection.End(xlDown).Select
Dim Trw As Long
Trw = ActiveCell.Row + 1
Dim Toszlop As String
Toszlop = "t"
Dim Tkoord As String
Tkoord = Toszlop & Trw
Range(Tkoord).Value = "vatta"

' - Visszaadásidátum   "U:U" - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim Urw As Long
Urw = ActiveCell.Row + 1
Dim Uoszlop As String
Uoszlop = "u"
Dim Ukoord As String
Ukoord = Uoszlop & Urw
Range(Ukoord).Value = "vatta"

' - Megjegyzés   "V:V" - '

Sheets("adatok").Select
Columns("v:v").Select
Selection.End(xlDown).Select
Dim Vrw As Long
Vrw = ActiveCell.Row + 1
Dim Voszlop As String
Voszlop = "v"
Dim Vkoord As String
Vkoord = Voszlop & Vrw
If AppWindow.TextBox78 = "" Then
Range(Vkoord).Value = " n/a "
Else
Range(Vkoord).Value = AppWindow.TextBox78.Value
End If

' - Megjegyzés   "X:X" - '

Sheets("adatok").Select
Columns("x:x").Select
Selection.End(xlDown).Select
Dim Xrw As Long
Xrw = ActiveCell.Row + 1
Dim Xoszlop As String
Xoszlop = "x"
Dim Xkoord As String
Xkoord = Xoszlop & Xrw
Range(Xkoord).Value = AppWindow.ComboBox8.Value
    
    
Sheets("Start").Select
Range("b2").Select

'JelszóRejtés
End Sub
