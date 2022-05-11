Attribute VB_Name = "FunctionAdatokMásolása2"
Sub AdatokMásolása2()

' - Bárcaszám   "B:B" - '

Sheets("adatok").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = AppWindow.TextBox54

' - Dátum   "C:C" - '

Sheets("adatok").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim Crw As Long
Crw = ActiveCell.row + 1
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
Drw = ActiveCell.row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Drw
Range(Dkoord).Value = AppWindow.TextBox63

' - RÁBAszám   "E:E" - '

Sheets("adatok").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Erw
Range(Ekoord).Value = AppWindow.TextBox64

' - Gép "F:F" - '

Sheets("adatok").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Frw
Range(Fkoord).Value = AppWindow.TextBox65

' - Kulcs "G:G" - '

Sheets("adatok").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Grw
Range(Gkoord).Value = AppWindow.TextBox66

' - Terület   "H:H" - '

Sheets("adatok").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw
Range(Hkoord).Value = AppWindow.TextBox67

' - Csapat   "I:I" - '

Sheets("adatok").Select
Columns("i:i").Select
Selection.End(xlDown).Select
Dim Irw As Long
Irw = ActiveCell.row + 1
Dim Ioszlop As String
Ioszlop = "i"
Dim Ikoord As String
Ikoord = Ioszlop & Irw
Range(Ikoord).Value = AppWindow.TextBox68

' - Mûszak   "M:M" - '

Sheets("adatok").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim Mrw As Long
Mrw = ActiveCell.row + 1
Dim Moszlop As String
Moszlop = "m"
Dim Mkoord As String
Mkoord = Moszlop & Mrw
Range(Mkoord).Value = AppWindow.TextBox69

' - -tól   "J:J" - '

Sheets("adatok").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Jrw As Long
Jrw = ActiveCell.row + 1
Dim Joszlop As String
Joszlop = "j"
Dim Jkoord As String
Jkoord = Joszlop & Jrw
Range(Jkoord).Value = AppWindow.TextBox70

' - -ig "K:K" - '

Sheets("adatok").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Krw As Long
Krw = ActiveCell.row + 1
Dim Koszlop As String
Koszlop = "k"
Dim Kkoord As String
Kkoord = Koszlop & Krw
Range(Kkoord).Value = AppWindow.TextBox49

' - Idõ "L:L" - '

Sheets("adatok").Select
Columns("l:l").Select
Selection.End(xlDown).Select
Dim Lrw As Long
Lrw = ActiveCell.row + 1
Dim Loszlop As String
Loszlop = "l"
Dim Lkoord As String
Lkoord = Loszlop & Lrw
Range(Lkoord).Value = AppWindow.TextBox71

' - Probléma "N:N" - '

Sheets("adatok").Select
Columns("n:n").Select
Selection.End(xlDown).Select
Dim Nrw As Long
Nrw = ActiveCell.row + 1
Dim Noszlop As String
Noszlop = "n"
Dim Nkoord As String
Nkoord = Noszlop & Nrw
Range(Nkoord).Value = AppWindow.TextBox72

' - Megoldás   "O:O" - '

Sheets("adatok").Select
Columns("o:o").Select
Selection.End(xlDown).Select
Dim Orw As Long
Orw = ActiveCell.row + 1
Dim Ooszlop As String
Ooszlop = "o"
Dim Okoord As String
Okoord = Ooszlop & Orw
Range(Okoord).Value = AppWindow.TextBox57

' - Státusz   "P:P" - '

Sheets("adatok").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim Prw As Long
Prw = ActiveCell.row + 1
Dim Poszlop As String
Poszlop = "p"
Dim Pkoord As String
Pkoord = Poszlop & Prw
Range(Pkoord).Value = AppWindow.TextBox56

' - Mérés   "Q:Q" - '

Sheets("adatok").Select
Columns("q:q").Select
Selection.End(xlDown).Select
Dim Qrw As Long
Qrw = ActiveCell.row + 1
Dim Qoszlop As String
Qoszlop = "q"
Dim Qkoord As String
Qkoord = Qoszlop & Qrw
Range(Qkoord).Value = AppWindow.TextBox55

' - Felelõs   "R:R" - '

Sheets("adatok").Select
Columns("r:r").Select
Selection.End(xlDown).Select
Dim Rrw As Long
Rrw = ActiveCell.row + 1
Dim Roszlop As String
Roszlop = "r"
Dim Rkoord As String
Rkoord = Roszlop & Rrw
Range(Rkoord).Value = AppWindow.TextBox58

' - Becsültdátum   "S:S" - '

Sheets("adatok").Select
Columns("s:s").Select
Selection.End(xlDown).Select
Dim Srw As Long
Srw = ActiveCell.row + 1
Dim Soszlop As String
Soszlop = "s"
Dim Skoord As String
Skoord = Soszlop & Srw
Range(Skoord).Value = AppWindow.TextBox59

' - Visszaigazoltdátum   "T:T" - '

Sheets("adatok").Select
Columns("t:t").Select
Selection.End(xlDown).Select
Dim Trw As Long
Trw = ActiveCell.row + 1
Dim Toszlop As String
Toszlop = "t"
Dim Tkoord As String
Tkoord = Toszlop & Trw
Range(Tkoord).Value = AppWindow.TextBox60

' - Visszaadásidátum   "U:U" - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim Urw As Long
Urw = ActiveCell.row + 1
Dim Uoszlop As String
Uoszlop = "u"
Dim Ukoord As String
Ukoord = Uoszlop & Urw
Range(Ukoord).Value = AppWindow.TextBox61

End Sub

