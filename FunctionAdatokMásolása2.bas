Attribute VB_Name = "FunctionAdatokM?sol?sa2"
Sub AdatokM?sol?sa2()
' - B?rcasz?m   "B:B" - '

Sheets("adatok").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.Row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = AppWindow.TextBox54.Value

' - D?tum   "C:C" - '

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

' - Munkasz?m   "D:D" - '

Sheets("adatok").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim Drw As Long
Drw = ActiveCell.Row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Drw
Range(Dkoord).Value = AppWindow.TextBox63.Value

' - R?BAsz?m   "E:E" - '

Sheets("adatok").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.Row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Erw
Range(Ekoord).Value = AppWindow.TextBox64.Value

' - G?p "F:F" - '

Sheets("adatok").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.Row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Frw
Range(Fkoord).Value = AppWindow.TextBox65.Value

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
Range(Gkoord).Value = AppWindow.TextBox66.Value

' - Ter?let   "H:H" - '

Sheets("adatok").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.Row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw
Range(Hkoord).Value = AppWindow.TextBox67.Value

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
Range(Ikoord).Value = AppWindow.TextBox68.Value

' - M?szak   "M:M" - '

Sheets("adatok").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim Mrw As Long
Mrw = ActiveCell.Row + 1
Dim Moszlop As String
Moszlop = "m"
Dim Mkoord As String
Mkoord = Moszlop & Mrw
Range(Mkoord).Value = AppWindow.TextBox69.Value

Id?Kalkul?tor2

' - Probl?ma "N:N" - '

Sheets("adatok").Select
Columns("n:n").Select
Selection.End(xlDown).Select
Dim Nrw As Long
Nrw = ActiveCell.Row + 1
Dim Noszlop As String
Noszlop = "n"
Dim Nkoord As String
Nkoord = Noszlop & Nrw
Range(Nkoord).Value = AppWindow.TextBox72.Value

' - Megold?s   "O:O" - '

Sheets("adatok").Select
Columns("o:o").Select
Selection.End(xlDown).Select
Dim Orw As Long
Orw = ActiveCell.Row + 1
Dim Ooszlop As String
Ooszlop = "o"
Dim Okoord As String
Okoord = Ooszlop & Orw
Range(Okoord).Value = AppWindow.TextBox57.Value

' - St?tusz   "P:P" - '

Sheets("adatok").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim Prw As Long
Prw = ActiveCell.Row + 1
Dim Poszlop As String
Poszlop = "p"
Dim Pkoord As String
Pkoord = Poszlop & Prw
Range(Pkoord).Value = AppWindow.ComboBox5.Value

' - M?r?s   "Q:Q" - '

Sheets("adatok").Select
Columns("q:q").Select
Selection.End(xlDown).Select
Dim Qrw As Long
Qrw = ActiveCell.Row + 1
Dim Qoszlop As String
Qoszlop = "q"
Dim Qkoord As String
Qkoord = Qoszlop & Qrw
Range(Qkoord).Value = AppWindow.ComboBox6.Value

' - Felel?s   "R:R" - '

Sheets("adatok").Select
Columns("r:r").Select
Selection.End(xlDown).Select
Dim Rrw As Long
Rrw = ActiveCell.Row + 1
Dim Roszlop As String
Roszlop = "r"
Dim Rkoord As String
Rkoord = Roszlop & Rrw
Range(Rkoord).Value = AppWindow.ComboBox7.Value

' - Becs?ltd?tum   "S:S" - '

Sheets("adatok").Select
Columns("s:s").Select
Selection.End(xlDown).Select
Dim Srw As Long
Srw = ActiveCell.Row + 1
Dim Soszlop As String
Soszlop = "s"
Dim Skoord As String
Skoord = Soszlop & Srw
Range(Skoord).Value = AppWindow.TextBox59.Value

' - Visszaigazoltd?tum   "T:T" - '

Sheets("adatok").Select
Columns("t:t").Select
Selection.End(xlDown).Select
Dim Trw As Long
Trw = ActiveCell.Row + 1
Dim Toszlop As String
Toszlop = "t"
Dim Tkoord As String
Tkoord = Toszlop & Trw
Range(Tkoord).Value = AppWindow.TextBox60.Value

' - Visszaad?sid?tum   "U:U" - '

Sheets("adatok").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim Urw As Long
Urw = ActiveCell.Row + 1
Dim Uoszlop As String
Uoszlop = "u"
Dim Ukoord As String
Ukoord = Uoszlop & Urw
Range(Ukoord).Value = AppWindow.TextBox61.Value

' - Megjegyz?s   "V:V" - '

Sheets("adatok").Select
Columns("v:v").Select
Selection.End(xlDown).Select
Dim Vrw As Long
Vrw = ActiveCell.Row + 1
Dim Voszlop As String
Voszlop = "v"
Dim Vkoord As String
Vkoord = Voszlop & Vrw
If AppWindow.TextBox79 = "" Then
Range(Vkoord).Value = " n/a "
Else
Range(Vkoord).Value = AppWindow.TextBox79.Value
End If

Sheets("Start").Select
Range("b2").Select

End Sub

