Attribute VB_Name = "FunctionAdatokM�sol�sa3"
Option Explicit

Sub AdatokM�sol�sa3()
'Jelsz�Rejt�s2

' - D�tum   "B:B" - '

Sheets("Megbesz�l�s").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = Date


' - D�lel�tt - Forr�s: Team1 "C:C" - '

Sheets("Megbesz�l�s").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim Crw As Long
Crw = ActiveCell.row + 1
Dim Coszlop As String
Coszlop = "c"
Dim Ckoord As String
Ckoord = Coszlop & Crw
Range(Ckoord).Value = AppWindow.ListBox40.Value


' - D�lel�tt - Jegyzet: Team1 "D:D" - '

Sheets("Megbesz�l�s").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim Drw As Long
Drw = ActiveCell.row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Drw
Range(Dkoord).Value = AppWindow.TextBox111.Value

' - D�lel�tt - Forr�s: Team2 "E:E" - '

Sheets("Megbesz�l�s").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Erw
Range(Ekoord).Value = AppWindow.ListBox41.Value

' - D�lel�tt - Jegyzet: Team2 "F:F" - '

Sheets("Megbesz�l�s").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Frw
Range(Fkoord).Value = AppWindow.TextBox116.Value

' - D�lel�tt - Forr�s: Team3 "G:G" - '

Sheets("Megbesz�l�s").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Grw
Range(Gkoord).Value = AppWindow.ListBox42.Value

' - D�lel�tt - Jegyzet: Team3 "H:H" - '

Sheets("Megbesz�l�s").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw
Range(Hkoord).Value = AppWindow.TextBox120.Value

' - D�lut�n - Forr�s: Team1 "I:I" - '

Sheets("Megbesz�l�s").Select
Columns("i:i").Select
Selection.End(xlDown).Select
Dim Irw As Long
Irw = ActiveCell.row + 1
Dim Ioszlop As String
Ioszlop = "i"
Dim Ikoord As String
Ikoord = Ioszlop & Irw
Range(Ikoord).Value = AppWindow.ListBox43.Value

' - D�lut�n - Jegyzet: Team1 "J:J" - '

Sheets("Megbesz�l�s").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Jrw As Long
Jrw = ActiveCell.row + 1
Dim Joszlop As String
Joszlop = "j"
Dim Jkoord As String
Jkoord = Joszlop & Jrw
Range(Jkoord).Value = AppWindow.TextBox124.Value

' - D�lut�n - Forr�s: Team2 "K:K" - '

Sheets("Megbesz�l�s").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Krw As Long
Krw = ActiveCell.row + 1
Dim Koszlop As String
Koszlop = "k"
Dim Kkoord As String
Kkoord = Koszlop & Krw
Range(Kkoord).Value = AppWindow.ListBox44.Value

' - D�lut�n - Jegyzet: Team2 "L:L" - '

Sheets("Megbesz�l�s").Select
Columns("l:l").Select
Selection.End(xlDown).Select
Dim Lrw As Long
Lrw = ActiveCell.row + 1
Dim Loszlop As String
Loszlop = "l"
Dim Lkoord As String
Lkoord = Loszlop & Lrw
Range(Lkoord).Value = AppWindow.TextBox128.Value

' - D�lut�n - Forr�s: Team3 "M:M" - '

Sheets("Megbesz�l�s").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim Mrw As Long
Mrw = ActiveCell.row + 1
Dim Moszlop As String
Moszlop = "m"
Dim Mkoord As String
Mkoord = Moszlop & Mrw
Range(Mkoord).Value = AppWindow.ListBox45.Value

' - D�lut�n - Jegyzet: Team3 "N:N" - '

Sheets("Megbesz�l�s").Select
Columns("n:n").Select
Selection.End(xlDown).Select
Dim Nrw As Long
Nrw = ActiveCell.row + 1
Dim Noszlop As String
Noszlop = "n"
Dim Nkoord As String
Nkoord = Noszlop & Nrw
Range(Nkoord).Value = AppWindow.TextBox132.Value


Sheets("Start").Select
Range("b2").Select
'Jelsz�Rejt�s
End Sub
