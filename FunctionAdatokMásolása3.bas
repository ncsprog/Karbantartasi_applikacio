Attribute VB_Name = "FunctionAdatokMásolása3"
Option Explicit

Sub AdatokMásolása3()

' - Dátum   "B:B" - '

Sheets("Megbeszélés").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = Date


' - Létszám - Futómû   "C:C" - '

Sheets("Megbeszélés").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim Crw As Long
Crw = ActiveCell.row + 1
Dim Coszlop As String
Coszlop = "c"
Dim Ckoord As String
Ckoord = Coszlop & Crw
Range(Ckoord).Value = AppWindow.TextBox81.Value


' - Létszám - Hídtest   "D:D" - '

Sheets("Megbeszélés").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim Drw As Long
Drw = ActiveCell.row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Drw
Range(Dkoord).Value = AppWindow.TextBox83.Value

' - Létszám - Kovács   "E:E" - '

Sheets("Megbeszélés").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Erw
Range(Ekoord).Value = AppWindow.TextBox85.Value

' - Délelõtt - Futómû "F:F" - '

Sheets("Megbeszélés").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Frw
Range(Fkoord).Value = AppWindow.TextBox82.Value

' - Délelõtt - Hídtest "G:G" - '

Sheets("Megbeszélés").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Grw
Range(Gkoord).Value = AppWindow.TextBox84.Value

' - Délelõtt - Kovács   "H:H" - '

Sheets("Megbeszélés").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw
Range(Hkoord).Value = AppWindow.TextBox86.Value

' - Délután - Futómû   "I:I" - '

Sheets("Megbeszélés").Select
Columns("i:i").Select
Selection.End(xlDown).Select
Dim Irw As Long
Irw = ActiveCell.row + 1
Dim Ioszlop As String
Ioszlop = "i"
Dim Ikoord As String
Ikoord = Ioszlop & Irw
Range(Ikoord).Value = AppWindow.TextBox88.Value

' - Délután - Hídtest   "J:J" - '

Sheets("Megbeszélés").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Jrw As Long
Jrw = ActiveCell.row + 1
Dim Joszlop As String
Joszlop = "j"
Dim Jkoord As String
Jkoord = Joszlop & Jrw
Range(Jkoord).Value = AppWindow.TextBox90.Value

' - Délután - Hídtest   "K:K" - '

Sheets("Megbeszélés").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Krw As Long
Krw = ActiveCell.row + 1
Dim Koszlop As String
Koszlop = "k"
Dim Kkoord As String
Kkoord = Koszlop & Krw
Range(Kkoord).Value = AppWindow.TextBox92.Value




Sheets("Start").Select
Range("b2").Select
End Sub
