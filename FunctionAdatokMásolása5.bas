Attribute VB_Name = "FunctionAdatokMásolása5"
Option Explicit

Sub AdatokMásolása5()
' - Dátum   "B:B" - '

Sheets("létszám").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.Row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = Date


' - Mérnök - Team 1 "C:C" - '

Sheets("létszám").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim Crw As Long
Crw = ActiveCell.Row + 1
Dim Coszlop As String
Coszlop = "c"
Dim Ckoord As String
Ckoord = Coszlop & Crw
Range(Ckoord).Value = AppWindow.TextBox113.Value


' - Lakatos - Team 1 "D:D" - '

Sheets("létszám").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim Drw As Long
Drw = ActiveCell.Row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Drw
Range(Dkoord).Value = AppWindow.TextBox114.Value

' - Villanyszerelõ - Team 1 "E:E" - '

Sheets("létszám").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.Row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Erw
Range(Ekoord).Value = AppWindow.TextBox115.Value

' - Mérnök - Team 2 "F: F " - '"

Sheets("létszám").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.Row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Frw
Range(Fkoord).Value = AppWindow.TextBox117.Value

' - Lakatos - Team 2 "G:G" - '

Sheets("létszám").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.Row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Grw
Range(Gkoord).Value = AppWindow.TextBox118.Value

' - Villanyszerelõ - Team 2 "H:H" - '

Sheets("létszám").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.Row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw
Range(Hkoord).Value = AppWindow.TextBox119.Value

' - Mérnök - Team 3 "I:I" - '

Sheets("létszám").Select
Columns("i:i").Select
Selection.End(xlDown).Select
Dim Irw As Long
Irw = ActiveCell.Row + 1
Dim Ioszlop As String
Ioszlop = "i"
Dim Ikoord As String
Ikoord = Ioszlop & Irw
Range(Ikoord).Value = AppWindow.TextBox121.Value

' - Lakatos - Team 3 "J:J" - '

Sheets("létszám").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Jrw As Long
Jrw = ActiveCell.Row + 1
Dim Joszlop As String
Joszlop = "j"
Dim Jkoord As String
Jkoord = Joszlop & Jrw
Range(Jkoord).Value = AppWindow.TextBox122.Value

' - Villanyszerelõ - Team 3 "K:K" - '

Sheets("létszám").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Krw As Long
Krw = ActiveCell.Row + 1
Dim Koszlop As String
Koszlop = "k"
Dim Kkoord As String
Kkoord = Koszlop & Krw
Range(Kkoord).Value = AppWindow.TextBox123.Value


End Sub
