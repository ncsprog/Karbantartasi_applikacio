Attribute VB_Name = "FunctionLétszámMásolás"
Option Explicit

Sub LétszámMásolás()

Sheets("Létszám").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Dátumellenõrzés As Date
Dátumellenõrzés = ActiveCell.Value

If Date <> Dátumellenõrzés Then

IDgenerálás2
ID_generálás2

Sheets("Létszám").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.row + 1
Dim Boszlop As String
Boszlop = "b"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = Date

                                        ' - Délelõtt - Mérnök - '

' - Team I. "C:C" - '

Sheets("Létszám").Select
Columns("c:c").Select
Selection.End(xlDown).Select
Dim Crw As Long
Crw = ActiveCell.row + 1
Dim Coszlop As String
Coszlop = "c"
Dim Ckoord As String
Ckoord = Coszlop & Crw

If AppWindow.TextBox41 = "" Then
Range(Ckoord).Value = "0"
Else
Range(Ckoord).Value = AppWindow.TextBox41.Value
End If



' - Team II. "D:D" - '

Sheets("Létszám").Select
Columns("d:d").Select
Selection.End(xlDown).Select
Dim Drw As Long
Drw = ActiveCell.row + 1
Dim Doszlop As String
Doszlop = "d"
Dim Dkoord As String
Dkoord = Doszlop & Drw

If AppWindow.TextBox38 = "" Then
Range(Dkoord).Value = "0"
Else
Range(Dkoord).Value = AppWindow.TextBox38.Value
End If

' - Team III. "E:E" - '

Sheets("Létszám").Select
Columns("e:e").Select
Selection.End(xlDown).Select
Dim Erw As Long
Erw = ActiveCell.row + 1
Dim Eoszlop As String
Eoszlop = "e"
Dim Ekoord As String
Ekoord = Eoszlop & Erw

If AppWindow.TextBox37 = "" Then
Range(Ekoord).Value = "0"
Else
Range(Ekoord).Value = AppWindow.TextBox37.Value
End If

                                        ' - Délelõtt - Lakatos - '

' - Team I. "F:F" - '

Sheets("Létszám").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Frw As Long
Frw = ActiveCell.row + 1
Dim Foszlop As String
Foszlop = "f"
Dim Fkoord As String
Fkoord = Foszlop & Frw

If AppWindow.TextBox40 = "" Then
Range(Fkoord).Value = "0"
Else
Range(Fkoord).Value = AppWindow.TextBox40.Value
End If

' - Team II. "G:G" - '

Sheets("Létszám").Select
Columns("g:g").Select
Selection.End(xlDown).Select
Dim Grw As Long
Grw = ActiveCell.row + 1
Dim Goszlop As String
Goszlop = "g"
Dim Gkoord As String
Gkoord = Goszlop & Grw

If AppWindow.TextBox31 = "" Then
Range(Gkoord).Value = "0"
Else
Range(Gkoord).Value = AppWindow.TextBox31.Value
End If

' - Team III. "H:H" - '

Sheets("Létszám").Select
Columns("h:h").Select
Selection.End(xlDown).Select
Dim Hrw As Long
Hrw = ActiveCell.row + 1
Dim Hoszlop As String
Hoszlop = "h"
Dim Hkoord As String
Hkoord = Hoszlop & Hrw

If AppWindow.TextBox30 = "" Then
Range(Hkoord).Value = "0"
Else
Range(Hkoord).Value = AppWindow.TextBox30.Value
End If


                                        ' - Délelõtt - Villanyszerelõ - '

' - Team I. "I:I" - '

Sheets("Létszám").Select
Columns("i:i").Select
Selection.End(xlDown).Select
Dim Irw As Long
Irw = ActiveCell.row + 1
Dim Ioszlop As String
Ioszlop = "i"
Dim Ikoord As String
Ikoord = Ioszlop & Irw

If AppWindow.TextBox39 = "" Then
Range(Ikoord).Value = "0"
Else
Range(Ikoord).Value = AppWindow.TextBox39.Value
End If

' - Team II. "J:J" - '

Sheets("Létszám").Select
Columns("j:j").Select
Selection.End(xlDown).Select
Dim Jrw As Long
Jrw = ActiveCell.row + 1
Dim Joszlop As String
Joszlop = "j"
Dim Jkoord As String
Jkoord = Joszlop & Jrw


If AppWindow.TextBox29 = "" Then
Range(Jkoord).Value = "0"
Else
Range(Jkoord).Value = AppWindow.TextBox29.Value
End If

' - Team III. "K:K" - '

Sheets("Létszám").Select
Columns("k:k").Select
Selection.End(xlDown).Select
Dim Krw As Long
Krw = ActiveCell.row + 1
Dim Koszlop As String
Koszlop = "k"
Dim Kkoord As String
Kkoord = Koszlop & Krw

If AppWindow.TextBox28 = "" Then
Range(Kkoord).Value = "0"
Else
Range(Kkoord).Value = AppWindow.TextBox28.Value
End If

                                        ' - Délután - Mérnök - '

' - Team I. "L:L" - '

Sheets("Létszám").Select
Columns("l:l").Select
Selection.End(xlDown).Select
Dim Lrw As Long
Lrw = ActiveCell.row + 1
Dim Loszlop As String
Loszlop = "l"
Dim Lkoord As String
Lkoord = Loszlop & Lrw

If AppWindow.TextBox36 = "" Then
Range(Lkoord).Value = "0"
Else
Range(Lkoord).Value = AppWindow.TextBox36.Value
End If

' - Team II. "M:M" - '

Sheets("Létszám").Select
Columns("m:m").Select
Selection.End(xlDown).Select
Dim Mrw As Long
Mrw = ActiveCell.row + 1
Dim Moszlop As String
Moszlop = "m"
Dim Mkoord As String
Mkoord = Moszlop & Mrw

If AppWindow.TextBox35 = "" Then
Range(Mkoord).Value = "0"
Else
Range(Mkoord).Value = AppWindow.TextBox35.Value
End If

' - Team III. "N:N" - '

Sheets("Létszám").Select
Columns("n:n").Select
Selection.End(xlDown).Select
Dim Nrw As Long
Nrw = ActiveCell.row + 1
Dim Noszlop As String
Noszlop = "n"
Dim Nkoord As String
Nkoord = Noszlop & Nrw

If AppWindow.TextBox34 = "" Then
Range(Nkoord).Value = "0"
Else
Range(Nkoord).Value = AppWindow.TextBox34.Value
End If

                                        ' - Délután - Lakatos - '

' - Team I. "O:O" - '

Sheets("Létszám").Select
Columns("o:o").Select
Selection.End(xlDown).Select
Dim Orw As Long
Orw = ActiveCell.row + 1
Dim Ooszlop As String
Ooszlop = "o"
Dim Okoord As String
Okoord = Ooszlop & Orw

If AppWindow.TextBox27 = "" Then
Range(Okoord).Value = "0"
Else
Range(Okoord).Value = AppWindow.TextBox27.Value
End If

' - Team II. "P:P" - '

Sheets("Létszám").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Dim Prw As Long
Prw = ActiveCell.row + 1
Dim Poszlop As String
Poszlop = "p"
Dim Pkoord As String
Pkoord = Poszlop & Prw

If AppWindow.TextBox25 = "" Then
Range(Pkoord).Value = "0"
Else
Range(Pkoord).Value = AppWindow.TextBox25.Value
End If

' - Team III. "Q:Q" - '

Sheets("Létszám").Select
Columns("q:q").Select
Selection.End(xlDown).Select
Dim Qrw As Long
Qrw = ActiveCell.row + 1
Dim Qoszlop As String
Qoszlop = "q"
Dim Qkoord As String
Qkoord = Qoszlop & Qrw

If AppWindow.TextBox19 = "" Then
Range(Qkoord).Value = "0"
Else
Range(Qkoord).Value = AppWindow.TextBox19.Value
End If


                                        ' - Délután - Villanyszerelõ - '

' - Team I. "R:R" - '

Sheets("Létszám").Select
Columns("r:r").Select
Selection.End(xlDown).Select
Dim Rrw As Long
Rrw = ActiveCell.row + 1
Dim Roszlop As String
Roszlop = "r"
Dim Rkoord As String
Rkoord = Roszlop & Rrw

If AppWindow.TextBox26 = "" Then
Range(Rkoord).Value = "0"
Else
Range(Rkoord).Value = AppWindow.TextBox26.Value
End If

' - Team II. "S:S" - '

Sheets("Létszám").Select
Columns("s:s").Select
Selection.End(xlDown).Select
Dim Srw As Long
Srw = ActiveCell.row + 1
Dim Soszlop As String
Soszlop = "s"
Dim Skoord As String
Skoord = Soszlop & Srw

If AppWindow.TextBox24 = "" Then
Range(Skoord).Value = "0"
Else
Range(Skoord).Value = AppWindow.TextBox24.Value
End If

' - Team III. "T:T" - '

Sheets("Létszám").Select
Columns("t:t").Select
Selection.End(xlDown).Select
Dim Trw As Long
Trw = ActiveCell.row + 1
Dim Toszlop As String
Toszlop = "t"
Dim Tkoord As String
Tkoord = Toszlop & Trw

If AppWindow.TextBox18 = "" Then
Range(Tkoord).Value = "0"
Else
Range(Tkoord).Value = AppWindow.TextBox18.Value
End If

                                        ' - Éjjel - Mérnök - '

' - Team I. "U:U" - '

Sheets("Létszám").Select
Columns("u:u").Select
Selection.End(xlDown).Select
Dim Urw As Long
Urw = ActiveCell.row + 1
Dim Uoszlop As String
Uoszlop = "u"
Dim Ukoord As String
Ukoord = Uoszlop & Urw

If AppWindow.TextBox33 = "" Then
Range(Ukoord).Value = "0"
Else
Range(Ukoord).Value = AppWindow.TextBox33.Value
End If

' - Team II. "V:V" - '

Sheets("Létszám").Select
Columns("v:v").Select
Selection.End(xlDown).Select
Dim Vrw As Long
Vrw = ActiveCell.row + 1
Dim Voszlop As String
Voszlop = "v"
Dim Vkoord As String
Vkoord = Voszlop & Vrw

If AppWindow.TextBox44 = "" Then
Range(Vkoord).Value = "0"
Else
Range(Vkoord).Value = AppWindow.TextBox44.Value
End If

' - Team III. "W:W" - '

Sheets("Létszám").Select
Columns("w:w").Select
Selection.End(xlDown).Select
Dim Wrw As Long
Wrw = ActiveCell.row + 1
Dim Woszlop As String
Woszlop = "w"
Dim Wkoord As String
Wkoord = Woszlop & Wrw

If AppWindow.TextBox47 = "" Then
Range(Wkoord).Value = "0"
Else
Range(Wkoord).Value = AppWindow.TextBox47.Value
End If

                                        ' - Éjjel - Lakatos - '

' - Team I. "X:X" - '

Sheets("Létszám").Select
Columns("x:x").Select
Selection.End(xlDown).Select
Dim Xrw As Long
Xrw = ActiveCell.row + 1
Dim Xoszlop As String
Xoszlop = "x"
Dim Xkoord As String
Xkoord = Xoszlop & Xrw

If AppWindow.TextBox21 = "" Then
Range(Xkoord).Value = "0"
Else
Range(Xkoord).Value = AppWindow.TextBox21.Value
End If

' - Team II. "Y:Y" - '

Sheets("Létszám").Select
Columns("y:y").Select
Selection.End(xlDown).Select
Dim Yrw As Long
Yrw = ActiveCell.row + 1
Dim Yoszlop As String
Yoszlop = "y"
Dim Ykoord As String
Ykoord = Yoszlop & Yrw

If AppWindow.TextBox43 = "" Then
Range(Ykoord).Value = "0"
Else
Range(Ykoord).Value = AppWindow.TextBox43.Value
End If

' - Team III. "Z:Z" - '

Sheets("Létszám").Select
Columns("z:z").Select
Selection.End(xlDown).Select
Dim Zrw As Long
Zrw = ActiveCell.row + 1
Dim Zoszlop As String
Zoszlop = "z"
Dim Zkoord As String
Zkoord = Zoszlop & Zrw

If AppWindow.TextBox46 = "" Then
Range(Zkoord).Value = "0"
Else
Range(Zkoord).Value = AppWindow.TextBox46.Value
End If


                                        ' - Éjjel - Villanyszerelõ - '

' - Team I. "AA:AA" - '

Sheets("Létszám").Select
Columns("aa:aa").Select
Selection.End(xlDown).Select
Dim AArw As Long
AArw = ActiveCell.row + 1
Dim AAoszlop As String
AAoszlop = "aa"
Dim AAkoord As String
AAkoord = AAoszlop & AArw

If AppWindow.TextBox20 = "" Then
Range(AAkoord).Value = "0"
Else
Range(AAkoord).Value = AppWindow.TextBox20.Value
End If

' - Team II. "AB:AB" - '

Sheets("Létszám").Select
Columns("ab:ab").Select
Selection.End(xlDown).Select
Dim ABrw As Long
ABrw = ActiveCell.row + 1
Dim ABoszlop As String
ABoszlop = "ab"
Dim ABkoord As String
ABkoord = ABoszlop & ABrw

If AppWindow.TextBox42 = "" Then
Range(ABkoord).Value = "0"
Else
Range(ABkoord).Value = AppWindow.TextBox42.Value
End If

' - Team III. "AC:AC" - '

Sheets("Létszám").Select
Columns("ac:ac").Select
Selection.End(xlDown).Select
Dim ACrw As Long
ACrw = ActiveCell.row + 1
Dim ACoszlop As String
ACoszlop = "ac"
Dim ACkoord As String
ACkoord = ACoszlop & ACrw

If AppWindow.TextBox45 = "" Then
Range(ACkoord).Value = "0"
Else
Range(ACkoord).Value = AppWindow.TextBox45.Value
End If

                                        ' - TPM - '

' - Mérnök "AD:AD" - '

Sheets("Létszám").Select
Columns("ad:ad").Select
Selection.End(xlDown).Select
Dim ADrw As Long
ADrw = ActiveCell.row + 1
Dim ADoszlop As String
ADoszlop = "ad"
Dim ADkoord As String
ADkoord = ADoszlop & ADrw

If AppWindow.TextBox32 = "" Then
Range(ADkoord).Value = "0"
Else
Range(ADkoord).Value = AppWindow.TextBox32.Value
End If

' - Team II. "AE:AE" - '

Sheets("Létszám").Select
Columns("ae:ae").Select
Selection.End(xlDown).Select
Dim AErw As Long
AErw = ActiveCell.row + 1
Dim AEoszlop As String
AEoszlop = "ae"
Dim AEkoord As String
AEkoord = AEoszlop & AErw

If AppWindow.TextBox23 = "" Then
Range(AEkoord).Value = "0"
Else
Range(AEkoord).Value = AppWindow.TextBox23.Value
End If

' - Team III. "AF:AF" - '

Sheets("Létszám").Select
Columns("af:af").Select
Selection.End(xlDown).Select
Dim AFrw As Long
AFrw = ActiveCell.row + 1
Dim AFoszlop As String
AFoszlop = "af"
Dim AFkoord As String
AFkoord = AFoszlop & AFrw

If AppWindow.TextBox22 = "" Then
Range(AFkoord).Value = "0"
Else
Range(AFkoord).Value = AppWindow.TextBox22.Value
End If


Else
Exit Sub
End If

End Sub
