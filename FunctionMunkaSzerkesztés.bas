Attribute VB_Name = "FunctionMunkaSzerkesztés"
Option Explicit

Sub MunkaSzerkesztés()

'ez vissza adja a kijelölt sor ID-t.

Dim JelöltSor As Long
JelöltSor = AppWindow.ListBox20.Value
Dim rownr As Long
rownr = JelöltSor + 1

' - Bárcaszám - B'

Dim Bclm As String
Bclm = "b"
Dim Bkoord As String
Bkoord = Bclm & rownr
AppWindow.TextBox54 = Munka1.Range(Bkoord).Value

' - Dátum - C'

Dim Cclm As String
Cclm = "c"
Dim Ckoord As String
Ckoord = Cclm & rownr
AppWindow.TextBox62 = Date

' - Munkaszám - D'

Dim Dclm As String
Dclm = "d"
Dim Dkoord As String
Dkoord = Dclm & rownr
AppWindow.TextBox63 = Munka1.Range(Dkoord).Value

' - Rábaszám - E'

Dim Eclm As String
Eclm = "e"
Dim Ekoord As String
Ekoord = Eclm & rownr
AppWindow.TextBox64 = Munka1.Range(Ekoord).Value

' - Gépszám - F'

Dim Fclm As String
Fclm = "f"
Dim Fkoord As String
Fkoord = Fclm & rownr
AppWindow.TextBox65 = Munka1.Range(Fkoord).Value

' - Kulcsgép - G'

Dim Gclm As String
Gclm = "g"
Dim Gkoord As String
Gkoord = Gclm & rownr
AppWindow.TextBox66 = Munka1.Range(Gkoord).Value

' - Terület - H'

Dim Hclm As String
Hclm = "h"
Dim Hkoord As String
Hkoord = Hclm & rownr
AppWindow.TextBox67 = Munka1.Range(Hkoord).Value

' - Csapat - I'

Dim Iclm As String
Iclm = "i"
Dim Ikoord As String
Ikoord = Iclm & rownr
AppWindow.TextBox68 = Munka1.Range(Ikoord).Value

' - Mûszak - M'

Dim Mclm As String
Mclm = "m"
Dim Mkoord As String
Mkoord = Mclm & rownr
AppWindow.TextBox69 = Munka1.Range(Mkoord).Value

' - -tól - J'

Dim Jclm As String
Jclm = "j"
Dim Jkoord As String
Jkoord = Jclm & rownr
AppWindow.TextBox74 = Munka1.Range(Jkoord).Value

' - -ig - K'

Dim Kclm As String
Kclm = "k"
Dim Kkoord As String
Kkoord = Kclm & rownr
AppWindow.TextBox49 = Munka1.Range(Kkoord).Value

' - Idõ - L'

Dim Lclm As String
Lclm = "l"
Dim Lkoord As String
Lkoord = Lclm & rownr
AppWindow.TextBox71 = Munka1.Range(Lkoord).Value

' - Probléma - N'

Dim Nclm As String
Nclm = "n"
Dim Nkoord As String
Nkoord = Nclm & rownr
AppWindow.TextBox72 = Munka1.Range(Nkoord).Value

' - Megoldás - O'

Dim Oclm As String
Oclm = "o"
Dim Okoord As String
Okoord = Oclm & rownr
AppWindow.TextBox57 = Munka1.Range(Okoord).Value

' - Státusz - P'

Dim Pclm As String
Pclm = "p"
Dim Pkoord As String
Pkoord = Pclm & rownr
AppWindow.ComboBox5 = Munka1.Range(Pkoord).Value

' - Mérés - Q'

Dim Qclm As String
Qclm = "q"
Dim Qkoord As String
Qkoord = Qclm & rownr
AppWindow.ComboBox6 = Munka1.Range(Qkoord).Value

' - Felelõs - R'

Dim Rclm As String
Rclm = "r"
Dim Rkoord As String
Rkoord = Rclm & rownr
AppWindow.ComboBox7 = Munka1.Range(Rkoord).Value

' - Becsült visszaadás - S'

Dim Sclm As String
Sclm = "s"
Dim Skoord As String
Skoord = Sclm & rownr
AppWindow.TextBox59 = Munka1.Range(Skoord).Value

' - Visszaigazolás - T'

Dim Tclm As String
Tclm = "t"
Dim Tkoord As String
Tkoord = Tclm & rownr
AppWindow.TextBox60 = Munka1.Range(Tkoord).Value

' - Visszaadás tény - U'

Dim Uclm As String
Uclm = "u"
Dim Ukoord As String
Ukoord = Uclm & rownr
AppWindow.TextBox61 = Munka1.Range(Ukoord).Value

Sheets("Start").Select
Range("b2").Select
End Sub
