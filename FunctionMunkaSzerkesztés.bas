Attribute VB_Name = "FunctionMunkaSzerkeszt�s"
Option Explicit

Sub MunkaSzerkeszt�s()

'ez vissza adja a kijel�lt sor ID-t.

Dim Jel�ltSor As Long
Jel�ltSor = AppWindow.ListBox20.Value
Dim rownr As Long
rownr = Jel�ltSor + 1

' - B�rcasz�m - B'

Dim Bclm As String
Bclm = "b"
Dim Bkoord As String
Bkoord = Bclm & rownr
AppWindow.TextBox54 = Munka1.Range(Bkoord).Value

' - D�tum - C'

Dim Cclm As String
Cclm = "c"
Dim Ckoord As String
Ckoord = Cclm & rownr
AppWindow.TextBox62 = Date

' - Munkasz�m - D'

Dim Dclm As String
Dclm = "d"
Dim Dkoord As String
Dkoord = Dclm & rownr
AppWindow.TextBox63 = Munka1.Range(Dkoord).Value

' - R�basz�m - E'

Dim Eclm As String
Eclm = "e"
Dim Ekoord As String
Ekoord = Eclm & rownr
AppWindow.TextBox64 = Munka1.Range(Ekoord).Value

' - G�psz�m - F'

Dim Fclm As String
Fclm = "f"
Dim Fkoord As String
Fkoord = Fclm & rownr
AppWindow.TextBox65 = Munka1.Range(Fkoord).Value

' - Kulcsg�p - G'

Dim Gclm As String
Gclm = "g"
Dim Gkoord As String
Gkoord = Gclm & rownr
AppWindow.TextBox66 = Munka1.Range(Gkoord).Value

' - Ter�let - H'

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

' - M�szak - M'

Dim Mclm As String
Mclm = "m"
Dim Mkoord As String
Mkoord = Mclm & rownr
AppWindow.TextBox69 = Munka1.Range(Mkoord).Value

' - -t�l - J'

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

' - Id� - L'

Dim Lclm As String
Lclm = "l"
Dim Lkoord As String
Lkoord = Lclm & rownr
AppWindow.TextBox71 = Munka1.Range(Lkoord).Value

' - Probl�ma - N'

Dim Nclm As String
Nclm = "n"
Dim Nkoord As String
Nkoord = Nclm & rownr
AppWindow.TextBox72 = Munka1.Range(Nkoord).Value

' - Megold�s - O'

Dim Oclm As String
Oclm = "o"
Dim Okoord As String
Okoord = Oclm & rownr
AppWindow.TextBox57 = Munka1.Range(Okoord).Value

' - St�tusz - P'

Dim Pclm As String
Pclm = "p"
Dim Pkoord As String
Pkoord = Pclm & rownr
AppWindow.ComboBox5 = Munka1.Range(Pkoord).Value

' - M�r�s - Q'

Dim Qclm As String
Qclm = "q"
Dim Qkoord As String
Qkoord = Qclm & rownr
AppWindow.ComboBox6 = Munka1.Range(Qkoord).Value

' - Felel�s - R'

Dim Rclm As String
Rclm = "r"
Dim Rkoord As String
Rkoord = Rclm & rownr
AppWindow.ComboBox7 = Munka1.Range(Rkoord).Value

' - Becs�lt visszaad�s - S'

Dim Sclm As String
Sclm = "s"
Dim Skoord As String
Skoord = Sclm & rownr
AppWindow.TextBox59 = Munka1.Range(Skoord).Value

' - Visszaigazol�s - T'

Dim Tclm As String
Tclm = "t"
Dim Tkoord As String
Tkoord = Tclm & rownr
AppWindow.TextBox60 = Munka1.Range(Tkoord).Value

' - Visszaad�s t�ny - U'

Dim Uclm As String
Uclm = "u"
Dim Ukoord As String
Ukoord = Uclm & rownr
AppWindow.TextBox61 = Munka1.Range(Ukoord).Value

Sheets("Start").Select
Range("b2").Select
End Sub
