Attribute VB_Name = "FunctionMunkaSzerkeszt�s2"
Option Explicit

Sub MunkaSzerkeszt�s2()

'ez vissza adja a kijel�lt sor ID-t.

Dim Jel�ltSor As Long
Jel�ltSor = AppWindow.ListBox7.Value
Dim rownr As Long
rownr = Jel�ltSor + 1

' - B�rcasz�m - B'

Dim Bclm As String
Bclm = "b"
Dim Bkoord As String
Bkoord = Bclm & rownr
AppWindow.TextBox11 = Munka1.Range(Bkoord).Value

' - Munkasz�m - D'

Dim Dclm As String
Dclm = "d"
Dim Dkoord As String
Dkoord = Dclm & rownr
AppWindow.TextBox1 = Munka1.Range(Dkoord).Value

' - R�basz�m - E'

Dim Eclm As String
Eclm = "e"
Dim Ekoord As String
Ekoord = Eclm & rownr
AppWindow.TextBox10 = Munka1.Range(Ekoord).Value

' - Ter�let - H'

Dim Hclm As String
Hclm = "h"
Dim Hkoord As String
Hkoord = Hclm & rownr
AppWindow.ComboBox1 = Munka1.Range(Hkoord).Value

' - Csapat - I'

Dim Iclm As String
Iclm = "i"
Dim Ikoord As String
Ikoord = Iclm & rownr
AppWindow.ComboBox2 = Munka1.Range(Ikoord).Value

' - -t�l - J'

Dim Jclm As String
Jclm = "j"
Dim Jkoord As String
Jkoord = Jclm & rownr

AppWindow.TextBox7.Value = Munka1.Range(Jkoord).Value

' - -ig - K'

Dim Kclm As String
Kclm = "k"
Dim Kkoord As String
Kkoord = Kclm & rownr
AppWindow.TextBox6.Value = Munka1.Range(Kkoord).Value

' - Probl�ma - N'

Dim Nclm As String
Nclm = "n"
Dim Nkoord As String
Nkoord = Nclm & rownr
AppWindow.TextBox5 = Munka1.Range(Nkoord).Value

' - Megold�s - O'

Dim Oclm As String
Oclm = "o"
Dim Okoord As String
Okoord = Oclm & rownr
AppWindow.TextBox4 = Munka1.Range(Okoord).Value

' - St�tusz - P'

Dim Pclm As String
Pclm = "p"
Dim Pkoord As String
Pkoord = Pclm & rownr
AppWindow.ComboBox4 = Munka1.Range(Pkoord).Value

' - M�r�s - Q'

Dim Qclm As String
Qclm = "q"
Dim Qkoord As String
Qkoord = Qclm & rownr
AppWindow.ComboBox3 = Munka1.Range(Qkoord).Value

' - Kateg�ria - X'

Dim Xclm As String
Xclm = "x"
Dim Xkoord As String
Xkoord = Xclm & rownr
AppWindow.ComboBox8 = Munka1.Range(Xkoord).Value

' - Megjegyz�s - V'

Dim Vclm As String
Vclm = "v"
Dim Vkoord As String
Vkoord = Vclm & rownr
AppWindow.TextBox78 = Munka1.Range(Vkoord).Value

Sheets("Start").Select
Range("b2").Select
End Sub

