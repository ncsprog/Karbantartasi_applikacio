VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantart�si adatgy�jt�"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "AppWindow.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adatfelv�telMent�s_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox11 = "" Then
MsgBox "B�rcasz�m megad�sa k�telez�!"
Adatfelv�telLista
Exit Sub
End If

'If AppWindow.ComboBox8 = "" Then
If AppWindow.TextBox11 <> "" Then
ElseIf AppWindow.TextBox10 = "" Then
MsgBox "Kateg�ri�t v�lasztani k�telez�!" & vbCrLf & "vagy" & vbCrLf & "R�basz�m hi�nyzik" & vbCrLf & vbCrLf & "Nem t�rt�nt adatment�s."
Exit Sub
ElseIf AppWindow.ComboBox8 = "" Then
MsgBox "Kateg�ri�t v�lasztani k�telez�!" & vbCrLf & "vagy" & vbCrLf & "R�basz�m hi�nyzik" & vbCrLf & vbCrLf & "Nem t�rt�nt adatment�s."
Adatfelv�telLista
Jelsz�Rejt�s
Exit Sub
End If

'If AppWindow.TextBox11 <> "" Then
'TartalomEllen�rz�s
AdatokM�sol�sa
IDgener�l�s
ID_gener�l�s
Id�Kalkul�tor
AdatokMent�se
Adatfelv�telLista
T�rl�s
'Id�sz�m�t�s
'End If

BoxInakt�v�l�

Jelsz�Rejt�s
End Sub


Private Sub CommandButton1_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox73 = "" Then
MsgBox "k�rek egy g�psz�mot"
'Adatfelv�telLista3
End If
If AppWindow.TextBox73 <> "" Then
G�pT�rt�net
End If

Jelsz�Rejt�s
End Sub





Private Sub CommandButton10_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox101 = "" Then
MsgBox "Nincs megadva �j st�tusz."
Else
St�tuszSzerkeszt�s
Rendez�s
Adatfelv�telLista10
CbFelt�lt�s
End If

Jelsz�Rejt�s
End Sub

Private Sub CommandButton12_Click()
Jelsz�Rejt�s2
If ListBox29.ListCount = 0 Or ListBox29.ListCount < 0 Or ListBox29.ListCount > 21 Then
MsgBox "Nincs kiv�lasztott st�tusz."
Exit Sub
Else
St�tuszSzerkeszt2
Rendez�s
CbFelt�lt�s
End If
Adatfelv�telLista10

Jelsz�Rejt�s
End Sub

Private Sub CommandButton13_Click()


Jelsz�Rejt�s2
If ListBox30.ListCount = 0 Then
MsgBox "Nincs kiv�lasztott sor."
Else
Felel�sSzerkeszt�s2
Adatfelv�telLista11
Rendez�s2
CbFelt�lt�s
End If

Jelsz�Rejt�s
End Sub

Private Sub CommandButton14_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox102 = "" Then
MsgBox "Nincs megadva �j st�tusz."
Else
Felel�sSzerkeszt�s
Rendez�s2
Adatfelv�telLista11
CbFelt�lt�s
End If
Jelsz�Rejt�s
End Sub

Private Sub CommandButton15_Click()

If AppWindow.TextBox105 = "" Then
MsgBox "K�rlek add meg az aktu�lis �sszl�tsz�mot."
Jelsz�Rejt�s2
Adatfelv�telLista12
Jelsz�Rejt�s
Exit Sub
Else
Jelsz�Rejt�s2
AdatokM�sol�sa4
ID_gener�l�s4
AppWindow.TextBox105 = ""
Adatfelv�telLista12
End If
Jelsz�Rejt�s
End Sub

Private Sub CommandButton16_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox106 = "" Then
B�rcaKeres_Friss�t
Else
B�rcaKeres_Keres
End If
Jelsz�Rejt�s
End Sub

Private Sub CommandButton17_Click()

Jelsz�Rejt�s2
If AppWindow.TextBox107 = "" Then
G�pKeres
Else
G�pKeres2
'B�rcaKeres_Keres
End If
Jelsz�Rejt�s




End Sub

Private Sub CommandButton18_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox108 = "" Then
MsgBox "Nincs megadva �j st�tusz."
Else
Kateg�riaSzerkeszt
Rendez�s3
Adatfelv�telLista14
CbFelt�lt�s
End If

Jelsz�Rejt�s
End Sub

Private Sub CommandButton19_Click()
Jelsz�Rejt�s2
If ListBox37.ListCount = 0 Or ListBox37.ListCount < 0 Or ListBox37.ListCount > 21 Then
MsgBox "Nincs kiv�lasztott kateg�ria."
Exit Sub
Else
Kateg�riaSzerkeszt2
Rendez�s3
CbFelt�lt�s
End If
Adatfelv�telLista14

Jelsz�Rejt�s
End Sub

Private Sub CommandButton2_Click()
Jelsz�Rejt�s2
G�p�ll�sokSz�ma
'tb75
Jelsz�Rejt�s
End Sub

Private Sub CommandButton20_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox109 = "" Then
MsgBox "Nincs megadva �j ter�let."
Else
Ter�letSzerkeszt
Rendez�s5
Adatfelv�telLista15
CbFelt�lt�s
End If

Jelsz�Rejt�s
End Sub

Private Sub CommandButton21_Click()
Jelsz�Rejt�s2
If ListBox38.ListCount = 0 Or ListBox38.ListCount < 0 Or ListBox38.ListCount > 21 Then
MsgBox "Nincs kiv�lasztott st�tusz."
Exit Sub
Else
Ter�letSzerkeszt2
Rendez�s
CbFelt�lt�s
End If
Adatfelv�telLista15

Jelsz�Rejt�s
End Sub

Private Sub CommandButton22_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox110 = "" Then
MsgBox "Nincs megadva �j csapat."
Else
CsapatSzerkeszt
Rendez�s4
Adatfelv�telLista16
CbFelt�lt�s
End If

Jelsz�Rejt�s
End Sub

Private Sub CommandButton23_Click()
Jelsz�Rejt�s2
If ListBox39.ListCount = 0 Or ListBox39.ListCount < 0 Or ListBox39.ListCount > 21 Then
MsgBox "Nincs kiv�lasztott csapat."
Exit Sub
Else
CsapatSzerkeszt2
Rendez�s4
CbFelt�lt�s
End If
Adatfelv�telLista16

Jelsz�Rejt�s
End Sub

Private Sub CommandButton24_Click()
Jelsz�Rejt�s2
CbR�gz�t
Jelsz�Rejt�s
End Sub

Private Sub CommandButton25_Click()
Jelsz�Rejt�s2
CbFelt�lt�s
CbVisszaad�s
Jelsz�Rejt�s
End Sub

Private Sub CommandButton26_Click()
Jelsz�Rejt�s2
Megbesz�l�sD�tumok
Megbesz�l�sM�sol�
Jelsz�Rejt�s
End Sub

Private Sub CommandButton27_Click()
Jelsz�Rejt�s2
MunkaSzerkeszt�s2
BoxAkt�v�l�
Jelsz�Rejt�s
End Sub

Private Sub CommandButton3_Click()
Jelsz�Rejt�s2
IdegenXL2
IdegenXL3
Adatfelv�telLista8_�1
Adatfelv�telLista9_R1
Adatfelv�telLista13
'AppWindow.TextBox76 = "M�sol�s k�sz"
'tb76
Jelsz�Rejt�s
End Sub

Private Sub CommandButton4_Click()
Jelsz�Rejt�s2
IdegenXL
Adatfelv�telLista4
Adatfelf�telLista5
Adatfelv�telLista6
Adatfelv�telLista7
'AppWindow.TextBox77.Value = "M�sol�s k�sz"

'tb77
Jelsz�Rejt�s
End Sub


Private Sub CommandButton5_Click()
Jelsz�Rejt�s2
IDgener�l�s2
IDgener�l�s3
ID_gener�l�s2
ID_gener�l�s3
AdatokM�sol�sa3
AdatokM�sol�sa5
L�tsz�m�sszes�t�s
Jelsz�Rejt�s
End Sub

Private Sub CommandButton6_Click()
Jelsz�Rejt�s2
If AppWindow.ComboBox5 = "" Then
MsgBox "K�rlek v�laszd ki a st�tuszt."
Exit Sub
Else
Adatfelv�telLista2
End If
Jelsz�Rejt�s
End Sub


Private Sub CommandButton8_Click()
Jelsz�Rejt�s3
End Sub

Private Sub CommandButton9_Click()

If AppWindow.TextBox100 = "smj266" Then
Jelsz�Rejt�s2
TbBet�lt�s
AppWindow.MultiPage1.page4.Visible = True
AppWindow.MultiPage1.page5.Visible = True
AppWindow.Frame19.Visible = True
AppWindow.Frame20.Visible = True
AppWindow.Frame21.Visible = True
AppWindow.Frame23.Visible = True
AppWindow.Frame26.Visible = True
AppWindow.Frame27.Visible = True
AppWindow.Frame28.Visible = True
AppWindow.Frame17.Visible = True
Adatfelv�telLista10
Adatfelv�telLista11
Adatfelv�telLista14
Adatfelv�telLista15
Adatfelv�telLista16
'CbFelt�lt�s
'CbVisszaad�s
Jelsz�Rejt�s
AppWindow.TextBox100 = ""
Else
MsgBox "Nem megfelel� jelsz�!"
AppWindow.MultiPage1.page4.Visible = False
AppWindow.MultiPage1.page5.Visible = False
AppWindow.Frame19.Visible = False
AppWindow.Frame20.Visible = False
AppWindow.Frame21.Visible = False
AppWindow.Frame23.Visible = False
AppWindow.Frame26.Visible = False
AppWindow.Frame27.Visible = False
AppWindow.Frame28.Visible = False
AppWindow.Frame17.Visible = False


AppWindow.TextBox100 = ""


End If

End Sub



Private Sub Nyomonk�vet�Friss�t�s_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox62 <> "" Then

'TartalomEllen�rz�s2
IDgener�l�s
ID_gener�l�s
AdatokM�sol�sa2
AdatokMent�se
T�rl�s2
Adatfelv�telLista2
End If
Jelsz�Rejt�s
End Sub

Private Sub Nyomonk�vet�Szerkeszt�s_Click()
Jelsz�Rejt�s2
If ListBox20.ListCount = 0 Then
MsgBox "Nincs kiv�lasztott sor." & vbCrLf & "A lista megjen�t�shez kattints a FRISS�T�S gombra."
Exit Sub
Else
MunkaSzerkeszt�s
End If
Jelsz�Rejt�s
End Sub




Private Sub TextBox103_Change()

End Sub

Private Sub TextBox112_Change()

AppWindow.TextBox112 = TextBox115 + TextBox114 + TextBox113

End Sub


Private Sub UserForm_Initialize()
Jelsz�Rejt�s2
    ' - Adatfelv�telTer�let - '

    ' - Adatfelv�telCsapat - '

    ' - Adatfelv�telM�r�s - '
    ComboBox3.AddItem "Igen"
    ComboBox3.AddItem "Nem"
    ' - Adatfelv�telSt�tusz - '
    ComboBox4.AddItem Munka12.Range("b2").Value
    ComboBox4.AddItem Munka12.Range("b3").Value
    ComboBox4.AddItem Munka12.Range("b4").Value
    ComboBox4.AddItem Munka12.Range("b5").Value
    ComboBox4.AddItem Munka12.Range("b6").Value
    ComboBox4.AddItem Munka12.Range("b7").Value
    ComboBox4.AddItem Munka12.Range("b8").Value
    ComboBox4.AddItem Munka12.Range("b9").Value
    ComboBox4.AddItem Munka12.Range("b10").Value
    ComboBox4.AddItem Munka12.Range("b11").Value
    ComboBox4.AddItem Munka12.Range("b12").Value
    ComboBox4.AddItem Munka12.Range("b13").Value
    ComboBox4.AddItem Munka12.Range("b14").Value
    ComboBox4.AddItem Munka12.Range("b15").Value
    ComboBox4.AddItem Munka12.Range("b16").Value
    ComboBox4.AddItem Munka12.Range("b17").Value
    ComboBox4.AddItem Munka12.Range("b18").Value
    ComboBox4.AddItem Munka12.Range("b19").Value
    ComboBox4.AddItem Munka12.Range("b20").Value
    ComboBox4.AddItem Munka12.Range("b21").Value
    
  
    ' - Nyomonk�vet�St�tusz - '
    ComboBox5.AddItem Munka12.Range("b2").Value
    ComboBox5.AddItem Munka12.Range("b3").Value
    ComboBox5.AddItem Munka12.Range("b4").Value
    ComboBox5.AddItem Munka12.Range("b5").Value
    ComboBox5.AddItem Munka12.Range("b6").Value
    ComboBox5.AddItem Munka12.Range("b7").Value
    ComboBox5.AddItem Munka12.Range("b8").Value
    ComboBox5.AddItem Munka12.Range("b9").Value
    ComboBox5.AddItem Munka12.Range("b10").Value
    ComboBox5.AddItem Munka12.Range("b11").Value
    ComboBox5.AddItem Munka12.Range("b12").Value
    ComboBox5.AddItem Munka12.Range("b13").Value
    ComboBox5.AddItem Munka12.Range("b14").Value
    ComboBox5.AddItem Munka12.Range("b15").Value
    ComboBox5.AddItem Munka12.Range("b16").Value
    ComboBox5.AddItem Munka12.Range("b17").Value
    ComboBox5.AddItem Munka12.Range("b18").Value
    ComboBox5.AddItem Munka12.Range("b19").Value
    ComboBox5.AddItem Munka12.Range("b20").Value
    ComboBox5.AddItem Munka12.Range("b21").Value
    

    ' - Nyomonk�vet�M�r�s - '
    ComboBox6.AddItem "Igen"
    ComboBox6.AddItem "Nem"
    ' - Nyomonk�vet�Felel�s - '
    ComboBox7.AddItem Munka12.Range("d2").Value
    ComboBox7.AddItem Munka12.Range("d3").Value
    ComboBox7.AddItem Munka12.Range("d4").Value
    ComboBox7.AddItem Munka12.Range("d5").Value
    ComboBox7.AddItem Munka12.Range("d6").Value
    ComboBox7.AddItem Munka12.Range("d7").Value
    ComboBox7.AddItem Munka12.Range("d8").Value
    ComboBox7.AddItem Munka12.Range("d9").Value
    ComboBox7.AddItem Munka12.Range("d10").Value
    ComboBox7.AddItem Munka12.Range("d11").Value
    ComboBox7.AddItem Munka12.Range("d12").Value
    ComboBox7.AddItem Munka12.Range("d13").Value
    ComboBox7.AddItem Munka12.Range("d14").Value
    ComboBox7.AddItem Munka12.Range("d15").Value
    ComboBox7.AddItem Munka12.Range("d16").Value
    ComboBox7.AddItem Munka12.Range("d17").Value
    ComboBox7.AddItem Munka12.Range("d18").Value
    ComboBox7.AddItem Munka12.Range("d19").Value
    ComboBox7.AddItem Munka12.Range("d20").Value
    ComboBox7.AddItem Munka12.Range("d21").Value
    ComboBox7.AddItem Munka12.Range("d22").Value
    ComboBox7.AddItem Munka12.Range("d23").Value
    ComboBox7.AddItem Munka12.Range("d24").Value
    ComboBox7.AddItem Munka12.Range("d25").Value
    ComboBox7.AddItem Munka12.Range("d26").Value
    ComboBox7.AddItem Munka12.Range("d27").Value
    ComboBox7.AddItem Munka12.Range("d28").Value
    ComboBox7.AddItem Munka12.Range("d29").Value
    ComboBox7.AddItem Munka12.Range("d30").Value
    ComboBox7.AddItem Munka12.Range("d31").Value
    
    ' - Kateg�ria - '
    
 ComboBox8.AddItem Munka12.Range("j2").Value
    ComboBox8.AddItem Munka12.Range("j3").Value
    ComboBox8.AddItem Munka12.Range("j4").Value
    ComboBox8.AddItem Munka12.Range("j5").Value
    ComboBox8.AddItem Munka12.Range("j6").Value
    ComboBox8.AddItem Munka12.Range("j7").Value
    ComboBox8.AddItem Munka12.Range("j8").Value
    ComboBox8.AddItem Munka12.Range("j9").Value
    ComboBox8.AddItem Munka12.Range("j10").Value
    ComboBox8.AddItem Munka12.Range("j11").Value
    ComboBox8.AddItem Munka12.Range("j12").Value
    ComboBox8.AddItem Munka12.Range("j13").Value
    ComboBox8.AddItem Munka12.Range("j14").Value
    ComboBox8.AddItem Munka12.Range("j15").Value
    ComboBox8.AddItem Munka12.Range("j16").Value
    ComboBox8.AddItem Munka12.Range("j17").Value
    ComboBox8.AddItem Munka12.Range("j18").Value
    ComboBox8.AddItem Munka12.Range("j19").Value
    ComboBox8.AddItem Munka12.Range("j20").Value
    ComboBox8.AddItem Munka12.Range("j21").Value
   
     ' - Ter�let - '
    
 ComboBox1.AddItem Munka12.Range("p2").Value
    ComboBox1.AddItem Munka12.Range("p3").Value
    ComboBox1.AddItem Munka12.Range("p4").Value
    ComboBox1.AddItem Munka12.Range("p5").Value
    ComboBox1.AddItem Munka12.Range("p6").Value
    ComboBox1.AddItem Munka12.Range("p7").Value
    ComboBox1.AddItem Munka12.Range("p8").Value
    ComboBox1.AddItem Munka12.Range("p9").Value
    ComboBox1.AddItem Munka12.Range("p10").Value
    ComboBox1.AddItem Munka12.Range("p11").Value
    ComboBox1.AddItem Munka12.Range("p12").Value
    ComboBox1.AddItem Munka12.Range("p13").Value
    ComboBox1.AddItem Munka12.Range("p14").Value
    ComboBox1.AddItem Munka12.Range("p15").Value
    ComboBox1.AddItem Munka12.Range("p16").Value
    ComboBox1.AddItem Munka12.Range("p17").Value
    ComboBox1.AddItem Munka12.Range("p18").Value
    ComboBox1.AddItem Munka12.Range("p19").Value
    ComboBox1.AddItem Munka12.Range("p20").Value
    ComboBox1.AddItem Munka12.Range("p21").Value
    
      ' - Csapat - '
    
 ComboBox2.AddItem Munka12.Range("m2").Value
    ComboBox2.AddItem Munka12.Range("m3").Value
    ComboBox2.AddItem Munka12.Range("m4").Value
    ComboBox2.AddItem Munka12.Range("m5").Value
    ComboBox2.AddItem Munka12.Range("m6").Value
    ComboBox2.AddItem Munka12.Range("m7").Value
    ComboBox2.AddItem Munka12.Range("m8").Value
    ComboBox2.AddItem Munka12.Range("m9").Value
    ComboBox2.AddItem Munka12.Range("m10").Value
    ComboBox2.AddItem Munka12.Range("m11").Value
    ComboBox2.AddItem Munka12.Range("m12").Value
    ComboBox2.AddItem Munka12.Range("m13").Value
    ComboBox2.AddItem Munka12.Range("m14").Value
    ComboBox2.AddItem Munka12.Range("m15").Value
    ComboBox2.AddItem Munka12.Range("m16").Value
    ComboBox2.AddItem Munka12.Range("m17").Value
    ComboBox2.AddItem Munka12.Range("m18").Value
    ComboBox2.AddItem Munka12.Range("m19").Value
    ComboBox2.AddItem Munka12.Range("m20").Value
    ComboBox2.AddItem Munka12.Range("m21").Value
    
  Jelsz�Rejt�s
End Sub
