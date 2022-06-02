VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantartási adatgyûjtõ"
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

Private Sub AdatfelvételMentés_Click()
JelszóRejtés2
If AppWindow.TextBox11 = "" Then
MsgBox "Bárcaszám megadása kötelezõ!"
AdatfelvételLista
Exit Sub
End If

'If AppWindow.ComboBox8 = "" Then
If AppWindow.TextBox11 <> "" Then
ElseIf AppWindow.TextBox10 = "" Then
MsgBox "Kategóriát választani kötelezõ!" & vbCrLf & "vagy" & vbCrLf & "Rábaszám hiányzik" & vbCrLf & vbCrLf & "Nem történt adatmentés."
Exit Sub
ElseIf AppWindow.ComboBox8 = "" Then
MsgBox "Kategóriát választani kötelezõ!" & vbCrLf & "vagy" & vbCrLf & "Rábaszám hiányzik" & vbCrLf & vbCrLf & "Nem történt adatmentés."
AdatfelvételLista
JelszóRejtés
Exit Sub
End If

'If AppWindow.TextBox11 <> "" Then
'TartalomEllenõrzés
AdatokMásolása
IDgenerálás
ID_generálás
IdõKalkulátor
AdatokMentése
AdatfelvételLista
Törlés
'Idõszámítás
'End If

BoxInaktíváló

JelszóRejtés
End Sub


Private Sub CommandButton1_Click()
JelszóRejtés2
If AppWindow.TextBox73 = "" Then
MsgBox "kérek egy gépszámot"
'AdatfelvételLista3
End If
If AppWindow.TextBox73 <> "" Then
GépTörténet
End If

JelszóRejtés
End Sub





Private Sub CommandButton10_Click()
JelszóRejtés2
If AppWindow.TextBox101 = "" Then
MsgBox "Nincs megadva új státusz."
Else
StátuszSzerkesztés
Rendezés
AdatfelvételLista10
CbFeltöltés
End If

JelszóRejtés
End Sub

Private Sub CommandButton12_Click()
JelszóRejtés2
If ListBox29.ListCount = 0 Or ListBox29.ListCount < 0 Or ListBox29.ListCount > 21 Then
MsgBox "Nincs kiválasztott státusz."
Exit Sub
Else
StátuszSzerkeszt2
Rendezés
CbFeltöltés
End If
AdatfelvételLista10

JelszóRejtés
End Sub

Private Sub CommandButton13_Click()


JelszóRejtés2
If ListBox30.ListCount = 0 Then
MsgBox "Nincs kiválasztott sor."
Else
FelelõsSzerkesztés2
AdatfelvételLista11
Rendezés2
CbFeltöltés
End If

JelszóRejtés
End Sub

Private Sub CommandButton14_Click()
JelszóRejtés2
If AppWindow.TextBox102 = "" Then
MsgBox "Nincs megadva új státusz."
Else
FelelõsSzerkesztés
Rendezés2
AdatfelvételLista11
CbFeltöltés
End If
JelszóRejtés
End Sub

Private Sub CommandButton15_Click()

If AppWindow.TextBox105 = "" Then
MsgBox "Kérlek add meg az aktuális összlétszámot."
JelszóRejtés2
AdatfelvételLista12
JelszóRejtés
Exit Sub
Else
JelszóRejtés2
AdatokMásolása4
ID_generálás4
AppWindow.TextBox105 = ""
AdatfelvételLista12
End If
JelszóRejtés
End Sub

Private Sub CommandButton16_Click()
JelszóRejtés2
If AppWindow.TextBox106 = "" Then
BárcaKeres_Frissít
Else
BárcaKeres_Keres
End If
JelszóRejtés
End Sub

Private Sub CommandButton17_Click()

JelszóRejtés2
If AppWindow.TextBox107 = "" Then
GépKeres
Else
GépKeres2
'BárcaKeres_Keres
End If
JelszóRejtés




End Sub

Private Sub CommandButton18_Click()
JelszóRejtés2
If AppWindow.TextBox108 = "" Then
MsgBox "Nincs megadva új státusz."
Else
KategóriaSzerkeszt
Rendezés3
AdatfelvételLista14
CbFeltöltés
End If

JelszóRejtés
End Sub

Private Sub CommandButton19_Click()
JelszóRejtés2
If ListBox37.ListCount = 0 Or ListBox37.ListCount < 0 Or ListBox37.ListCount > 21 Then
MsgBox "Nincs kiválasztott kategória."
Exit Sub
Else
KategóriaSzerkeszt2
Rendezés3
CbFeltöltés
End If
AdatfelvételLista14

JelszóRejtés
End Sub

Private Sub CommandButton2_Click()
JelszóRejtés2
GépállásokSzáma
'tb75
JelszóRejtés
End Sub

Private Sub CommandButton20_Click()
JelszóRejtés2
If AppWindow.TextBox109 = "" Then
MsgBox "Nincs megadva új terület."
Else
TerületSzerkeszt
Rendezés5
AdatfelvételLista15
CbFeltöltés
End If

JelszóRejtés
End Sub

Private Sub CommandButton21_Click()
JelszóRejtés2
If ListBox38.ListCount = 0 Or ListBox38.ListCount < 0 Or ListBox38.ListCount > 21 Then
MsgBox "Nincs kiválasztott státusz."
Exit Sub
Else
TerületSzerkeszt2
Rendezés
CbFeltöltés
End If
AdatfelvételLista15

JelszóRejtés
End Sub

Private Sub CommandButton22_Click()
JelszóRejtés2
If AppWindow.TextBox110 = "" Then
MsgBox "Nincs megadva új csapat."
Else
CsapatSzerkeszt
Rendezés4
AdatfelvételLista16
CbFeltöltés
End If

JelszóRejtés
End Sub

Private Sub CommandButton23_Click()
JelszóRejtés2
If ListBox39.ListCount = 0 Or ListBox39.ListCount < 0 Or ListBox39.ListCount > 21 Then
MsgBox "Nincs kiválasztott csapat."
Exit Sub
Else
CsapatSzerkeszt2
Rendezés4
CbFeltöltés
End If
AdatfelvételLista16

JelszóRejtés
End Sub

Private Sub CommandButton24_Click()
JelszóRejtés2
CbRögzít
JelszóRejtés
End Sub

Private Sub CommandButton25_Click()
JelszóRejtés2
CbFeltöltés
CbVisszaadás
JelszóRejtés
End Sub

Private Sub CommandButton26_Click()
JelszóRejtés2
MegbeszélésDátumok
MegbeszélésMásoló
JelszóRejtés
End Sub

Private Sub CommandButton27_Click()
JelszóRejtés2
MunkaSzerkesztés2
BoxAktíváló
JelszóRejtés
End Sub

Private Sub CommandButton3_Click()
JelszóRejtés2
IdegenXL2
IdegenXL3
AdatfelvételLista8_Á1
AdatfelvételLista9_R1
AdatfelvételLista13
'AppWindow.TextBox76 = "Másolás kész"
'tb76
JelszóRejtés
End Sub

Private Sub CommandButton4_Click()
JelszóRejtés2
IdegenXL
AdatfelvételLista4
AdatfelfételLista5
AdatfelvételLista6
AdatfelvételLista7
'AppWindow.TextBox77.Value = "Másolás kész"

'tb77
JelszóRejtés
End Sub


Private Sub CommandButton5_Click()
JelszóRejtés2
IDgenerálás2
IDgenerálás3
ID_generálás2
ID_generálás3
AdatokMásolása3
AdatokMásolása5
LétszámÖsszesítés
JelszóRejtés
End Sub

Private Sub CommandButton6_Click()
JelszóRejtés2
If AppWindow.ComboBox5 = "" Then
MsgBox "Kérlek válaszd ki a státuszt."
Exit Sub
Else
AdatfelvételLista2
End If
JelszóRejtés
End Sub


Private Sub CommandButton8_Click()
JelszóRejtés3
End Sub

Private Sub CommandButton9_Click()

If AppWindow.TextBox100 = "smj266" Then
JelszóRejtés2
TbBetöltés
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
AdatfelvételLista10
AdatfelvételLista11
AdatfelvételLista14
AdatfelvételLista15
AdatfelvételLista16
'CbFeltöltés
'CbVisszaadás
JelszóRejtés
AppWindow.TextBox100 = ""
Else
MsgBox "Nem megfelelõ jelszó!"
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



Private Sub NyomonkövetõFrissítés_Click()
JelszóRejtés2
If AppWindow.TextBox62 <> "" Then

'TartalomEllenõrzés2
IDgenerálás
ID_generálás
AdatokMásolása2
AdatokMentése
Törlés2
AdatfelvételLista2
End If
JelszóRejtés
End Sub

Private Sub NyomonkövetõSzerkesztés_Click()
JelszóRejtés2
If ListBox20.ListCount = 0 Then
MsgBox "Nincs kiválasztott sor." & vbCrLf & "A lista megjenítéshez kattints a FRISSÍTÉS gombra."
Exit Sub
Else
MunkaSzerkesztés
End If
JelszóRejtés
End Sub




Private Sub TextBox103_Change()

End Sub

Private Sub TextBox112_Change()

AppWindow.TextBox112 = TextBox115 + TextBox114 + TextBox113

End Sub


Private Sub UserForm_Initialize()
JelszóRejtés2
    ' - AdatfelvételTerület - '

    ' - AdatfelvételCsapat - '

    ' - AdatfelvételMérés - '
    ComboBox3.AddItem "Igen"
    ComboBox3.AddItem "Nem"
    ' - AdatfelvételStátusz - '
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
    
  
    ' - NyomonkövetõStátusz - '
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
    

    ' - NyomonkövetõMérés - '
    ComboBox6.AddItem "Igen"
    ComboBox6.AddItem "Nem"
    ' - NyomonkövetõFelelõs - '
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
    
    ' - Kategória - '
    
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
   
     ' - Terület - '
    
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
    
  JelszóRejtés
End Sub
