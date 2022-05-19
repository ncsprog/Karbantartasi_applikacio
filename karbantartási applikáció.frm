VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantartási kimutatások"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "karbantartási applikáció.frx":0000
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
LétszámMásolás
LétszámÖsszesítés

If AppWindow.TextBox11 = "" Then
MsgBox "Bárcaszám megadása kötelezõ!" & vbCrLf & "Nem történt adatmentés."
AdatfelvételLista
JelszóRejtés
Exit Sub
End If

If AppWindow.TextBox11 <> "" Then
'TartalomEllenõrzés
AdatokMásolása
IDgenerálás
ID_generálás
AdatokMentése
AdatfelvételLista
Törlés
Idõszámítás
End If

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

Private Sub CommandButton2_Click()
JelszóRejtés2
GépállásokSzáma
'tb75
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
IDgenerálás3
ID_generálás3
AdatokMásolása3
JelszóRejtés
End Sub

Private Sub CommandButton6_Click()
JelszóRejtés2
If AppWindow.TextBox54 = "" Then
AdatfelvételLista2
End If
JelszóRejtés
End Sub


Private Sub CommandButton8_Click()
If AppWindow.TextBox99 = "smj266" Then
JelszóRejtés3
AppWindow.TextBox99 = ""
Else
MsgBox "Nem megfelelõ betekintési jelszó!"
AppWindow.TextBox99 = ""
End If
End Sub

Private Sub CommandButton9_Click()

If AppWindow.TextBox100 = "smj266" Then
JelszóRejtés2
AppWindow.MultiPage1.page4.Visible = True
AppWindow.MultiPage1.page5.Visible = True
AppWindow.Frame19.Visible = True
AppWindow.Frame20.Visible = True
AppWindow.Frame21.Visible = True
AppWindow.Frame23.Visible = True

AdatfelvételLista10
AdatfelvételLista11
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
AppWindow.TextBox100 = ""


End If

End Sub

Private Sub NévsorFrissítés_Click()
JelszóRejtés2
BárcaKeres
JelszóRejtés
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

Private Sub UserForm_Initialize()
JelszóRejtés2
    ' - AdatfelvételTerület - '
    ComboBox1.AddItem "67000"
    ComboBox1.AddItem "28000"
    ComboBox1.AddItem "Kovács"
    ' - AdatfelvételCsapat - '
    ComboBox2.AddItem "Team I."
    ComboBox2.AddItem "Team II."
    ComboBox2.AddItem "Team III."
    ComboBox2.AddItem "TPM"
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
    
    
  JelszóRejtés
End Sub
