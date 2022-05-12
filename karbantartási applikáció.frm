VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantartási kimutatások"
   ClientHeight    =   13404
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   21768
   OleObjectBlob   =   "karbantartási applikáció.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AdatfelvételMentés_Click()


LétszámMásolás
LétszámÖsszesítés

If AppWindow.TextBox11 = "" Then
MsgBox "Bárcaszám megadása kötelezõ!" & vbCrLf & "Nem történt adatmentés."
AdatfelvételLista
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
End If


End Sub

Private Sub CommandButton1_Click()

If AppWindow.TextBox73 = "" Then
MsgBox "kérek egy gépszámot"
'AdatfelvételLista3
End If
If AppWindow.TextBox73 <> "" Then
GépTörténet
End If


End Sub


Private Sub CommandButton2_Click()
'tb75

End Sub

Private Sub CommandButton3_Click()
'tb76

End Sub

Private Sub CommandButton4_Click()
'tb77

End Sub


Private Sub NévsorFrissítés_Click()

BárcaKeres

End Sub

Private Sub NyomonkövetõFrissítés_Click()

If AppWindow.TextBox54 = "" Then
AdatfelvételLista2
End If
If AppWindow.TextBox62 <> "" Then
'TartalomEllenõrzés2
IDgenerálás
ID_generálás
AdatokMásolása2
AdatokMentése
Törlés2
AdatfelvételLista2
End If

End Sub

Private Sub NyomonkövetõSzerkesztés_Click()

If ListBox20.ListCount = 0 Then
MsgBox "Nincs kiválasztott sor." & vbCrLf & "A lista megjenítéshez kattints a FRISSÍTÉS gombra."
Exit Sub
Else
MunkaSzerkesztés
End If
End Sub


Private Sub UserForm_Initialize()

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
    ComboBox4.AddItem "Alkatrészre vár"
    ComboBox4.AddItem "Azonosított hiba - javítás alatt"
    ComboBox4.AddItem "Gépkezelõ megoldotta"
    ComboBox4.AddItem "Hibakeresés"
    ComboBox4.AddItem "Javítás nélkül lezárva"
    ComboBox4.AddItem "Javított"
    ComboBox4.AddItem "Nem megkezdett javítás"
    ComboBox4.AddItem "Összeszerelésre vár"
    ComboBox4.AddItem "Összeszerelve tesztre vár"
    ComboBox4.AddItem "Törölve"
    ' - NyomonkövetõStátusz - '
    ComboBox5.AddItem "Alkatrészre vár"
    ComboBox5.AddItem "Azonosított hiba - javítás alatt"
    ComboBox5.AddItem "Gépkezelõ megoldotta"
    ComboBox5.AddItem "Hibakeresés"
    ComboBox5.AddItem "Javítás nélkül lezárva"
    ComboBox5.AddItem "Javított"
    ComboBox5.AddItem "Nem megkezdett javítás"
    ComboBox5.AddItem "Összeszerelésre vár"
    ComboBox5.AddItem "Összeszerelve tesztre vár"
    ComboBox5.AddItem "Törölve"
    ' - NyomonkövetõMérés - '
    ComboBox6.AddItem "Igen"
    ComboBox6.AddItem "Nem"
    ' - NyomonkövetõFelelõs - '
    ComboBox7.AddItem "Csorba Dávid"
    ComboBox7.AddItem "Dénes András"
    ComboBox7.AddItem "Gajdos Péter"
    ComboBox7.AddItem "Haik Dávid"
    ComboBox7.AddItem "Kiss Máté"
    ComboBox7.AddItem "Kónyi István"
    ComboBox7.AddItem "Nedvesi Csaba Péter"
    ComboBox7.AddItem "Németh Szilárd"
    ComboBox7.AddItem "Takács Dávid"
    ComboBox7.AddItem "Takács Róbert"
    ComboBox7.AddItem "Takács Tivadar"
    ComboBox7.AddItem "Toth Sándor"
  
End Sub
