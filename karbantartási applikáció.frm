VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantartási kimutatások"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "karbantartási applikáció.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AdatfelvételMentés_Click()

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
