VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantart�si kimutat�sok"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "karbantart�si applik�ci�.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adatfelv�telMent�s_Click()

If AppWindow.TextBox11 = "" Then
MsgBox "B�rcasz�m megad�sa k�telez�!" & vbCrLf & "Nem t�rt�nt adatment�s."
Adatfelv�telLista
Exit Sub
End If

If AppWindow.TextBox11 <> "" Then
'TartalomEllen�rz�s
AdatokM�sol�sa
IDgener�l�s
ID_gener�l�s
AdatokMent�se
Adatfelv�telLista
T�rl�s
End If


End Sub

Private Sub CommandButton1_Click()

If AppWindow.TextBox73 = "" Then
MsgBox "k�rek egy g�psz�mot"
'Adatfelv�telLista3
End If
If AppWindow.TextBox73 <> "" Then
G�pT�rt�net
End If


End Sub


Private Sub N�vsorFriss�t�s_Click()

B�rcaKeres

End Sub

Private Sub Nyomonk�vet�Friss�t�s_Click()

If AppWindow.TextBox54 = "" Then
Adatfelv�telLista2
End If
If AppWindow.TextBox62 <> "" Then
'TartalomEllen�rz�s2
IDgener�l�s
ID_gener�l�s
AdatokM�sol�sa2
AdatokMent�se
T�rl�s2
Adatfelv�telLista2
End If

End Sub

Private Sub Nyomonk�vet�Szerkeszt�s_Click()

If ListBox20.ListCount = 0 Then
MsgBox "Nincs kiv�lasztott sor." & vbCrLf & "A lista megjen�t�shez kattints a FRISS�T�S gombra."
Exit Sub
Else
MunkaSzerkeszt�s
End If
End Sub
