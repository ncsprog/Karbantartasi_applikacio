VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantart�si kimutat�sok"
   ClientHeight    =   13404
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   21768
   OleObjectBlob   =   "karbantart�si applik�ci�.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adatfelv�telMent�s_Click()


L�tsz�mM�sol�s
L�tsz�m�sszes�t�s

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


Private Sub CommandButton2_Click()
'tb75

End Sub

Private Sub CommandButton3_Click()
'tb76

End Sub

Private Sub CommandButton4_Click()
'tb77

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


Private Sub UserForm_Initialize()

    ' - Adatfelv�telTer�let - '
    ComboBox1.AddItem "67000"
    ComboBox1.AddItem "28000"
    ComboBox1.AddItem "Kov�cs"
    ' - Adatfelv�telCsapat - '
    ComboBox2.AddItem "Team I."
    ComboBox2.AddItem "Team II."
    ComboBox2.AddItem "Team III."
    ComboBox2.AddItem "TPM"
    ' - Adatfelv�telM�r�s - '
    ComboBox3.AddItem "Igen"
    ComboBox3.AddItem "Nem"
    ' - Adatfelv�telSt�tusz - '
    ComboBox4.AddItem "Alkatr�szre v�r"
    ComboBox4.AddItem "Azonos�tott hiba - jav�t�s alatt"
    ComboBox4.AddItem "G�pkezel� megoldotta"
    ComboBox4.AddItem "Hibakeres�s"
    ComboBox4.AddItem "Jav�t�s n�lk�l lez�rva"
    ComboBox4.AddItem "Jav�tott"
    ComboBox4.AddItem "Nem megkezdett jav�t�s"
    ComboBox4.AddItem "�sszeszerel�sre v�r"
    ComboBox4.AddItem "�sszeszerelve tesztre v�r"
    ComboBox4.AddItem "T�r�lve"
    ' - Nyomonk�vet�St�tusz - '
    ComboBox5.AddItem "Alkatr�szre v�r"
    ComboBox5.AddItem "Azonos�tott hiba - jav�t�s alatt"
    ComboBox5.AddItem "G�pkezel� megoldotta"
    ComboBox5.AddItem "Hibakeres�s"
    ComboBox5.AddItem "Jav�t�s n�lk�l lez�rva"
    ComboBox5.AddItem "Jav�tott"
    ComboBox5.AddItem "Nem megkezdett jav�t�s"
    ComboBox5.AddItem "�sszeszerel�sre v�r"
    ComboBox5.AddItem "�sszeszerelve tesztre v�r"
    ComboBox5.AddItem "T�r�lve"
    ' - Nyomonk�vet�M�r�s - '
    ComboBox6.AddItem "Igen"
    ComboBox6.AddItem "Nem"
    ' - Nyomonk�vet�Felel�s - '
    ComboBox7.AddItem "Csorba D�vid"
    ComboBox7.AddItem "D�nes Andr�s"
    ComboBox7.AddItem "Gajdos P�ter"
    ComboBox7.AddItem "Haik D�vid"
    ComboBox7.AddItem "Kiss M�t�"
    ComboBox7.AddItem "K�nyi Istv�n"
    ComboBox7.AddItem "Nedvesi Csaba P�ter"
    ComboBox7.AddItem "N�meth Szil�rd"
    ComboBox7.AddItem "Tak�cs D�vid"
    ComboBox7.AddItem "Tak�cs R�bert"
    ComboBox7.AddItem "Tak�cs Tivadar"
    ComboBox7.AddItem "Toth S�ndor"
  
End Sub
