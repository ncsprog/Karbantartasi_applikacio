VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantart�si kimutat�sok"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "karbantart�si applik�ci�.frx":0000
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
L�tsz�mM�sol�s
L�tsz�m�sszes�t�s

If AppWindow.TextBox11 = "" Then
MsgBox "B�rcasz�m megad�sa k�telez�!" & vbCrLf & "Nem t�rt�nt adatment�s."
Adatfelv�telLista
Jelsz�Rejt�s
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
Id�sz�m�t�s
End If

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
End If
Jelsz�Rejt�s
End Sub

Private Sub CommandButton2_Click()
Jelsz�Rejt�s2
G�p�ll�sokSz�ma
'tb75
Jelsz�Rejt�s
End Sub

Private Sub CommandButton3_Click()
Jelsz�Rejt�s2
IdegenXL2
Adatfelv�telLista8_�1
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
IDgener�l�s3
ID_gener�l�s3
AdatokM�sol�sa3
Jelsz�Rejt�s
End Sub

Private Sub CommandButton6_Click()
Jelsz�Rejt�s2
If AppWindow.TextBox54 = "" Then
Adatfelv�telLista2
End If
Jelsz�Rejt�s
End Sub


Private Sub CommandButton7_Click()
If AppWindow.TextBox98 = "jelsz�" Or AppWindow.TextBox98 = "password" Then

AppWindow.MultiPage1.page4.Visible = True
Else
AppWindow.MultiPage1.page4.Visible = False
End If

End Sub

Private Sub CommandButton8_Click()
If AppWindow.TextBox99 = "smj266" Then
Jelsz�Rejt�s3
Else
MsgBox "Nem megfelel� betekint�si jelsz�!"
End If
End Sub

Private Sub CommandButton9_Click()

If AppWindow.TextBox100 = "smj266" Then
Jelsz�Rejt�s2
AppWindow.MultiPage1.page4.Visible = True
AppWindow.MultiPage1.page5.Visible = True
AppWindow.Frame19.Visible = True
AppWindow.Frame20.Visible = True
Adatfelv�telLista10
Adatfelv�telLista11
Jelsz�Rejt�s
Else
MsgBox "Nem megfelel� jelsz�!"
AppWindow.MultiPage1.page4.Visible = False
AppWindow.MultiPage1.page5.Visible = False
AppWindow.Frame19.Visible = False
AppWindow.Frame20.Visible = False
End If

End Sub

Private Sub N�vsorFriss�t�s_Click()
Jelsz�Rejt�s2
B�rcaKeres
Jelsz�Rejt�s
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




Private Sub UserForm_Initialize()
Jelsz�Rejt�s2
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
    ComboBox4.AddItem Munka12.Range("a2").Value
    ComboBox4.AddItem Munka12.Range("a3").Value
    ComboBox4.AddItem Munka12.Range("a4").Value
    ComboBox4.AddItem Munka12.Range("a5").Value
    ComboBox4.AddItem Munka12.Range("a6").Value
    ComboBox4.AddItem Munka12.Range("a7").Value
    ComboBox4.AddItem Munka12.Range("a8").Value
    ComboBox4.AddItem Munka12.Range("a9").Value
    ComboBox4.AddItem Munka12.Range("a10").Value
    ComboBox4.AddItem Munka12.Range("a11").Value
    ComboBox4.AddItem Munka12.Range("a12").Value
    ComboBox4.AddItem Munka12.Range("a13").Value
  
    ' - Nyomonk�vet�St�tusz - '
    ComboBox5.AddItem Munka12.Range("a2").Value
    ComboBox5.AddItem Munka12.Range("a3").Value
    ComboBox5.AddItem Munka12.Range("a4").Value
    ComboBox5.AddItem Munka12.Range("a5").Value
    ComboBox5.AddItem Munka12.Range("a6").Value
    ComboBox5.AddItem Munka12.Range("a7").Value
    ComboBox5.AddItem Munka12.Range("a8").Value
    ComboBox5.AddItem Munka12.Range("a9").Value
    ComboBox5.AddItem Munka12.Range("a10").Value
    ComboBox5.AddItem Munka12.Range("a11").Value
    ComboBox5.AddItem Munka12.Range("a12").Value

    ' - Nyomonk�vet�M�r�s - '
    ComboBox6.AddItem "Igen"
    ComboBox6.AddItem "Nem"
    ' - Nyomonk�vet�Felel�s - '
    ComboBox7.AddItem Munka12.Range("c2").Value
    ComboBox7.AddItem Munka12.Range("c3").Value
    ComboBox7.AddItem Munka12.Range("c4").Value
    ComboBox7.AddItem Munka12.Range("c5").Value
    ComboBox7.AddItem Munka12.Range("c6").Value
    ComboBox7.AddItem Munka12.Range("c7").Value
    ComboBox7.AddItem Munka12.Range("c8").Value
    ComboBox7.AddItem Munka12.Range("c9").Value
    ComboBox7.AddItem Munka12.Range("c10").Value
    ComboBox7.AddItem Munka12.Range("c11").Value
    ComboBox7.AddItem Munka12.Range("c12").Value
    ComboBox7.AddItem Munka12.Range("c13").Value
    
    
  Jelsz�Rejt�s
End Sub
