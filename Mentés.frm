VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppWindow 
   Caption         =   "Karbantartási kimutatások"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   OleObjectBlob   =   "Mentés.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
' --- Ürlap kitöltésének ellenõrzése --- '
TartalomEllenõrzés

' --- Kitöltött munkalap mentése --- '
' - IDgenerálás - '
IDgenerálás

' - ID_generálás - '
ID_generálás

' - Adatok másolása - '
AdatokMásolása

' --- Kitöltött munkalap mentése --- '

End Sub
