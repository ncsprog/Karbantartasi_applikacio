Attribute VB_Name = "FunctionAdatokMentése"
Option Explicit

Sub AdatokMentése()
'JelszóRejtés2

ActiveWorkbook.Save
MsgBox "Másolás, mentés kész!"

Sheets("Start").Select
Range("b2").Select
'JelszóRejtés
End Sub
