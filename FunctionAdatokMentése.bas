Attribute VB_Name = "FunctionAdatokMent�se"
Option Explicit

Sub AdatokMent�se()

ActiveWorkbook.Save
MsgBox "M�sol�s, ment�s k�sz!"

Sheets("Start").Select
Range("b2").Select
End Sub
