Attribute VB_Name = "FunctionAdatokM�sol�sa4"
Option Explicit

Sub AdatokM�sol�sa4()

Sheets("alapadatok").Select
Columns("f:f").Select
Selection.End(xlDown).Select
Dim Brw As Long
Brw = ActiveCell.Row + 1
Dim Boszlop As String
Boszlop = "f"
Dim Bkoord As String
Bkoord = Boszlop & Brw
Range(Bkoord).Value = AppWindow.TextBox105.Value

End Sub
