Attribute VB_Name = "FunctionAdatokMásolása4"
Option Explicit

Sub AdatokMásolása4()

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
