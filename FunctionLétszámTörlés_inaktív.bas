Attribute VB_Name = "FunctionL�tsz�mT�rl�s_inakt�v"
Option Explicit

Sub L�tsz�mT�rl�s()


Sheets("L�tsz�m").Select
Columns("b:b").Select
Selection.End(xlDown).Select
Dim D�tumellen�rz�s As Date
D�tumellen�rz�s = ActiveCell.Value

If Date <> D�tumellen�rz�s Then

AppWindow.TextBox41 = ""
AppWindow.TextBox38 = ""
AppWindow.TextBox37 = ""
AppWindow.TextBox40 = ""
AppWindow.TextBox31 = ""
AppWindow.TextBox30 = ""
AppWindow.TextBox39 = ""
AppWindow.TextBox29 = ""
AppWindow.TextBox28 = ""
AppWindow.TextBox36 = ""
AppWindow.TextBox35 = ""
AppWindow.TextBox34 = ""
AppWindow.TextBox27 = ""
AppWindow.TextBox25 = ""
AppWindow.TextBox19 = ""
AppWindow.TextBox26 = ""
AppWindow.TextBox24 = ""
AppWindow.TextBox18 = ""
AppWindow.TextBox33 = ""
AppWindow.TextBox44 = ""
AppWindow.TextBox47 = ""
AppWindow.TextBox21 = ""
AppWindow.TextBox43 = ""
AppWindow.TextBox46 = ""
AppWindow.TextBox20 = ""
AppWindow.TextBox42 = ""
AppWindow.TextBox45 = ""
AppWindow.TextBox32 = ""
AppWindow.TextBox23 = ""
AppWindow.TextBox22 = ""
Else
Exit Sub
End If


Sheets("Start").Select
Range("b2").Select
End Sub
