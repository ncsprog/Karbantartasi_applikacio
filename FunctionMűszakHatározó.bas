Attribute VB_Name = "FunctionMûszakHatározó"
Option Explicit

Sub MûszakHatározó()

Dim Jrw As Long, ColJ As String, Jkoord As String, Htõl As Integer, Mkoord As String, ColM As String

ColJ = "j"
ColM = "M"

' - J - '
Munka1.Select
Columns("j:j").Select
Selection.End(xlDown).Select
Jrw = ActiveCell.Row
Jkoord = ColJ & Jrw
' - M - '
Mkoord = ColM & Jrw

Htõl = Left(AppWindow.TextBox7, 2)

End Sub
