Attribute VB_Name = "FunctionM�szakHat�roz�"
Option Explicit

Sub M�szakHat�roz�()

Dim Jrw As Long, ColJ As String, Jkoord As String, Ht�l As Integer, Mkoord As String, ColM As String

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

Ht�l = Left(AppWindow.TextBox7, 2)

End Sub
