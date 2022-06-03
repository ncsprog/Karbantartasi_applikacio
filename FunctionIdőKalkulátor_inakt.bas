Attribute VB_Name = "FunctionIdõKalkulátor_inakt"
Option Explicit

Sub IdõKalkulátor()

Dim Drw As Integer, Dkoord As String, T1rw As Integer, T1koord As String, _
T2rw As Integer, T2koord As String, T_tõl As Integer, M_tõl As Integer, T_ig As Integer, _
M_ig As Integer, Dh As Integer, Dm As String, T3rw As Integer, T3koord As String, H As Integer, _
M As Integer, H_24 As Integer, M_60 As Integer, Csek1 As Integer, Csek2 As Integer, _
D As Integer


'idõ_tól

Munka1.Select
Columns("j:j").Select
Selection.End(xlDown).Select
T1rw = ActiveCell.Row + 1
T1koord = "j" & T1rw
Range(T1koord) = AppWindow.TextBox7

'idõ_ig

Munka1.Select
Columns("k:k").Select
Selection.End(xlDown).Select
T2rw = ActiveCell.Row + 1
T2koord = "k" & T2rw
Range(T2koord) = AppWindow.TextBox6

'Csekkolás

Csek1 = Len(AppWindow.TextBox7)
Csek2 = Len(AppWindow.TextBox6)

If Csek1 <> 5 Then
MsgBox "Kezdõ idõpont formátuma nem megfelelõ! óó:pp"
Exit Sub
End If
If Csek2 <> 5 Then
MsgBox "Befejezõ idõpont formátuma nem megfelelõ! óó:pp"
Exit Sub
End If


'Tól-óra, perc

T_tõl = Left(AppWindow.TextBox7, 2)
M_tõl = Right(AppWindow.TextBox7, 2)

'Ig-óra, perc

T_ig = Left(AppWindow.TextBox6, 2)
M_ig = Right(AppWindow.TextBox6, 2)


'Óra, perc

H = 24
M = 60

'DeltaT

Munka1.Select
Columns("l:l").Select
Selection.End(xlDown).Select
T3rw = ActiveCell.Row + 1
T3koord = "l" & T3rw

'Számítás_óra

If T_tõl = T_ig Then
Dh = T_ig - T_tõl
ElseIf T_tõl > T_ig Then
    If M_tõl = 0 Then
    Dh = H - T_tõl + T_ig
    Else
    Dh = H - T_tõl + T_ig - 1
    End If
ElseIf T_ig > T_tõl Then
Dh = T_ig - T_tõl
End If

'Számítás_perc

If M_tõl = M_ig Then
D = M_ig - M_tõl
ElseIf M_tõl > M_ig Then
D = M - M_tõl + M_ig
ElseIf M_ig > M_tõl Then
D = M_ig - M_tõl
End If

If D < 10 Then
Dm = "0" & D
Else
Dm = D
End If

Range(T3koord) = Dh & ":" & Dm & " óra"

End Sub
