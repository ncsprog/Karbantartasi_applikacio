Attribute VB_Name = "FunctionId�Kalkul�tor_inakt"
Option Explicit

Sub Id�Kalkul�tor()

Dim Drw As Integer, Dkoord As String, T1rw As Integer, T1koord As String, _
T2rw As Integer, T2koord As String, T_t�l As Integer, M_t�l As Integer, T_ig As Integer, _
M_ig As Integer, Dh As Integer, Dm As String, T3rw As Integer, T3koord As String, H As Integer, _
M As Integer, H_24 As Integer, M_60 As Integer, Csek1 As Integer, Csek2 As Integer, _
D As Integer


'id�_t�l

Munka1.Select
Columns("j:j").Select
Selection.End(xlDown).Select
T1rw = ActiveCell.Row + 1
T1koord = "j" & T1rw
Range(T1koord) = AppWindow.TextBox7

'id�_ig

Munka1.Select
Columns("k:k").Select
Selection.End(xlDown).Select
T2rw = ActiveCell.Row + 1
T2koord = "k" & T2rw
Range(T2koord) = AppWindow.TextBox6

'Csekkol�s

Csek1 = Len(AppWindow.TextBox7)
Csek2 = Len(AppWindow.TextBox6)

If Csek1 <> 5 Then
MsgBox "Kezd� id�pont form�tuma nem megfelel�! ��:pp"
Exit Sub
End If
If Csek2 <> 5 Then
MsgBox "Befejez� id�pont form�tuma nem megfelel�! ��:pp"
Exit Sub
End If


'T�l-�ra, perc

T_t�l = Left(AppWindow.TextBox7, 2)
M_t�l = Right(AppWindow.TextBox7, 2)

'Ig-�ra, perc

T_ig = Left(AppWindow.TextBox6, 2)
M_ig = Right(AppWindow.TextBox6, 2)


'�ra, perc

H = 24
M = 60

'DeltaT

Munka1.Select
Columns("l:l").Select
Selection.End(xlDown).Select
T3rw = ActiveCell.Row + 1
T3koord = "l" & T3rw

'Sz�m�t�s_�ra

If T_t�l = T_ig Then
Dh = T_ig - T_t�l
ElseIf T_t�l > T_ig Then
    If M_t�l = 0 Then
    Dh = H - T_t�l + T_ig
    Else
    Dh = H - T_t�l + T_ig - 1
    End If
ElseIf T_ig > T_t�l Then
Dh = T_ig - T_t�l
End If

'Sz�m�t�s_perc

If M_t�l = M_ig Then
D = M_ig - M_t�l
ElseIf M_t�l > M_ig Then
D = M - M_t�l + M_ig
ElseIf M_ig > M_t�l Then
D = M_ig - M_t�l
End If

If D < 10 Then
Dm = "0" & D
Else
Dm = D
End If

Range(T3koord) = Dh & ":" & Dm & " �ra"

End Sub
