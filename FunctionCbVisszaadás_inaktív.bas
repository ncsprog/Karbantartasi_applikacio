Attribute VB_Name = "FunctionCbVisszaad�s_inakt�v"
Option Explicit

Sub CbVisszaad�s()

Munka16.Select
Range("a:ax").Value = ""

' - m�sol - '
Munka1.Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim A As Long
A = ActiveCell.Row
Dim B As String
B = "x"
Dim C As String
C = B & A
Range("a1", C).Copy
Munka16.Select
Range("a1").PasteSpecial xlPasteValues

Columns("x:x").Select
Selection.End(xlDown).Select
Dim D As Long
D = ActiveCell.Row
Dim Dkoord As String
Dkoord = "x" & D

Range("a1").Select
    Selection.AutoFilter
    ' - st�tusz - '
    If Munka12.Range("s2") = True Then 'true/false ellen�rz�s
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b2").Value 'keres�si �rt�k
    End If
    Megbesz�l�sM�sol�
    
    If Munka12.Range("s3").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b3").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s4").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b4").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s5").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b5").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s6").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b6").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s7").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b7").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s8").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b8").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s9").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b9").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s10").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b10").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s11").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b11").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s12").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b12").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s13").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b13").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s14").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b14").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s15").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b15").Value
    End If
    Megbesz�l�sM�sol�
           
    If Munka12.Range("s16").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=16, Criteria1:=Munka12.Range("b16").Value
    End If
    Megbesz�l�sM�sol�

' - kateg�ria - '
    If Munka12.Range("s42").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j2").Value
    End If
    Megbesz�l�sM�sol�

    If Munka12.Range("s43").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j3").Value
    End If
    Megbesz�l�sM�sol�
    
    If Munka12.Range("s44").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j4").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s45").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j5").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s46").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j6").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s47").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j7").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s48").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j8").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s49").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j9").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s50").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j10").Value
    End If
    Megbesz�l�sM�sol�
    
        If Munka12.Range("s51").Value = True Then
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=24, Criteria1:=Munka12.Range("j11").Value
    End If
    Megbesz�l�sM�sol�
  ' - ter�let - '
    ActiveSheet.Range("a1", Dkoord).AutoFilter Field:=8, Criteria1:=Munka12.Range("p2").Value
    Megbesz�l�sM�sol�
  
Range("a1").Select
Selection.AutoFilter

End Sub
