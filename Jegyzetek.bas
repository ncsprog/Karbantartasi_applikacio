Attribute VB_Name = "Jegyzetek"
Option Explicit

'listbox tabul�l�s

'full    65;65;95;65;65;240;40;65;65;50;50;85;35;50;50;180;40;195;95;95;95;165
'adatfelv�tel    0;65;95;65;65;240;40;65;65;50;50;85;35;50;50;180;40;0;0;0;0;0
'nyomonk�vet�    0;65;95;65;65;240;40;65;65;50;50;85;35;50;50;180;40;195;95;95;95;0

'minden, nem publikus f�let rejtsen, jelszavazzon a k�d oda-vissza

'IDgener�l�s3 - megbesz�l�shez
'ID_gener�l�s3 - megbesz�l�shez
'AdatokM�sol�sa3 - megbesz�l�shez



'Kulcsg�pek ==> Rendelkez�sre �ll�s �s �ssz�ll�sid� id�szakra.xlsx  ==> FunctionIdegenXL2_�ll�sid�  ==> transfer_kulcsg�p   ==> LB27    ==> FunctionAdatfelv�telLista8_�1
'Rendelkez�sre�ll�s ==> �ll�sid� adott id�szakban.xlsx  ==> FunctionIdegenXL3_Rendelkez�s   ==> transfer_rendelkez�s    ==> LB31    ==> FunctionAdatfelv�telLista9_R1
'Gazdas�gi  ==> gazdas�gi lek�rdezett adatok.xlsx   ==> FunctionIdegenXL_gazdas�gi  ==> transfer_gazdas�gi  ==> LB23,   LB24,   LB25,   LB26,   ==> FunctionAdatfelv�telLista4_G1,  FunctionAdatfelv�telLista5_2,  FunctionAdatfelv�telList6_G3,  FunctionAdatfelv�telLista7_G4


'K�ls� filok el�rhet�s�g�t a v�gleges�t�skor pontosan megadni!
Sub pw()
'Jelsz�Rejt�s2

Worksheets("adatok").Visible = True
Munka3.Unprotect "asguard"
Worksheets("sz�r�_transfer").Visible = True
Munka3.Unprotect "asguard"

Worksheets("sz�r�_transfer").Visible = False
Munka3.Protect "asguard"
Worksheets("adatok").Visible = False
Munka3.Protect "asguard"
'Jelsz�Rejt�s

Range("a1").Select
ActiveSheet.Range("A1", A).AutoFilter Field:=3, Operator:= _
        xlFilterValues, Criteria2:=Dif
' - ter�let - TEAM 1. - '
ActiveSheet.Range("A1", A).AutoFilter Field:=9, Operator:= _
        xlFilterValues, Criteria1:="Team1."
        
  Range("C14").Select
    Selection.AutoFilter
Range("a1", A).Copy
Range("z1").PasteSpecial xlPasteValues


If Munka12.Range("y2").Value = "" And Munka12.Range("y3").Value = "" Then
MsgBox "Nincs meghat�rozva id� intervallum. Alap�rt�k 1 nap."

Range("a1").Select
ActiveSheet.Range("A1", A).AutoFilter Field:=3, Operator:= _
        xlFilterValues, Criteria1:=Array(Dday, D0)
' - ter�let - TEAM 1. - '
ActiveSheet.Range("A1", A).AutoFilter Field:=9, Operator:= _
        xlFilterValues, Criteria1:="Team1."
        
  Range("C14").Select
    Selection.AutoFilter
        
ElseIf Munka12.Range("y2").Value <> "" And Munka12.Range("y3").Value <> "" Then
MsgBox "Vizsg�lt intervallum: " & DR & " nap."







End Sub



