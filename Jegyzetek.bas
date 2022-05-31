Attribute VB_Name = "Jegyzetek"
Option Explicit

'listbox tabulálás

'full    65;65;95;65;65;240;40;65;65;50;50;85;35;50;50;180;40;195;95;95;95;165
'adatfelvétel    0;65;95;65;65;240;40;65;65;50;50;85;35;50;50;180;40;0;0;0;0;0
'nyomonkövetõ    0;65;95;65;65;240;40;65;65;50;50;85;35;50;50;180;40;195;95;95;95;0

'minden, nem publikus fület rejtsen, jelszavazzon a kód oda-vissza

'IDgenerálás3 - megbeszéléshez
'ID_generálás3 - megbeszéléshez
'AdatokMásolása3 - megbeszéléshez



'Kulcsgépek ==> Rendelkezésre állás és összállásidõ idõszakra.xlsx  ==> FunctionIdegenXL2_Állásidõ  ==> transfer_kulcsgép   ==> LB27    ==> FunctionAdatfelvételLista8_Á1
'Rendelkezésreállás ==> Állásidõ adott idõszakban.xlsx  ==> FunctionIdegenXL3_Rendelkezés   ==> transfer_rendelkezés    ==> LB31    ==> FunctionAdatfelvételLista9_R1
'Gazdasági  ==> gazdasági lekérdezett adatok.xlsx   ==> FunctionIdegenXL_gazdasági  ==> transfer_gazdasági  ==> LB23,   LB24,   LB25,   LB26,   ==> FunctionAdatfelvételLista4_G1,  FunctionAdatfelvételLista5_2,  FunctionAdatfelvételList6_G3,  FunctionAdatfelvételLista7_G4


'Külsõ filok elérhetõségét a véglegesítéskor pontosan megadni!
Sub pw()
'JelszóRejtés2

Worksheets("adatok").Visible = True
Munka3.Unprotect "asguard"
Worksheets("szûrõ_transfer").Visible = True
Munka3.Unprotect "asguard"

Worksheets("szûrõ_transfer").Visible = False
Munka3.Protect "asguard"
Worksheets("adatok").Visible = False
Munka3.Protect "asguard"
'JelszóRejtés

Range("a1").Select
ActiveSheet.Range("A1", A).AutoFilter Field:=3, Operator:= _
        xlFilterValues, Criteria2:=Dif
' - terület - TEAM 1. - '
ActiveSheet.Range("A1", A).AutoFilter Field:=9, Operator:= _
        xlFilterValues, Criteria1:="Team1."
        
  Range("C14").Select
    Selection.AutoFilter
Range("a1", A).Copy
Range("z1").PasteSpecial xlPasteValues


If Munka12.Range("y2").Value = "" And Munka12.Range("y3").Value = "" Then
MsgBox "Nincs meghatározva idõ intervallum. Alapérték 1 nap."

Range("a1").Select
ActiveSheet.Range("A1", A).AutoFilter Field:=3, Operator:= _
        xlFilterValues, Criteria1:=Array(Dday, D0)
' - terület - TEAM 1. - '
ActiveSheet.Range("A1", A).AutoFilter Field:=9, Operator:= _
        xlFilterValues, Criteria1:="Team1."
        
  Range("C14").Select
    Selection.AutoFilter
        
ElseIf Munka12.Range("y2").Value <> "" And Munka12.Range("y3").Value <> "" Then
MsgBox "Vizsgált intervallum: " & DR & " nap."







End Sub



