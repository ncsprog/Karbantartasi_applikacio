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
End Sub



