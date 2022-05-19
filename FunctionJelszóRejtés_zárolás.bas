Attribute VB_Name = "FunctionJelszóRejtés_zárolás"
Option Explicit

Sub JelszóRejtés()
'JelszóRejtés2
' - rejtés - '

Munka1.Visible = xlSheetVeryHidden
Munka2.Visible = xlSheetVeryHidden
Munka3.Visible = xlSheetVeryHidden
Munka4.Visible = xlSheetVeryHidden
Munka5.Visible = xlSheetVeryHidden
Munka6.Visible = xlSheetVeryHidden
Munka7.Visible = xlSheetVeryHidden
Munka8.Visible = xlSheetVeryHidden
'Munka9.Visible = xlSheetVeryHidden  - ez a start lap, ezt ne rejtsd el!
Munka10.Visible = xlSheetVeryHidden
Munka11.Visible = xlSheetVeryHidden
Munka12.Visible = xlSheetVeryHidden
Munka14.Visible = xlSheetVeryHidden
Munka15.Visible = xlSheetVeryHidden

' - blokkolás - '

Munka1.Protect "asaguard"
Munka2.Protect "asaguard"
Munka3.Protect "asaguard"
Munka4.Protect "asaguard"
Munka5.Protect "asaguard"
Munka6.Protect "asaguard"
Munka7.Protect "asaguard"
Munka8.Protect "asaguard"
'Munka9.Protect "asaguard"  - ez a start lap, ennek külön jelszava van!
Munka10.Protect "asaguard"
Munka11.Protect "asaguard"
Munka12.Protect "asaguard"
Munka14.Protect "asaguard"
Munka15.Protect "asaguard"

'JelszóRejtés
End Sub
