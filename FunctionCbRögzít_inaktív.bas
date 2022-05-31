Attribute VB_Name = "FunctionCbRögzít_inaktív"
Option Explicit

Sub CbRögzít()
'  - intervallum dátumok - '
If AppWindow.TextBox133 And AppWindow.TextBox134 = "" Then
Munka12.Range("y2").Value = ""
Munka12.Range("y3").Value = ""
Else
Munka12.Range("y2").Value = AppWindow.TextBox133.Value
Munka12.Range("y3").Value = AppWindow.TextBox134.Value
End If

' - checkbox rögzítés - '
' - 1 - '
If AppWindow.CheckBox1.Value = True Then
Munka12.Range("s2").Value = "True"
Else
Munka12.Range("s2").Value = "False"
End If
' - 2 - '
If AppWindow.CheckBox2.Value = True Then
Munka12.Range("s3").Value = "True"
Else
Munka12.Range("s3").Value = "False"
End If
' - 3 - '
If AppWindow.CheckBox3.Value = True Then
Munka12.Range("s4").Value = "True"
Else
Munka12.Range("s4").Value = "False"
End If
' - 4 - '
If AppWindow.CheckBox4.Value = True Then
Munka12.Range("s5").Value = "True"
Else
Munka12.Range("s5").Value = "False"
End If
' - 5 - '
If AppWindow.CheckBox5.Value = True Then
Munka12.Range("s6").Value = "True"
Else
Munka12.Range("s6").Value = "False"
End If
' - 6 - '
If AppWindow.CheckBox6.Value = True Then
Munka12.Range("s7").Value = "True"
Else
Munka12.Range("s7").Value = "False"
End If
' - 7 - '
If AppWindow.CheckBox7.Value = True Then
Munka12.Range("s8").Value = "True"
Else
Munka12.Range("s8").Value = "False"
End If
' - 8 - '
If AppWindow.CheckBox8.Value = True Then
Munka12.Range("s9").Value = "True"
Else
Munka12.Range("s9").Value = "False"
End If
' - 9 - '
If AppWindow.CheckBox9.Value = True Then
Munka12.Range("s10").Value = "True"
Else
Munka12.Range("s10").Value = "False"
End If
' - 10 - '
If AppWindow.CheckBox10.Value = True Then
Munka12.Range("s11").Value = "True"
Else
Munka12.Range("s11").Value = "False"
End If
' - 11 - '
If AppWindow.CheckBox11.Value = True Then
Munka12.Range("s12").Value = "True"
Else
Munka12.Range("s12").Value = "False"
End If
' - 12 - '
If AppWindow.CheckBox12.Value = True Then
Munka12.Range("s13").Value = "True"
Else
Munka12.Range("s13").Value = "False"
End If
' - 13 - '
If AppWindow.CheckBox13.Value = True Then
Munka12.Range("s14").Value = "True"
Else
Munka12.Range("s14").Value = "False"
End If
' - 14 - '
If AppWindow.CheckBox14.Value = True Then
Munka12.Range("s15").Value = "True"
Else
Munka12.Range("s15").Value = "False"
End If
' - 15 - '
If AppWindow.CheckBox15.Value = True Then
Munka12.Range("s16").Value = "True"
Else
Munka12.Range("s16").Value = "False"
End If
' - 16 - '
If AppWindow.CheckBox16.Value = True Then
Munka12.Range("s17").Value = "True"
Else
Munka12.Range("s17").Value = "False"
End If
' - 17 - '
If AppWindow.CheckBox17.Value = True Then
Munka12.Range("s18").Value = "True"
Else
Munka12.Range("s18").Value = "False"
End If
' - 18 - '
If AppWindow.CheckBox18.Value = True Then
Munka12.Range("s19").Value = "True"
Else
Munka12.Range("s19").Value = "False"
End If
' - 19 - '
If AppWindow.CheckBox19.Value = True Then
Munka12.Range("s20").Value = "True"
Else
Munka12.Range("s20").Value = "False"
End If
' - 20 - '
If AppWindow.CheckBox20.Value = True Then
Munka12.Range("s21").Value = "True"
Else
Munka12.Range("s21").Value = "False"
End If
' - 21 - '
If AppWindow.CheckBox21.Value = True Then
Munka12.Range("s22").Value = "True"
Else
Munka12.Range("s22").Value = "False"
End If
' - 22 - '
If AppWindow.CheckBox22.Value = True Then
Munka12.Range("s23").Value = "True"
Else
Munka12.Range("s23").Value = "False"
End If
' - 23 - '
If AppWindow.CheckBox23.Value = True Then
Munka12.Range("s24").Value = "True"
Else
Munka12.Range("s24").Value = "False"
End If
' - 24 - '
If AppWindow.CheckBox24.Value = True Then
Munka12.Range("s25").Value = "True"
Else
Munka12.Range("s25").Value = "False"
End If
' - 25 - '
If AppWindow.CheckBox25.Value = True Then
Munka12.Range("s26").Value = "True"
Else
Munka12.Range("s26").Value = "False"
End If
' - 26 - '
If AppWindow.CheckBox26.Value = True Then
Munka12.Range("s27").Value = "True"
Else
Munka12.Range("s27").Value = "False"
End If
' - 27 - '
If AppWindow.CheckBox27.Value = True Then
Munka12.Range("s28").Value = "True"
Else
Munka12.Range("s28").Value = "False"
End If
' - 28 - '
If AppWindow.CheckBox28.Value = True Then
Munka12.Range("s29").Value = "True"
Else
Munka12.Range("s29").Value = "False"
End If
' - 29 - '
If AppWindow.CheckBox29.Value = True Then
Munka12.Range("s30").Value = "True"
Else
Munka12.Range("s30").Value = "False"
End If
' - 30 - '
If AppWindow.CheckBox30.Value = True Then
Munka12.Range("s31").Value = "True"
Else
Munka12.Range("s31").Value = "False"
End If
' - 31 - '
If AppWindow.CheckBox31.Value = True Then
Munka12.Range("s32").Value = "True"
Else
Munka12.Range("s32").Value = "False"
End If
' - 32 - '
If AppWindow.CheckBox32.Value = True Then
Munka12.Range("s33").Value = "True"
Else
Munka12.Range("s33").Value = "False"
End If
' - 33 - '
If AppWindow.CheckBox33.Value = True Then
Munka12.Range("s34").Value = "True"
Else
Munka12.Range("s34").Value = "False"
End If
' - 35 - '
If AppWindow.CheckBox35.Value = True Then
Munka12.Range("s36").Value = "True"
Else
Munka12.Range("s36").Value = "False"
End If
' - 36 - '
If AppWindow.CheckBox36.Value = True Then
Munka12.Range("s37").Value = "True"
Else
Munka12.Range("s37").Value = "False"
End If
' - 37 - '
If AppWindow.CheckBox37.Value = True Then
Munka12.Range("s38").Value = "True"
Else
Munka12.Range("s38").Value = "False"
End If
' - 38 - '
If AppWindow.CheckBox38.Value = True Then
Munka12.Range("s39").Value = "True"
Else
Munka12.Range("s39").Value = "False"
End If
' - 39 - '
If AppWindow.CheckBox39.Value = True Then
Munka12.Range("s40").Value = "True"
Else
Munka12.Range("s40").Value = "False"
End If
' - 40 - '
If AppWindow.CheckBox40.Value = True Then
Munka12.Range("s41").Value = "True"
Else
Munka12.Range("s41").Value = "False"
End If
' - 41 - '
If AppWindow.CheckBox41.Value = True Then
Munka12.Range("s42").Value = "True"
Else
Munka12.Range("s42").Value = "False"
End If
' - 42 - '
If AppWindow.CheckBox42.Value = True Then
Munka12.Range("s43").Value = "True"
Else
Munka12.Range("s43").Value = "False"
End If
' - 43 - '
If AppWindow.CheckBox43.Value = True Then
Munka12.Range("s44").Value = "True"
Else
Munka12.Range("s44").Value = "False"
End If
' - 44 - '
If AppWindow.CheckBox44.Value = True Then
Munka12.Range("s45").Value = "True"
Else
Munka12.Range("s45").Value = "False"
End If
' - 45 - '
If AppWindow.CheckBox45.Value = True Then
Munka12.Range("s46").Value = "True"
Else
Munka12.Range("s46").Value = "False"
End If
' - 46 - '
If AppWindow.CheckBox46.Value = True Then
Munka12.Range("s47").Value = "True"
Else
Munka12.Range("s47").Value = "False"
End If
' - 47 - '
If AppWindow.CheckBox47.Value = True Then
Munka12.Range("s48").Value = "True"
Else
Munka12.Range("s48").Value = "False"
End If
' - 48 - '
If AppWindow.CheckBox48.Value = True Then
Munka12.Range("s49").Value = "True"
Else
Munka12.Range("s49").Value = "False"
End If
' - 49 - '
If AppWindow.CheckBox49.Value = True Then
Munka12.Range("s50").Value = "True"
Else
Munka12.Range("s50").Value = "False"
End If
' - 50 - '
If AppWindow.CheckBox50.Value = True Then
Munka12.Range("s51").Value = "True"
Else
Munka12.Range("s51").Value = "False"
End If
' - 51 - '
If AppWindow.CheckBox51.Value = True Then
Munka12.Range("s52").Value = "True"
Else
Munka12.Range("s52").Value = "False"
End If
' - 52 - '
If AppWindow.CheckBox52.Value = True Then
Munka12.Range("s52").Value = "True"
Else
Munka12.Range("s52").Value = "False"
End If
' - 52 - '
If AppWindow.CheckBox52.Value = True Then
Munka12.Range("s53").Value = "True"
Else
Munka12.Range("s53").Value = "False"
End If
' - 53 - '
If AppWindow.CheckBox53.Value = True Then
Munka12.Range("s54").Value = "True"
Else
Munka12.Range("s54").Value = "False"
End If
' - 54 - '
If AppWindow.CheckBox54.Value = True Then
Munka12.Range("s55").Value = "True"
Else
Munka12.Range("s55").Value = "False"
End If
' - 55 - '
If AppWindow.CheckBox55.Value = True Then
Munka12.Range("s56").Value = "True"
Else
Munka12.Range("s56").Value = "False"
End If
' - 56 - '
If AppWindow.CheckBox56.Value = True Then
Munka12.Range("s57").Value = "True"
Else
Munka12.Range("s57").Value = "False"
End If
' - 57 - '
If AppWindow.CheckBox57.Value = True Then
Munka12.Range("s58").Value = "True"
Else
Munka12.Range("s58").Value = "False"
End If
' - 58 - '
If AppWindow.CheckBox58.Value = True Then
Munka12.Range("s59").Value = "True"
Else
Munka12.Range("s59").Value = "False"
End If
' - 59 - '
If AppWindow.CheckBox59.Value = True Then
Munka12.Range("s60").Value = "True"
Else
Munka12.Range("s60").Value = "False"
End If
' - 60 - '
If AppWindow.CheckBox60.Value = True Then
Munka12.Range("s61").Value = "True"
Else
Munka12.Range("s61").Value = "False"
End If
' - 61 - '
If AppWindow.CheckBox61.Value = True Then
Munka12.Range("s62").Value = "True"
Else
Munka12.Range("s62").Value = "False"
End If
' - 62 - '
If AppWindow.CheckBox62.Value = True Then
Munka12.Range("s63").Value = "True"
Else
Munka12.Range("s63").Value = "False"
End If
' - 63 - '
If AppWindow.CheckBox63.Value = True Then
Munka12.Range("s64").Value = "True"
Else
Munka12.Range("s64").Value = "False"
End If
' - 64 - '
If AppWindow.CheckBox64.Value = True Then
Munka12.Range("s65").Value = "True"
Else
Munka12.Range("s65").Value = "False"
End If
' - 65 - '
If AppWindow.CheckBox65.Value = True Then
Munka12.Range("s66").Value = "True"
Else
Munka12.Range("s66").Value = "False"
End If

End Sub
