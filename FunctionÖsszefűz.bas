Attribute VB_Name = "Function�sszef�z"
Option Explicit

Sub �sszef�z()

Munka16.Select
Range("a:y").Value = ""

' - m�sol - '
Munka1.Select
Columns("a:a").Select
Selection.End(xlDown).Select
Dim A As Long
A = ActiveCell.Row
Dim B As String
B = "w"
Dim C As String
C = B & A
Range("a1", C).Copy
Munka16.Select
Range("a1").PasteSpecial xlPasteValues

' - meddig - '
Columns("a:a").Select
Selection.End(xlDown).Select
Dim lastnr As Long
lastnr = ActiveCell.Row
Dim lastcl As String
lastcl = "a"
Dim last As String
last = lastcl & lastnr
' - l�p�sek - '
Dim Row As Integer
For Row = Range("a2") To Range(last) Step 1
If Range("y" & Row + 1) = "" Then
' - SORONK�NTI m�velet - '
Range("y" & Row + 1).Value = _
Range("b" & Row + 1) & " - " & Range("c" & Row + 1) & " - " & _
Range("d" & Row + 1) & " - " & Range("e" & Row + 1) & " - " & _
Range("f" & Row + 1) & " - " & _
Range("h" & Row + 1) & " - " & Range("i" & Row + 1) & " - " & _
Range("n" & Row + 1) & " - " & Range("o" & Row + 1) & " - " & _
Range("p" & Row + 1)
End If
Next

' - sz�r - '

End Sub
