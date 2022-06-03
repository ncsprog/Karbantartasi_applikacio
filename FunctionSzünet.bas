Attribute VB_Name = "FunctionSzünet"
Option Explicit

Sub Szünet()

' - az itt megadott értékekkel lehet a szünet hosszát meghatározni

Dim newHour As Date, newMinute As Date, newSecond As Date, waitTime As Date

newHour = Hour(Now())           '+1 = 1óra
newMinute = Minute(Now())   '+1 = 1 perc
newSecond = Second(Now()) + 1   '+1=1 másodperc

waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime

End Sub
