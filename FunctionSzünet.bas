Attribute VB_Name = "FunctionSz�net"
Option Explicit

Sub Sz�net()

' - az itt megadott �rt�kekkel lehet a sz�net hossz�t meghat�rozni

Dim newHour As Date, newMinute As Date, newSecond As Date, waitTime As Date

newHour = Hour(Now())           '+1 = 1�ra
newMinute = Minute(Now())   '+1 = 1 perc
newSecond = Second(Now()) + 1   '+1=1 m�sodperc

waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime

End Sub
