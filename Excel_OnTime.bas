Attribute VB_Name = "Excel_OnTime"
Public RunTimer As Double

Sub Start_OnTime()
    RunTimer = Now + TimeSerial(0, 0, 5) 'Format (Hours, Minutes, Seconds)
    Application.OnTime EarliestTime:=RunTimer, Procedure:="Procedure1", _
        Schedule:=True
End Sub

Sub Stop_OnTime()
    On Error Resume Next
    Application.OnTime EarliestTime:=RunTimer, Procedure:="Procedure1", _
        Schedule:=False
End Sub

Sub Procedure1()
    Range("A1") = Time ' Enter code here
    Start_OnTime  ' Reschedule Procedure1
End Sub
