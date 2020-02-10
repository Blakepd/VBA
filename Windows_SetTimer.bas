Attribute VB_Name = "Windows_SetTimer"
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal HWnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As LongPtr, _
    ByVal lpTimerFunc As LongPtr) As Long

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal HWnd As LongPtr, _
    ByVal nIDEvent As LongPtr) As Long

Private Declare PtrSafe Function SetThreadExecutionState _
    Lib "Kernel32.dll" _
        (ByVal esFlags As LongPtr) _
    As LongPtr
    
Public iCounter As Double
Public TimerID As Long
Public TimerSeconds As Single
    
Sub KeepPowerOn()
' If you are unable to or do not wish to change sleep/power settings this will stop machine from entering sleep mode

    Const ES_SYSTEM_REQUIRED As LongPtr = &H1
    Const ES_DISPLAY_REQUIRED As LongPtr = &H2
    Const ES_CONTINUOUS As LongPtr = &H80000000
    
    Dim PrevState As LongPtr
    
        PrevState = SetThreadExecutionState(ES_CONTINUOUS Or ES_DISPLAY_REQUIRED Or ES_SYSTEM_REQUIRED)
        Application.OnTime Now() + TimeValue("00:01:00"), "KeepPowerOn", , True
        
End Sub

Sub StartTimer()
    TimerSeconds = 60 ' Time between runs in seconds
    TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc) ' AddressOf is Win API callback
End Sub

Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
End Sub

Sub TimerProc(ByVal HWnd As LongPtr, ByVal uMsg As LongPtr, _
        ByVal nIDEvent As LongPtr, ByVal dwTimer As LongPtr)
    
    ' Do whatever task here...
    KeepPowerOn
    
    If Time > TimeValue("19:00:00") Then
    Range("A1") = Now
    EndTimer ' Will run task once and exit
    End If
    
    If iCounter >= 90 Then
    EndTimer ' After 90 TimerSeconds exit
    End If
    
    iCounter = iCounter + 1
End Sub


