VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagingTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
' Timer class
'''
Private Type singleTimer
    Name As String
    TimeStart As Double
End Type

Private Timers() As singleTimer
Private TimerNames As Dictionary

Private Sub MsgNoTimer(Name As String)
    MsgBox "Timer: Timer " & Name & " has not been defined"
End Sub

Public Sub addTimer(Name As String)
    Dim iTimer As Integer
    ' only create once the timer with a specific name
    If TimerNames.Exists(Name) Then
        Exit Sub
    End If
    If isArrayEmpty(Timers) Then
        ReDim Grids(0)
    Else
        ReDim Preserve Timers(0 To UBound(Timers) + 1)
    End If
    TimerNames.Add Name, UBound(Timers)
    iTimer = TimerNames.Item(Name)
    Timers.Name = Name
    Timers.startTime = CDbl(GetTickCount) * 0.001
End Sub

Public Function checkTimerName(Name As String) As Boolean
    If TimerNames.Exists(Name) Then
        checkGridName = True
    Else
        MsgNoTimer Name
    End If
End Function


Public Sub updateTimeStart(Name As String)
    If checkTimerName(Name) Then
        Timers(TimerNames.Item(Name)).TimeStart = CDbl(GetTickCount) * 0.001
    End If
End Sub

Public Function wait(Name As String, Delay As Double) As Double
    If checkTimerName(Name) Then
        wait = Timers(TimerNames.Item(Name)).TimeStart + Delay - CDbl(GetTickCount) * 0.001
    End If
End Function
