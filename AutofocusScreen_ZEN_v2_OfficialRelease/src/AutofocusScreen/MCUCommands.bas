Attribute VB_Name = "MCUCommands"

Public Abort As Boolean
Private Interface As Object
Private Frequency As Double
Private Resolution As Double

Public ExchangeXY As Boolean
Public MirrorX As Boolean
Public MirrorY As Boolean
  

Public Sub InitializeStageProperties()
    Set Interface = Lsm5.ExternalCpObject.pHardwareObjects.pInterfaces
    Set Interface = Interface.pItem("CANN")
    
    Resolution = 0.00000025
    Frequency = 2000000
  
End Sub

Public Function GetStagePositionX(CANN As Boolean) As Double
  
   
  
    On Error GoTo nostage
   
    
    If ExchangeXY Then
        GetStagePositionX = StageGetPositionY(CANN) * 1000000#
    Else
        GetStagePositionX = StageGetPositionX(CANN) * 1000000#
    End If
    If MirrorX Then
        GetStagePositionX = -GetStagePositionX
    End If
nostage:
End Function

Public Function GetStagePositionY(CANN As Boolean) As Double
    

    On Error GoTo nostage
    
    If ExchangeXY Then
        GetStagePositionY = StageGetPositionX(CANN) * 1000000#
    Else
        GetStagePositionY = StageGetPositionY(CANN) * 1000000#
    End If
    If MirrorY Then
    Else
        GetStagePositionY = -GetStagePositionY
    End If
nostage:
End Function

Public Function SetStagePositionX(PositionMicrons As Double, CANN As Boolean)

    Dim PositionMetre As Double
    
    PositionMetre = PositionMicrons * 0.000001
    
  
    If MirrorX Then
        PositionMetre = -PositionMetre
    Else
    End If
    If ExchangeXY Then
        StageMoveToPositionY PositionMetre, CANN
        Exit Function
    End If
    On Error GoTo nostage
    StageMoveToPositionX PositionMetre, CANN
nostage:
End Function

Public Function SetStagePositionY(PositionMicrons As Double, CANN As Boolean)
   
    Dim PositionMetre As Double
    
    PositionMetre = PositionMicrons * 0.000001

    If MirrorY Then
    Else
        PositionMetre = -PositionMetre
    End If
    If ExchangeXY Then
        StageMoveToPositionX PositionMetre, CANN
        Exit Function
    End If
    On Error GoTo nostage
    StageMoveToPositionY PositionMetre, CANN
nostage:
End Function

Public Function SetStageSpeed(StageSpeed As Double, CANN As Boolean) As Boolean

    Dim SampleTimer As Double
    Dim SamplingTime As Double
    Dim v As Long
    
    SampleTimer = SendStageCommandWaitForAnswer("Xn" + Chr(13))
    SamplingTime = 16 * (SampleTimer + 1) * (1 / Frequency)

    v = CLng(StageSpeed * SamplingTime / Resolution)
    If v = 0 Then
        SetStageSpeed = False
        Exit Function
    End If
    
    If CANN Then
        SendStageCommand "XV" + CStr(v) + Chr(13)
        SendStageCommand "YV" + CStr(v) + Chr(13)
        SetStageSpeed = True
    Else
        SendCommand "NPXV" + CStr(v) + Chr(13)
        SendCommand "NPYV" + CStr(v) + Chr(13)
    End If
End Function
    
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
    
Public Sub SetStageAcceleration(Accelearation As Double, CANN As Boolean)
    If CANN Then
        SendStageCommand "XA" + CStr(CLng(Accelearation)) + Chr(13)
        SendStageCommand "YA" + CStr(CLng(Accelearation)) + Chr(13)
    Else
        SendCommand "NPXA" + CStr(CLng(Accelearation)) + Chr(13)
        SendCommand "NPYA" + CStr(CLng(Accelearation)) + Chr(13)
    End If
End Sub

Public Function GetStageAcceleration(CANN As Boolean) As Long
    If CANN Then
        GetStageAcceleration = SendStageCommandWaitForAnswer("Xa" + Chr(13))
    Else
        GetStageAcceleration = SendCommandWaitForAnswer("NPXa" + Chr(13))
    End If
End Function

Public Function GetStageSpeed(CANN As Boolean) As Double
    Dim SampleTimer As Double
    Dim SamplingTime As Double

    SampleTimer = SendStageCommandWaitForAnswer("Xn" + Chr(13))
    SamplingTime = 16 * (SampleTimer + 1) * (1 / Frequency)
    
    If CANN Then
        GetStageSpeed = SendStageCommandWaitForAnswer("Xv" + Chr(13)) * Resolution / SamplingTime
    Else
        GetStageSpeed = SendCommandWaitForAnswer("NPXv" + Chr(13)) * Resolution / SamplingTime
    End If
End Function

Public Function IsStageBusy(CANN As Boolean) As Boolean
    If CANN Then
        IsStageBusy = (SendStageCommandWaitForAnswer("Xt" + Chr(13)) <> 0) _
                   Or (SendStageCommandWaitForAnswer("Yt" + Chr(13)) <> 0)
    Else
        IsStageBusy = (SendCommandWaitForAnswer("NPXt" + Chr(13)) <> 0) _
                   Or (SendCommandWaitForAnswer("NPYt" + Chr(13)) <> 0)
    End If
End Function

Public Function GetMaximumStageSpeed() As Double
    GetMaximumStageSpeed = 0.02703786166
End Function

Public Function GetMinimumStageSpeed() As Double
    GetMinimumStageSpeed = 0.0003004206851
End Function

Private Function StageMoveToPositionX(PositionMetre As Double, CANN As Boolean)
    Dim Position As String
    
    Position = Hex(CLng(-PositionMetre / Resolution))
    While Len(Position) < 6
        Position = "0" + Position
    Wend
    If Len(Position) > 6 Then
        Position = Strings.Right(Position, 6)
    End If
    If CANN Then
        SendStageCommand ("XT" + Position + Chr(13))
    Else
         SendCommand ("NPXT" + Position + Chr(13))
    End If
End Function

Private Function StageMoveToPositionY(PositionMetre As Double, CANN As Boolean)
    Dim Position As String
        Position = Hex(CLng(PositionMetre / Resolution))
    While Len(Position) < 6
        Position = "0" + Position
    Wend
    If Len(Position) > 6 Then
        Position = Strings.Right(Position, 6)
    End If
    If CANN Then
        SendStageCommand ("YT" + Position + Chr(13))
    Else
        SendCommand ("NPYT" + Position + Chr(13))
    End If
    
End Function

Private Function StageGetPositionX(CANN As Boolean) As Double
    If CANN Then
        StageGetPositionX = -SendStageCommandWaitForHexAnswer("Xp" + Chr(13)) * Resolution
    Else
        StageGetPositionX = -SendCommandWaitForHexAnswer("NPXp" + Chr(13)) * Resolution
    End If
End Function

Private Function StageGetPositionY(CANN As Boolean) As Double
    If CANN Then
        StageGetPositionY = SendStageCommandWaitForHexAnswer("Yp" + Chr(13)) * Resolution
    Else
        StageGetPositionY = SendCommandWaitForHexAnswer("NPYp" + Chr(13)) * Resolution
    End If
End Function

Private Sub SendStageCommand(command As String)
    If Not Interface Is Nothing Then
        Interface.bSendCmd (command)
    End If
End Sub
            
Private Function SendStageCommandWaitForAnswer(command As String) As Long
    Dim Answer As String
    
    SendStageCommandWaitForAnswer = 0
On Error GoTo ErrorExit
    If Not Interface Is Nothing Then
        Interface.bSendCmdWait4Answer command, Answer
        If Answer <> "" Then
            SendStageCommandWaitForAnswer = CLng(Strings.Right(Answer, Len(Answer) - 2))
        End If
    End If
ErrorExit:
End Function

Private Function SendStageCommandWaitForHexAnswer(command As String) As Long
    Dim Answer As String
    
    SendStageCommandWaitForHexAnswer = 0
On Error GoTo ErrorExit
    If Not Interface Is Nothing Then
        Interface.bSendCmdWait4Answer command, Answer
        If Answer <> "" Then
            SendStageCommandWaitForHexAnswer = Val("&H" + Strings.Right(Answer, Len(Answer) - 2))
            If SendStageCommandWaitForHexAnswer < 0 Then
            SendStageCommandWaitForHexAnswer = SendStageCommandWaitForHexAnswer + &H10000
            End If
            If SendStageCommandWaitForHexAnswer > &H7FFFFF Then
                SendStageCommandWaitForHexAnswer = SendStageCommandWaitForHexAnswer - &H1000000
            End If
        End If
    End If
ErrorExit:
End Function

Private Function SendCommand(command As String) As Long
    Lsm5.DsRecording.StartScanTriggerOut = Lsm5.DsRecording.StartScanTriggerOut + command
End Function

Private Function SendCommandWaitForAnswer(command As String) As Long

End Function

Private Function SendCommandWaitForHexAnswer(command As String) As Long

End Function

Public Sub NoImageAxisChange()
 ExchangeXY = False
 MirrorX = False
 MirrorY = False
End Sub

Public Sub ImageAxisChange()
Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisStateS 1, ExchangeXY, MirrorX, MirrorY
End Sub



Public Function AreStageCoordinateExchanged() As Boolean
    Dim ExchangeXY As Boolean
    Dim MirrorX As Boolean
    Dim MirrorY As Boolean
    Dim bLSM As Boolean
    Dim bLIVE As Boolean
    Dim bCamera As Boolean
    Dim lsystem As Integer
    
     UsedDevices40 bLSM, bLIVE, bCamera
        If bLSM Then
            lsystem = 0
        ElseIf bLIVE Then
            lsystem = 1
        ElseIf bCamera Then
            lsystem = 3
        End If

    Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisStateS lsystem, ExchangeXY, MirrorX, MirrorY
    AreStageCoordinateExchanged = ExchangeXY
End Function

