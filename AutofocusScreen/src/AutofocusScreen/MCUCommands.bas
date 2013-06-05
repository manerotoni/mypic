Attribute VB_Name = "MCUCommands"

Public Abort As Boolean
Private Interface As Object
Private Frequency As Double
Private Resolution As Double

Public ExchangeXY As Boolean
Public MirrorX As Boolean
Public MirrorY As Boolean
  

Public Sub InitializeStageProperties(Optional Tmp As Boolean) ' tmp is a hack so that function does not appear in menu
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
    
    SampleTimer = SendStageCommandWaitForAnswer("Xn" + Strings.Chr(13))
    SamplingTime = 16 * (SampleTimer + 1) * (1 / Frequency)

    v = CLng(StageSpeed * SamplingTime / Resolution)
    If v = 0 Then
        SetStageSpeed = False
        Exit Function
    End If
    
    If CANN Then
        SendStageCommand "XV" + CStr(v) + Strings.Chr(13)
        SendStageCommand "YV" + CStr(v) + Strings.Chr(13)
        SetStageSpeed = True
    Else
        SendCommand "NPXV" + CStr(v) + Strings.Chr(13)
        SendCommand "NPYV" + CStr(v) + Strings.Chr(13)
    End If
End Function
    
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
    
Public Sub SetStageAcceleration(Accelearation As Double, CANN As Boolean)
    If CANN Then
        SendStageCommand "XA" + CStr(CLng(Accelearation)) + Strings.Chr(13)
        SendStageCommand "YA" + CStr(CLng(Accelearation)) + Strings.Chr(13)
    Else
        SendCommand "NPXA" + CStr(CLng(Accelearation)) + Strings.Chr(13)
        SendCommand "NPYA" + CStr(CLng(Accelearation)) + Strings.Chr(13)
    End If
End Sub

Public Function GetStageAcceleration(CANN As Boolean) As Long
    If CANN Then
        GetStageAcceleration = SendStageCommandWaitForAnswer("Xa" + Strings.Chr(13))
    Else
        GetStageAcceleration = SendCommandWaitForAnswer("NPXa" + Strings.Chr(13))
    End If
End Function

Public Function GetStageSpeed(CANN As Boolean) As Double
    Dim SampleTimer As Double
    Dim SamplingTime As Double

    SampleTimer = SendStageCommandWaitForAnswer("Xn" + Strings.Chr(13))
    SamplingTime = 16 * (SampleTimer + 1) * (1 / Frequency)
    
    If CANN Then
        GetStageSpeed = SendStageCommandWaitForAnswer("Xv" + Strings.Chr(13)) * Resolution / SamplingTime
    Else
        GetStageSpeed = SendCommandWaitForAnswer("NPXv" + Strings.Chr(13)) * Resolution / SamplingTime
    End If
End Function

Public Function IsStageBusy(CANN As Boolean) As Boolean
    If CANN Then
        IsStageBusy = (SendStageCommandWaitForAnswer("Xt" + Strings.Chr(13)) <> 0) _
                   Or (SendStageCommandWaitForAnswer("Yt" + Strings.Chr(13)) <> 0)
    Else
        IsStageBusy = (SendCommandWaitForAnswer("NPXt" + Strings.Chr(13)) <> 0) _
                   Or (SendCommandWaitForAnswer("NPYt" + Strings.Chr(13)) <> 0)
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
        SendStageCommand ("XT" + Position + Strings.Chr(13))
    Else
         SendCommand ("NPXT" + Position + Strings.Chr(13))
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
        SendStageCommand ("YT" + Position + Strings.Chr(13))
    Else
        SendCommand ("NPYT" + Position + Strings.Chr(13))
    End If
    
End Function

Private Function StageGetPositionX(CANN As Boolean) As Double
    If CANN Then
        StageGetPositionX = -SendStageCommandWaitForHexAnswer("Xp" + Strings.Chr(13)) * Resolution
    Else
        StageGetPositionX = -SendCommandWaitForHexAnswer("NPXp" + Strings.Chr(13)) * Resolution
    End If
End Function

Private Function StageGetPositionY(CANN As Boolean) As Double
    If CANN Then
        StageGetPositionY = SendStageCommandWaitForHexAnswer("Yp" + Strings.Chr(13)) * Resolution
    Else
        StageGetPositionY = SendCommandWaitForHexAnswer("NPYp" + Strings.Chr(13)) * Resolution
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

'''
' Function is not used
'''
Public Sub NoImageAxisChange(Optional Tmp As Boolean) ' tmp is used so that the function does not appear in the menu. Not very clean
    ExchangeXY = False
    MirrorX = False
    MirrorY = False
End Sub

'''
' Function is not used
'''
Public Sub ImageAxisChange(Optional Tmp As Boolean)
    Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisStateS 1, ExchangeXY, MirrorX, MirrorY
End Sub


''''''
'   AreStageCoordinateExchanged() As Boolean
'       Check weather X and Y axis are exchanged and return True if yes.
'       Todo: Could also return weather axis are mirrored. Intrestingly althouh axes are not mirrored we use -X??
''''''
Public Sub StageSettings(MirrorX As Boolean, MirrorY As Boolean, ExchangeXY As Boolean)
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
End Sub


