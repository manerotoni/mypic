VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Generic class to control position of Stage, focus and for SPIM the rotator
Public Function GetPositionX() As Double
    Dim posX As Double
    Dim posY As Double
    If Lsm5.Info.IsSPIM Then
        StageDevice.ReadPosition posX, posY
        GetPositionX = posX
    Else
        Lsm5.Hardware.CpStages.GetXYPosition posX, posY
        GetPositionX = posX
    End If
End Function

Public Function SetPositionX(value As Double)
    If Lsm5.Info.IsSPIM Then
        StageDevice.MoveToPosition value, GetPositionY
    Else
        Lsm5.Hardware.CpStages.SetXYPosition value, GetPositionY
    End If
    
End Function


Public Function GetPositionY() As Double
    Dim posX As Double
    Dim posY As Double
    If Lsm5.Info.IsSPIM Then
        StageDevice.ReadPosition posX, posY
        GetPositionY = posY
    Else
        Lsm5.Hardware.CpStages.GetXYPosition posX, posY
        GetPositionY = posY
    End If
End Function

Public Function SetPositionY(value As Double)
    If Lsm5.Info.IsSPIM Then
        StageDevice.MoveToPosition GetPositionX, value
    Else
         Lsm5.Hardware.CpStages.SetXYPosition GetPositionX, value
    End If
End Function

Public Function GetPositionZ() As Double
    If Lsm5.Info.IsSPIM Then
        GetPositionZ = FocusDevice.dPosition
    Else
        GetPositionZ = Lsm5.Hardware.CpFocus.position
    End If
End Function

Public Function SetPositionZ(value As Double)
    If Lsm5.Info.IsSPIM Then
        FocusDevice.dPosition = value
    Else
        Lsm5.Hardware.CpFocus.position = Z
    End If
End Function

Public Function GetPositionAlpha() As Double
    GetPositionAlpha = RotatorDevice.dPosition
End Function

Public Function SetPositionAlpha(value As Double)
    RotatorDevice.dPosition = value
End Function

Private Property Get StageDevice() As CP.Stage
    Set StageDevice = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem("Stage")
End Property

Private Property Get FocusDevice() As CP.Focus
    Set FocusDevice = Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem("Focus")
End Property

Private Property Get RotatorDevice() As CP.Servo
    Set RotatorDevice = Lsm5.ExternalCpObject.pHardwareObjects.pServos.pItem("Rotator")
End Property
