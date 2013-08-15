Attribute VB_Name = "FCSFunctions"
''''
' Module contains Functions used during Fcs
''''

Option Explicit 'force to declare all variables

Public Declare Function GetInputState Lib "user32" () As Long ' Check if mouse or keyboard has been pushed


Public FcsControl As AimFcsController
Public viewerGuiServer As AimViewerGuiServer
Public FcsPositions As AimFcsSamplePositionParameters
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public FcsData As AimFcsData
Public Const Pause = 100 'pause in ms




'''''
'   GetFcsPosition(PosX As Double, PosY As Double, PosZ As Double)
'   reads position of small crosshair
'''''
Public Sub GetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Optional pos As Long = -1)
    If pos = -1 Then
        'read actual position of crosshair
        Set viewerGuiServer = Lsm5.viewerGuiServer
        viewerGuiServer.FcsGetLsmCoordinates PosX, PosY, PosZ
    Else
        Set FcsControl = Fcs
        Set FcsPositions = FcsControl.SamplePositionParameters
        PosX = FcsPositions.PositionX(pos)
        PosY = FcsPositions.PositionY(pos)
        PosZ = FcsPositions.PositionZ(pos)
    End If
End Sub



'''''
'   SetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Pos As Long)
'   Create a new position if Pos > FcsPositions.PositionListSize
'   then all positions inbetween are set to 0
'
'''''
Public Function SetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, pos As Long) As Boolean
    Set FcsControl = Fcs
    Set FcsPositions = FcsControl.SamplePositionParameters
    FcsPositions.PositionX(pos) = PosX
    FcsPositions.PositionY(pos) = PosY
    FcsPositions.PositionZ(pos) = PosZ
    'this shows the small crosshair
    viewerGuiServer.UpdateFcsPositions
End Function

'''''
'   SetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Pos As Long)
'   Create a new position if Pos > FcsPositions.PositionListSize
'   then all positions inbetween are set to 0
'''''
Public Function setFcsPositions(positions() As Vector) As Boolean
    Dim pos As Integer
    Set FcsControl = Fcs
    Set FcsPositions = FcsControl.SamplePositionParameters
    For pos = 0 To UBound(positions)
        FcsPositions.PositionX(pos) = positions(pos).X
        FcsPositions.PositionY(pos) = positions(pos).Y
        FcsPositions.PositionZ(pos) = positions(pos).Z
    Next pos
    Debug.Print FcsPositions.PositionZ(0)
    'this shows the small crosshair
    viewerGuiServer.UpdateFcsPositions
End Function

''''
'   GetFcsListPositionLength()
'   Maximal number of positions
''''
Public Function GetFcsPositionListLength() As Long
    Set FcsControl = Fcs
    Set FcsPositions = FcsControl.SamplePositionParameters
    GetFcsPositionListLength = FcsPositions.PositionListSize
End Function

''''
'   ClearFcsPositionList()
'   Remove all FCSpositions stored
''''
Public Function ClearFcsPositionList()
    'This clear the positions
    Set FcsControl = Fcs
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Set FcsPositions = FcsControl.SamplePositionParameters

    FcsPositions.PositionListSize = 0
    viewerGuiServer.UpdateFcsPositions
End Function








