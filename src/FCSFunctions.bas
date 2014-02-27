Attribute VB_Name = "FCSFunctions"
''''
' Module contains Functions used during Fcs
''''

Option Explicit 'force to declare all variables

Public FcsControl As AimFcsController
Public viewerGuiServer As AimViewerGuiServer
Public FcsPositions As AimFcsSamplePositionParameters



'''''
'   GetFcsPosition(PosX As Double, PosY As Double, PosZ As Double)
'   reads position of small crosshair
'''''
Public Sub GetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Optional Pos As Long = -1)
    If Pos = -1 Then
        'read actual position of crosshair
        Set viewerGuiServer = Lsm5.viewerGuiServer
        viewerGuiServer.FcsGetLsmCoordinates PosX, PosY, PosZ
    Else
        Set FcsControl = Fcs
        Set FcsPositions = FcsControl.SamplePositionParameters
        PosX = FcsPositions.PositionX(Pos)
        PosY = FcsPositions.PositionY(Pos)
        PosZ = FcsPositions.PositionZ(Pos)
    End If
End Sub




'''''
'   SetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Pos As Long)
'   Create a new position if Pos > FcsPositions.PositionListSize
'   then all positions inbetween are set to 0
'''''
Public Function setFcsPositions(Positions() As Vector) As Boolean
    Dim Pos As Integer
    Set FcsControl = Fcs
    Set FcsPositions = FcsControl.SamplePositionParameters
    For Pos = 0 To UBound(Positions)
        FcsPositions.PositionX(Pos) = Positions(Pos).X
        FcsPositions.PositionY(Pos) = Positions(Pos).Y
        FcsPositions.PositionZ(Pos) = Positions(Pos).Z
    Next Pos
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




''''''
''   SetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Pos As Long)
''   Create a new position if Pos > FcsPositions.PositionListSize
''   then all positions inbetween are set to 0
''
''''''
'Public Function SetFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Pos As Long) As Boolean
'    Set FcsControl = Fcs
'    Set FcsPositions = FcsControl.SamplePositionParameters
'    FcsPositions.PositionX(Pos) = PosX
'    FcsPositions.PositionY(Pos) = PosY
'    FcsPositions.PositionZ(Pos) = PosZ
'    'this shows the small crosshair
'    viewerGuiServer.UpdateFcsPositions
'End Function




