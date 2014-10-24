Attribute VB_Name = "FCSFunctions"
''''
' Module contains Functions used during Fcs
''''

Option Explicit 'force to declare all variables

Public FcsControl As AimFcsController
Public ViewerGuiServer As AimViewerGuiServer
Public FcsPositions As AimFcsSamplePositionParameters



'''''
'   GetFcsPosition(PosX As Double, PosY As Double, PosZ As Double)
'   reads position of small crosshair
'''''
Public Sub getFcsPosition(PosX As Double, PosY As Double, PosZ As Double, Optional pos As Long = -1)
    If pos = -1 Then
        'read actual position of crosshair
        Set ViewerGuiServer = Lsm5.ViewerGuiServer
        ViewerGuiServer.FcsGetLsmCoordinates PosX, PosY, PosZ
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
'''''
Public Function setFcsPositions(Positions() As Vector) As Boolean
    Dim pos As Integer
    Set FcsControl = Fcs
    Set FcsPositions = FcsControl.SamplePositionParameters
    For pos = 0 To UBound(Positions)
        FcsPositions.PositionX(pos) = Positions(pos).X
        FcsPositions.PositionY(pos) = Positions(pos).Y
        FcsPositions.PositionZ(pos) = Positions(pos).Z
    Next pos
    Debug.Print FcsPositions.PositionZ(0)
    'this shows the small crosshair
    ViewerGuiServer.UpdateFcsPositions
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
    Set ViewerGuiServer = Lsm5.ViewerGuiServer
    Set FcsPositions = FcsControl.SamplePositionParameters

    FcsPositions.PositionListSize = 0
    ViewerGuiServer.UpdateFcsPositions
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




