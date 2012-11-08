Attribute VB_Name = "TileScan"

Dim TileX As Integer
Dim TileY As Integer
Dim GlobalPositionsStageOld As Integer
Dim FrameWidth As Double
Dim FrameHeight As Double
Dim RelFrameHeight As Double
Dim RelFrameWidth As Double
Dim Overlap As Double
Dim Xnew As Double
Dim Ynew As Double


Public Sub Tile()
TileX = AutofocusForm.TextBoxTileX.Value
TileY = AutofocusForm.TextBoxTileY.Value
Overlap = AutofocusForm.TextBoxOverlap.Value

FrameWidth = Lsm5.DsRecording.FrameWidth
FrameHeight = Lsm5.DsRecording.FrameHeight

RelFrameWidth = FrameWidth - (FrameWidth * Overlap) / 100
RelFrameHeight = FrameHeight - (FrameHeight * Overlap) / 100

If (TileX > 1 Or TileY > 1) Then CalculateTileLocs

End Sub


Public Sub CalculateTileLocs()
CopyPositionArrays

End Sub


Public Sub CopyPositionArrays()
GlobalXposOld() = GlobalXpos()
GlobalYposOld() = GlobalYpos()
GlobalZposOld() = GlobalZpos()
GlobalLocationsNameOld() = GlobalLocationsName()
GlobalPositionsStageOld = GlobalPositionsStage
GlobalPositionsStage = GlobalPositionsStage * TileX * TileY
ReDim GlobalXpos(GlobalPositionsStage)
ReDim GlobalYpos(GlobalPositionsStage)
ReDim GlobalZpos(GlobalPositionsStage)
ReDim GlobalLocationsName(GlobalPositionsStage)
Dim n As Integer
Dim TX As Integer
Dim TY As Integer

For n = 0 To GlobalPositionsStageOld - 1
Xnew = GlobalXposOld(n + 1) - ((TileX - 1) / 2) * RelFrameWidth
Ynew = GlobalYposOld(n + 1) - ((TileY - 1) / 2) * RelFrameHeight
    For TY = 0 To TileY - 1
        For TX = 0 To TileX - 1
            GlobalXpos(n * TileX * TileY + (TY * TileX) + TX + 1) = Xnew + TX * RelFrameWidth
            GlobalYpos(n * TileX * TileY + (TY * TileX) + TX + 1) = Ynew + TY * RelFrameHeight
            GlobalZpos(n * TileX * TileY + (TY * TileX) + TX + 1) = GlobalZposOld(n + 1)
            If Grid Then
            GlobalLocationsName(n * TileX * TileY + (TY * TileX) + TX + 1) = GlobalLocationsNameOld(n + 1)
            End If
         Next TX
     Next TY
Next n
End Sub
