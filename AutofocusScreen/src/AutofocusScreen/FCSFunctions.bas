Attribute VB_Name = "FCSFunctions"
Option Explicit 'force to declare all variables

Public Declare Function GetInputState Lib "user32" () As Long ' Check if mouse or keyboard has been pushed


Public FcsControl As AimFcsController
Public viewerGuiServer As AimViewerGuiServer
Public FcsPositions As AimFcsSamplePositionParameters
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public FcsData As AimFcsData
'Public ScanStop As Boolean
'Public AcquisitionController  As AimScanController
'Public Const PrecXY = 3
Public Const Pause = 100 'pause in ms


'''''
'   Set the FCS controller and data stuff
'''''
Private Sub Initialize_Controller()
    Set FcsControl = Fcs 'member of Lsm5VBAProject
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Set FcsPositions = FcsControl.SamplePositionParameters
    viewerGuiServer.FcsSelectLsmImagePositions = True
End Sub

Public Sub NewRecord(RecordingDoc As DsRecordingDoc, Optional name As String, Optional Container As Long = 0)
    Dim node As AimExperimentTreeNode
    Set viewerGuiServer = Lsm5.viewerGuiServer
    If RecordingDoc Is Nothing Then
        Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
        node.Type = eExperimentTeeeNodeTypeLsm
        viewerGuiServer.InsertExperimentTreeNode node, True, Container
        Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
        While RecordingDoc.IsBusy
            Sleep (Pause)
            DoEvents
        Wend
        RecordingDoc.SetTitle name
    End If
End Sub


Public Sub NewFcsRecord(FcsData As AimFcsData, Optional name As String, Optional Container As Long = 0)
    Dim node As AimExperimentTreeNode
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Dim Recording As DsRecordingDoc
    If FcsData Is Nothing Then
        Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
        node.Type = eExperimentTeeeNodeTypeConfoCor
        viewerGuiServer.InsertExperimentTreeNode node, True, Container
        ' Insert an FCS document into ZEN
        Set FcsData = node.FcsData
        FcsData.name = name
        Set Recording = Lsm5.DsRecordingActiveDocObject
        Recording.SetTitle name
    End If

End Sub
''''
' Start Fcs Measurment
''''
Public Function FcsMeasurement(Optional FcsData As AimFcsData) As Boolean
    Dim node As AimExperimentTreeNode
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Set FcsControl = Fcs
    Set viewerGuiServer = Lsm5.viewerGuiServer
    
    If FcsData Is Nothing Then
       NewFcsRecord FcsData
    End If
    'FcsData.name = "Bla"
    FcsControl.StopAcquisitionAndWait
    FcsControl.StartMeasurement FcsData
    Sleep (Pause)
    While FcsControl.IsAcquisitionRunning(1)
        Sleep (Pause)
        If ScanStop Then
            FcsControl.StopAcquisitionAndWait
            Exit Function
        End If
        DoEvents
    Wend
    FcsControl.StopAcquisitionAndWait
    FcsMeasurement = True
End Function


''''''
''   ScanToImage ( RecordingDoc As DsRecordingDoc) As Boolean
''   scan overwrite the same image, even with several z-slices
''''''
'Public Function ScanToImage(RecordingDoc As DsRecordingDoc) As Boolean
'    Dim ProgressFifo As IAimProgressFifo ' what is this?
'    Dim gui As Object, treenode As Object
'    'Set gui = Lsm5.ViewerGuiServer
'    If RecordingDoc Is Nothing Then
'        NewRecord RecordingDoc
'    End If
'    If Not RecordingDoc Is Nothing Then
'        Set treenode = RecordingDoc.RecordingDocument.image(0, True)
'        'Set treenode = Lsm5.NewDocument why not this?
'        Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller ' public variable
'        AcquisitionController.DestinationImage(0) = treenode 'EngelImageToHechtImage(GlobalSingleImage).Image(0, True)
'        AcquisitionController.DestinationImage(1) = Nothing
'        Set ProgressFifo = AcquisitionController.DestinationImage(0)
'        Lsm5.tools.CheckLockControllers True
'        AcquisitionController.StartGrab eGrabModeSingle
'        'Set RecordingDoc = Lsm5.StartScan this does not overwrite
'        If Not ProgressFifo Is Nothing Then ProgressFifo.Append AcquisitionController
'    End If
'    Sleep (Pause)
'    While AcquisitionController.IsGrabbing
'        Sleep (Pause) ' this sometimes hangs if we use GetInputState. Try now without it and test if it does not hang
'        DoEvents
'        If ScanStop Then
'            Lsm5.StopAcquisition
'            Exit Function
'        End If
'    Wend
'    ScanToImage = True
'End Function


'''''''
'' SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String) As Boolean
'' Copied and adapted from MultiTimeSeries macro
'''''''
'Public Function SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String) As Boolean
'    Dim Export As AimImageExport
'    Dim image As AimImageMemory
'    Dim Error As AimError
'    Dim Planes As Long
'    Dim Plane As Long
'    Dim Horizontal As enumAimImportExportCoordinate
'    Dim Vertical As enumAimImportExportCoordinate
'
'
'    'Set Image = EngelImageToHechtImage(Document).Image(0, True)
'    If Not Document Is Nothing Then
'        Set image = Document.RecordingDocument.image(0, True)
'    End If
'
'    Set Export = Lsm5.CreateObject("AimImageImportExport.Export.4.5")
'    'Set Export = New AimImageExport
'    Export.FileName = FileName
'    Export.Format = eAimExportFormatLsm5
'    Export.StartExport image, image
'    Set Error = Export
'    Error.LastErrorMessage
'
'    Planes = 1
'    Export.GetPlaneDimensions Horizontal, Vertical
'
'    Select Case Vertical
'        Case eAimImportExportCoordinateY:
'             Planes = image.GetDimensionZ * image.GetDimensionT
'        Case eAimImportExportCoordinateZ:
'            Planes = image.GetDimensionT
'    End Select
'
'    'TODO check. what happens here with Export.ExportPlane Nothing why Nothing (thumbnails)
'    For Plane = 0 To Planes - 1
'        If GetInputState() <> 0 Then
'            DoEvents
'             If ScanStop Then
'                Export.FinishExport
'                Exit Function
'            End If
'        End If
'        Export.ExportPlane Nothing
'    Next Plane
'    Export.FinishExport
'    SaveDsRecordingDoc = True
'
'End Function

''''
' SaveFcsMeasurment to File
''''
Public Sub SaveFcsMeasurement(FcsData As AimFcsData, fileName As String)
    
    If FcsData Is Nothing Then
        MsgBox "No Fcs Recording to Save"
        Exit Sub
    End If
    ' Write to file
    Dim writer As AimFcsFileWrite
    Set writer = Lsm5.CreateObject("AimFcsFile.Write")
    writer.fileName = fileName
    writer.FileWriteType = eFcsFileWriteTypeAll
    writer.Format = eFcsFileFormatConfoCor3WithRawData

    writer.Source = FcsData
    writer.Run
       
    If Not writer.DestinationFilesExist(fileName) Then
        writer.fileName = fileName
        writer.FileWriteType = eFcsFileWriteTypeAll
        writer.Format = eFcsFileFormatConfoCor3WithRawData
    
        writer.Source = FcsData
        writer.Run
    Else
    
    End If
End Sub

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

Public Sub SaveFcsPositionList(sFile As String, pixelSizeXY As Double, pixelSizeZ As Double)
    On Error GoTo ErrorHandle
    Close
    Dim iFileNum As Integer
    Dim i As Long
    Dim PosX As Double
    Dim PosY As Double
    Dim PosZ As Double
    iFileNum = FreeFile()
    Open sFile For Output As iFileNum
    If pixelSizeXY > 0 And pixelSizeZ > 0 Then
        Print #iFileNum, "%X Y Z (um) X Y Z (px); 0 0 is center of image"
    Else
        Print #iFileNum, "%X Y Z (um); 0 0 is center of image"
    End If
    For i = 0 To GetFcsPositionListLength - 1
        GetFcsPosition PosX, PosY, PosZ, i
        If pixelSizeXY > 0 And pixelSizeZ > 0 Then
            Print #iFileNum, Round(PosX * 1000000, PrecXY) & " " & Round(PosY * 1000000, PrecXY) & " " & Round(PosZ * 1000000, PrecXY) & " " & PosX / pixelSizeXY & " " & PosY / pixelSizeXY & " " & PosZ / pixelSizeZ
        Else
            Print #iFileNum, Round(PosX * 1000000, PrecXY) & " " & Round(PosY * 1000000, PrecXY) & " " & Round(PosZ * 1000000, PrecXY)
        End If
    Next i
    Close
    Exit Sub
ErrorHandle:
    MsgBox "Can't write " & sFile & " for the FcsPositions"
End Sub






