Attribute VB_Name = "Module1"

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub Macro1()
    
    
    Dim viewerGuiServer As AimViewerGuiServer
    Dim node As AimExperimentTreeNode
    
    ' Insert an FCS document into ZEN
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
    node.Type = eExperimentTeeeNodeTypeConfoCor
    viewerGuiServer.InsertExperimentTreeNode node, True, 0
    
    ' Get the fcs controller
    Dim FcsControl As AimFcsController
    Set FcsControl = Fcs
    
    ' Get the fcs data object
    Dim FcsData1 As AimFcsData
    Set FcsData1 = node.FcsData
    
    ' Example: Set up measurement parameters
    Dim AcqParams As AimFcsAcquisitionParameters
    Set AcqParams = FcsControl.AcquisitionParameters
    AcqParams.ChannelDetectorA(0) = 1 ' use detector 1 for channel 0 and vice versa
    AcqParams.ChannelDetectorA(1) = 0
    AcqParams.MeasurementTime = 1
    AcqParams.ChannelEnabled(0) = True
    AcqParams.ChannelEnabled(1) = True
    ' measure
  
    Dim FcsPositions As AimFcsDataSamplePositions
       Dim PosX As Double
    Dim PosY As Double
    Dim PosList As Long
    Set FcsPositions = FcsData1.DataSet(0).AcquisitionParameters.SamplePositions
    PosX = FcsPositions.PositionX(0)

    FcsControl.StopAcquisitionAndWait
   FcsControl.StartMeasurement FcsData1
    While FcsControl.IsAcquisitionRunning(1)
        Sleep (2000)
        DoEvents
    Wend
    Sleep (2000)
     Set FcsPositions = FcsData1.DataSet(1).AcquisitionParameters.SamplePositions
    PosX = FcsPositions.PositionX(0)
'    node.NumberImages = 0
'
'    ' Example: Retrieve fcs data into arrays
'    Dim dataArraySize As Long
'    dataArraySize = FcsData1.DataSet(0).dataArraySize(eFcsDataTypeCorrelation)
'    Dim arrayD1 As Variant
'    Dim arrayD2 As Variant
'    FcsData1.DataSet(0).GetDataSafeArray eFcsDataTypeCorrelation, dataArraySize, arrayD1, arrayD2
'
'    'MsgBox (UBound(arrayD1) - LBound(arrayD1) + 1)
'    'MsgBox (arrayD1(10))
'    'MsgBox (UBound(arrayD2) + 1)
'
'    ' Write to file
'    Dim writer As AimFcsFileWrite
'    Set writer = Lsm5.CreateObject("AimFcsFile.Write")
'    writer.FileName = "C:\\Data\tmp2\testFile.fcs"
'    writer.FileWriteType = eFcsFileWriteTypeAll
'    writer.Format = eFcsFileFormatConfoCor3WithRawData
'
'    writer.Source = FcsData1
'    writer.Run
'

    
End Sub

'Sub Macro2()
'    '**************************************
'    'Recorded: 13/03/2013
'    'Description:
'    '**************************************
'    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
'    Set ZEN = Application.ApplicationInterface
'    'ZEN.GUI.Fcs.Positions.PositionListRemoveAll.Execute
'    Dim pos As ZEN.GUI.Fcs.Positions.SelectedPosition
'    Set pos = ZEN.GUI.Fcs.Positions.SelectedPosition
'    ZEN.GUI.Fcs.Positions.AddPosition.Execute
'
''    ZEN.GUI.Fcs.Positions.EnableCurrentPosition.Value = True
''    ZEN.GUI.Fcs.Positions.CurrentPositionX.Value = -150
''    ZEN.GUI.Fcs.Positions.CurrentPositionY.Value = 20
''    ZEN.GUI.Fcs.Positions.CurrentPositionCrossHair.Value = False
''    ZEN.GUI.Fcs.Positions.CurrentPositionCrossHair.Value = True
''
''
''    'ZEN.GUI.Fcs.Positions.AddPosition.Execute
'
'
'End Sub



