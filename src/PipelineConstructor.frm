VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PipelineConstructor 
   Caption         =   "Pipeline Constructor"
   ClientHeight    =   17775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16005
   OleObjectBlob   =   "PipelineConstructor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PipelineConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents EventMng As EventAdmin
Attribute EventMng.VB_VarHelpID = -1
Private currPipeline As Integer
Const NrPipelines = 3
Private PipelineCaption(0 To NrPipelines - 1) As String
Private Lett() As Variant

Public Version As String


Private Sub PauseExpButton_Click()
    If Not Pipelines(0).Grid.isRunning Then
        ScanPause = False
        PauseExpButton.value = False
        PauseExpButton.Caption = ""
        PauseExpButton.BackColor = &H8000000F
    Else
        If PauseExpButton.value Then
            ScanPause = True
            PauseExpButton.Caption = "RESUME"
            PauseExpButton.BackColor = 12648447
        Else
            ScanPause = False
            PauseExpButton.Caption = ""
            PauseExpButton.BackColor = &H8000000F
        End If
    End If
End Sub

Public Sub UserForm_Initialize()
    Dim i As Integer
    'Contains name of the Grids two letter code
    Dim GridNames(2) As String
    GridNames(0) = "DE"
    GridNames(1) = "TR1"
    GridNames(2) = "TR2"
    
    ZenV = getVersionNr
    'find the version of the software and load ZEN object
    If ZenV > 2010 Then
        On Error GoTo errorMsg
        'in some cases this does not register properly
        'Set ZEN = Lsm5.CreateObject("Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface")
        'this should always work
        Set ZEN = Application.ApplicationInterface
        Dim TestBool As Boolean
        'Check if it works
        TestBool = ZEN.GUI.Acquisition.EnableTimeSeries.value
        ZEN.GUI.Acquisition.EnableTimeSeries.value = Not TestBool
        ZEN.GUI.Acquisition.EnableTimeSeries.value = TestBool
        GoTo NoError
errorMsg:
        MsgBox "Version is ZEN" & ZenV & " but can't find Zeiss.Micro.AIM.ApplicationInterface." & vbCrLf & "Using ZEN2010 settings instead." & vbCrLf & "Check if Zeiss.Micro.AIM.ApplicationInterface.dll is registered?" & "See also the manual how to register a dll into windows."
        ZenV = 2010
NoError:
    End If
    
    Version = "v0.2"
    Me.Caption = Me.Caption + " " + Version
    FormatUserForm Me.Caption
    Set EventMng = New EventAdmin
    EventMng.initialize
        
    ''Pipeline settings
    ReDim Pipelines(0 To NrPipelines - 1)
    For i = 0 To NrPipelines - 1
        Set Pipelines(i) = New APipeline
        Set Pipelines(i).Repetition = New ARepetition
        Set Pipelines(i).Grid = New AGrid
        Pipelines(i).Repetition.interval = True
        Pipelines(i).Grid.NameGrid = GridNames(i)
        Pipelines(i).keepParent = True
    Next i
    PipelineCaption(0) = "Default"
    PipelineCaption(1) = "Trigger1"
    PipelineCaption(2) = "Trigger2"
    Erase ImgJobs
    Erase FcsJobs
    Set OiaSettings = New OnlineIASettings
    OiaSettings.initializeDefault
    OiaSettings.resetRegistry
    imgFileFormat = eAimExportFormatLsm5
    imgFileExtension = ".lsm"
    
    ''Form layout

    CurrentPipelineList.ColumnCount = 3
    CurrentPipelineList.ColumnWidths = "20;30;200"
    JobChoiceList.ColumnCount = 2
    JobChoiceList.ColumnWidths = "30;200"
    JobChoiceFrame.Visible = False
    PositionsList.ColumnCount = 5
    PositionsList.ColumnWidths = "20;25;35;35;50"

    Set FocusMethods = New Dictionary
    FocusMethods.Add NoAnalyse, "None"
    FocusMethods.Add AnalyseCenterOfMassThr, "Center of Mass (thr)"
    FocusMethods.Add AnalysePeak, "Peak"
    FocusMethods.Add AnalyseCenterOfMass, "Center of Mass"
    FocusMethods.Add AnalyseOnline, "Online img. analysis"
    For i = 0 To FocusMethods.count - 1
        FocusMethod.AddItem FocusMethods.item(i), i
    Next i
    FocusMethod.ListIndex = 0
    PlateType.AddItem "None"
    PlateType.AddItem "Single Well"
    PlateType.AddItem "2 Wells"
    PlateType.AddItem "4 Wells (1x4)"
    PlateType.AddItem "8 Wells (2x4)"
    PlateType.AddItem "96 Wells (8x12)"
    PlateType.ListIndex = 0
    Lett = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")

    PositionButton2.value = True
    PositionButton1.value = True
    currentImgJob = -1
    currentFcsJob = -1
    ToggleFrameButton (1)
    Me.Height = 465
    Me.Width = 430
    
    Load JobSetter
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim exitPipCon As Integer
    exitPipCon = MsgBox("Exit PipelineConstructor?", VbOKCancel + VbQuestion, "PipCon exit")
    If exitPipCon = vbOK Then
        Unload JobSetter
    Else
        Cancel = True
    End If
End Sub


Private Sub GridScanPositionFileButton_Click()
    Dim fso As New FileSystemObject
    Dim Filter As String, fileName As String
    Dim Flags As Long
    Dim DefDir As String
    Dim locGrid As AGrid
    Dim gridDim() As Long
    Set locGrid = New AGrid
    
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "position files (*.pos)" & Chr$(0) & "*.pos" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    
    fileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DefDir, "Select position file")
    If fileName = "" Then
        Exit Sub
    End If
    If Not FileExist(fileName) Then
        MsgBox "Load positions from file failed. File " & fileName & " does not exist"
        Exit Sub
    End If
    WorkingDir = fso.GetParentFolderName(fileName) & "\"
    gridDim = locGrid.getGridDimFromFile(fileName)
    If Not locGrid.loadPositionGridFile(fileName) Then
        Exit Sub
    End If
    
    GridScanPositionFile = fileName
    UpdatePositionsListFromGrid locGrid
End Sub

Private Sub UpdatePositionsListFromGrid(locGrid As AGrid)
    Dim index As Long
    locGrid.setIndeces 1, 1, 1, 1
    
    PositionsList.Clear
    If locGrid.getNrValidPts > 100 Then
        MsgBox "Warning: Position file contains more than 100 positions!" & vbCrLf & _
        "All positions will be loaded but only the first 100 are shown in list"
    End If
    Do ''Cycle all positions defined in grid
        If locGrid.getThisValid Then
            index = index + 1
            AddPosition WellID.value, locGrid.getThisX, locGrid.getThisY, locGrid.getThisZ
        End If
    Loop While (locGrid.nextGridPt(False) And index < 100)
End Sub





Private Sub SavePositionsButton_Click()
    Dim fso As New FileSystemObject
    Dim Filter As String, fileName As String, DefDir As String
    Dim Flags As Long
    Dim locGrid As AGrid
    Set locGrid = New AGrid
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
    Filter = "Position file (*.pos)" & Chr$(0) & "*.pos" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    fileName = CommonDialogAPI.ShowSave(Filter, Flags, "*.pos", DefDir, "Save positions")
    DisplayProgress Me.ProgressLabel, "Saving positions..", RGB(0, &HC0, 0)
    
    If fileName <> "" Then
        If VBA.Right(fileName, 4) <> ".pos" Then
            fileName = fileName & ".pos"
        End If
    Else
        GoTo ExitSub
    End If
    WorkingDir = fso.GetParentFolderName(fileName) & "\"
    If Not setGridFromPositionChoice(locGrid) Then
        MsgBox "Saving of positions failed"
    End If
    If Not locGrid.writePositionGridFile(fileName) Then
         MsgBox "Saving of positions failed"
    End If
ExitSub:
    DisplayProgress Me.ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)
End Sub

Private Sub SaveSettings_Click()
    Dim fso As FileSystemObject
    Dim Filter As String, fileName As String
    Dim Flags As Long
    Dim DefDir As String
   
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Images (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    
    fileName = CommonDialogAPI.ShowSave(Filter, Flags, "PipelineConstructor.ini", DefDir, "Save PipelineConstructor settings")
    If fileName = "" Then
        Exit Sub
    End If
    Set fso = New FileSystemObject
    WorkingDir = fso.GetParentFolderName(fileName) & "\"
    If Len(fileName) > 3 And VBA.Right(fileName, 4) <> ".ini" Then
        fileName = fileName & ".ini"
    End If
    SaveFormSettings fileName
End Sub

Private Sub LoadSettings_Click()
    Dim fso As FileSystemObject
    Dim Filter As String, fileName As String
    Dim Flags As Long
    Dim DefDir As String

    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Images (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    
    fileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DefDir, "Load PipelineConstructor settings")
    If fileName = "" Then
        Exit Sub
    End If
    Set fso = New FileSystemObject
    WorkingDir = fso.GetParentFolderName(fileName) & "\"
    LoadFormSettings fileName
End Sub

Private Sub StopExpButton_Click()
    ScanStop = True
    StopAcquisition

    If Pipelines(0).Grid.getNrPts > 1 And CheckDir(GlobalDataBaseName) And Pipelines(0).Grid.isRunning Then
        Pipelines(0).Grid.writePositionGridFile (GlobalDataBaseName + "positionsAfterStop.pos")
    End If
    Pipelines(0).Grid.isRunning = False
End Sub






Private Sub StartExpButton_Click()
    'Do some check for consistency
    StartSetting
End Sub

''''''
'   StartSetting()
'   Setups and controls before start of experiment
'       Create list of positions for Grid or Multiposition
''''''
Private Function StartSetting() As Boolean
    Dim i As Integer
    Dim initPos As Boolean   'if False and gridsize correspond positions are taken from file positionsGrid.csv
    Dim SuccessRecenter As Boolean
    Dim Job As Variant
    Dim gridDim() As Long
    Dim pos() As Vector
    Dim posCurr As Vector   'current position
    ScanStop = False
    StageSettings MirrorX, MirrorY, ExchangeXY
    If Not GlobalRecordingDoc Is Nothing Then
        GlobalRecordingDoc.BringToTop
    End If
    NewRecordGui GlobalRecordingDoc, Pipelines(currPipeline).Grid.NameGrid, ZEN, ZenV
    Lsm5.Hardware.CpStages.GetXYPosition posCurr.X, posCurr.Y
    posCurr.Z = Lsm5.Hardware.CpFocus.position
    

    Set FileSystem = New FileSystemObject
    If Pipelines(0).count = 0 Then
        MsgBox ("Nothing to do! Add at least one task to Default pipeline!")
        Exit Function
    End If
    ''Create and check directory for output and log files
    SetDatabase
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start acquisition.")
        Exit Function
    Else
        If Not CheckDir(GlobalDataBaseName) Then
            Exit Function
        End If
        'initialize logFiles
        If LogFileNameBase <> "" Then
            'On Error GoTo ErrorHandleLogFile
            LogFileName = LogFileNameBase
            ErrFileName = ErrFileNameBase
            Close
            If SafeOpenTextFile(LogFileName, LogFile, FileSystem) And SafeOpenTextFile(ErrFileName, ErrFile, FileSystem) Then
                LogFile.WriteLine "% ZEN software version " & ZenV & " PipelineConstructor " & Version
                ErrFile.WriteLine "% ZEN software version " & ZenV & " PipelineConstructor " & Version
                LogFile.Close
                ErrFile.Close
                Log = True
            Else
                Log = False
            End If
        Else
            Log = False
        End If
    End If
    SetFileName
    
    For Each Job In ImgJobs
        Job.timeToAcquire = 0
    Next Job
    For Each Job In FcsJobs
        Job.timeToAcquire = 0
    Next Job
    DisplayProgress Me.ProgressLabel, "Initialize all grid positions...", RGB(0, &HC0, 0)
    
    
    ''Single position
    For i = 0 To UBound(Pipelines)
        Pipelines(i).Grid.initializeToZero
    Next i
    Set TimersGridCreation = Nothing
    If Not setGridFromPositionChoice(Pipelines(0).Grid) Then
        GoTo ExitStart
    End If
    Pipelines(0).Grid.writePositionGridFile GlobalDataBaseName & "PipelineConstructor.pos"
    Pipelines(0).Grid.setAllParentPath GlobalDataBaseName
    SaveFormSettings GlobalDataBaseName & "PipelineConstructor.ini"
    StartPipeline 0, GlobalRecordingDoc, GlobalFcsRecordingDoc, GlobalFcsData, GlobalDataBaseName
ExitStart:
    LogManager.UpdateLog "End of Global pipeline", -1
    DisplayProgress PipelineConstructor.ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)


End Function


Function setGridFromPositionChoice(locGrid As AGrid) As Boolean
    Dim posCurr As Vector
    Dim i As Integer
    posCurr.X = Lsm5.Hardware.CpStages.PositionX
    posCurr.Y = Lsm5.Hardware.CpStages.PositionX
    posCurr.Z = Lsm5.Hardware.CpFocus.position
    If PositionButton1 Then
        locGrid.initialize 1, 1, 1, 1
        locGrid.setPt posCurr, True, 1, 1, 1, 1
    End If
    
    ''Multiple positions
    If PositionButton2 Then
        If PositionsList.ListCount <= 0 Then
            MsgBox "No positions defined for multiple position! Add positions to default positions!"
            Exit Function
        Else
            locGrid.initialize 1, PositionsList.ListCount, 1, 1
            For i = 0 To PositionsList.ListCount - 1
                posCurr.X = PositionsList.List(i, 2)
                posCurr.Y = PositionsList.List(i, 3)
                posCurr.Z = PositionsList.List(i, 4)
                locGrid.setPt posCurr, True, 1, i + 1, 1, 1
            Next i
        End If
    End If
    
    ''Grid
    If PositionButton3 Then
        If PositionsList.ListCount <= 0 Then
            MsgBox "No positions defined for Grid! First position is used as reference!"
            Exit Function
        Else
            posCurr.X = PositionsList.List(0, 2)
            posCurr.Y = PositionsList.List(0, 3)
            posCurr.Z = PositionsList.List(0, 4)
            locGrid.makeGridFromOnePt posCurr, GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub, GridScan_dRow, GridScan_dColumn, GridScan_dRowsub, GridScan_dColumnsub
        End If
    End If
    
    
    ''Grid from multiple positions
    If PositionButton4 Then
        If PositionsList.ListCount <= 0 Then
            MsgBox "No positions defined for multiple positions + grid! Main grid Positions are marked positions, subpositions are made accordingly!"
            Exit Function
        Else
            Dim posVec() As Vector
            ReDim posVec(0 To PositionsList.ListCount - 1)
            For i = 0 To PositionsList.ListCount - 1
                posVec(i).X = PositionsList.List(i, 2)
                posVec(i).Y = PositionsList.List(i, 3)
                posVec(i).Z = PositionsList.List(i, 4)
            Next i
            locGrid.makeGridFromManyPts posVec, 1, PositionsList.ListCount, GridScan_nRowsub, GridScan_nColumnsub, GridScan_dRowsub, GridScan_dColumnsub
        End If
    End If
    
    If PositionButton5 Then
        If Not FileExist(GridScanPositionFile) Then
            MsgBox "Load positions from file failed. Could not find " & GridScanPositionFile
            Exit Function
        End If

        If Not locGrid.loadPositionGridFile(GridScanPositionFile) Then
            Exit Function
        End If
    End If
    setGridFromPositionChoice = True
End Function

'''
' Run the current pipeline
'''
Private Sub AcquirePipelineButton_Click()
    Dim stgPos As Vector
    Dim RepNum As Long
    ScanStop = False
    If Pipelines(currPipeline).count > 0 Then
        If Not GlobalRecordingDoc Is Nothing Then
            GlobalRecordingDoc.BringToTop
        End If
        NewRecordGui GlobalRecordingDoc, Pipelines(currPipeline).Grid.NameGrid, ZEN, ZenV
        Lsm5.Hardware.CpStages.GetXYPosition stgPos.X, stgPos.Y
        stgPos.Z = Lsm5.Hardware.CpFocus.position
        Pipelines(currPipeline).Grid.initialize 1, 1, 1, 1
        Pipelines(currPipeline).Grid.setPt stgPos, True, 1, 1, 1, 1
        UpdateRepetitionSettings currPipeline
        Debug.Print Pipelines(currPipeline).Grid.getNrValidPts
        Pipelines(currPipeline).Grid.setAllParentPath "C:\Antonio\"
        RepNum = Pipelines(currPipeline).Repetition.number
        Pipelines(currPipeline).Repetition.number = 1
        StartPipeline CInt(currPipeline), GlobalRecordingDoc, GlobalFcsRecordingDoc, GlobalFcsData, "C:\Antonio\"
        Pipelines(currPipeline).Repetition.number = RepNum
        DisplayProgress ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)
        
    Else
        MsgBox "You need to add a task to the pipeline. Click on + button"
         For RepNum = 0 To 10
            AddJobToPipelineButton.BackColor = "&H0080FFFF&"
            SleepWithEvents (200)
            AddJobToPipelineButton.BackColor = "&H0000C000&"
        Next RepNum
        AddJobToPipelineButton.BackColor = "&H0000C000&"
    End If
End Sub


Private Sub StopPipelineButton_Click()
     StopAcquisition
End Sub

Private Sub fileFormatczi_Click()
#If ZENvC > 2010 Then
    imgFileFormat = eAimExportFormatCzi
    imgFileExtension = ".czi"
#Else
    MsgBox "Your ZEN version does not support czi files", VbInformation
    fileFormatlsm.value = True
#End If
End Sub

Private Sub fileFormatlsm_Click()
    imgFileFormat = eAimExportFormatLsm5
    imgFileExtension = ".lsm"
End Sub

Private Sub JobSetterButton_Click()
    JobSetter.Show
    JobSetter.Repaint
    DoEvents
End Sub

Private Sub KeepParentButton_Click()
    Pipelines(currPipeline).keepParent = keepParentButton.value
End Sub

Private Sub maxWait_Click()
    Pipelines(currPipeline).maxWait = CDbl(maxWait.value)
End Sub

Private Sub optPtNumber_Click()
    Pipelines(currPipeline).optPtNumber = CInt(optPtNumber.value)
End Sub


Private Sub TimeOutButton_Click()
    TimeOut = TimeOutButton.value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''POSITIONS MANAGEMENT'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''single position''
Private Sub PositionButton1_Click()
    Dim i As Integer
    If PositionButton1 Then
        enableFrame FramePositionsControl, False
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, False
        enableFrame FrameGridLoad, False
        
    End If
End Sub

''multiple position''
Private Sub PositionButton2_Click()
    Dim i As Integer
    If PositionButton2 Then
        enableFrame FramePositionsControl, True
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, False
        enableFrame FrameGridLoad, False
    End If
End Sub

''grid''
Private Sub PositionButton3_Click()
    Dim i As Integer
    If PositionButton3 Then
        enableFrame FramePositionsControl, True
        enableFrame FrameGridControl, True
        enableFrame FrameSubGridControl, True
        enableFrame FrameGridLoad, False
    End If
End Sub

''grid based on multiple positions''
Private Sub PositionButton4_Click()
    Dim i As Integer
    If PositionButton4 Then
        enableFrame FramePositionsControl, True
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, True
        enableFrame FrameGridLoad, False
    End If
End Sub

''grid based on multiple positions''
Private Sub PositionButton5_Click()
    Dim i As Integer
    If PositionButton5 Then
        enableFrame FramePositionsControl, False
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, False
        enableFrame FrameGridLoad, True
    End If
End Sub

Private Sub AddPositionButton_Click()
    AddPosition WellID.value, Lsm5.Hardware.CpStages.PositionX, _
    Lsm5.Hardware.CpStages.PositionY, Lsm5.Hardware.CpFocus.position
End Sub

Private Sub AddPosition(ID As String, X As Double, Y As Double, Z As Double)
    With PositionsList
        .AddItem
        .List(.ListCount - 1, 0) = .ListCount
        .List(.ListCount - 1, 1) = WellID.value
        .List(.ListCount - 1, 2) = X
        .List(.ListCount - 1, 3) = Y
        .List(.ListCount - 1, 4) = Z
        .ListIndex = .ListCount - 1
    End With
End Sub

Private Sub MoveToPositionButton_Click()
    With PositionsList
        If .ListIndex > -1 Then
            FailSafeMoveStageXY CDbl(.List(.ListIndex, 2)), CDbl(.List(.ListIndex, 3))
            FailSafeMoveStageZ CDbl(.List(.ListIndex, 4))
        End If
    End With
End Sub

Private Sub UpdatePositionButton_Click()
    With PositionsList
        If .ListIndex > -1 Then
            .List(.ListIndex, 1) = WellID.value
            .List(.ListIndex, 2) = Lsm5.Hardware.CpStages.PositionX
            .List(.ListIndex, 3) = Lsm5.Hardware.CpStages.PositionY
            .List(.ListIndex, 4) = Lsm5.Hardware.CpFocus.position
        End If
    End With
End Sub

Private Sub SwitchPosition_SpinUp()
    Dim i As Integer
    Dim currIndex As Integer
    With PositionsList
        currIndex = .ListIndex
        If .ListIndex > 0 Then
            MoveListboxItem PositionsList, .ListIndex, .ListIndex - 1
            For i = 0 To .ListCount - 1
                .List(i, 0) = i + 1
            Next i
            .ListIndex = currIndex - 1
        End If
    End With
End Sub

Private Sub SwitchPosition_SpinDown()
    Dim i As Integer
    Dim currIndex As Integer
    With PositionsList
        currIndex = .ListIndex
        If .ListIndex < .ListCount - 1 And .ListIndex > -1 Then
            MoveListboxItem PositionsList, .ListIndex, .ListIndex + 1
            For i = 0 To .ListCount - 1
                .List(i, 0) = i + 1
            Next i
            .ListIndex = currIndex + 1
        End If
    End With
End Sub


Private Sub RemovePositionButton_Click()
    Dim i As Integer
    With PositionsList
        If .ListIndex > -1 Then
            .RemoveItem .ListIndex
        End If
        For i = 0 To .ListCount - 1
            .List(i, 0) = i + 1
        Next i
    End With
End Sub

Private Sub PlateType_Change()
    Dim iL As Integer
    Dim iNum As Integer
    Dim MaxINum As Integer
    Dim MaxIL As Integer
    WellID.Clear
    Select Case PlateType.ListIndex
        Case 1
            MaxIL = 0
            MaxINum = 1
        Case 2
            MaxIL = 0
            MaxINum = 2
        Case 3
            MaxIL = 1
            MaxINum = 4
        Case 4
            MaxIL = 2
            MaxINum = 4
        Case 5
            MaxIL = 8
            MaxINum = 12
    End Select
    If PlateType.ListIndex > 0 Then
        For iL = 0 To MaxIL
            For iNum = 1 To MaxINum
                WellID.AddItem "" & Lett(iL) & iNum
            Next iNum
        Next iL
        WellID.ListIndex = 0
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END POSITIONS MANAGEMENT'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''
'   fills popup menu for chosing a track for post-acquisition tracking for Job with JobName
''''
Public Sub FillTrackingChannelList(tmpTask As Task)
    Dim Success As Integer
    Dim iTrack As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim Track As DsTrack
    Dim TrackOn As Boolean
    CenterOfMassChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    If tmpTask.jobType = 0 Then
        With ImgJobs(tmpTask.jobNr)
            For iTrack = 0 To .TrackNumber - 1
                Set Track = .GetRecording.TrackObjectByMultiplexOrder(iTrack, Success)
                    If .getAcquireTrack(iTrack) Then
                    For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
                        If Track.DetectionChannelObjectByIndex(c - 1, Success).Acquire Then 'if channel is activated
                            ca = ca + 1 'counter for active channels will increase by one
                            CenterOfMassChannel.AddItem Track.Name & " " & Track.DetectionChannelObjectByIndex(c - 1, Success).Name & "-T" & iTrack + 1   'entry is added to combo box to chose track for post-acquisition tracking
                            TrackOn = True
                        End If
                    Next c
                End If
            Next iTrack
        End With
    End If
End Sub
    
Private Sub TrackXY_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setTrackXY index, TrackXY
    End If
End Sub

Private Sub TrackZ_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setTrackZ index, TrackZ
    End If
End Sub

Private Sub CenterOfMassChannel_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setTrackChannel index, CenterOfMassChannel.ListIndex
    End If
End Sub

Private Sub UpdateFocusEnabled()
    Dim index As Integer
    TrackingFrame.Visible = True
    index = CurrentPipelineList.ListIndex
    If index = -1 Then
        enableFrame TrackingFrame, False
        Exit Sub
    ElseIf Pipelines(currPipeline).getTask(index).jobType = 1 Then
        enableFrame TrackingFrame, False
        Exit Sub
    End If
    enableFrame TrackingFrame, True
    FocusMethod.Enabled = True
    CenterOfMassChannel.Enabled = True And (FocusMethod.ListIndex > NoAnalyse) And (Not FocusMethod.ListIndex = AnalyseOnline)
    TrackZ.value = Pipelines(currPipeline).getTrackZ(index)
    TrackXY.value = Pipelines(currPipeline).getTrackXY(index)
    With ImgJobs(Pipelines(currPipeline).getTask(index).jobNr)
        TrackZ.Enabled = .isZStack And (FocusMethod.ListIndex > NoAnalyse)
        TrackXY.Enabled = (FocusMethod.ListIndex > NoAnalyse) And (.Recording.ScanMode <> "ZScan") And (.Recording.ScanMode <> "Line")
    End With
    
End Sub


Private Sub FocusMethod_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index < 0 Then
        Exit Sub
    End If
    Pipelines(currPipeline).setAnalyse index, FocusMethod.ListIndex

    UpdateFocusEnabled
    If Pipelines(currPipeline).getAnalyse(index) = AnalyseOnline Then
        SaveImage = True
        Pipelines(currPipeline).setSaveImage index, True
    End If
    
End Sub


Private Sub CurrentPipelineList_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    getPeriod
    getZOffset
    getSaveImage
    If index > -1 And Pipelines(currPipeline).getTask(index).jobType = 0 Then
        FillTrackingChannelList Pipelines(currPipeline).getTask(index)
        CenterOfMassChannel.ListIndex = Pipelines(currPipeline).getTrackChannel(index)
        FocusMethod.ListIndex = Pipelines(currPipeline).getAnalyse(index)
    Else
        CenterOfMassChannel.Clear
    End If
    enableFrame FrameTaskOptions, True
    enableFrame FramePipelineRepetitions, True
    enableFrame FramePipelineTrigger, True
    UpdateFocusEnabled
    
End Sub



Private Sub AddJobToPipelineButton_Click()
    If isArrayEmpty(ImgJobs) And isArrayEmpty(FcsJobs) < 0 Then
        MsgBox "First define jobs. Press JobSetter"
        Exit Sub
    End If
    JobChoiceList.Clear
    JobChoiceFrame.Height = 126
    JobChoiceFrame.Visible = True
    JobChoiceList.SetFocus
    AddJobsToList JobChoiceList, ImgJobs
    AddJobsToList JobChoiceList, FcsJobs
    FrameTaskOptions.Visible = False
End Sub

Private Sub DelJobPipelineButton_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).delTask index
    Else
        Exit Sub
    End If
    UpdatePipelineList CurrentPipelineList, currPipeline
    If CurrentPipelineList.ListCount = 0 Then
        TrackingFrame.Visible = False
    End If
End Sub


Private Sub JobChoiceList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim index As Integer
    Dim indexImg As Integer
    Dim tmpTask As Task
    index = JobChoiceList.ListIndex
    
    If isArrayEmpty(ImgJobs) Then
        indexImg = -1
    Else
        indexImg = UBound(ImgJobs)
    End If
    
    If index > -1 Then
        With CurrentPipelineList
                .AddItem
                .List(.ListCount - 1, 0) = .ListCount
                .List(.ListCount - 1, 1) = JobChoiceList.List(index, 0)
                .List(.ListCount - 1, 2) = JobChoiceList.List(index, 1)
        End With
        If JobChoiceList.List(index, 0) = "Img" Then
            tmpTask.jobType = 0
            tmpTask.jobNr = index
        ElseIf JobChoiceList.List(index, 0) = "Fcs" Then
            tmpTask.jobType = 1
            tmpTask.jobNr = index - (indexImg + 1)
        End If
        tmpTask.SaveImage = True
        tmpTask.Period = 1
        Pipelines(currPipeline).addTask tmpTask
        If Pipelines(currPipeline).count = 1 Then
            Pipelines(currPipeline).Repetition.number = CInt(RepetitionNumber.value)
            RepetitionTimeUpdate (index)
            Pipelines(currPipeline).maxWait = CDbl(maxWait.value)
            Pipelines(currPipeline).optPtNumber = CInt(optPtNumber.value)
        End If
    End If
    JobChoiceFrame.Visible = False
    If CurrentPipelineList.ListCount > 0 Then
        If CurrentPipelineList.ListIndex < 0 Then
            CurrentPipelineList.ListIndex = 0
        End If
        TrackingFrame.Visible = True
        enableFrame FramePipelineRepetitions, True
        enableFrame FramePipelineTrigger, True
        enableFrame FrameTaskOptions, True
    End If
    FrameTaskOptions.Visible = True
End Sub


Private Sub JobUpDown_SpinDown()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index >= 0 And index < CurrentPipelineList.ListCount - 1 Then
        Pipelines(currPipeline).swapTask index, index + 1
    Else
        Exit Sub
    End If
    UpdatePipelineList CurrentPipelineList, currPipeline
    CurrentPipelineList.Selected(index + 1) = True
End Sub

Private Sub JobUpDown_SpinUp()
    Dim index As Integer
    
    index = CurrentPipelineList.ListIndex
    Debug.Print index
    If index >= 1 Then
        Pipelines(currPipeline).swapTask index, index - 1
    Else
        Exit Sub
    End If
    UpdatePipelineList CurrentPipelineList, currPipeline
    CurrentPipelineList.Selected(index - 1) = True
End Sub

Private Sub FrameButton1_Click()
    ToggleFrameButton 1
End Sub

Private Sub FrameButton2_Click()
    ToggleFrameButton 2
End Sub

Private Sub FrameButton3_Click()
    ToggleFrameButton 3
End Sub

Private Sub FrameButton4_Click()
    ToggleFrameButton 4
End Sub

Private Sub FrameButton5_Click()
    ToggleFrameButton 5
End Sub

Public Sub ToggleFrameButton(ButtonNumber As Integer)
    Dim i As Integer

    For i = 1 To NrPipelines + 2
        Me.Controls("FrameButton" & i).value = False
        Me.Controls("FrameButton" & i).BackColor = &H8000000A
    Next i
    Me.Controls("FrameButton" & ButtonNumber).value = True
    Me.Controls("FrameButton" & ButtonNumber).BackColor = &HC000&
    If ButtonNumber <= NrPipelines Then
        FramePipeline.Visible = True
        FramePositions.Visible = False
        FrameSaving.Visible = False
        currPipeline = ButtonNumber - 1
        FramePipelineTask.Caption = "Pipeline " & PipelineCaption(currPipeline) & " tasks"
        FramePipelineRepetitions.Caption = "Pipeline " & PipelineCaption(currPipeline) & " repetitions"
        FramePipelineTrigger.Caption = "Pipeline " & PipelineCaption(currPipeline) & " start/end conditions"
        UpdatePipelineList CurrentPipelineList, currPipeline
        UpdateRepetitionSettings currPipeline
        If CurrentPipelineList.ListCount > 0 Then
            If CurrentPipelineList.ListIndex = -1 Then
                CurrentPipelineList.ListIndex = 0
            End If
            enableFrame FrameTaskOptions, True
            enableFrame FramePipelineRepetitions, True
            enableFrame FramePipelineTrigger, True
            getPeriod
        Else
            enableFrame FrameTaskOptions, False
            enableFrame FramePipelineRepetitions, False
            enableFrame FramePipelineTrigger, False
        End If

        UpdateFocusEnabled
        keepParentButton.value = Pipelines(currPipeline).keepParent
        maxWait.value = Pipelines(currPipeline).maxWait
        optPtNumber.value = Pipelines(currPipeline).optPtNumber
    End If
    
    If ButtonNumber = 4 Then
        FramePipeline.Visible = False
        FrameSaving.Visible = False
        FramePositions.Visible = True
        FramePositions.Left = 65
        FramePositions.Top = 25
    End If
    
    If ButtonNumber = 5 Then
        FrameSaving.Visible = True
        FramePipeline.Visible = False
        FramePositions.Visible = False
        
        FrameSaving.Left = 73
        FrameSaving.Top = 25
    End If
    
    If ButtonNumber = 1 Then
        FramePipelineTrigger.Visible = False
    Else
        FramePipelineTrigger.Visible = True
    End If
End Sub



Private Sub AddJobsToList(List As ListBox, Jobs)
    Dim jobNr As Integer
    Dim prefix As String
    With List
        If Not isArrayEmpty(Jobs) Then
            If TypeOf Jobs(0) Is AJob Then
                prefix = "Img"
            End If
            If TypeOf Jobs(0) Is AFcsJob Then
                prefix = "Fcs"
            End If
            For jobNr = 0 To UBound(Jobs)
                .AddItem
                .List(.ListCount - 1, 0) = prefix
                .List(.ListCount - 1, 1) = Jobs(jobNr).Name
            Next jobNr
        End If
    End With
End Sub

Public Sub UpdatePipelineList(List As ListBox, index As Integer)
    Dim jobType As Integer
    Dim jobNr As Integer
    
    Dim i As Integer
    List.Clear
    If Pipelines(index).isEmpty Then
        Exit Sub
    End If
    Debug.Print "Counts " & Pipelines(index).count
    
    For i = 0 To Pipelines(index).count - 1
        If i > Pipelines(index).count - 1 Then
            GoTo Nexti
        End If
        jobType = Pipelines(index).getTask(i).jobType
        jobNr = Pipelines(index).getTask(i).jobNr
        If jobType = 0 Then
            If isArrayEmpty(ImgJobs) Then
                Pipelines(index).delTask (i)
                GoTo Nexti
            End If
        ElseIf jobType = 1 Then
            If isArrayEmpty(FcsJobs) Then
                Pipelines(index).delTask (i)
                GoTo Nexti
            End If
        End If
        With List
            If jobType = 0 Then
                If UBound(ImgJobs) >= jobNr Then
                    .AddItem
                    .List(.ListCount - 1, 0) = .ListCount
                    .List(.ListCount - 1, 1) = "Img"
                    .List(.ListCount - 1, 2) = ImgJobs(jobNr).Name
                Else
                    Pipelines(index).delTask (i)
                    GoTo Nexti
                End If
            ElseIf jobType = 1 Then
                If UBound(FcsJobs) >= jobNr Then
                    .AddItem
                    .List(.ListCount - 1, 0) = .ListCount
                    .List(.ListCount - 1, 1) = "Fcs"
                    .List(.ListCount - 1, 2) = FcsJobs(jobNr).Name
                End If
            Else
                Pipelines(index).delTask (i)
                GoTo Nexti
            End If
        End With
Nexti:
    Next i
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Looping/RepetitionSettings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RepetitionTimeUpdate(index As Integer)
    If RepetitionSec.value Then
        Pipelines(index).Repetition.Time = CDbl(RepetitionTime.value)
    ElseIf RepetitionMin.value Then
        Pipelines(index).Repetition.Time = CDbl(RepetitionTime.value) * 60
    End If
End Sub

Private Sub RepetitionMinSecUpdate(Min As Boolean)
    If Min Then
        RepetitionMin.BackColor = &HFF8080
        RepetitionSec.BackColor = &H8000000F
    Else
        RepetitionSec.BackColor = &HFF8080
        RepetitionMin.BackColor = &H8000000F
    End If
    RepetitionTime.MAX = 360
    RepetitionTimeUpdate (currPipeline)
End Sub

Private Sub RepetitionTime_Click()
    RepetitionTimeUpdate (currPipeline)
End Sub

Private Sub RepetitionInterval_Click()
    Pipelines(currPipeline).Repetition.interval = RepetitionInterval.value
End Sub

Private Sub RepetitionMin_Click()
    RepetitionSec.value = Not RepetitionMin.value
    If RepetitionMin.value Then
        RepetitionMinSecUpdate (True)
    Else
        RepetitionMinSecUpdate (False)
    End If
End Sub

Private Sub RepetitionNumber_Click()
   Pipelines(currPipeline).Repetition.number = CInt(RepetitionNumber.value)
End Sub

Private Sub RepetitionNumber_Change()
    Pipelines(currPipeline).Repetition.number = CInt(RepetitionNumber.value)
End Sub

Private Sub RepetitionSec_Click()
    RepetitionMin.value = Not RepetitionSec.value
End Sub

'''
' update form from pipeline index
'''
Private Sub UpdateRepetitionSettings(index As Integer)
    If Pipelines(index).Repetition.Time > 0 And ((Pipelines(index).Repetition.Time Mod 60) = 0 Or Pipelines(index).Repetition.Time > 360) Then
        RepetitionTime.value = Pipelines(index).Repetition.Time / 60
        RepetitionMin.value = True
    Else
        RepetitionTime.value = Pipelines(index).Repetition.Time
        RepetitionSec.value = True
    End If
    RepetitionNumber.value = Pipelines(index).Repetition.number
    RepetitionInterval.value = Pipelines(index).Repetition.interval
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' END Repetitions/Looping
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub saveImage_Click()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setSaveImage index, SaveImage.value
    End If
End Sub

Private Sub getSaveImage()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        SaveImage.value = Pipelines(currPipeline).getSaveImage(index)
    End If
End Sub


Private Sub StartOption_Click()
    setPeriod
    Period.Enabled = False
    PeriodButton.Enabled = False
End Sub

Private Sub EndOption_Click()
    setPeriod
    Period.Enabled = False
    PeriodButton.Enabled = False
End Sub

Private Sub PeriodOption_Click()
    setPeriod
    Period.Enabled = True
    PeriodButton.Enabled = True
End Sub

Private Sub PeriodButton_SpinUp()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        If Period.value < RepetitionNumber - 1 Then
            Period.value = Period.value + 1
        Else
            MsgBox "The period of acquisition cannot be higher than number of Repetitions , i.e. " & RepetitionNumber, VbInformation
            Exit Sub
        End If
        Pipelines(currPipeline).setPeriod index, Period.value
    End If
End Sub

Private Sub PeriodButton_SpinDown()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        If Period.value > 1 Then
            Period.value = Period.value - 1
        End If
        Pipelines(currPipeline).setPeriod index, Period.value
    End If
End Sub

Private Sub setPeriod()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        If PeriodOption Then
            Pipelines(currPipeline).setPeriod index, Period.value
        End If
        If StartOption Then
            Pipelines(currPipeline).setPeriod index, 0
        End If
        If EndOption Then
            Pipelines(currPipeline).setPeriod index, -1
        End If
    End If
End Sub

Private Sub getPeriod()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        With Pipelines(currPipeline)
            If .getPeriod(index) > 0 Then
                Period.value = .getPeriod(index)
                PeriodOption.value = True
                Period.Enabled = True
                
            End If
            If .getPeriod(index) = 0 Then
                StartOption.value = True
                Period.Enabled = False
            End If
            If .getPeriod(index) = -1 Then
                EndOption.value = True
                Period.Enabled = False
            End If
        End With
    End If
End Sub

Private Sub getZOffset()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        ZOffset.value = Pipelines(currPipeline).getZOffset(index)
    End If
End Sub





Private Sub ZOffset_Change()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setZOffset index, ZOffset.value
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub EventMng_Ready()
    MicStatus.Caption = "READY"
    MicStatus.ForeColor = "&H00008000"
End Sub

Private Sub EventMng_Busy(Task As Integer)
    MicStatus.Caption = "BUSY"
    MicStatus.ForeColor = "&H00000080"
End Sub







Private Sub StopButton_Click()
    StopAcquisition
End Sub

Private Sub StopFcsButton_Click()
    StopAcquisition
End Sub
















Private Function UniqueListName(List As ListBox, JobName As String) As Boolean
    Dim ListEntry As Variant
    If List.ListCount > 0 Then
        For Each ListEntry In List.List
            If StrComp(ListEntry, JobName) = 0 Then
                Exit Function
            End If
        Next
    End If
    UniqueListName = True
End Function
    

Public Function HideShowForms(OpenForms() As Boolean) As Boolean()
    Dim UForm As Object
    Dim i As Integer
    If isArrayEmpty(OpenForms) Then
    
        For Each UForm In VBA.UserForms
            If isArrayEmpty(OpenForms) Then
                 ReDim OpenForms(0)
                 OpenForms(0) = UForm.Visible
            Else
                ReDim Preserve OpenForms(UBound(OpenForms) + 1)
                OpenForms(UBound(OpenForms)) = UForm.Visible
            End If
            If UForm.Visible = True Then
                UForm.Hide
            End If
        Next
        HideShowForms = OpenForms
    Else
        i = 0
        For Each UForm In VBA.UserForms
            If OpenForms(i) Then
                UForm.Show
            End If
            i = i + 1
        Next
    End If
End Function


Private Sub setTrackNames(index As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim iTrack As Integer
    Dim Track As DsTrack
    Dim ChannelOK As Boolean
    Dim AcquireTrack() As Boolean
    Dim MaxTracks As Long
    MaxTracks = ImgJobs(index).Recording.GetNormalTrackCount
    AcquireTrack = ImgJobs(index).AcquireTrack
    For i = 0 To 3
        If iTrack < 5 Then
            ChannelOK = False
            Set Track = ImgJobs(index).Recording.TrackObjectByMultiplexOrder(i, 1)
            For j = 0 To Track.DataChannelCount - 1
                If Track.DataChannelObjectByIndex(j, 1).Acquire = True Then
                    ChannelOK = True
                End If
            Next j
            If ChannelOK And (Not Track.IsLambdaTrack) And (Not Track.IsBleachTrack) Then
                Me.Controls("Track" + CStr(iTrack + 1)).Visible = True
                Me.Controls("Track" + CStr(iTrack + 1)).value = AcquireTrack(iTrack)
                Me.Controls("Track" + CStr(iTrack + 1)).Caption = Track.Name
                Me.Controls("Track" + CStr(iTrack + 1)).Enabled = True
                iTrack = iTrack + 1
            Else
                Me.Controls("Track" + CStr(iTrack + 1)).Visible = False
                iTrack = iTrack + 1
            End If
        End If
    Next i
End Sub



Public Sub AddJob(JobsV() As AJob, Name As String, Recording As DsRecording, ZEN As Object)
    If isArrayEmpty(JobsV) Then
        ReDim JobsV(0)
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) + 1)
    End If
    Set JobsV(UBound(JobsV)) = New AJob
    JobsV(UBound(JobsV)).Name = Name
    JobsV(UBound(JobsV)).SetJob Lsm5.DsRecording, ZEN
End Sub


Public Sub AddFcsJob(JobsV() As AFcsJob, Name As String, ZEN As Object)
    If isArrayEmpty(JobsV) Then
        ReDim JobsV(0)
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) + 1)
    End If
    Set JobsV(UBound(JobsV)) = New AFcsJob
    JobsV(UBound(JobsV)).Name = Name
    JobsV(UBound(JobsV)).SetJob ZEN, ZenV
End Sub










''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output Folder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CommandButtonNewDataBase_Click()
    Dim Filter As String, fileName As String
    Dim Flags As Long
    Dim DefDir As String

    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE

    'Filter = "All Data (*.*)" & Chr$(0) & "*.*"
    If GlobalDataBaseName = "" Then
        DefDir = "C:\"
    Else
        DefDir = GlobalDataBaseName
    End If
    
    fileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", DefDir, "Select output folder")
    If Len(fileName) > 3 Then
        fileName = VBA.Left(fileName, Len(fileName) - 3)
        DatabaseTextbox.value = fileName
        SetDatabase
    End If
End Sub

'''''
'   Only update the outputfolder when enter is pressed. This avoids creating a folder at every keystroke
'''''
Private Sub DatabaseTextbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'this is the enter key
        SetDatabase
    End If
End Sub

Private Sub TextBoxFileName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'this is the enter key
        SetFileName
    End If
End Sub

'''''
'   Set global variables for files and check if we can create Outputfolder
'       [GlobalDataBaseName] Out/Global - The name of Outputfolder
'       [LogFileNameBase]    Out/Global - The name of the LogfileName
'       Log]                Out/Global - If yes results are logged
'''''
Private Sub SetDatabase()

    GlobalDataBaseName = DatabaseTextbox.value

    If Not GlobalDataBaseName = "" Then
        If VBA.Right(GlobalDataBaseName, 1) <> "\" Then
            DatabaseTextbox.value = DatabaseTextbox.value + "\"
            GlobalDataBaseName = DatabaseTextbox.value
        End If
        On Error GoTo ErrorHandleDataBase
        If Not CheckDir(GlobalDataBaseName) Then
            Exit Sub
        End If
        OiaSettings.writeKeyToRegistry "OutputFolder", GlobalDataBaseName
        LogFileNameBase = GlobalDataBaseName & "\PipelineConstructor.log"
        ErrFileNameBase = GlobalDataBaseName & "\PipelineConstructor.err"
        If VBA.Right(GlobalDataBaseName, 1) = "\" Then
            BackSlash = ""
        Else
            BackSlash = "\"
        End If
    End If

    If LogFileNameBase <> "" Then
        On Error GoTo ErrorHandleLogFile
        LogFileName = LogFileNameBase
        ErrFileName = ErrFileNameBase
        Log = True
    Else
        Log = False
    End If
    Exit Sub
ErrorHandleDataBase:
    MsgBox "Could not create output Directory " & GlobalDataBaseName
    Exit Sub
ErrorHandleLogFile:
    MsgBox "Could not create LogFile " & LogFileName
End Sub



Private Sub SetFileName()
    If TextBoxFileName.value <> "" Then
        If VBA.Right(TextBoxFileName.value, Len(FNSep)) <> FNSep Then
            TextBoxFileName.value = TextBoxFileName.value & FNSep
        End If
    End If
End Sub

Public Sub LoadFormSettings(fileName As String)
    Dim iFileNum As Integer, ipip As Integer, iSet As Integer
    Dim tsk As Task
    Dim arr() As Variant
    Dim Fields As String
    Dim JobName As String
    Dim FieldEntries() As String
    Close
    'On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open fileName For Input As iFileNum
    arr = TaskToArray(tsk)
    Pipelines(0).delAllTasks
    Pipelines(1).delAllTasks
    Pipelines(2).delAllTasks
    Do While Not EOF(iFileNum)
            Line Input #iFileNum, Fields
            While VBA.Left(Fields, 1) = "%"
                Line Input #iFileNum, Fields
            Wend
            If Fields <> "" Then
                FieldEntries = Split(Fields, " ")
                ipip = CInt(FieldEntries(1))
                If FieldEntries(2) = "Reptime" Then
                    Pipelines(ipip).Repetition.Time = CDbl(FieldEntries(3))
                    Pipelines(ipip).Repetition.number = CInt(FieldEntries(5))
                    Pipelines(ipip).Repetition.interval = CBool(FieldEntries(7))
                End If
                If FieldEntries(2) = "Tsk" Then
                    For iSet = 0 To UBound(arr)
                        Select Case VarType(arr(iSet))
                            Case vbInteger
                                arr(iSet) = CInt(FieldEntries(iSet * 2 + 5))
                            Case vbDouble
                                arr(iSet) = CDbl(FieldEntries(iSet * 2 + 5))
                            Case vbBoolean
                                arr(iSet) = CBool(FieldEntries(iSet * 2 + 5))
                            Case vbLong
                                arr(iSet) = CLng(FieldEntries(iSet * 2 + 5))
                        End Select
                    Next iSet
                    Pipelines(ipip).addTask ArrayToTask(arr)
                End If
            End If
    Loop
    
    UpdatePipelineList PipelineConstructor.CurrentPipelineList, currPipeline
    UpdateRepetitionSettings currPipeline
    UpdateFocusEnabled
    getPeriod
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & fileName & " for AutofocusScreen settings"
End Sub
