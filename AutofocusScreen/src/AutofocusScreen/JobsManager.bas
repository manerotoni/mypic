Attribute VB_Name = "JobsManager"
''''
'' A Class to manage and imagingJob with different Settings and Tracks
''''
''Public Type AJob
''    Name As String
''    Recording As DsRecording
''    AcquireTrack() As Boolean
''    TrackNumber As Integer
''    TimeBetweenStacks As Double
''End Type
'
Public Reps As ImagingRepetitions
Public Grids As ImagingGrids
Public Jobs As ImagingJobs
Private Const UpdateFormAsRunning = True
Public Type StagePosition
  X As Double
  Y As Double
  Z As Double
End Type


'''
'   Execute an imaging Job
'       JobName: The name of the Job to execute
'       RecordingDoc: the dsRecording where image is stored
'       X, Y, and Z: The position of stage and focus (central slice)
'''
Public Function AcquireJob(JobName As String, RecordingDoc As DsRecordingDoc, RecordingName As String, StgPos As StagePosition) As Boolean
    Dim SuccessRecenter As Boolean
    'stop any running jobs
    AutofocusForm.StopScanCheck

    'Creates a NewRecord if required
    NewRecord RecordingDoc, RecordingName, 0
    'move stage if required
    If Round(Lsm5.Hardware.CpStages.PositionX, PrecXY) <> Round(StgPos.X, PrecXY) Or Round(Lsm5.Hardware.CpStages.PositionY, PrecXY) <> Round(StgPos.Y, PrecXY) Then
        If Not FailSafeMoveStageXY(StgPos.X, StgPos.Y) Then
            AutofocusForm.StopAcquisition
            Exit Function
        End If
    End If
    
    'Change settings for new Job
    Jobs.PutJob JobName, ZEN
    
    'Not sure if this is required
    If Jobs.GetSpecialScanMode(JobName) = "ZScanner" Then
        Lsm5.Hardware.CpHrz.Leveling
    End If
    If Not Recenter_pre(StgPos.Z, SuccessRecenter, ZENv) Then
        Exit Function
    End If

    'Acquire the image
    If Not ScanToImage(RecordingDoc) Then
        Exit Function
    End If
    
    'wait that slice recentered after acquisition
    If Not Recenter_post(StgPos.Z, SuccessRecenter, ZENv) Then
       Exit Function
    End If
    

    AcquireJob = True
End Function



'''''
' This executes part of the Job save the file compute offline tracking and set the registry
'''''
Public Function ExecuteJob(JobName As String, RecordingDoc As DsRecordingDoc, FilePath As String, FileName As String, _
StgPos As StagePosition, Optional deltaZ As Integer = -1)
    
    If Not AcquireJob(JobName, RecordingDoc, FileName, StgPos) Then
        Exit Function
    End If
    'this is a dummy variable used for consistencey except for autofocus the default is saving of all images
    
    If AutofocusForm.Controls(JobName & "SaveImage") Then
        If Not SaveDsRecordingDoc(RecordingDoc, FilePath & FileName) Then
            Exit Function
        End If
    End If
    
    StgPos = TrackOffLine(JobName, RecordingDoc, StgPos)
    
    If ScanStop Then
        Exit Function
    End If
    ExecuteJob = True
End Function



''''
'   Update StgPos to StgPosNew according the weather Coordinates are tracked or not
''''
Public Function TrackJob(JobName As String, StgPos As StagePosition, StgPosNew As StagePosition) As StgPos
    TrackJob = StgPos
    If AutofocusForm.Controls(JobName & "TrackZ") Then
        TrackJob.Z = StgPosNew.Z
    End If
    If AutofocusForm.Controls(JobName & "TrackXY") Then
        TrackJob.X = StgPosNew.X
        TrackJob.Y = StgPosNew.Y
    End If
End Function


''''''
'   StartAcquisition(BleachingActivated)
'   Perform many things (TODO: write more). Pretty much the whole macro runs through here
''''''
Public Sub StartJobOnGrid(GridName As String, JobName As String, ParentPath As String)
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    Dim StgPos As StagePosition
    Dim FileName As String
    Dim deltaZ As Integer
    deltaZ = -1
    Dim SuccessRecenter As Boolean
      
    'Stop all running acquisitions (maybe to strong)
    AutofocusForm.StopScanCheck
    
    'coordinates
    Dim previousZ As Double   'remember position of previous position in Z
    
    'block usage of grid during acquisition
    AutofocusForm.SwitchEnableGridScanPage False
    
    
    'Coordinates
    Dim X As Double              ' x value where to move the stage (this is used as reference)
    Dim Y As Double              ' y value where to move the stage
    Dim Z As Double              ' z value where to move the stage
    Dim Xold As Double
    Dim Yold As Double
    Dim Zold As Double
    
    'test variables
    Dim Success As Integer       ' Check if something was sucessfull
    
    'Recording stuff
    Dim FilePath As String   ' full path of file to save (changes through function)
    Dim Scancontroller As AimScanController ' the controller
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    
    OiaSettings.resetRegistry
    OiaSettings.readFromRegistry
    
    NewRecord GlobalRecordingDoc, TextBoxFileName & Grids.getName(JobName, 1, 1, 1, 1) & Grids.suffix(JobName, 1, 1, 1, 1) & Reps.suffix(JobName, 1), 0
        
    
    InitializeStageProperties
    SetStageSpeed 9, True    'What do ou do here

    Running = True  'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.
    
    
    previousZ = Grids.getZ(JobName, 1, 1, 1, 1)
    Reps.resetIndex (JobName)
    While Reps.nextRep(JobName) ' cycle all repetitions
        Grids.setIndeces JobName, 1, 1, 1, 1
        Do ''Cycle all positions defined in grid
            If Grids.getThisValid(JobName) Then
                'Do some positional Job
                StgPos.X = Grids.getThisX(JobName)
                StgPos.Y = Grids.getThisY(JobName)
                StgPos.Z = Grids.getThisZ(JobName)
                
                If Reps.getIndex(JobName) = 1 And GridScanActive Then
                    StgPos.Z = previousZ
                End If

                    
                ' Recenter and move where it should be
                If JobName = "Global" Then
                    'recenter only after activation of new track
                    If AutofocusActive Then
                        FileName = FileNameFromGrid(GridName, "Autofocus")
                        FilePath = ParentPath & FilePathSuffix(GridName, "Autofocus") & "\"
                        If Not ExecuteJob("Autofocus", GlobalRecordingDoc, FilePath, FileName, StgPos) Then
                            GoTo StopAcquisition
                        End If
                        newStgPos = ComputeJobSequential("Autofocus", GlobalRecordingDoc, FilePath, FileName, StgPos)
                        StgPos = TrackJob("Autofocus", StgPos, newStgPos)
                    End If
                    
                    If AcquisitionActive Then
                        FileName = FileNameFromGrid(GridName, "Acquisition")
                        FilePath = ParentPath & FilePathSuffix(GridName, "Acquisition") & "\"
                        StgPos.Z = StgPos.Z + AcquisitionZOffset.Value
                        If Not ExecuteJob("Acquisition", GlobalRecordingDoc, FilePath, FileName, StgPos) Then
                            GoTo StopAcquisition
                        End If
                        StgPos.Z = StgPos.Z - AcquisitionZOffset.Value
                    End If
                    
                    If AlterAcquisitionActive Then
                        FileName = FileNameFromGrid(GridName, "AlterAcquisition")
                        FilePath = ParentPath & FilePathSuffix(GridName, "AlterAcquisition") & "\"
                        StgPos.Z = StgPos.Z + AlterAcquisitionZOffset.Value
                        If Not ExecuteJob("AlterAcquisition", GlobalRecordingDoc, FilePath, FileName, StgPos) Then
                            GoTo StopAcquisition
                        End If
                        StgPos.Z = StgPos.Z - AlterAcquisitionZOffset.Value
                    End If
                Else
                    FileName = FileNameFromGrid(GridName, JobName)
                    FilePath = ParentPath & FilePathSuffix(GridName, JobName) & "\"
                    StgPos.Z = StgPos.Z + AutofocusForm.Controls(JobName + "ZOffset").Value
                    If Not ExecuteJob(JobName, GlobalRecordingDoc, FilePath, FileName, StgPos) Then
                        GoTo StopAcquisition
                    End If
                    StgPos.Z = StgPos.Z - AutofocusForm.Controls(JobName + "ZOffset").Value
                    
                End If
                Grids.setThisX = StgPos.X
                Grids.setThisY = StgPos.Y
                Grids.setThisZ = StgPos.Z
                previousZ = Grids.getThisZ(JobName)
                
            End If
        Loop While Grids.nextGridPt(JobName)
        ''Wait till next repetition
        Reps.updateTimeStart (JobName)
        If Reps.wait(JobName) > 0 Then
            DisplayProgress "Waiting " & CStr(CInt(Reps.wait(JobName))) & " s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
            DoEvents
        Else
            DisplayProgress "Waiting " & "0 s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
        End If
        
        While Reps.wait(JobName) > 0
            Sleep (100)
            DoEvents
            If ScanPause = True Then
                If Not Pause Then ' Pause is true is Resume
                    ScanStop = True
                    AutofocusForm.StopAcquisition
                    Exit Sub
                End If
            End If
            If ScanStop Then
                AutofocusForm.StopAcquisition
                Exit Sub
            End If
            DisplayProgress "Waiting " & CStr(CInt(Reps.wait(JobName))) & " s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
        Wend
    Wend
StopAcquisition:
    AutofocusForm.StopAcquisition
    AutofocusForm.SwitchEnableGridScanPage True
End Sub

'''
' Derive filename from Grid and repetition
'''
Private Function FileNameFromGrid(GridName As String, JobName As String) As String
     FileNameFromGrid = TextBoxFileName.Value & Grids.getThisName(GridName) & JobShortNames(JobName) & "_" & Grids.thisSuffix(GridName) & Reps.thisSuffix(GridName)
End Function

'''
' Derive filepath Suffix from Grid and repetition
'''
Private Function FilePathSuffix(GridName As String, JobName As String) As String
    FilePathSuffix = TextBoxFileName.Value & Grids.getThisName(GridName) & JobShortNames(JobName)
    If (Grids.numCol(GridName) * Grids.numRow(GridName) = 1 And Grids.numColSub(GridName) * Grids.numRowSub(GridName) = 1) Then
        FilePathSuffix = FilePathSuffix & "_" & Grids.thisSuffix(GridName)
        Exit Function
    End If
    If (Grids.numCol(GridName) * Grids.numRow(GridName) > 1 And Not Grids.numColSub(GridName) * Grids.numRowSub(GridName) > 1) _
    Or (Not Grids.numCol(GridName) * Grids.numRow(GridName) > 1 And Grids.numColSub(GridName) * Grids.numRowSub(GridName) > 1) Then
        FilePathSuffix = FilePathSuffix & "_" & Grids.thisSuffix(GridName)
    Else
        FilePathSuffix = FilePathSuffix & "_" & Grids.thisSuffixWell(GridName) & "\" & FilePathSuffix & "_" & Grids.thisSuffix(GridName)
    End If
End Function





''''
' SetDefaultRecordings()
' Load default recording settings from ZEN
' Obsolete
'''
Public Sub SetDefaultRecordings()
    
    Set GlobalAutoFocusRecording = Lsm5.CreateBackupRecording
    Set GlobalAcquisitionRecording = Lsm5.CreateBackupRecording
    Set GlobalTrigger1Recording = Lsm5.CreateBackupRecording
    Set GlobalBleachRecording = Lsm5.CreateBackupRecording
    Set GlobalAltRecording = Lsm5.CreateBackupRecording
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    GlobalAutoFocusRecording.Copy Lsm5.DsRecording
    GlobalAcquisitionRecording.Copy Lsm5.DsRecording
    GlobalTrigger1Recording.Copy Lsm5.DsRecording
    GlobalBleachRecording.Copy Lsm5.DsRecording
    GlobalAltRecording.Copy Lsm5.DsRecording
    GlobalBackupRecording.Copy Lsm5.DsRecording ' this will not be changed remains always the same
    GlobalBackupSampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    Dim i As Long
    Dim NrTracks As Long
    ReDim GlobalBackupActiveTracks(Lsm5.DsRecording.TrackCount)
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       GlobalBackupActiveTracks(i) = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i

End Sub


'''
'   Update the settings of the corresponding Formpage from the Job
'''
Public Sub UpdateFormFromJob(Jobs As ImagingJobs, JobName As String)
    
    'update form for any new tracks
    'AutofocusForm.AutoFindTracks
    Dim i As Integer
    Dim Record As DsRecording
    Set Record = Jobs.GetRecording(JobName)
    
    For i = 0 To TrackNumber - 1
       AutofocusForm.Controls(JobName + "Track" + CStr(i + 1)).Value = Jobs.GetAcquireTrack(JobName, i)
    Next i

End Sub

'''
'   Update the settings of Job with JobName from corresponding Formpage
'''
Public Sub UpdateJobFromForm(Jobs As ImagingJobs, JobName As String)
    Dim i As Integer
    
    For i = 0 To TrackNumber - 1
       Jobs.SetAcquireTrack JobName, i, AutofocusForm.Controls(JobName + "Track" + CStr(i + 1)).Value
    Next i

End Sub


Public Function TrackOffLine(JobName As String, RecordingDoc As DsRecordingDoc, StgPos As StagePosition) As StagePosition
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    TrackOffLine = StgPos
    If AutofocusForm.Controls(JobName & "OfflineTrack") Then
        TrackingChannel = AutofocusForm.Controls(JobName & "OfflineTrackChannel").List(AutofocusForm.Controls(JobName & "OfflineTrackChannel").ListIndex)
        MassCenter RecordingDoc, TrackingChannel, XMass, YMass, ZMass
        ComputeShiftedCoordinates XMass, YMass, ZMass, X, Y, Z
    Else
        X = StgPos.X
        Y = StgPos.Y
        Z = StgPos.Z
    End If
    If AutofocusForm.Controls(JobName & "TrackZ") Then
        TrackOffLine.Z = Z
    End If
    If AutofocusForm.Controls(JobName & "TrackXY") Then
        TrackOffLine.X = X
        TrackOffLine.Y = Y
    End If
End Function

''''
'   Perform computation for tracking and wait for Onlineimageanalysis if sequential Oia is on
''''
Public Function ComputeJobSequential(JobName As String, RecordingDoc As DsRecordingDoc, FilePath As String, FileName As String, X As Double, _
Y As Double, Z As Double, Optional deltaZ As Integer = -1) As Double()
    
    If AutofocusForm.Controls(JobName & "OiaActive") And AutofocusForm.Controls(JobName & "OiaSequential") Then
        SaveSetting "OnlineImageAnalysis", "macro", "code", "wait"
        SaveSetting "OnlineImageAnalysis", "macro", "filePath", FilePath & FileName
    Else
    
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    Dim NewCoord() As Double
    ReDim NewCoord(3)
    'Defaults we dont change anything
    deltaZ = -1
    NewCoord(0) = X
    NewCoord(1) = Y
    NewCoord(2) = Z
    NewCoord(3) = deltaZ
    Dim XMass As Double
    Dim YMass As Double
    Dim ZMass As Double
    Dim TrackingChannel As String
    
    OiaSettings.readFromRegistry

    If AutofocusForm.Controls(JobName & "OiaActive") And AutofocusForm.Controls(JobName & "OiaSequential") Then
        'getValues from registry
        
        'NewCoord = ListenToRegistry
        ComputeJobSequential = NewCoord
    End If
End Function

Public Function ComputeJobSequential(CurrentJob As String, CurrentPath As String, RecordingDoc As DsRecordingDoc, Coord() As Double) As Double()
    Dim NewCoord() As Double
    ReDim NewCoord(3)
    NewCoord(0) = Coord(0)
    NewCoord(1) = Coord(1)
    NewCoord(2) = Coord(2)
    NewCoord(3) = Coord(3)

    Dim code As String
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    Dim StgPos As StagePosition
    code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="code")
    Dim TimeWait, timeStart, MaxTimeWait As Double
    MaxTimeWait = 100

    Select Case code
        Case "1", "wait":
            'Wait for image analysis to finish
            DisplayProgress "Waiting for image analysis...", RGB(0, &HC0, 0)
            timeStart = CDbl(GetTickCount) * 0.001
            Do While ((TimeWait < MaxTimeWait) And (code = "1" Or code = "wait" Or code = "Wait" Or code = "0"))
                Sleep (50)
                TimeWait = CDbl(GetTickCount) * 0.001 - timeStart
                code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", _
                          key:="Code")
                DoEvents
                If ScanStop Then
                    Exit Function
                End If
            Loop

            If TimeWait > MaxTimeWait Then
                code = "nothing"
                SaveSetting "OnlineImageAnalysis", "macro", "code", code
            End If
    End Select

    ''Read all settings at once
    OiaSettings.readFromRegistry
    
    'read coordinates
    
    
    If OiaSettings.Settings.Item("X") <> "" And OiaSettings.Settings.Item("Y") <> "" Then
        OiaSettings.getPositions (StgPos)
    End If
    
    Select Case OiaSettings.Settings.Item("code")
        Case "2", "nothing", "Nothing", "DoNothing", "doNothing", "donothing":  'Nothing to do
            ListenToRegistry = NewCoord
        Case "3", "Trigger1Position", "trigger1Position": 'store positions for later processing
            SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
            DisplayProgress "Registry Code 3 (Trigger1Position): store positions and do nothing ...", RGB(0, &HC0, 0)
            If OiaSettings.getPositions(StgPos) Then
                computeCoordinatesImaging StgPos, NewCoord
            End If
                
                
        Case "5", "trigger1", "Trigger1":
            SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
            DisplayProgress "Registry Code 4 (Trigger1): store positions and do imaging Trigger1...", RGB(0, &HC0, 0)

            StorePositionsFromRegistry Xref, Yref, Zref, X, Y, Z, deltaZ
            ' BatchHighresImagingRoutine
            ' HERE THE IMAGES ARE ACQUIRED
            Repetitions.number = CInt(AutofocusForm.Trigger1RepetitionNumber.Value)
            Repetitions.time = CDbl(AutofocusForm.Trigger1RepetitionTime.Value)
            Repetitions.interval = True
            SubImagingWorkFlow RecordingDoc, GlobalTrigger1Recording, "Trigger1", AutofocusForm.Trigger1Autofocus, AutofocusForm.Trigger1ZOffset, Repetitions, _
            X, Y, Z, deltaZ, GridPos, FileNameID
            'create empty arrays
            Erase X
            Erase Y
            Erase Z
            Erase deltaZ

        Case "6", "fcs":
            SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
            DisplayProgress "Registry Code 6 (fcs): peform a fcs measurment ...", RGB(0, &HC0, 0)
            'create empty arrays
            Erase locX
            Erase locY
            Erase locZ
            Erase locDeltaZ
            DisplayProgress "Registry Code 6 (fcs): perform FCS measurement ...", RGB(0, &HC0, 0)
            StorePositionsFromRegistry Xref, Yref, Zref, locX, locY, locZ, locDeltaZ
            DisplayProgress "Registry Code 6 (fcs): perform FCS measurement ...", RGB(0, &HC0, 0)
            SubImagingWorkFlowFcs FcsData, locX, locY, locZ, GlobalDataBaseName & BackSlash, AutofocusForm.TextBoxFileName.Value & UnderScore & FileNameID, _
            RecordingDoc.Recording.SampleSpacing, RecordingDoc.Recording.FrameSpacing

        Case "7", "Trigger2", "trigger2":
            'This only specify position of ROIs the stage is not moved
            SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
            DisplayProgress "Registry Code 7 (bleach): peform a bleach measurment ...", RGB(0, &HC0, 0)
            ' read potitions from
            ReDim locX(0)
            ReDim locY(0)
            ReDim locZ(0)
            ReDim locDeltaZ(0)
            locX(0) = Xref
            locY(0) = Yref
            locZ(0) = Zref + AutofocusForm.Trigger2ZOffset.Value
            locDeltaZ(0) = -1

            Repetitions.number = 1
            Repetitions.time = 0
            Repetitions.interval = True
            DisplayProgress "Registry Code 8 (bleach): perform Bleaching ...", RGB(0, &HC0, 0)
            'Delete All ROI's
            Dim AcquisitionController As AimAcquisitionController40.AimScanController
            Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
            Dim vo As AimImageVectorOverlay
            Set vo = AcquisitionController.AcquisitionRegions
            vo.Cleanup
            CreateRoisFromRegistry

            'StorePositionsFromRegistry Xref, Yref, Zref, HighResArrayX, HighResArrayY, HighResArrayZ, HighResArrayDeltaZ
            DisplayProgress "Registry Code 8 (bleach): perform Bleaching ...", RGB(0, &HC0, 0)
            SubImagingWorkFlow RecordingDoc, GlobalBleachRecording, "Bleach", AutofocusForm.Trigger1Autofocus, AutofocusForm.Trigger1ZOffset, Repetitions, _
            locX, locY, locZ, locDeltaZ, GridPos, FileNameID, iRepStart
            Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
            Track.UseBleachParameters = False  'switch off the bleaching

        Case Else
            MsgBox ("Invalid OnlineImageAnalysis Code = " & code)
            Exit Function
    End Select
    MicroscopePilot = True

End Function




Public Function ComputeJobParallel(JobName As String, Recording As DsRecording, FilePath As String, FileName As String, X As Double, _
Y As Double, Z As Double, Optional deltaZ As Integer = -1) As Double()
    Dim OiaSettings As Dictionary
    Dim NewCoord() As Double
    ReDim NewCoord(3)
    'Defaults we dont change anything
    deltaZ = -1
    NewCoord(0) = X
    NewCoord(1) = Y
    NewCoord(2) = Z
    NewCoord(3) = deltaZ
    If AutofocusForm.Controls(JobName & "OiaActive") And AutofocusForm.Controls(JobName & "OiaParalle") Then
        ComputeJobParallel = NewCoord
    End If
End Function

'
'
'
'
'Public Sub GetJob(Jobs As Collection, JobName As String)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).GetJob
'End Sub
'
'Public Sub SetJob(Jobs As Collection, JobName As String)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).SetJob
'End Sub
'
'Public Sub UpdateJobLinesPerFrame(Jobs As Collection, JobName As String, Value As Integer)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).LinesPerFrame Value
'End Sub
'
'Public Sub UpdateJobSamplesPerLine(Jobs As Collection, JobName As String, Value As Integer)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).SamplesPerLine Value
'End Sub
'
'Public Sub UpdateJobSpecialScanMode(Jobs As Collection, JobName As String, Value As String)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).SpecialScanMode Value
'End Sub
'
'Public Sub UpdateJobScanDirection(Jobs As Collection, JobName As String, Value As Integer)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).ScanDirection Value
'End Sub
'
'Public Sub UpdateJobStacksPerRecord(Jobs As Collection, JobName As String, Value As Integer)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).StacksPerRecord Value
'End Sub
'
'Public Sub UpdateJobZoom(Jobs As Collection, JobName As String, Value As Double)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).Zoom Value
'End Sub
'
'Public Sub UpdateJobStacks(Jobs As Collection, JobName As String, ZRange As Double, ZStep As Double)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).FramesPerStack CLng(ZRange / ZStep) + 1
'    Jobs(iJob).FrameSpacing ZStep
'End Sub
'
'Public Sub UpdateJobFramesPerStack(Jobs As Collection, JobName As String, FramesPerStack As Long)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).FramesPerStack FramesPerStack
'End Sub
'
'Public Sub UpdateJobFrameSpacing(Jobs As Collection, JobName As String, FrameSpacing As Double)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).FrameSpacing FrameSpacing
'End Sub
'
'Public Sub UpdateJobFrameSize(Jobs As Collection, JobName As String, FrameSize As Integer)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).FrameSize FrameSize
'End Sub
'
'''''''
''    UpdateJobTimeSeries: if True the Job also has time Series
''''''
'Public Sub UpdateJobTimeSeries(Jobs As Collection, JobName As String, Value As Boolean)
'    Dim iJob As Integer
'    iJob = JobsDic(JobName)
'    Jobs(iJob).TimeSeries Value
'End Sub
'
'
'Public Function TestImgJob()
'    Dim JobNames(1) As String
'    JobNames(0) = "Laser"
'    JobNames(1) = "Space"
'    Dim JobsTest As ImagingJobs
'    Set JobsTest = New ImagingJobs
'    JobsTest.Initialize JobNames, Lsm5.DsRecording
'    JobsTest.SetAcquireTrack "Laser", 1, True
'    JobsTest.SetAcquireTrack "Laser", 0, False
'    JobsTest.SetFramesPerStack "Laser", 3
'    JobsTest.PutJob "Laser"
'
'
''    Dim Name As Variant
''    For Each Name In JobNames
''        MsgBox Name
''    Next Name
''    Set ZEN = Lsm5.CreateObject("Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface")
''    Set Jobs = New Collection
''    Dim Job1 As ImagingJob
''    Dim Record As DsRecording
''    Dim Track As DsTrack
''    Set Job1 = New ImagingJob
''    Job1.Name = "Job1"
''    Job1.GetJob
''    Job1.TimeBetweenStacks = 5
''    Set Job2 = New ImagingJob
''
''    Job2.Name = "Job2"
''    Job2.GetJob
''    Set Record = Job2.Recording
''    Set Track = Job2.Tracks(0)
''    Job2.TimeBetweenStacks = 10
''    Job1.StacksPerRecord 2
''    Job2.StacksPerRecord 3
''    'ZEN.gui.Acquisition.TimeSeries.Interval.Value = Job2.TimeBetweenStacks
''
''    Job1.SetJob
''    'ZEN.gui.Acquisition.TimeSeries.Interval.Value = Job2.TimeBetweenStacks
'    'ZEN.SetDouble "TimeSeries.End.Duration", 18
'
'    Job2.SetJob
'
'    Job1.SetJob
''    Dim Job2 As ImagingJob
''    Set Job2 = New ImagingJob
''    Job2.Name = "Job2"
''    Jobs.Add Job1
''    Jobs.Add Job2
''    ReDim Jobs(1)
''    Jobs(0).Name = "Hello"
''    Jobs(0).GetSettings
''    Jobs(0).Recording.FramesPerStack = 5
''    Jobs(0).SetSettings
' '   MsgBox Jobs.Item(2).Name
'End Function
