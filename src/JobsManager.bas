Attribute VB_Name = "JobsManager"
'---------------------------------------------------------------------------------------
' Module    : JobsManager
' Author    : Antonio Politi
' Date      : 29/08/2013
' Purpose   : Functions to perform imaging and fcs using the Imging, Fcs, Grid, repetitions classes. The functions
'             also access the form AutofocusForm using the same name identifier for the jobs
'---------------------------------------------------------------------------------------

Option Explicit
Public Type Task
    Analyse As Integer
    jobType As Integer
    jobNr As Integer
    Period As Long
    SaveImage As Boolean
    TrackXY  As Boolean
    TrackZ As Boolean
    TrackChannel As Integer
    ZOffset As Double
End Type

Public wellPt() As WellPoint
Public Const LogLevel = 0


'name of the repetitions
Public RepNames() As String
Public Enum jobTypes
    imgjob = 0
    fcsjob = 1
    gotoPip = 2
End Enum



Public Enum AnalyseImage
    No = 0
    CenterOfMassThr = 1
    Peak = 2
    CenterOfMass = 3
    Online = 4
    FcsLoop = 5 'debug mode to automatically start an fcs measurment after an image
End Enum
'Public Const NoAnalyse As Integer = 0
'Public Const AnalyseCenterOfMassThr As Integer = 1
'Public Const AnalysePeak As Integer = 2
'Public Const AnalyseCenterOfMass As Integer = 3
'Public Const AnalyseOnline As Integer = 4
'Public Const AnalyseFcsLoop As Integer = 5
Public FocusMethods As Dictionary


Public ImgJobs() As AJob
Public FcsJobs() As AFcsJob
Public Pipelines() As APipeline
Public Repetitions() As ARepetition

'Determines if pumping should be on or off
Public Pump As Boolean
'lastTimePump occurred
Public lastTimePump As Double
'Parameters for the water pump duration, waiting time after a pump event, time between pump events, distance between positions where to pump
Public PumpTime As Double
Public PumpWait As Double
Public PumpIntervalTime As Double
Public PumpIntervalDistance As Double



'the name of the job that is currently loaded
Public currentImgJob As Long
Public currentFcsJob As Long
Public TimeOut As Boolean


'Contains name of the Jonbs
Public JobFcsNames() As String
'short name of the jobs (prefix to the file)
Public JobFcsShortNames As Collection
'the name of the Fcsjob that is currently loaded
Public CurrentJobFcs As String
'Name of file to be saved (used as reference for other functions)
Public CurrentFileName As String

Public OiaSettings As OnlineIASettings


''' Timers initiated when great is created, reinitialized if recquired
Public TimersGridCreation As Timers

Private Const TimeOutOverHead = 1
Public Function TaskFieldNames() As String()
    Dim names(8) As String
    names(0) = "analyse"
    names(1) = "jobNr"
    names(2) = "jobType"
    names(3) = "Period"
    names(4) = "SaveImage"
    names(5) = "TrackChannel"
    names(6) = "TrackXY"
    names(7) = "TrackZ"
    names(8) = "ZOffset"
    TaskFieldNames = names
End Function
Public Function TaskToArray(tsk As Task) As Variant()
    Dim arr(8) As Variant
    With tsk
        arr(0) = .Analyse
        arr(1) = .jobNr
        arr(2) = .jobType
        arr(3) = .Period
        arr(4) = .SaveImage
        arr(5) = .TrackChannel
        arr(6) = .TrackXY
        arr(7) = .TrackZ
        arr(8) = .ZOffset
    End With
    TaskToArray = arr
End Function

Public Function ArrayToTask(arr() As Variant) As Task
    Dim tsk As Task
    With tsk
        .Analyse = arr(0)
        .jobNr = arr(1)
        .jobType = arr(2)
        .Period = arr(3)
        .SaveImage = arr(4)
        .TrackChannel = arr(5)
        .TrackXY = arr(6)
        .TrackZ = arr(7)
        .ZOffset = arr(8)
    End With
    ArrayToTask = tsk
End Function

'---------------------------------------------------------------------------------------
' Procedure : AcquireJob
' Purpose   : Sets and execute an imaging Job
' Variables : JobName - The name of the Job to execute
'             RecordingDoc - the dsRecording where image is stored
'             RocordingName - The name of the recording (also for the GUI)
'             position - A vector with stage position where to acquire image X, Y, and Z (cental slice) in um
'---------------------------------------------------------------------------------------

Public Function AcquireJob(jobNr As Integer, Job As AJob, RecordingDoc As DsRecordingDoc, RecordingName As String, position As Vector) As Boolean
On Error GoTo AcquireJob_Error
    Dim SuccessRecenter As Boolean
    Dim Time As Double
    Dim cStgPos As Vector 'current stage position
    cStgPos = getCurrentPosition
    
    'stop any running jobs
    StopAcquisition
    Time = Timer
    'Create a NewRecord if required
    NewRecord RecordingDoc, "IMG:" & RecordingName, 0
    'move stage if required
    If Round(cStgPos.X, PrecXY) <> Round(position.X, PrecXY) Or Round(cStgPos.Y, PrecXY) <> Round(position.Y, PrecXY) Then
        If ZSafeDown <> 0 Then
            If Not FailSafeMoveStageZ(cStgPos.Z - ZSafeDown) Then
                Exit Function
            End If
        End If
        If Not FailSafeMoveStageXY(position.X, position.Y) Then
            Exit Function
        End If
        If ZSafeDown <> 0 Then
            If Not FailSafeMoveStageZ(cStgPos.Z) Then
                Exit Function
            End If
        End If
        'pump some water after large movement
        If Pump Then
            lastTimePump = waitForPump(PumpTime, PumpWait, lastTimePump, normVector2D(diffVector(position, cStgPos)), 0, _
            PumpIntervalDistance * 1000, 10)
        End If
    End If
        
    'Change settings for new Job if it is different from currentJob (global variable)
    If jobNr <> currentImgJob Then
        Job.PutJob ZEN
    End If
      
    currentImgJob = jobNr
    
    
    ''' recenter before acquisition
    'Time = Timer
    If Not Recenter_pre(position.Z, SuccessRecenter, ZenV) Then
        Exit Function
    End If
    Debug.Print "Time to put job and recenter pre " & Round(Timer - Time, 3)
    If DebugCode And isZStack(Lsm5.DsRecording) Then
        'SleepWithEvents (1000)
        cStgPos.Z = Lsm5.Hardware.CpFocus.position
#If ZENvC >= 2012 Then
        If (Abs(Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z)) > 0.01) Or (Abs(Lsm5.DsRecording.ReferenceZ - position.Z) > 0.01) Then
            LogManager.UpdateWarningLog " Problems in settings ZStack before imaging. Sample0Z_diff " _
            & Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z) & " um , ReferenceZ_diff " & Lsm5.DsRecording.ReferenceZ - position.Z & " um"
        End If
#Else
        If (Abs(Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z)) > 0.01) Then
            LogManager.UpdateWarningLog " Problems in settings ZStack before imaging. Sample0Z_diff " _
            & Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z) & " um"
        End If
#End If
    End If
    'Time = Timer
    Application.ThrowEvent tag_Events.eEventScanStart, 0 'notify that acquisition is started

    If Job.isAcquiring Then
        If Not ScanToImage(RecordingDoc, Job.timeToAcquire) Then
            Exit Function
        End If
    Else
        GoTo ErrorTrack
    End If
    Application.ThrowEvent tag_Events.eEventScanStop, 0 'notify that acquisition is finished

    'wait that slice recentered after acquisition
    'Time = Timer
    If Not Recenter_post(position.Z, SuccessRecenter, ZenV, False) Then
       Exit Function
    End If
    
    If isZStack(Lsm5.DsRecording) Then
        'Warning if there are any issues with the central slice
        cStgPos.Z = Lsm5.Hardware.CpFocus.position
#If ZENvC >= 2012 Then
        If (Abs(Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z)) > 0.01) Or (Abs(Lsm5.DsRecording.ReferenceZ - position.Z) > 0.01) Then
            LogManager.UpdateWarningLog " Problems returning to rest position after imaging. Sample0Z_diff " _
            & Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z) & " um , ReferenceZ_diff " & Lsm5.DsRecording.ReferenceZ - position.Z & " um"
        End If
#Else
        If (Abs(Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z)) > 0.01) Then
            LogManager.UpdateWarningLog " Problems returning to rest position after imaging. Sample0Z_diff " _
            & Lsm5.DsRecording.Sample0Z - (getHalfZRange(Lsm5.DsRecording) + position.Z - cStgPos.Z) & " um"
        End If
#End If
    End If
    AcquireJob = True
    Exit Function
ErrorTrack:
    MsgBox "No active track for Job " & jobNr & " defined. Exit now!"
    Exit Function

   On Error GoTo 0
   Exit Function
AcquireJob_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & ") in procedure AcquireJob of Module JobsManager at line " & Erl
End Function

''---------------------------------------------------------------------------------------
'' Procedure : AcquireFcsJob
'' Purpose   : Sets and execute an FCS Job at specified position
'' Variables : JobName  -  The name of the Job to execute
''             RecordingDoc - the DsRecordingDoc of the Fcs measurements
''             FcsData -  the AimFcsData containing the Fcs
''             FileName - Name appearing on top of RecordingDoc
''             positions -  A vector array with position where to acquire Fcs X, Y (relative to center of image), and Z (absolute). Unit are in meter!!
''---------------------------------------------------------------------------------------
''
Public Function AcquireFcsJob(jobNr As Integer, Job As AFcsJob, RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, fileName As String, Positions() As Vector) As Boolean
On Error GoTo AcquireFcsJob_Error

    Dim Time As Double
    Dim i As Integer
    Dim posTxt() As String
    Set FcsControl = Fcs

    'Stop Fcs acquisition
    StopAcquisition
    Time = Timer
    If Not NewFcsRecord(RecordingDoc, FcsData, "FCS:" & fileName, 0) Then
        GoTo WarningHandle
    End If
    If Not CleanFcsData(RecordingDoc, FcsData) Then
        Exit Function
    End If
    'Use position list mode
    FcsControl.SamplePositionParameters.SamplePositionMode = eFcsSamplePositionModeList

    '''clear previous positions
    ClearFcsPositionList

    '''update positions
    setFcsPositions Positions

    If jobNr <> currentFcsJob Then
        If Not Job.PutJob(ZEN, ZenV) Then
           Exit Function
        End If
    End If
    currentFcsJob = jobNr
    If Not ScanToFcs(RecordingDoc, FcsData, Job.timeToAcquire) Then
        Exit Function
    End If

    AcquireFcsJob = True
    posTxt = VectorList2String(scaleVectorList(Positions, 1000000#), 2)

    LogManager.UpdateLog " Acquire Fcsjob " & jobNr & " " & fileName & " at X = " & posTxt(0) & " Y = " & posTxt(1) & " Z = " & posTxt(2) & ". Acquisitiontime " & Round(Timer - Time, 3) & " sec" & ". Relative position to center in um"
    Exit Function

WarningHandle:
    MsgBox "AcquireFcsJob for job " & jobNr & ". Not able to create document!", VbExclamation
    Exit Function

    On Error GoTo 0
    Exit Function

AcquireFcsJob_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & ") in procedure AcquireFcsJob of Module JobsManager at line " & Erl & " " & fileName
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExecuteFcsJob
' Purpose   : Executes the AcquireFcsJob and save data and positions
' Variables : JobName  -  The name of the Job to execute
'             RecordingDoc - the DsRecordingDoc of the Fcs measurements
'             FcsData -  the AimFcsData containing the Fcs
'             FilePath - Path to store file
'             FileName - Name of file
'             positions -  A vector array with position where to acquire Fcs X, Y (relative to center of image), and Z (absolute). Unit are in meter!!
'             positionsPx - A vector array with position in px relative to upper corner  of image. Z = 0 bottom of stack. Used for logging the position
'---------------------------------------------------------------------------------------
'
'Public Function ExecuteFcsJob(JobName As String, RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, FilePath As String, fileName As String, _
'Positions() As Vector, positionsPx() As Vector) As Boolean
'    Dim OiaSettings As OnlineIASettings
'    Set OiaSettings = New OnlineIASettings
'    OiaSettings.initializeDefault
'On Error GoTo ExecuteFcsJob_Error
'
'    Dim i As Integer
'    Dim Time As Double
'
'    For i = 0 To UBound(Positions)
'        Positions(i).Z = Positions(i).Z + AutofocusForm.Controls(JobName + "ZOffset").value * 0.000001
'    Next i
'
'    If Not CleanFcsData(RecordingDoc, FcsData) Then
'        Exit Function
'    End If
'    Time = Timer
'    If Not AcquireFcsJob(JobName, RecordingDoc, FcsData, fileName, Positions) Then
'        Exit Function
'    End If
'
'    CurrentFileName = fileName
'
'    If AutofocusForm.Controls(JobName + "TimeOut") Then
'        If JobsFcs.getTimeToAcquire(JobName) <= 0 Then
'            JobsFcs.setTimeToAcquire JobName, Timer - Time + TimeOutOverHead
'        End If
'    Else
'        JobsFcs.setTimeToAcquire JobName, 0
'    End If
'
'
'
'    Sleep (500)
'    If Not SaveFcsMeasurement(FcsData, FilePath & fileName & ".fcs") Then
'         Exit Function
'    End If
'    While RecordingDoc.IsBusy
'        Sleep (50)
'    Wend
'    LogManager.UpdateLog " save Fcsjob " & JobName & " " & FilePath & fileName & ".fcs"
'    SaveFcsPositionList FilePath & fileName & ".txt", positionsPx
'
'    OiaSettings.writeKeyToRegistry "filePath", FilePath & fileName & ".fcs"
'    If ScanStop Then
'        Exit Function
'    End If
'    ExecuteFcsJob = True
'    Exit Function
'
'   On Error GoTo 0
'   Exit Function
'
'ExecuteFcsJob_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure ExecuteFcsJob of Module JobsManager at line " & Erl & " " & fileName
'End Function

'---------------------------------------------------------------------------------------
' Procedure : ExecuteJob
' Purpose   : Executes part of imaging Job and save the file (no tracking)
' Variables : JobName  -  The name of the Job to execute
'             RecordingDoc - the DsRecordingDoc of the Fcs measurements
'             FilePath - Path to store file
'             FileName - Name of file
'             StgPos -  stage position where to acquire image X, Y (absolute), and Z (absolute). Unit are in micrometer!!
'             delatZ - size of Z stack. Not currently used
'---------------------------------------------------------------------------------------
'
'Public Function ExecuteJob(JobName As String, RecordingDoc As DsRecordingDoc, FilePath As String, FileName As String, _
'stgPos As Vector, Optional deltaZ As Integer = -1) As Boolean
'On Error GoTo ExecuteJob_Error
'    Dim Time As Double
'    Dim OiaSettings As OnlineIASettings
'    Set OiaSettings = New OnlineIASettings
'    OiaSettings.initializeDefault
'
'    Time = Timer
'    If Not AcquireJob(JobName, RecordingDoc, FileName, stgPos) Then
'        Exit Function
'    End If
'
'    CurrentFileName = FileName
'    If AutofocusForm.Controls(JobName + "TimeOut") Then
'        If Jobs.getTimeToAcquire(JobName) <= 0 Then
'            Jobs.setTimeToAcquire JobName, Timer - Time + TimeOutOverHead
'        End If
'    Else
'        Jobs.setTimeToAcquire JobName, 0
'    End If
'
'
'    If AutofocusForm.Controls(JobName & "SaveImage") Then
'        If Not SaveDsRecordingDoc(RecordingDoc, FilePath & FileName & imgFileExtension, imgFileFormat) Then
'            Exit Function
'        End If
'        'we set the waiting after writing the file this may still be a problem if we do the analysis on the run
'        If AutofocusForm.Controls(JobName + "OiaActive") And AutofocusForm.Controls(JobName + "OiaSequential") Then
'            OiaSettings.writeKeyToRegistry "codeMic", "wait"
'        End If
'        OiaSettings.writeKeyToRegistry "filePath", FilePath & FileName & imgFileExtension
'        LogManager.UpdateLog " save job " & JobName & " " & FilePath & FileName & imgFileExtension
'    End If
'
'    If ScanStop Then
'        Exit Function
'    End If
'    ExecuteJob = True
'    Exit Function
'
'   On Error GoTo 0
'   Exit Function
'
'ExecuteJob_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure ExecuteJob of Module JobsManager at line " & Erl & " " & FileName
'End Function


'---------------------------------------------------------------------------------------
' Procedure : TrackOffLine
' Purpose   : Compute new positions according to center of mass
' Variables : JobName - Origin job of image
'             RecordingDoc - the Recording where image is store
'             currentPosition - current absolute stage position (in um)
' Returns   : a new stage position
'---------------------------------------------------------------------------------------
'
Public Function TrackOffLine(tsk As Task, RecordingDoc As DsRecordingDoc, currentPosition As Vector) As Vector
On Error GoTo TrackOffLine_Error
    Dim newPosition() As Vector
    ReDim newPosition(0)
    Dim TrackingChannel As String
    newPosition(0) = currentPosition
    TrackOffLine = currentPosition
    If tsk.Analyse = AnalyseImage.No Or tsk.Analyse = AnalyseImage.Online Then
        Exit Function
    End If
    newPosition(0) = MassCenter(RecordingDoc, tsk.TrackChannel, tsk.Analyse)
    If Not checkForMaximalDisplacementVecPixels(ImgJobs(tsk.jobNr), newPosition) Then
        LogManager.UpdateWarningLog "TrackOffline for ImgJob " & tsk.jobNr & " computed position differs from possible range. Use current position!"
        GoTo Abort
    End If
    'transform it in um
    newPosition = computeCoordinatesImaging(ImgJobs(tsk.jobNr), currentPosition, newPosition)
    
    If tsk.TrackZ Then
        TrackOffLine.Z = newPosition(0).Z
    End If
    
    If tsk.TrackXY Then
        TrackOffLine.X = newPosition(0).X
        TrackOffLine.Y = newPosition(0).Y
    End If
    
    If Not checkForMaximalDisplacement(ImgJobs(tsk.jobNr), TrackOffLine, currentPosition) Then
        TrackOffLine = currentPosition
    End If
    
    On Error GoTo 0
    Exit Function
Abort:

   On Error GoTo 0
   Exit Function

TrackOffLine_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure TrackOffLine of Module JobsManager at line " & Erl & " "
End Function
''---------------------------------------------------------------------------------------
'' Procedure : TrackJob
'' Purpose   : Update a position with new position according to task specified (either none, Z, XY, or XYZ)
'' Variables : JobName - Name of job (refer to access of AutofocusForm)
''             StgPos - Current stage position (absolute in um)
''             StgPosNew - New stage position
'' Returns :   A stage position
''---------------------------------------------------------------------------------------
''
'Public Function TrackJob(JobName As String, stgPos As Vector, StgPosNew As Vector) As Vector
'On Error GoTo TrackJob_Error
'
'    TrackJob = stgPos
'    If AutofocusForm.Controls(JobName & "TrackZ") Then
'        TrackJob.Z = StgPosNew.Z
'    End If
'    If AutofocusForm.Controls(JobName & "TrackXY") Then
'        TrackJob.X = StgPosNew.X
'        TrackJob.Y = StgPosNew.Y
'    End If
'    Exit Function
'
'   On Error GoTo 0
'   Exit Function
'
'TrackJob_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure TrackJob of Module JobsManager at line " & Erl & " " & JobName
'End Function

'''
'
'''
'---------------------------------------------------------------------------------------
' Procedure : ExecuteJobAndTrack
' Purpose   : Execute a imaging job and perform a tracking. Returns an updated position X, Y, and Z
' Variables : GridName - Name of position grid
'             JobName - Name of imaging Job
'             RecordingDoc - The recording Doc
'             parentPath - the main imaging path
'             StgPos - current stage position (absolute in um)
'             Success - True if function finishes
' Returns : an updated stage position (absolute in um)
'---------------------------------------------------------------------------------------
'Public Function ExecuteJobAndTrack(GridName As String, JobName As String, RecordingDoc As DsRecordingDoc, ParentPath As String, stgPos As Vector, _
'Success As Boolean) As Vector
'On Error GoTo ExecuteJobAndTrack_Error
'
'    Dim Time As Double
'    Dim ScanMode As String
'    Dim newStgPos As Vector
'    Dim FileName As String
'    Dim FilePath As String
'    Dim OiaSettings As OnlineIASettings
'    Set OiaSettings = New OnlineIASettings
'    Success = False
'    'Acquire if active and at periodicity JobNamePeriod
'    If AutofocusForm.Controls(JobName + "Active") And _
'    Not CBool(CInt(Reps.thisIndex(GridName) - 1) Mod AutofocusForm.Controls(JobName + "Period")) Then
'         DisplayProgress "Job " & JobName & ", Row " & Grids.thisRow(GridName) & ", Col " & Grids.thisColumn(GridName) & vbCrLf & _
'        "subRow " & Grids.thisSubRow(GridName) & ", subCol " & Grids.thisSubColumn(GridName) & ", Rep " & Reps.thisIndex(GridName), RGB(&HC0, &HC0, 0)
'
'        ScanMode = Jobs.getScanMode(JobName)
'        If ScanMode = "ZScan" Or ScanMode = "Line" Then
'            AutofocusForm.Controls(JobName & "TrackXY").value = False
'        End If
'        FileName = FileNameFromGrid(GridName, JobName)
'        FilePath = Grids.getThisParentPath(GridName) & FilePathSuffix(GridName, JobName) & "\"
'        'FilePath = GridSet
'        stgPos.Z = stgPos.Z + AutofocusForm.Controls(JobName + "ZOffset").value
'
'
'        Time = Timer
'        If Not AcquireJob(JobName, RecordingDoc, FileName, stgPos) Then
'            Exit Function
'        End If
'
'        CurrentFileName = FileName
'        If AutofocusForm.Controls(JobName + "TimeOut") Then
'            If Jobs.getTimeToAcquire(JobName) <= 0 Then
'                Jobs.setTimeToAcquire JobName, Timer - Time + TimeOutOverHead
'            End If
'        Else
'            Jobs.setTimeToAcquire JobName, 0
'        End If
'
'
'    If AutofocusForm.Controls(JobName & "SaveImage") Then
'        If Not SaveDsRecordingDoc(RecordingDoc, FilePath & FileName & imgFileExtension, imgFileFormat) Then
'            Exit Function
'        End If
'        'we set the waiting after writing the file this may still be a problem if we do the analysis on the run
'        If AutofocusForm.Controls(JobName + "OiaActive") And AutofocusForm.Controls(JobName + "OiaSequential") Then
'            OiaSettings.writeKeyToRegistry "codeMic", "wait"
'        End If
'        OiaSettings.writeKeyToRegistry "filePath", FilePath & FileName & imgFileExtension
'        LogManager.UpdateLog " save job " & JobName & " " & FilePath & FileName & imgFileExtension
'    End If
'
'    If ScanStop Then
'        Exit Function
'    End If
'    ExecuteJob = True
'        If Not ExecuteJob(JobName, RecordingDoc, FilePath, FileName, stgPos) Then
'            Exit Function
'        End If
'        'do any recquired computation
'        Time = Timer
'        stgPos = TrackOffLine(JobName, RecordingDoc, stgPos)
'
'        Debug.Print "Time to TrackOffLine " & Timer - Time
'        If AutofocusForm.Controls(JobName + "OiaActive") And AutofocusForm.Controls(JobName + "OiaSequential") Then
'            OiaSettings.writeKeyToRegistry "codeOia", "newImage"
'            newStgPos = ComputeJobSequential(JobName, GridName, stgPos, FilePath, FileName, RecordingDoc)
'            If Not checkForMaximalDisplacement(JobName, stgPos, newStgPos) Then
'                newStgPos = stgPos
'            End If
'
'            Debug.Print "X =" & stgPos.x & ", " & newStgPos.x & ", " & stgPos.y & ", " & newStgPos.y & ", " & stgPos.Z & ", " & newStgPos.Z
'            stgPos = TrackJob(JobName, stgPos, newStgPos)
'        End If
'
'        If Not AutofocusForm.Controls(JobName & "TrackZ").value Then
'            stgPos.Z = stgPos.Z - AutofocusForm.Controls(JobName + "ZOffset").value
'        End If
'    End If
'    ExecuteJobAndTrack = stgPos
'    Success = True
'    Exit Function
'    OiaSettings.readKeyFromRegistry ("filePath")
'   On Error GoTo 0
'   Exit Function
'
'ExecuteJobAndTrack_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure ExecuteJobAndTrack of Module JobsManager at line " & Erl & " " & GridName & " " & JobName & " " & ParentPath & " " & OiaSettings.getSettings("filePath")
'End Function

Public Function ExecuteTask(indexPl As Integer, indexTsk As Integer, RecordingDoc As DsRecordingDoc, _
    FcsRecordingDoc As DsRecordingDoc, FcsData As AimFcsData, ParentPath As String, _
    stgPos As Vector, Success As Boolean) As Vector
On Error GoTo ExecuteJobAndTrack_Error
    'Default return is input position
    ExecuteTask = stgPos
    Dim Time As Double
    Dim fcsPos() As Vector
    Dim fcsPosPx() As Vector
    Dim ScanMode As String
    Dim newStgPos As Vector
    Dim fileName As String
    Dim FilePath As String
    Dim Period As Long
    Dim Rep As Long
    Dim jobNr As Long
    jobNr = Pipelines(indexPl).getTask(indexTsk).jobNr
    Rep = Pipelines(indexPl).Repetition.index
    Period = Pipelines(indexPl).getTask(indexTsk).Period
    
    'Acquire if at periodicity
    If Period = 0 And Rep > 1 Then 'only acquire at beginning
        GoTo NoProcess
    End If
    If Period = -1 And Rep <> Pipelines(indexPl).Repetition.number Then  'only acquire at the end
        GoTo NoProcess
    End If
    If Period > 0 Then
        If CBool(CInt(Rep - 1) Mod Period) Then
            GoTo NoProcess
        End If
    End If
    With Pipelines(indexPl)
         DisplayProgress PipelineConstructor.ProgressLabel, _
         "Pipeline " & .Grid.NameGrid & " Task " & indexTsk + 1 & "/" & .count & vbCrLf & _
         "Row " & .Grid.iRow & ", Col " & .Grid.iCol & vbCrLf & _
         "subRow " & .Grid.iRowSub & ", subCol " & .Grid.iColSub & vbCrLf & _
         "Repetition " & .Repetition.index & "/" & .Repetition.number, RGB(&HC0, &HC0, 0)
    End With
    fileName = FileNameFromPipeline(indexPl, indexTsk)
    FilePath = Pipelines(indexPl).Grid.getThisParentPath & FilePathSuffixFromPipeline(indexPl) & "\"
    stgPos.Z = stgPos.Z + Pipelines(indexPl).getTask(indexTsk).ZOffset
    With Pipelines(indexPl).getTask(indexTsk)
        Select Case .jobType
            Case jobTypes.imgjob
                Time = Timer
                If Not AcquireJob(.jobNr, ImgJobs(.jobNr), RecordingDoc, fileName, stgPos) Then
                    Exit Function
                End If
                LogManager.UpdateLog "Pipeline " & Pipelines(indexPl).Grid.NameGrid & " task " & indexTsk + 1 & " ImgJob " & jobNr + 1 & " " & fileName & " at X = " & stgPos.X & ", Y =  " & stgPos.Y & ", Z =  " & stgPos.Z & " in " & Round(Timer - Time, 3) & " sec"
                If .SaveImage Then
                    If Not SaveDsRecordingDoc(RecordingDoc, FilePath & fileName & imgFileExtension, imgFileFormat) Then
                        Exit Function
                    End If
                    OiaSettings.writeKeyToRegistry "filePath", FilePath & fileName & imgFileExtension
                End If
                Select Case .Analyse
                    Case AnalyseImage.No
                    Case AnalyseImage.Online
                        OiaSettings.writeKeyToRegistry "codeMic", "wait"
                        OiaSettings.writeKeyToRegistry "codeOia", "newImage"
                        newStgPos = ComputeJobSequential(indexPl, indexTsk, stgPos, FilePath, fileName, RecordingDoc)
                        If .TrackZ Then
                           stgPos.Z = newStgPos.Z
                        End If
                        If .TrackXY Then
                            stgPos.X = newStgPos.X
                            stgPos.Y = newStgPos.Y
                        End If
                    Case AnalyseImage.FcsLoop
                        ReDim fcsPos(0 To 2)
                        'position in pixels
                        fcsPos(0) = ImgJobs(indexTsk).getCentralPointPx
                        fcsPos(1) = ImgJobs(indexTsk).getCentralPointPx
                        fcsPos(2) = ImgJobs(indexTsk).getCentralPointPx
                        Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPos
                        'position of fcs pt with respect to center of image  (XY) and absolute Z in meter
                        fcsPos(0) = Double2Vector(0, 0, Lsm5.Hardware.CpFocus.position * 0.000001)
                        fcsPos(1) = Double2Vector(0, 0, Lsm5.Hardware.CpFocus.position * 0.000001)
                        fcsPos(2) = Double2Vector(0, 0, Lsm5.Hardware.CpFocus.position * 0.000001)
                        Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                    Case Else
                        stgPos = TrackOffLine(Pipelines(indexPl).getTask(indexTsk), RecordingDoc, stgPos)
                        LogManager.UpdateLog " Time to TrackOffline " & Round(Timer - Time, 2), 1
                End Select
                If Not .TrackZ Then
                    stgPos.Z = stgPos.Z - Pipelines(indexPl).getTask(indexTsk).ZOffset
                End If
            Case jobTypes.fcsjob
                Time = Timer
                If isPosArrayEmpty(Pipelines(indexPl).Grid.getThisFcsPositions) Then
                    ReDim fcsPos(0)
                    ReDim fcsPosPx(0)
                    fcsPos(0) = Double2Vector(0, 0, Lsm5.Hardware.CpFocus.position * 0.000001)
                    LogManager.UpdateErrorLog "No fcs Positions have been defined for " & Pipelines(indexPl).Grid.NameGrid & "_" & indexTsk & " use center of image and current Z!"
                    Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                    Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPosPx
                End If
                Pipelines(indexPl).Grid.setThisFcsPositionsZOffset .ZOffset * 0.000001

                If Not AcquireFcsJob(.jobNr, FcsJobs(.jobNr), FcsRecordingDoc, FcsData, fileName, Pipelines(indexPl).Grid.getThisFcsPositions) Then
                    Exit Function
                End If
                
                If .SaveImage Then
                    If Not SaveFcsMeasurement(FcsData, FcsRecordingDoc, FilePath & fileName & ".fcs") Then
                        Exit Function
                    End If
                    SaveFcsPositionList FilePath & fileName, Pipelines(indexPl).Grid.getThisFcsPositionsPx, VBA.Right(RecordingDoc.Title, Len(RecordingDoc.Title) - 4) & imgFileExtension
                End If
            Case jobTypes.gotoPip
                updateSubPipelineGrid .jobNr, Vector2Array(stgPos), fcsPos, fcsPosPx, "", FilePath & fileName & "\"
            End Select
            
    End With

NoProcess:
    ExecuteTask = stgPos
    Success = True
    Exit Function
    OiaSettings.readKeyFromRegistry ("filePath")
   On Error GoTo 0
   Exit Function

ExecuteJobAndTrack_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ExecuteTask of Module JobsManager at line " & Erl
End Function

'---------------------------------------------------------------------------------------
' Procedure : StartJobOnGrid
' Purpose   : Performs imaging/fcs on a grid. Pretty much the whole macro runs through here
' Variables : GridName -
'             JobName -
'             parentPath - Path from where job has been initiated
'---------------------------------------------------------------------------------------
'
'Public Function StartJobOnGrid(GridName As String, JobName As String, RecordingDoc As DsRecordingDoc, ParentPath As String) As Boolean
'On Error GoTo StartJobOnGrid_Error
'
'    Dim OiaSettings As OnlineIASettings
'    Set OiaSettings = New OnlineIASettings
'    Dim i As Integer
'    Dim stgPos As Vector
'    '''The name of jobs run for the global mode
'    Dim JobNamesGlobal(2) As String
'    Dim iJobGlobal As Integer
'
'    JobNamesGlobal(0) = "Autofocus"
'    JobNamesGlobal(1) = "Acquisition"
'    JobNamesGlobal(2) = "AlterAcquisition"
'
'    Dim FileName As String
'    Dim deltaZ As Integer
'    deltaZ = -1
'    Dim SuccessExecute As Boolean
'    'Stop all running acquisitions (maybe to strong)
'    StopAcquisition
'
'    'coordinates
'    Dim previousZ As Double   'remember position of previous position in Z
'
'
'    OiaSettings.resetRegistry
'    OiaSettings.readFromRegistry
'
'    FileName = AutofocusForm.TextBoxFileName.value & Grids.getName(JobName, 1, 1, 1, 1) & Grids.suffix(JobName, 1, 1, 1, 1) & Reps.suffix(JobName, 1)
'    'create a new Gui document if recquired
'    NewRecord RecordingDoc, FileName
'
'    CurrentJob = ""
'    Running = True  'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.
'
'
'
'    previousZ = Grids.getZ(JobName, 1, 1, 1, 1)
'    Reps.resetIndex (JobName)
'
'    '''
'    ' Check if there are any valid positions
'    ''''
'    If Grids.getNrValidPts(GridName) = 0 Then
'        DisplayProgress "Job " & JobName & ", on grid " & GridName & " has no valid positions !", RGB(&HC0, &HC0, 0)
'        Sleep (500)
'        Exit Function
'    End If
'
'    Grids.setIsRunning GridName, True
'
'    While Reps.nextRep(GridName) ' cycle all repetitions
'        Grids.setIndeces GridName, 1, 1, 1, 1
'        Do ''Cycle all positions defined in grid
'            If Grids.getThisValid(GridName) Then
'               DisplayProgress "Job " & JobName & ", Row " & Grids.thisRow(GridName) & ", Col " & Grids.thisColumn(JobName) & vbCrLf & _
'                "subRow " & Grids.thisSubRow(GridName) & ", subCol " & Grids.thisSubColumn(GridName) & ", Rep " & Reps.thisIndex(GridName), RGB(&HC0, &HC0, 0)
'
'                'Do some positional Job
'                stgPos.x = Grids.getThisX(GridName)
'                stgPos.y = Grids.getThisY(GridName)
'                stgPos.Z = Grids.getThisZ(GridName)
'
'                'For first repetition and globalgrid we use previous position to prime next position (this is not the optimal way of doing it, better is a focusMap)
'                If Reps.getIndex(GridName) = 1 And AutofocusForm.GridScanActive And AutofocusForm.SingleLocationToggle And GridName = "Global" And AutofocusForm.GridScanPositionFile = "" Then
'                    stgPos.Z = previousZ
'                End If
'                'pump if time elapsed before starting imaging on a specific point
'                If Pump Then
'                    lastTimePump = waitForPump(PumpTime, PumpWait, lastTimePump, 0, PumpIntervalTime * 60, _
'                    0, 10)
'                End If
'                ' Recenter and move where it should be. Job global is a series of jobs
'                ' TODO move into one single function per task
'                If JobName = "Global" Then
'                    For iJobGlobal = 0 To UBound(JobNamesGlobal)
'                        ' run subJobs for global setting
'                        stgPos = ExecuteJobAndTrack(GridName, JobNamesGlobal(iJobGlobal), RecordingDoc, ParentPath, stgPos, SuccessExecute)
'                        If Not SuccessExecute Then
'                            GoTo StopJob
'                        End If
'                    Next iJobGlobal
'                Else
'                    If AutofocusForm.Controls(JobName + "Autofocus") Then
'                        stgPos = ExecuteJobAndTrack(GridName, "Autofocus", RecordingDoc, ParentPath, stgPos, SuccessExecute)
'                    End If
'                    stgPos = ExecuteJobAndTrack(GridName, JobName, RecordingDoc, ParentPath, stgPos, SuccessExecute)
'                    If Not SuccessExecute Then
'                        GoTo StopJob
'                    End If
'                End If
'
'                Grids.setThisX GridName, stgPos.x
'                Grids.setThisY GridName, stgPos.y
'                Grids.setThisZ GridName, stgPos.Z
'                previousZ = Grids.getThisZ(GridName)
'            End If
'            If ScanPause = True Then
'                If Not AutofocusForm.Pause Then ' Pause is true if Resume
'                    GoTo StopJob
'                End If
'            End If
'        Loop While Grids.nextGridPt(JobName, AutofocusForm.GridScan_WellsFirst)
'        ''Wait till next repetition
'        Reps.updateTimeStart (JobName)
'
'        If Reps.wait(JobName) > 0 Then
'            DisplayProgress "Waiting " & CStr(CInt(Reps.wait(JobName))) & " s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
'            DoEvents
'        End If
'
'        If AutofocusForm.StopAfterRepetition Then
'            GoTo StopJob
'        End If
'
'        While ((Reps.wait(JobName) > 0) And (Reps.getIndex(JobName) < Reps.getRepetitionNumber(JobName)))
'            Sleep (100)
'            DoEvents
'            If Pump Then
'                lastTimePump = waitForPump(PumpTime, PumpWait, lastTimePump, 0, PumpIntervalTime * 60, _
'                0, 10)
'            End If
'            If ScanPause = True Then
'                If Not AutofocusForm.Pause Then ' Pause is true if Resume
'                    GoTo StopJob
'                    Exit Function
'                End If
'            End If
'            If ScanStop Then
'                GoTo StopJob
'            End If
'            DisplayProgress "Waiting " & CStr(CInt(Reps.wait(JobName))) & " s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
'
'            '''Check for extra jobs to run
'            For i = 3 To 4
'                 If Grids.getNrValidPts(JobNames(i)) > 0 And Not Grids.getIsRunning(JobNames(i)) Then
'                    If TimersGridCreation.wait(JobNames(i), CDbl(AutofocusForm.Controls(JobNames(i) + "maxWait").value)) < 0 Then
'                        LogManager.UpdateLog " OnlineImageAnalysis  execute job " & JobNames(i) & " after maximal time exceeded "
'                        'start acquisition of Job on grid named JobName
'                        If Not StartJobOnGrid(JobNames(i), JobNames(i), RecordingDoc, ParentPath & "\") Then
'                            GoTo StopJob
'                        End If
'                        'set all run positions to notValid
'                        Grids.setAllValid JobNames(i), False
'                        Grids.setIsRunning JobNames(i), False
'                    End If
'                End If
'            Next i
'        Wend
'        Sleep (100)
'        DoEvents
'        If ScanPause = True Then
'            If Not AutofocusForm.Pause Then ' Pause is true is Resume
'                GoTo StopJob
'            End If
'        End If
'        If ScanStop Then
'            GoTo StopJob
'        End If
'    Wend
'    StartJobOnGrid = True
'    Grids.setIsRunning GridName, False
'    Exit Function
'StopJob:
'    ScanStop = True
'    StopAcquisition
'    DisplayProgress "Stopped", RGB(&HC0, 0, 0)
'    Exit Function
'
'   On Error GoTo 0
'   Exit Function
'
'StartJobOnGrid_Error:
'    ScanStop = True
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure StartJobOnGrid of Module JobsManager at line " & Erl & " " & " Grid " & GridName & " Job " & JobName
'End Function


'''''
'   Pause()
'   Function called when ScanPause = True
'   Checks state and wait for action in Form
'''''
Public Function Pause() As Boolean
    
    Dim rettime As Double
    Dim GlobalPrvTime As Double
    Dim DiffTime As Double
    
    GlobalPrvTime = CDbl(GetTickCount) * 0.001
    rettime = GlobalPrvTime
    DiffTime = rettime - GlobalPrvTime
    Do While True
        SleepWithEvents 100
        DoEvents
        If ScanStop Then
            Exit Function
        End If
        If Not ScanPause Then
            Pause = True
            Exit Function
        End If

        DisplayProgress PipelineConstructor.ProgressLabel, "Pause " & CStr(CInt(DiffTime)) & " s", RGB(&HC0, &HC0, 0)
        rettime = CDbl(GetTickCount) * 0.001
        DiffTime = rettime - GlobalPrvTime
    Loop
End Function

Public Sub resetStopFlags(Optional i As Integer)
    ScanStop = False
    ScanStopAfterRepetition = False
    PipelineConstructor.StopAfterRepButton = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : StartJobOnGrid
' Purpose   : Performs imaging/fcs on a grid. Pretty much the whole macro runs through here
' Variables : GridName -
'             JobName -
'             parentPath - Path from where job has been initiated
'---------------------------------------------------------------------------------------
'
Public Function StartPipeline(index As Integer, RecordingDoc As DsRecordingDoc, FcsRecordingDoc As DsRecordingDoc, _
FcsData As AimFcsData, ParentPath As String) As Boolean
On Error GoTo StartPipeline_Error

    Dim i As Integer
    Dim ipip As Integer
    Dim iTask As Integer
    Dim stgPos As Vector
    
    Dim fileName As String
    Dim SuccessExecute As Boolean
    'Stop all running acquisitions (maybe to strong)
    StopAcquisition
    
    'coordinates
    Dim previousZ As Double   'remember position of previous position in Z
    
       
    OiaSettings.resetRegistry
    OiaSettings.readFromRegistry
      
    fileName = PipelineConstructor.TextBoxFileName.value & Pipelines(index).Grid.getName(1, 1, 1, 1) & Pipelines(index).Grid.suffix(1, 1, 1, 1) & Pipelines(index).Repetition.suffix(1)
    'create a new Gui document if recquired
    NewRecord RecordingDoc, "IMG:" & fileName
    
    currentImgJob = -1
     
    
    previousZ = Pipelines(index).Grid.getZ(1, 1, 1, 1)
    Pipelines(index).Repetition.index = 0
    
    '''
    ' Check if there are any valid positions
    ''''
    If Pipelines(index).Grid.getNrValidPts = 0 Then
        DisplayProgress PipelineConstructor.ProgressLabel, "Pipeline " & Pipelines(index).Grid.NameGrid & " has no valid positions !", RGB(&HC0, &HC0, 0)
        Sleep (500)
        Exit Function
    End If
    With Pipelines(index)
        .Grid.isRunning = True

        While .Repetition.nextRep ' cycle all repetitions
            .Grid.setIndeces 1, 1, 1, 1
        
            Do ''Cycle all positions defined in grid
                If .Grid.getThisValid Then
                    'set current position
                    stgPos = .Grid.getThisPosition
                    'pump if time elapsed before starting imaging on a specific point
                    If Pump Then
                        lastTimePump = waitForPump(PumpTime, PumpWait, lastTimePump, 0, PumpIntervalTime * 60, _
                        0, 10)
                    End If
                    ' Recenter and move where it should be. Job global is a series of jobs
                    For iTask = 0 To .count - 1
                        stgPos = ExecuteTask(index, iTask, RecordingDoc, FcsRecordingDoc, FcsData, ParentPath, stgPos, SuccessExecute)
                        If ScanStop Then
                            GoTo StopJob
                        End If
                        For ipip = 1 To UBound(Pipelines)
                            If Not Pipelines(ipip).Grid.isRunning And runSubPipeline(ipip) Then
                                StartPipeline ipip, RecordingDoc, FcsRecordingDoc, FcsData, ParentPath
                            End If
                        Next ipip
                    Next iTask
                    .Grid.setThisPosition stgPos
                End If
                If ScanPause Then
                    If Not Pause Then
                        GoTo StopJob
                    End If
                End If
                
            Loop While .Grid.nextGridPt(False)
            
            ''Wait till next repetition
            .Repetition.updateTimeStart
        
            If .Repetition.wait > 0 Then
                DisplayProgress PipelineConstructor.ProgressLabel, "Waiting " & CStr(CInt(.Repetition.wait)) & " s before scanning repetition  " & .Repetition.index + 1, RGB(&HC0, &HC0, 0)
                DoEvents
            End If
            
        
            While ((.Repetition.wait > 0) And (.Repetition.index < .Repetition.number))
                SleepWithEvents (200)
                If Pump Then
                    lastTimePump = waitForPump(PumpTime, PumpWait, lastTimePump, 0, PumpIntervalTime * 60, _
                    0, 10)
                End If
                For ipip = 1 To UBound(Pipelines)
                    If Not Pipelines(ipip).Grid.isRunning And runSubPipeline(ipip) Then
                        If Not StartPipeline(ipip, RecordingDoc, FcsRecordingDoc, FcsData, ParentPath) Then
                            GoTo StopJob
                        End If
                    End If
                Next ipip
                If ScanStop Then
                   GoTo StopJob
                End If
                If .Repetition.wait > 0 Then
                    DisplayProgress PipelineConstructor.ProgressLabel, "Waiting " & CStr(CInt(.Repetition.wait)) & " s before scanning repetition  " & .Repetition.index + 1, RGB(&HC0, &HC0, 0)
                    DoEvents
                End If
                If .Grid.getNrValidPts = 0 And index = 0 Then
                    LogManager.UpdateErrorLog "No more default active positions. use keep-parent positions in Trigger1/2 to keep default positions after trigger"
                    GoTo StopJob
                End If
                If ScanPause Then
                    If Not Pause Then
                        GoTo StopJob
                    End If
                End If
                If .Grid.getNrValidPts = 0 Then
                        
                End If
            Wend
            If ScanStopAfterRepetition Then
                GoTo StopJob
            End If
        DoEvents
    Wend
PipelineEnd:
    StartPipeline = True
    
    .Grid.isRunning = False
    .Grid.setAllValid False
    End With
    Exit Function
StopJob:
    ScanStop = True
    StopAcquisition
    DisplayProgress PipelineConstructor.ProgressLabel, "Stopped", RGB(&HC0, 0, 0)
    Exit Function
    
   On Error GoTo 0
   Exit Function

StartPipeline_Error:
    ScanStop = True
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StartPipelines of Module JobsManager at line " & Erl & " " & " Pipeline " & Pipelines(index).Grid.NameGrid & " TaskNr " & iTask
End Function

'---------------------------------------------------------------------------------------
' Procedure : waitForPump
' Purpose   : check if activation of pump is recquired if yes write command to registry
' Variables:
' Inputs:
'        timeToPump: time to activate the pump (in ms)
'        timeToWait: time to wait after pump event (in ms)
'        lastTimePump: time of last event (in ms): CDbl(GetTickCount) * 0.001
'        distDiff: a distance (in um)
'        timeMax: maximal timeDiff over which pump is activated
'        distmax: maximal distDiff over which pump is activated
'        maxTimeWaitRegistry: maximal time we wait for registry (sec)
' Outputs:
'        updated last time pump was active
'---------------------------------------------------------------------------------------
'
Public Function waitForPump(timeToPump As Double, TimeToWait As Double, lastTimePump As Double, distDiff As Double, timeMax As Double, distMax As Double, maxTimeWaitRegistry As Double) As Double
    
    Dim doPump As Boolean
    Dim OiaSettings As OnlineIASettings
    Dim TimeStart As Double
    Dim TimeWait As Double
    Set OiaSettings = New OnlineIASettings
    ''check if we need to pump
    If (distDiff <= distMax Or distMax = 0) And (CDbl(GetTickCount) * 0.001 - lastTimePump <= timeMax Or timeMax = 0) Then
        waitForPump = lastTimePump
        Exit Function
    End If
    
    OiaSettings.readFromRegistry
    OiaSettings.writeKeyToRegistry "codeMic", "wait"
    OiaSettings.writeKeyToRegistry "codePump", CStr(timeToPump)
    DoEvents
    Sleep (200)
    TimeStart = CDbl(GetTickCount) * 0.001
    DisplayProgress PipelineConstructor.ProgressLabel, "Waiting for pump...", RGB(0, &HC0, 0)
    Do While OiaSettings.readKeyFromRegistry("codeMic") = "wait" And (TimeWait < maxTimeWaitRegistry)
            TimeWait = CDbl(GetTickCount) * 0.001 - TimeStart
            Sleep (50)
            DoEvents
            If ScanStop Then
                GoTo Abort
            End If
    Loop
    
    If TimeWait > maxTimeWaitRegistry Then
        OiaSettings.writeKeyToRegistry "codeMic", "timeExpired"
    End If
    
    ''Read all settings at once
    OiaSettings.readFromRegistry
    If Not OiaSettings.checkKeyItem("codeMic", OiaSettings.getSettings("codeMic")) Then
        GoTo Abort
    End If
    

    Select Case OiaSettings.getSettings("codeMic")
        Case "nothing", "": 'Nothing to do
            LogManager.UpdateLog " Pump from was successfull "
        Case "error":
            OiaSettings.writeKeyToRegistry "codeMic", "nothing"
            OiaSettings.getSettings ("errorMsg")
            LogManager.UpdateErrorLog "codeMic error. Pump for job failed . " _
            & " Error from pump: " & OiaSettings.getSettings("errorMsg")
            LogManager.UpdateLog " Pump from failed. " & OiaSettings.getSettings("errorMsg")
            OiaSettings.writeKeyToRegistry "errorMsg", ""
        Case "timeExpired":
            OiaSettings.writeKeyToRegistry "codeMic", "nothing"
            'LogManager.UpdateErrorLog "codeMic timeExpired. Waiting for pump signal took more then " & maxTimeWaitRegistry & " sec. Have you started the PumpController"
            LogManager.UpdateLog " Waiting for pump signal took more then " & maxTimeWaitRegistry & " sec"
            MsgBox "Waiting for pump signal took more then " & maxTimeWaitRegistry & " sec. Have you started the PumpController?", VbCritical
            
    End Select
    
    waitForPump = CDbl(GetTickCount) * 0.001
    Sleep (TimeToWait)
    Exit Function

Abort:
    ScanStop = True ' global flag to stop everything
    StopAcquisition

   On Error GoTo 0
   Exit Function

End Function


''---------------------------------------------------------------------------------------
'' Procedure : FileNameFromGrid
'' Purpose   : Derive filename from Grid and repetition
''---------------------------------------------------------------------------------------
''
'Private Function FileNameFromGrid(GridName As String, JobName As String) As String
'On Error GoTo FileNameFromGrid_Error
'    FileNameFromGrid = AutofocusForm.TextBoxFileName.value & Grids.getThisName(GridName) & JobShortNames(JobName) & FNSep & Grids.thisSuffix(GridName) & Reps.thisSuffix(GridName)
'    Exit Function
'   On Error GoTo 0
'   Exit Function
'
'FileNameFromGrid_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure FileNameFromGrid of Module JobsManager at line " & Erl & " " & GridName & " " & JobName
'End Function

'---------------------------------------------------------------------------------------
' Procedure : FilePathSuffix
' Purpose   : Derive filepath Suffix from Grid and repetition
'---------------------------------------------------------------------------------------

'Private Function FilePathSuffix(GridName As String, JobName As String) As String
'On Error GoTo FilePathSuffix_Error
'
'    FilePathSuffix = AutofocusForm.TextBoxFileName.value & Grids.getThisName(GridName) & JobShortNames(JobName)
'    If (Grids.numCol(GridName) * Grids.numRow(GridName) = 1 And Grids.numColSub(GridName) * Grids.numRowSub(GridName) = 1) Then
'        FilePathSuffix = FilePathSuffix & FNSep & Grids.thisSuffix(GridName)
'        Exit Function
'    End If
'    If (Grids.numCol(GridName) * Grids.numRow(GridName) > 1 And Not Grids.numColSub(GridName) * Grids.numRowSub(GridName) > 1) _
'    Or (Not Grids.numCol(GridName) * Grids.numRow(GridName) > 1 And Grids.numColSub(GridName) * Grids.numRowSub(GridName) > 1) Then
'        FilePathSuffix = FilePathSuffix & FNSep & Grids.thisSuffix(GridName)
'    Else
'        FilePathSuffix = FilePathSuffix & FNSep & Grids.thisSuffixWell(GridName) & "\" & FilePathSuffix & FNSep & Grids.thisSuffix(GridName)
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'FilePathSuffix_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure FilePathSuffix of Module JobsManager at line " & Erl & " "
'End Function



Private Function FileNameFromPipeline(indexPl As Integer, indexTask As Integer) As String
On Error GoTo FileNameFromPipeline_Error
    With Pipelines(indexPl)
        FileNameFromPipeline = PipelineConstructor.TextBoxFileName.value & .Grid.getThisName & .Grid.NameGrid & FNSep & _
        CInt(indexTask + 1) & FNSep & .Grid.thisSuffix & .Repetition.thisSuffix
    End With
    Exit Function
   On Error GoTo 0
   Exit Function

FileNameFromPipeline_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure FileNameFromPipline of Module JobsManager at line " & Erl & " " & Pipelines(indexPl).Grid.NameGrid & " " & indexTask + 1
End Function

Private Function FilePathSuffixFromPipeline(indexPl As Integer) As String
On Error GoTo FilePathSuffixFromPipeline_Error
    With Pipelines(indexPl)
        FilePathSuffixFromPipeline = PipelineConstructor.TextBoxFileName.value & .Grid.getThisName & .Grid.NameGrid
        If .Grid.hasOneGridPoint Or Not .Grid.hasWellsAndSubwells Then 'only one position
            FilePathSuffixFromPipeline = FilePathSuffixFromPipeline & FNSep & .Grid.thisSuffix
        Else
            FilePathSuffixFromPipeline = FilePathSuffixFromPipeline & FNSep & .Grid.thisSuffixWell & "\" & FilePathSuffixFromPipeline & FNSep & .Grid.thisSuffix
        End If
    End With
   On Error GoTo 0
   Exit Function
FilePathSuffixFromPipeline_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure FilePathSuffixFromPipeline of Module JobsManager at line " & Erl & " "
End Function





'---------------------------------------------------------------------------------------
' Procedure : checkForMaximalDisplacement
' Purpose   : check  that newPos is not further away than the size of the image. In fact it should be half the image
' Variables : JobName -
'             currentPos - stage position in um
'             newPos - new stage position in um
'---------------------------------------------------------------------------------------
'
Public Function checkForMaximalDisplacement(IJob As AJob, currentPos As Vector, newPos As Vector) As Boolean
On Error GoTo checkForMaximalDisplacement_Error

    Dim MaxMovementXY As Double
    Dim MaxMovementZ As Double
    MaxMovementXY = MAX(IJob.Recording.SamplesPerLine, IJob.Recording.LinesPerFrame) * IJob.Recording.SampleSpacing
    MaxMovementZ = IJob.Recording.framesPerStack * IJob.Recording.frameSpacing
    
                                
    If Abs(newPos.X - currentPos.X) > MaxMovementXY Or Abs(newPos.Y - currentPos.Y) > MaxMovementXY Or Abs(newPos.Z - currentPos.Z) > MaxMovementZ Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "dX, dY, dZ = " & Abs(newPos.X - currentPos.X) & ", " & Abs(newPos.Y - currentPos.Y) & ", " & Abs(newPos.Z - currentPos.Z) & vbCrLf & _
        "accepted dX, dY, dZ = " & MaxMovementXY & ", " & MaxMovementXY & ", " & MaxMovementZ
        Exit Function
    End If
    checkForMaximalDisplacement = True

   On Error GoTo 0
   Exit Function

checkForMaximalDisplacement_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure checkForMaximalDisplacement of Module JobsManager at line " & Erl & " "
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkForMaximalDisplacementVec
' Purpose   : check  that newPos vectors are not further away than the size of the image
' Variables : JobName -
'             currentPos - stage position in um
'             newPos - vector of stage positions in um
'---------------------------------------------------------------------------------------
'
Private Function checkForMaximalDisplacementVec(IJob As AJob, currentPos As Vector, newPos() As Vector) As Boolean
On Error GoTo checkForMaximalDisplacementVec_Error
    Dim i As Integer

    For i = 0 To UBound(newPos)
        If Not checkForMaximalDisplacement(IJob, currentPos, newPos(i)) Then
            Exit Function
        End If
    Next i
    checkForMaximalDisplacementVec = True

   On Error GoTo 0
   Exit Function

checkForMaximalDisplacementVec_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure checkForMaximalDisplacementVec of Module JobsManager at line " & Erl & " "
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkForMaximalDisplacementPixels
' Purpose   : check  that newPos is within possible boundary using pixels
' Variables : JobName -
'             newPos - A new position in pixels 0,0,0 is upper left bottom slice
'---------------------------------------------------------------------------------------
'
Private Function checkForMaximalDisplacementPixels(IJob As AJob, newPos As Vector) As Boolean
On Error GoTo checkForMaximalDisplacementPixels_Error
    Dim MaxX As Long
    Dim MaxY As Long
    Dim MaxZ As Long
    Dim i As Integer

    MaxX = IJob.Recording.SamplesPerLine - 1
    If IJob.Recording.ScanMode = "ZScan" Then
        MaxY = 0
    Else
        MaxY = IJob.Recording.LinesPerFrame - 1
    End If
    If IJob.isZStack Then
        MaxZ = IJob.Recording.framesPerStack - 1
    Else
        MaxZ = 0
    End If

    If newPos.X + TolPx < 0 Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned negative pixel values " & _
        "X = " & newPos.X & ". VBA macro will set this to 0"
        newPos.X = 0
    End If
    
    If newPos.Y + TolPx < 0 Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned negative pixel values " & _
        "Y = " & newPos.Y & ". VBA macro will set this to 0"
        newPos.Y = 0
    End If
    
    If newPos.Z + TolPx < 0 Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned negative pixel values " & _
        "Z = " & newPos.Z & ". VBA macro will set this to 0"
        newPos.Z = 0
    End If
    
    If newPos.X - MaxX > TolPx Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "X = " & newPos.X & " accepted range is X = " & 0 & "-" & MaxX & ". VBA macro sets value to center of image" & MaxX / 2
        newPos.X = MaxX / 2
    End If
    
    If newPos.Y - MaxY > TolPx Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "Y = " & newPos.Y & " accepted range is Y = " & 0 & "-" & MaxY & ". VBA macro sets value to center of image" & MaxY / 2
        newPos.Y = MaxY / 2
    End If
    
    If newPos.Z - MaxZ > TolPx Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "Z = " & newPos.Z & " accepted range is Z = " & 0 & "-" & MaxZ & ". VBA macro sets value to center of image" & MaxZ / 2
        newPos.Z = MaxZ / 2
    End If
    checkForMaximalDisplacementPixels = True

   On Error GoTo 0
   Exit Function

checkForMaximalDisplacementPixels_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure checkForMaximalDisplacementPixels of Module JobsManager at line " & Erl & " "
End Function

'---------------------------------------------------------------------------------------
' Procedure : checkForMaximalDisplacementVecPixels
' Purpose   : check  that newPos is within possible boundary using pixels
' Variables : JobName -
'             newPos - A vector of new positions in pixels 0,0,0 is upper left bottom slice
'---------------------------------------------------------------------------------------
'
Private Function checkForMaximalDisplacementVecPixels(IJob As AJob, newPos() As Vector) As Boolean
On Error GoTo checkForMaximalDisplacementVecPixels_Error
    Dim i As Integer
    For i = 0 To UBound(newPos)
        If Not checkForMaximalDisplacementPixels(IJob, newPos(i)) Then
            Exit Function
        End If
    Next i
    checkForMaximalDisplacementVecPixels = True

   On Error GoTo 0
   Exit Function

checkForMaximalDisplacementVecPixels_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure checkForMaximalDisplacementVecPixels of Module JobsManager at line " & Erl & " "
End Function


'---------------------------------------------------------------------------------------
' Procedure : computeShiftedCoordinates
' Purpose   : given offsetPosition with (0,0,0) center of image central slice (in um)
'             Computes absolute stage/focus coordinates from currentPosition.
'             Considers mirror possible mirror of axis
' Variables : offsetPosition - position in um relative to 0,0,0 center of image and central slice
'             currentPosiotion -
' Returns   : new shifted position
'---------------------------------------------------------------------------------------
'
Public Function computeShiftedCoordinates(offsetPosition As Vector, currentPosition As Vector) As Vector
On Error GoTo computeShiftedCoordinates_Error

    Dim Xpre As Integer
    Dim Ypre As Integer

    If MirrorX Then
        Xpre = -1
    Else
        Xpre = 1
    End If
    
    If MirrorY Then
        Ypre = -1
    Else
        Ypre = 1
    End If
    
    If ExchangeXY Then ' not sure about this and needs to be properly tested
        computeShiftedCoordinates.X = currentPosition.X + Xpre * offsetPosition.Y
        computeShiftedCoordinates.Y = currentPosition.Y + Ypre * offsetPosition.X
    Else
        computeShiftedCoordinates.X = currentPosition.X + Xpre * offsetPosition.X
        computeShiftedCoordinates.Y = currentPosition.Y + Ypre * offsetPosition.Y
    End If
      
    computeShiftedCoordinates.Z = currentPosition.Z + offsetPosition.Z

    computeShiftedCoordinates.X = Round(computeShiftedCoordinates.X, PrecXY)
    computeShiftedCoordinates.Y = Round(computeShiftedCoordinates.Y, PrecXY)
    computeShiftedCoordinates.Z = Round(computeShiftedCoordinates.Z, PrecZ)

   On Error GoTo 0
   Exit Function

computeShiftedCoordinates_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure computeShiftedCoordinates of Module JobsManager at line " & Erl & " "
End Function


'---------------------------------------------------------------------------------------
' Procedure : computeCoordinatesImaging
' Purpose   : compute new stage coordinates for imaging from pixel coordinates
' Variables : JobName -
'             currentPosition - stage position in um
'             newPosition - Vector of positions in pixel (0,0,0) is upper left bottom slice
' Returns   : stage positions in um!
'---------------------------------------------------------------------------------------
'
Public Function computeCoordinatesImaging(IJob As AJob, currentPosition As Vector, newPosition() As Vector) As Vector()
On Error GoTo computeCoordinatesImaging_Error
    
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim MaxX As Integer
    Dim MaxY As Integer
    Dim MaxZ  As Integer
    Dim i As Integer
    Dim position() As Vector

    position = newPosition
    'pixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing 'This is in meter!!! be careful . Position for imaging is provided in um
    pixelSize = IJob.Recording.SampleSpacing ' this is in um
    'compute difference with respect to center
    MaxX = IJob.Recording.SamplesPerLine - 1
    If IJob.Recording.ScanMode = "ZScan" Then
        MaxY = 0
    Else
        MaxY = IJob.Recording.LinesPerFrame - 1
    End If
    If IJob.isZStack Then
        MaxZ = IJob.Recording.framesPerStack - 1
    Else
        MaxZ = 0
    End If
    frameSpacing = IJob.Recording.frameSpacing
    
    For i = 0 To UBound(newPosition)
        position(i).X = (position(i).X - MaxX / 2) * pixelSize
        position(i).Y = (position(i).Y - MaxY / 2) * pixelSize
        If IJob.isZStack Then
            position(i).Z = (position(i).Z - MaxZ / 2) * frameSpacing
        Else
            position(i).Z = 0
        End If
        'this accounts for any shifts in XY and mirroring of hardware!
        position(i) = computeShiftedCoordinates(position(i), currentPosition)
    Next i
    computeCoordinatesImaging = position

   On Error GoTo 0
   Exit Function

computeCoordinatesImaging_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure computeCoordinatesImaging of Module JobsManager at line " & Erl & " "
End Function

'---------------------------------------------------------------------------------------
' Procedure : computeCoordinatesFcs
' Purpose   : Compute coordinates for fcs from pixel coordinates
' Variables : JobName -
'             currentPosition - stage/focus position in um
'             newPosition - Vector of positions in pixel (0,0,0) is upper left bottom slice
' Returns   : stage positions in meter!!! (different from computeCoordinatesImaging which returns in um)
'---------------------------------------------------------------------------------------
'
Public Function computeCoordinatesFcs(IJob As AJob, currentPosition As Vector, newPosition() As Vector) As Vector()
On Error GoTo computeCoordinatesFcs_Error
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim MaxX As Integer
    Dim MaxY As Integer
    Dim MaxZ As Integer
    Dim framesPerStack As Integer
    Dim i As Integer
    Dim position() As Vector
    position = newPosition
    'pixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing 'This is in meter!!! be careful . Position for imaging is provided in um
    pixelSize = IJob.Recording.SampleSpacing ' this is in um
    'compute difference with respect to center
    MaxX = IJob.Recording.SamplesPerLine - 1
    If IJob.Recording.ScanMode = "ZScan" Then
        MaxY = 0
    Else
        MaxY = IJob.Recording.LinesPerFrame - 1
    End If
    If IJob.isZStack Then
        MaxZ = IJob.Recording.framesPerStack - 1
    Else
        MaxZ = 0
    End If
    frameSpacing = IJob.Recording.frameSpacing
    For i = 0 To UBound(newPosition)
        'for FCS position is with respect center of image in meter
        position(i).X = (position(i).X - MaxX / 2) * pixelSize * 0.000001
        position(i).Y = (position(i).Y - MaxY / 2) * pixelSize * 0.000001
        If IJob.isZStack Then
            position(i).Z = (position(i).Z - MaxZ / 2) * frameSpacing
        Else
            position(i).Z = 0
        End If
        'absolute position in meter
        position(i).Z = (currentPosition.Z + position(i).Z) * 0.000001
    Next i
    computeCoordinatesFcs = position

   On Error GoTo 0
   Exit Function

computeCoordinatesFcs_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure computeCoordinatesFcs of Module JobsManager at line " & Erl & " "
End Function


'---------------------------------------------------------------------------------------
' Procedure : runSubImagingJob
' Purpose   : create and update a subgrid and eventually decide whether to run Job
' Variables : GridName - Name of grid where to execute job
'             JobName -
'             newPositions - Array of stage/focus positions (in um)
'---------------------------------------------------------------------------------------
'
Public Function updateSubPipelineGrid(index As Integer, newPositions() As Vector, fcsPos() As Vector, fcsPosPx() As Vector, prefix As String, Optional ParentPath As String) As Boolean
On Error GoTo updateSubPipelineGrid_Error
    
    Dim i As Integer
    Dim GridLowBound As Integer
    
    With Pipelines(index)
        If .Grid.isGridEmpty Then
            .Grid.initialize 1, 1, 1, UBound(newPositions) + 1
            GridLowBound = 1
        Else
            GridLowBound = .Grid.numColSub + 1
            .Grid.updateGridSizePreserve 1, 1, 1, UBound(newPositions) + GridLowBound
        End If
        If .Grid.getNrValidPts = 0 Then
            If TimersGridCreation Is Nothing Then
                Set TimersGridCreation = New Timers
            End If
            TimersGridCreation.addTimer .Grid.NameGrid
            TimersGridCreation.updateTimeStart .Grid.NameGrid
        End If
        ''' input grid positions
        For i = 0 To UBound(newPositions)
            .Grid.setPt newPositions(i), True, 1, 1, 1, i + GridLowBound
            .Grid.setParentPath ParentPath, 1, 1, 1, i + GridLowBound
            .Grid.setFcsPositions fcsPos, 1, 1, 1, i + GridLowBound
            .Grid.setFcsPositionsPx fcsPosPx, 1, 1, 1, i + GridLowBound
            .Grid.setName prefix & .Grid.getName(1, 1, 1, i + GridLowBound), 1, 1, 1, i + GridLowBound
        Next i
    End With
    On Error GoTo 0
   Exit Function
updateSubPipelineGrid_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure updateSubPipelineGrid of Module JobsManager at line " & Erl & " "
End Function

Public Function runSubPipeline(index As Integer) As Boolean

On Error GoTo runSubPipeline_Error
    With Pipelines(index)
        If .Grid.getNrValidPts > 0 Then
            If Not TimersGridCreation Is Nothing Then
                If TimersGridCreation.checkTimerName(.Grid.NameGrid) Then
                    If (.optPtNumber <= .Grid.getNrValidPts Or TimersGridCreation.wait(.Grid.NameGrid, .maxWait) < 0) Then
                        runSubPipeline = True
                        Exit Function
                End If
            Else
                Exit Function
            End If
            Else
                If .optPtNumber <= .Grid.getNrValidPts And .Grid.getNrValidPts > 0 Then
                    runSubPipeline = True
                    Exit Function
                End If
            End If
        End If
    End With

   On Error GoTo 0
   Exit Function

runSubPipeline_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure runSubPipeline of Module JobsManager at line " & Erl & " "
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ComputeJobSequential
' Purpose   : Wait for image analysis and perform a specific task.
' Variables : parent variables define Job and grid from which one comes
'---------------------------------------------------------------------------------------
'
Public Function ComputeJobSequential(indexPl As Integer, indexTsk As Integer, parentPosition As Vector, ParentPath As String, parentFile As String, _
RecordingDoc As DsRecordingDoc) As Vector
On Error GoTo ComputeJobSequential_Error
    
    Dim i As Integer
    Dim newPositionsPx() As Vector 'from the registru one obtains positions in pixels
    Dim newPositions() As Vector
    Dim newPositionsAbs() As Vector
    Dim VectorString() As String
    Dim Rois() As Roi
    'Position for fcs as read fro registry these correspond to position in image
    Dim fcsPosPx() As Vector
    'position for fcs as used by the microscope. These are in meters! deviation with respect to current position and z in meters abosolute
    Dim fcsPos() As Vector
    
    Dim codeMic() As String
    Dim code As Variant
    Dim tsk As Task
    Dim JobName As String
    Dim prefix As String
    tsk = Pipelines(indexPl).getTask(indexTsk)
    JobName = Pipelines(indexPl).Grid.NameGrid & "_" & indexTsk + 1
    Dim codeMicToJobName As Dictionary 'use to convert codes of regisrty into Jobnames as used in the code
    Set codeMicToJobName = New Dictionary
    codeMicToJobName.Add "trigger1", 1
    codeMicToJobName.Add "trigger2", 2
    
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    
    codeMic = Split(Replace(OiaSettings.readKeyFromRegistry("codeMic"), " ", ""), ";")
    Dim TimeWait, TimeStart, maxTimeWait As Double
    
    maxTimeWait = 100
    
    'default return value is currentPosition
    ComputeJobSequential = parentPosition
    'helping variables giving the parentPosition in px
        
    Select Case codeMic(0)
        Case "wait":
            'Wait for image analysis to finish
            DisplayProgress PipelineConstructor.ProgressLabel, "Waiting for image analysis...", RGB(0, &HC0, 0)
            TimeStart = CDbl(GetTickCount) * 0.001
            Do While ((TimeWait < maxTimeWait) And (codeMic(0) = "wait"))
                Sleep (50)
                TimeWait = CDbl(GetTickCount) * 0.001 - TimeStart
                codeMic = Split(Replace(OiaSettings.readKeyFromRegistry("codeMic"), " ", ""), ";")
                DoEvents
                If ScanStop Then
                    GoTo Abort
                End If
            Loop

            If TimeWait > maxTimeWait Then
                SaveSetting "OnlineImageAnalysis", "macro", "codeMic", "timeExpired"
                SaveSetting "OnlineImageAnalysis", "macro", "codeOia", "nothing"
            End If
    End Select

    ''Read all settings at once
    OiaSettings.readFromRegistry
    
    ComputeJobSequential = parentPosition
    'read if it is the correct code
    If Not OiaSettings.checkKeyItem("codeMic", OiaSettings.getSettings("codeMic")) Then
        GoTo Abort
    End If
    
    codeMic = Split(Replace(OiaSettings.getSettings("codeMic"), " ", ""), ";")


    'Read positions and rois from registry the fcs positions are read with respect to
    If OiaSettings.getFcsPositions(fcsPosPx, ImgJobs(tsk.jobNr).getCentralPointPx) Then
        VectorString = VectorList2String(fcsPosPx)
        LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " obtained " & UBound(fcsPosPx) + 1 & " position(s) " & _
           " X = " & VectorString(0) & " Y = " & VectorString(1) & " Z = " & VectorString(2)
        If Not checkForMaximalDisplacementVecPixels(ImgJobs(tsk.jobNr), fcsPosPx) Then
            VectorString = VectorList2String(fcsPosPx)
            LogManager.UpdateLog "OnlineImageAnalysis position(s) exceeded boundaries and has been set to   X = " & VectorString(0) & " Y = " & VectorString(1) & " Z = " & VectorString(2)
        End If
        fcsPos = computeCoordinatesFcs(ImgJobs(tsk.jobNr), parentPosition, fcsPosPx)
        OiaSettings.writeKeyToRegistry "fcsX", ""
        OiaSettings.writeKeyToRegistry "fcsY", ""
        OiaSettings.writeKeyToRegistry "fcsZ", ""
    End If
    
    If OiaSettings.getPositions(newPositionsPx, ImgJobs(tsk.jobNr).getCentralPointPx) Then
        VectorString = VectorList2String(newPositionsPx)
        LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " obtained " & UBound(newPositionsPx) + 1 & " position(s) (in px)" & _
        " X = " & VectorString(0) & " Y = " & VectorString(1) & " Z = " & VectorString(2)
        If Not checkForMaximalDisplacementVecPixels(ImgJobs(tsk.jobNr), newPositionsPx) Then
            LogManager.UpdateErrorLog "OnlineImageAnalysis position exceed boundaries and has been set to  X = " & newPositionsPx(0).X & " Y = " & newPositionsPx(0).Y & " Z = " & newPositionsPx(0).Z
        End If
        newPositions = computeCoordinatesImaging(ImgJobs(tsk.jobNr), parentPosition, newPositionsPx)
        If Not checkForMaximalDisplacementVec(ImgJobs(tsk.jobNr), parentPosition, newPositions) Then
            GoTo ExitThis
        End If
    End If
    OiaSettings.getRois Rois
    prefix = OiaSettings.readKeyFromRegistry("prefix")

    
    ''for all commands in codeMic
    For Each code In codeMic
        LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " found " & code
        OiaSettings.writeKeyToRegistry "codeMic", "nothing"
        Select Case code
            Case "nothing", "": 'Nothing to do
                
                
            Case "error":
                OiaSettings.readKeyFromRegistry "errorMsg"
                OiaSettings.getSettings ("errorMsg")
                LogManager.UpdateErrorLog "codeMic error. Online image analysis for task " & JobName & " file " & ParentPath & parentFile & " failed . " _
                & " Error from Oia: " & OiaSettings.getSettings("errorMsg")
                LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " obtained an error. " & OiaSettings.getSettings("errorMsg")
            
            Case "timeExpired":
                LogManager.UpdateErrorLog "codeMic timeExpired. Online image analysis for job " & JobName & " file " & ParentPath & parentFile & " took more then " & maxTimeWait & " sec"
                LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " took more then " & maxTimeWait & " sec"
            
            Case "focus":
                If isPosArrayEmpty(newPositions) Then
                    LogManager.UpdateErrorLog "ComputeJobSequential: No position/wrong position for Job focus. " & ParentPath & parentFile & vbCrLf & _
                    "Specify one position in X, Y, Z of registry (in pixels, (X,Y) = (0,0) upper left corner image, Z = 0 -> central slice of current stack)!"
                    GoTo nextCode
                End If
                If UBound(newPositions) > 0 Then
                    LogManager.UpdateErrorLog " ComputeJobSequential: for Job focus " & ParentPath & parentFile & " passed only one point to X, Y, and Z of regisrty instead of " & UBound(newPositions) + 1 & ". Using the first point!"
                End If
'                Pipelines(indexPl).Grid.setThisFcsPosition fcsPos
'                If Not isArrayEmpty(Rois) Then
'                    ImgJobs(Pipelines(indexPl).getTask(indexTsk).jobNr).UseRoi = True 'this is not exactly how it should be a task should have associated a roi
'                    ImgJobs(Pipelines(indexPl).getTask(indexTsk).jobNr).setRois Rois  'this is not exactly how it should be a task should have associated a roi
'                End If
                ComputeJobSequential = newPositions(0)
                LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " focus at  " & " X = " & newPositions(0).X & " Y = " & newPositions(0).Y & " Z = " & newPositions(0).Z & ". Absolute position in um"
            
            Case "setFcsPos":
                If isPosArrayEmpty(fcsPos) Then
                    LogManager.UpdateErrorLog "ComputeJobSequential: No position/wrong position for settings FCS position of current point."
                    Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                    GoTo nextCode
                End If
                Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPosPx
            Case "trigger1", "trigger2":
                If Pipelines(codeMicToJobName.item(code)).count = 0 Then
                    LogManager.UpdateErrorLog " ComputeJobSequential:  Pipeline " & Pipelines(codeMicToJobName.item(code)).Grid.NameGrid & " has no task to do. Original file " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath")
                    GoTo nextCode
                End If
                
                Pipelines(indexPl).Grid.setThisValid Pipelines(codeMicToJobName.item(code)).keepParent
                If isPosArrayEmpty(newPositions) Then
                    LogManager.UpdateErrorLog " ComputeJobSequential: No position for pipeline " & Pipelines(codeMicToJobName.item(code)).Grid.NameGrid & " from file " & ParentPath & parentFile & " (key = " & code & ") has been specified! Imaging current position. "
                    ReDim newPositions(0)
                    newPositions(0) = parentPosition
                End If
'                If Not isArrayEmpty(Rois) And Pipelines(codeMicToJobName.item(code)).getTask(0).jobType = 0 Then
'                    ImgJobs(Pipelines(codeMicToJobName.item(code)).getTask(0).jobNr).UseRoi = True  'this is not exactly how it should be a task should have associated a roi
'                    ImgJobs(Pipelines(codeMicToJobName.item(code)).getTask(0).jobNr).setRois = Rois 'this is not exactly how it should be a task should have associated a roi
'                End If
                
                updateSubPipelineGrid codeMicToJobName.item(code), newPositions, fcsPos, fcsPosPx, prefix, ParentPath & parentFile & "\"
            Case Else
                MsgBox ("Invalid OnlineImageAnalysis codeMic = " & code)
                GoTo Abort
        End Select
nextCode:
    Next code
ExitThis:
Exit Function
Abort:
    ScanStop = True ' global flag to stop everything
    StopAcquisition
    Exit Function
   On Error GoTo 0
   Exit Function

ComputeJobSequential_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ComputeJobSequential of Module JobsManager at line " & Erl & " " & ParentPath & " " & parentFile
End Function

