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
    'tskRoi As roi
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
    Dim Names(8) As String
    Names(0) = "analyse"
    Names(1) = "jobNr"
    Names(2) = "jobType"
    Names(3) = "Period"
    Names(4) = "SaveImage"
    Names(5) = "TrackChannel"
    Names(6) = "TrackXY"
    Names(7) = "TrackZ"
    Names(8) = "ZOffset"
    TaskFieldNames = Names
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
    Dim PosUnit As New PositionUnit
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
        
    'Change settings for new Job if it is different from currentImgJob (global variable)
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
        cStgPos.Z = PosUnit.GetPositionZ
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
    Debug.Print (Lsm5.DsRecording.MultiPositionZ(0))
    Debug.Print (position.Z * 0.000001)
    
    If Lsm5.DsRecording.MultiPositionAcquisition Then
            Lsm5.DsRecording.MultiPositionZ(0) = position.Z * 0.000001
    End If
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
        cStgPos.Z = PosUnit.GetPositionZ
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
Public Function AcquireFcsJob(jobNr As Integer, Job As AFcsJob, RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, FileName As String, Positions() As Vector) As Boolean
On Error GoTo AcquireFcsJob_Error

    Dim Time As Double
    Dim i As Integer
    Dim posTxt() As String
    Set FcsControl = Fcs

    'Stop Fcs acquisition
    StopAcquisition
    Time = Timer
    If Not NewFcsRecord(RecordingDoc, FcsData, "FCS:" & FileName, 0) Then
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

    LogManager.UpdateLog " Acquire Fcsjob " & jobNr & " " & FileName & " at X = " & posTxt(0) & " Y = " & posTxt(1) & " Z = " & posTxt(2) & ". Acquisitiontime " & Round(Timer - Time, 3) & " sec" & ". Relative position to center in um"
    Exit Function

WarningHandle:
    MsgBox "AcquireFcsJob for job " & jobNr & ". Not able to create document!", VbExclamation
    Exit Function

    On Error GoTo 0
    Exit Function

AcquireFcsJob_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & ") in procedure AcquireFcsJob of Module JobsManager at line " & Erl & " " & FileName
End Function




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
    Dim FileName As String
    Dim FilePath As String
    Dim Period As Long
    Dim Rep As Long
    Dim jobNr As Long
    'Dim locimgFileFormat As enumAimExportFormat 'work around to save czi fro Airy
    'Dim locimgFileExtension As String
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
    FileName = FileNameFromPipeline(indexPl, indexTsk)
    FilePath = Pipelines(indexPl).Grid.getThisParentPath & FilePathSuffixFromPipeline(indexPl) & "\"
    stgPos.Z = stgPos.Z + Pipelines(indexPl).getTask(indexTsk).ZOffset
    With Pipelines(indexPl).getTask(indexTsk)
        Select Case .jobType
            Case jobTypes.imgjob
                Time = Timer
                If Not AcquireJob(.jobNr, ImgJobs(.jobNr), RecordingDoc, FileName, stgPos) Then
                    Exit Function
                End If
                LogManager.UpdateLog "Pipeline " & Pipelines(indexPl).Grid.NameGrid & " task " & indexTsk + 1 & " ImgJob " & jobNr + 1 & " " & FileName & " at X = " & stgPos.X & ", Y =  " & stgPos.Y & ", Z =  " & stgPos.Z & " in " & Round(Timer - Time, 3) & " sec"
                If .SaveImage Then
                    ''' Work around to use Airy (requires. czi) and Micronaut that can only read lsm
                    'If indexPl = 1 Then
                    '    locimgFileFormat = eAimExportFormatCzi
                    '    locimgFileExtension = ".czi"
                    'Else
                    '    locimgFileFormat = imgFileFormat
                    '    locimgFileExtension = imgFileExtension
                    'End If
                    'If Not SaveDsRecordingDoc(RecordingDoc, FilePath & fileName & locimgFileExtension, locimgFileFormat) Then
                    '    Exit Function
                    'End If
                    
                    If Not SaveDsRecordingDoc(RecordingDoc, FilePath & FileName & imgFileExtension, imgFileFormat) Then
                        Exit Function
                    End If
                    OiaSettings.writeKeyToRegistry "filePath", FilePath & FileName & imgFileExtension
                End If
                Select Case .Analyse
                    Case AnalyseImage.No
                    Case AnalyseImage.Online
                        OiaSettings.writeKeyToRegistry "codeMic", "wait"
                        OiaSettings.writeKeyToRegistry "codeOia", "newImage"
                        newStgPos = ComputeJobSequential(indexPl, indexTsk, stgPos, FilePath, FileName, RecordingDoc)
                        If .TrackZ Then
                           stgPos.Z = newStgPos.Z
                        End If
                        If .TrackXY Then
                            stgPos.X = newStgPos.X
                            stgPos.Y = newStgPos.Y
                        End If
                    Case AnalyseImage.FcsLoop
                        ReDim fcsPos(0 To 2)
                        ReDim fcsPosPx(0 To 2)
                        'position in pixels
                        fcsPosPx(0) = ImgJobs(indexTsk).getCentralPointPx
                        fcsPosPx(1) = ImgJobs(indexTsk).getCentralPointPx
                        fcsPosPx(2) = ImgJobs(indexTsk).getCentralPointPx
                        fcsPosPx(1).X = fcsPosPx(1).X + 10
                        fcsPosPx(1).Y = fcsPosPx(1).Y + 10
                        fcsPosPx(1).Z = fcsPosPx(1).Z + 1
                        fcsPosPx(2).X = fcsPosPx(1).X - 10
                        fcsPosPx(2).Y = fcsPosPx(1).Y - 10
                        fcsPosPx(2).Z = fcsPosPx(1).Z - 1
                        Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPosPx
                        Pipelines(indexPl).Grid.setThisFcsImage FilePath & FileName & imgFileExtension
                        Pipelines(indexPl).Grid.setThisFcsName "fcsLoop; testPt; testPt; testPt"
                        fcsPos = computeCoordinatesFcs(ImgJobs(Pipelines(indexPl).getTask(indexTsk).jobNr), stgPos, fcsPosPx)
                        Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                    Case Else
                        stgPos = TrackOffLine(Pipelines(indexPl).getTask(indexTsk), RecordingDoc, stgPos)
                        LogManager.UpdateLog " Time to TrackOffline " & Round(Timer - Time, 2), 1
                End Select
                If Not .TrackZ Then
                    stgPos.Z = stgPos.Z - Pipelines(indexPl).getTask(indexTsk).ZOffset
                End If
            Case jobTypes.fcsjob
                Dim prefix() As String
                prefix = Split(VBA.Replace(Pipelines(indexPl).Grid.getThisFcsName, " ", ""), ";")
                If UBound(prefix) < 0 Then
                    ReDim prefix(0)
                    prefix(0) = ""
                End If
                Time = Timer
                If isPosArrayEmpty(Pipelines(indexPl).Grid.getThisFcsPositions) Then
                    LogManager.UpdateWarningLog "No fcs Positions has been defined for " & Pipelines(indexPl).Grid.NameGrid & "_" & indexTsk _
                    & " fcs measurment is not performed! Analyse the previous image in the pipeline and pass the position via the registry!"
                    GoTo NoProcess
                End If
                
                Pipelines(indexPl).Grid.setThisFcsPositionsZOffset .ZOffset * 0.000001

                If Not AcquireFcsJob(.jobNr, FcsJobs(.jobNr), FcsRecordingDoc, FcsData, appendSep(prefix(0), FNSep) & _
                FileName, Pipelines(indexPl).Grid.getThisFcsPositions) Then
                    Exit Function
                End If
                
                If .SaveImage Then
                    If Not SaveFcsMeasurement(FcsData, FcsRecordingDoc, FilePath & appendSep(prefix(0), FNSep) & FileName & ".fcs") Then
                        Exit Function
                    End If
                    SaveFcsPositionList FilePath & appendSep(prefix(0), FNSep) & FileName, Pipelines(indexPl).Grid.getThisFcsPositionsPx, _
                                        Pipelines(indexPl).Grid.getThisFcsImage, prefix
                End If
            Case jobTypes.gotoPip
                updateSubPipelineGrid .jobNr, Vector2Array(stgPos), fcsPos, fcsPosPx, "", FilePath & FileName & "\"
            End Select
            
    End With

NoProcess:
    ExecuteTask = stgPos
    Success = True
    Exit Function
   On Error GoTo 0
   Exit Function

ExecuteJobAndTrack_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ExecuteTask of Module JobsManager at line " & Erl
End Function


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
' Procedure : StartPipeline
' Purpose   : Performs imaging/fcs on a grid. Pretty much the whole macro runs through here
' Variables : GridName -
'             JobName -
'             parentPath - Path from where job has been initiated
'---------------------------------------------------------------------------------------
'
Public Function StartPipeline(index As Integer, RecordingDoc As DsRecordingDoc, FcsRecordingDoc As DsRecordingDoc, _
FcsData As AimFcsData, ParentPath As String, Optional WellFirst As Boolean = False) As Boolean
On Error GoTo StartPipeline_Error

    Dim i As Integer
    Dim ipip As Integer
    Dim iTask As Integer
    Dim stgPos As Vector
    
    Dim FileName As String
    Dim SuccessExecute As Boolean
    'Stop all running acquisitions (maybe to strong)
    StopAcquisition
    
    'coordinates
    Dim previousZ As Double   'remember position of previous position in Z
    
       
    OiaSettings.resetRegistry
      
    FileName = PipelineConstructor.TextBoxFileName.value & Pipelines(index).Grid.getName(1, 1, 1, 1) & Pipelines(index).Grid.suffix(1, 1, 1, 1) & Pipelines(index).Repetition.suffix(1)
    'create a new Gui document if recquired
    NewRecord RecordingDoc, "IMG:" & FileName
    
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
                
            Loop While .Grid.nextGridPt(WellFirst)
            
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
    Dim TimeStart As Double
    Dim TimeWait As Double
    ''check if we need to pump
    If (distDiff <= distMax Or distMax = 0) And (CDbl(GetTickCount) * 0.001 - lastTimePump <= timeMax Or timeMax = 0) Then
        waitForPump = lastTimePump
        Exit Function
    End If
    
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




'''
' make name of file from pipeline and task
'''
Private Function FileNameFromPipeline(indexPl As Integer, indexTask As Integer) As String
On Error GoTo FileNameFromPipeline_Error
    With Pipelines(indexPl)
        FileNameFromPipeline = appendSep(PipelineConstructor.TextBoxFileName.value, FNSep) & appendSep(.Grid.getThisName, FNSep) & appendSep(.Grid.NameGrid, FNSep) & _
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
            FilePathSuffixFromPipeline = appendSep(FilePathSuffixFromPipeline, FNSep) & .Grid.thisSuffix
        Else
            FilePathSuffixFromPipeline = appendSep(FilePathSuffixFromPipeline, FNSep) & .Grid.thisSuffixWell & "\" & appendSep(FilePathSuffixFromPipeline, FNSep) & .Grid.thisSuffix
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
    Dim imgSize As Vector
    imgSize = IJob.imageSize
    Dim MaxMovementXY As Double
    Dim MaxMovementZ As Double
    MaxMovementXY = MAX(imgSize.X, imgSize.Y)
   
    
                                
    If Abs(newPos.X - currentPos.X) > MaxMovementXY Or Abs(newPos.Y - currentPos.Y) > MaxMovementXY Or (IJob.isZStack And Abs(newPos.Z - currentPos.Z) > imgSize.Z) Then
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
    Dim imgSizePx As Vector
    Dim MaxY As Long
    Dim MaxZ As Long
    Dim i As Integer
    
    imgSizePx = IJob.imageSizePx


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
    
    If newPos.X - (imgSizePx.X - 1) > TolPx Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "X = " & newPos.X & " accepted range is X = " & 0 & "-" & imgSizePx.X & ". VBA macro sets value to center of image " & imgSizePx.X / 2
        newPos.X = imgSizePx.X / 2
    End If
    
    If newPos.Y - (imgSizePx.Y - 1) > TolPx Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "Y = " & newPos.Y & " accepted range is Y = " & 0 & "-" & imgSizePx.Y & ". VBA macro sets value to center of image" & imgSizePx.Y / 2
        newPos.Y = imgSizePx.Y / 2
    End If
    
    If newPos.Z - (imgSizePx.Z - 1) > TolPx And IJob.isZStack Then
        LogManager.UpdateErrorLog "Job " & IJob.Name & " " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " online image analysis returned a too large displacement/focus " & _
        "Z = " & newPos.Z & " accepted range is Z = " & 0 & "-" & imgSizePx.Z & ". VBA macro sets value to center of image" & imgSizePx.Z / 2
        newPos.Z = imgSizePx.Z / 2
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
    Dim imgSize As Vector
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
    imgSize = IJob.imageSizePx
    frameSpacing = IJob.Recording.frameSpacing
    
    For i = 0 To UBound(newPosition)
        position(i).X = (position(i).X - (imgSize.X - 1) / 2) * pixelSize
        position(i).Y = (position(i).Y - (imgSize.Y - 1) / 2) * pixelSize
        If IJob.isZStack Then
            position(i).Z = (position(i).Z - (imgSize.Z - 1) / 2) * frameSpacing
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
    Dim imgSizePx  As Vector
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim i As Integer
    Dim position() As Vector
    position = newPosition
    'pixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing 'This is in meter!!! be careful . Position for imaging is provided in um
    pixelSize = IJob.Recording.SampleSpacing ' this is in um
    frameSpacing = IJob.Recording.frameSpacing
    imgSizePx = IJob.imageSizePx
    'compute difference with respect to center
    
    For i = 0 To UBound(newPosition)
        'for FCS position is with respect center of image in meter
        position(i).X = (position(i).X - (imgSizePx.X - 1) / 2) * pixelSize * 0.000001
        position(i).Y = (position(i).Y - (imgSizePx.Y - 1) / 2) * pixelSize * 0.000001
        If IJob.isZStack Then
            position(i).Z = (position(i).Z - (imgSizePx.Z - 1) / 2) * frameSpacing
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
        ''' input grid positions only up to max values
        For i = 0 To UBound(newPositions)
            If .Grid.getNrValidPts < .optPtNumber Then
                .Grid.setPt newPositions(i), True, 1, 1, 1, i + GridLowBound
                .Grid.setParentPath ParentPath, 1, 1, 1, i + GridLowBound
                .Grid.setFcsPositions fcsPos, 1, 1, 1, i + GridLowBound
                .Grid.setFcsPositionsPx fcsPosPx, 1, 1, 1, i + GridLowBound
                .Grid.setName prefix & .Grid.getName(1, 1, 1, i + GridLowBound), 1, 1, 1, i + GridLowBound
            End If
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
                    If (.optPtNumber <= .Grid.getNrValidPts Or TimersGridCreation.wait(.Grid.NameGrid, Round(.maxWait)) < 0 Or .maxWait <= 0) Then
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
    Dim iTsk As Integer
    Dim JobName As String
    Dim prefix As String
    tsk = Pipelines(indexPl).getTask(indexTsk)
    JobName = Pipelines(indexPl).Grid.NameGrid & "_" & indexTsk + 1
    Dim codeMicToJobName As Dictionary 'use to convert codes of regisrty into Jobnames as used in the code
    Set codeMicToJobName = New Dictionary
    codeMicToJobName.Add "trigger1", 1
    codeMicToJobName.Add "trigger2", 2
    
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

    'Read positions and rois from registry the fcs positions are read with respect to center of image
    If OiaSettings.getFcsPositions(fcsPosPx, ImgJobs(tsk.jobNr).getCentralPointPx) Then
        VectorString = VectorList2String(fcsPosPx)
        LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & imgFileExtension & " obtained " & UBound(fcsPosPx) + 1 & " position(s) " & _
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
        LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & imgFileExtension & " obtained " & UBound(newPositionsPx) + 1 & " position(s) (in px)" & _
        " X = " & VectorString(0) & " Y = " & VectorString(1) & " Z = " & VectorString(2)
        If Not checkForMaximalDisplacementVecPixels(ImgJobs(tsk.jobNr), newPositionsPx) Then
            LogManager.UpdateErrorLog "OnlineImageAnalysis position exceed boundaries and has been set to  X = " & newPositionsPx(0).X & " Y = " & newPositionsPx(0).Y & " Z = " & newPositionsPx(0).Z
        End If
        newPositions = computeCoordinatesImaging(ImgJobs(tsk.jobNr), parentPosition, newPositionsPx)
        If Not checkForMaximalDisplacementVec(ImgJobs(tsk.jobNr), parentPosition, newPositions) Then
            GoTo ExitThis
        End If
        OiaSettings.writeKeyToRegistry "X", ""
        OiaSettings.writeKeyToRegistry "Y", ""
        OiaSettings.writeKeyToRegistry "Z", ""
    End If
    OiaSettings.getRois Rois
    OiaSettings.writeKeyToRegistry "roiAim", ""
    OiaSettings.writeKeyToRegistry "roiType", ""
    OiaSettings.writeKeyToRegistry "roiX", ""
    OiaSettings.writeKeyToRegistry "roiY", ""
    prefix = OiaSettings.readKeyFromRegistry("prefix")
    OiaSettings.writeKeyToRegistry "prefix", ""
    'TODO find way for passing ROIS
    
    ''for all commands in codeMic
    For Each code In codeMic
        LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " found " & code
        OiaSettings.writeKeyToRegistry "codeMic", "nothing"
        Select Case code
            Case "nothing", "": 'Nothing to do
                Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPosPx
                Pipelines(indexPl).Grid.setThisFcsName ""
                Pipelines(indexPl).Grid.setThisFcsImage ""
            Case "error":
                OiaSettings.readKeyFromRegistry "errorMsg"
                OiaSettings.getSettings ("errorMsg")
                LogManager.UpdateErrorLog "codeMic error. Online image analysis for task " & JobName & " file " & ParentPath & parentFile & " failed . " _
                & " Error from Oia: " & OiaSettings.getSettings("errorMsg")
                LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " obtained an error. " & OiaSettings.getSettings("errorMsg")
                OiaSettings.writeKeyToRegistry "errorMsg", ""
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
                ComputeJobSequential = newPositions(0)
                LogManager.UpdateLog "OnlineImageAnalysis from " & ParentPath & parentFile & " focus at  " & " X = " & newPositions(0).X & " Y = " & newPositions(0).Y & " Z = " & newPositions(0).Z & ". Absolute position in um"
            
            Case "setFcsPos":
                If isPosArrayEmpty(fcsPos) Then
                    LogManager.UpdateErrorLog "ComputeJobSequential: No position/wrong position for settings FCS. No FCS pts are being set."
                    Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                    Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPosPx
                    Pipelines(indexPl).Grid.setThisFcsName prefix
                    Pipelines(indexPl).Grid.setThisFcsImage ""
                    GoTo nextCode
                End If

                Pipelines(indexPl).Grid.setThisFcsPositions fcsPos
                Pipelines(indexPl).Grid.setThisFcsPositionsPx fcsPosPx
                Pipelines(indexPl).Grid.setThisFcsName prefix
                Pipelines(indexPl).Grid.setThisFcsImage ParentPath & parentFile & imgFileExtension
                
            Case "setRoi":
                If indexTsk + 1 > Pipelines(indexPl).count - 1 Then
                    LogManager.UpdateErrorLog "ComputeJobSequential: No next imaging task to which associate a ROI."
                    GoTo nextCode
                End If
                If Pipelines(indexPl).getTask(indexTsk + 1).jobType <> 0 Then
                    LogManager.UpdateErrorLog "ComputeJobSequential: No next imaging task to which associate a ROI."
                    GoTo nextCode
                End If
                If isArrayEmpty(Rois) Then
                    LogManager.UpdateErrorLog "ComputeJobSequential: No ROI specified. ROIs of next imaging task will be removed."
                    ImgJobs(Pipelines(indexPl).getTask(indexTsk + 1).jobNr).UseRoi = False
                    ImgJobs(Pipelines(indexPl).getTask(indexTsk + 1).jobNr).clearRois
                    GoTo nextCode
                End If
                ImgJobs(Pipelines(indexPl).getTask(indexTsk + 1).jobNr).UseRoi = True
                ImgJobs(Pipelines(indexPl).getTask(indexTsk + 1).jobNr).setRois Rois
                currentImgJob = -1  'reset current imaging Job so that next job will be forced to be reloaded
            
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
                'this creates a rois for all jobs in pipeline not optimal!!
                If Not isArrayEmpty(Rois) Then
                    For iTsk = 0 To Pipelines(codeMicToJobName.item(code)).count - 1
                        If Pipelines(codeMicToJobName.item(code)).getTask(iTsk).jobType = 0 Then
                            ImgJobs(Pipelines(codeMicToJobName.item(code)).getTask(iTsk).jobNr).UseRoi = True  'this is not exactly how it should be a task should have associated a roi
                            ImgJobs(Pipelines(codeMicToJobName.item(code)).getTask(iTsk).jobNr).setRois Rois   'this is not exactly how it should be a task should have associated a roi
                        End If
                    Next iTsk
                End If
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

