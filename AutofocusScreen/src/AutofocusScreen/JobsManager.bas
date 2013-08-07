Attribute VB_Name = "JobsManager"
''''
'' A Class to manage and imagingJob with different Settings and Tracks
''''
Option Explicit
''' The repetition for tasks
Public Reps As ImagingRepetitions
'name of the repetitions
Public RepNames() As String


''' A collection of imaging jobs each defining a recording setting
Public Jobs As ImagingJobs
'Contains name of the Jonbs
Public JobNames() As String
'short name of the jobs (prefix to the file)
Public JobShortNames As Collection
'the name of the job that is currently loaded
Public CurrentJob As String

''' A collection of fcs jobs each defining a specific fcs config (smaller set settings stored for ZENv < 2011)
Public JobsFcs As FcsJobs
'Contains name of the Jonbs
Public JobFcsNames() As String
'short name of the jobs (prefix to the file)
Public JobFcsShortNames As Collection
'the name of the Fcsjob that is currently loaded
Public CurrentJobFcs As String


''' The grid for tasks
Public Grids As ImagingGrids
''' Timers initiated when great is created, reinitialized if recquired
Public TimersGridCreation As Timers


''' A vector
''' ToDo move it to another module
Public Type Vector
  x As Double
  y As Double
  Z As Double
End Type


'''
'   Sets and execute an imaging Job
'       JobName: The name of the Job to execute
'       RecordingDoc: the dsRecording where image is stored
'       RocordingName: The name of the recording (also for the GUI)
'       position: A vector with stage position where to acquire image X, Y, and Z (cental slice)
''''
Public Function AcquireJob(JobName As String, RecordingDoc As DsRecordingDoc, RecordingName As String, position As Vector) As Boolean
    On Error GoTo ErrorHandle:
    Dim SuccessRecenter As Boolean
    Dim Time As Double
    'stop any running jobs
    AutofocusForm.StopScanCheck
    'Create a NewRecord if required
    NewRecord RecordingDoc, RecordingName, 0
    'move stage if required
    Time = Timer
    If Round(Lsm5.Hardware.CpStages.PositionX, PrecXY) <> Round(position.x, PrecXY) Or Round(Lsm5.Hardware.CpStages.PositionY, PrecXY) <> Round(position.y, PrecXY) Then
        If Not FailSafeMoveStageXY(position.x, position.y) Then
            AutofocusForm.StopAcquisition
            Exit Function
        End If
    End If
    'Debug.Print "Time to move stage " & Round(Timer - Time, 3)
    
    
    Time = Timer
    'Change settings for new Job if it is different from currentJob (global variable)
    If JobName <> CurrentJob Then
        Jobs.putJob JobName, ZEN
    End If
    Debug.Print "Time to put Job " & Round(Timer - Time, 3)
    CurrentJob = JobName
    'Not sure if this is required
    Time = Timer
    If Jobs.GetSpecialScanMode(JobName) = "ZScanner" Then
        Lsm5.Hardware.CpHrz.Leveling
    End If
    'Debug.Print "Time to level Hrz " & Round(Timer - Time, 3)
    
    
    ''' recenter before acquisition
    Time = Timer
    If Not Recenter_pre(position.Z, SuccessRecenter, ZENv) Then
        Exit Function
    End If
    'Debug.Print "Time to recenter pre " & Round(Timer - Time, 3)


    Time = Timer
    If Jobs.isAcquiring(JobName) Then
        If Not ScanToImage(RecordingDoc) Then
            Exit Function
        End If
    Else
        GoTo ErrorTrack
    End If
    'Debug.Print "Time to scan image " & Round(Timer - Time, 3)
    
    'wait that slice recentered after acquisition
    Time = Timer
    If Not Recenter_post(position.Z, SuccessRecenter, ZENv) Then
       Exit Function
    End If
    'Debug.Print "Time to recenter post " & Round(Timer - Time, 3)
    AcquireJob = True
    Exit Function
ErrorHandle:
    MsgBox "Error: AcquireJob for Job " + JobName + " " + Err.Description
    
    Exit Function
ErrorTrack:
    MsgBox "Error: AcquireJob for job " + JobName + ". Exit now!"
    Exit Function
End Function



'''''
' This executes part of the Job save the file compute offline tracking and set the registry
'''''
Public Function ExecuteJob(JobName As String, RecordingDoc As DsRecordingDoc, FilePath As String, FileName As String, _
StgPos As Vector, Optional deltaZ As Integer = -1)
    On Error GoTo ErrorHandle:
    If Not AcquireJob(JobName, RecordingDoc, FileName, StgPos) Then
        Exit Function
    End If
    'this is a dummy variable used for consistencey except for autofocus the default is saving of all images
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    OiaSettings.initializeDefault
    
    If AutofocusForm.Controls(JobName & "SaveImage") Then
        If Not SaveDsRecordingDoc(RecordingDoc, FilePath & FileName & ".lsm") Then
            Exit Function
        End If
        OiaSettings.writeKeyToRegistry "filePath", FilePath & FileName & ".lsm"
    End If
    
    If ScanStop Then
        Exit Function
    End If
    ExecuteJob = True
    Exit Function
ErrorHandle:
    ErrorLog.Show
    ErrorLog.UpdateLog "Error in AcquireJob for Job " + JobName + " " + Err.Description
    MsgBox "Error in AcquireJob for Job " + JobName + " " + Err.Description
    AutofocusForm.StopAcquisition
End Function

'''
' Compute new positions according to center of mass
'''''
Public Function TrackOffLine(JobName As String, RecordingDoc As DsRecordingDoc, currentPosition As Vector) As Vector
    On Error GoTo ErrorHandle:
    Dim newPosition() As Vector
    ReDim newPosition(0)
    Dim TrackingChannel As String
    newPosition(0) = currentPosition
    TrackOffLine = currentPosition
    If AutofocusForm.Controls(JobName & "CenterOfMass") And (AutofocusForm.Controls(JobName & "TrackZ") Or AutofocusForm.Controls(JobName & "TrackXY")) Then
        TrackingChannel = AutofocusForm.Controls(JobName & "CenterOfMassChannel").List(AutofocusForm.Controls(JobName & "CenterOfMassChannel").ListIndex)
        ''compute center of mass in pixel
        newPosition(0) = MassCenter(RecordingDoc, TrackingChannel)
        'transform it in um
        newPosition = computeCoordinatesImaging(JobName, currentPosition, newPosition)
    End If
    If AutofocusForm.Controls(JobName & "TrackZ") Then
        TrackOffLine.Z = newPosition(0).Z
    End If
    If AutofocusForm.Controls(JobName & "TrackXY") Then
        TrackOffLine.x = newPosition(0).x
        TrackOffLine.y = newPosition(0).y
    End If
    If Not checkForMaximalDisplacement(JobName, TrackOffLine, currentPosition) Then
        TrackOffLine = currentPosition
    End If
    Debug.Print "X = " & currentPosition.x & ", " & newPosition(0).x & ", Y = " & currentPosition.y & ", " & newPosition(0).y & ", Z = " & currentPosition.Z & ", " & newPosition(0).Z
    Exit Function
ErrorHandle:
    MsgBox "Error in TrackOffLine " + JobName + " " + Err.Description
    ErrorLog.UpdateLog "Error in TrackOffLine " + JobName + " " + Err.Description
End Function



''''
'   Update positions according to track command
''''
Public Function TrackJob(JobName As String, StgPos As Vector, StgPosNew As Vector) As Vector
    On Error GoTo ErrorHandle:
    TrackJob = StgPos
    If AutofocusForm.Controls(JobName & "TrackZ") Then
        TrackJob.Z = StgPosNew.Z
    End If
    If AutofocusForm.Controls(JobName & "TrackXY") Then
        TrackJob.x = StgPosNew.x
        TrackJob.y = StgPosNew.y
    End If
    Exit Function
ErrorHandle:
    MsgBox "Error in TrackJob " + JobName + " " + Err.Description
    ErrorLog.UpdateLog "Error in TrackJob " + JobName + " " + Err.Description
End Function

'''
'   Execute a job and perform a tracking. Returns an updated position X, Y, and Z
'''
Public Function ExecuteJobAndTrack(GridName As String, JobName As String, parentPath As String, StgPos As Vector, Success As Boolean) As Vector
    
    On Error GoTo ErrorHandle:
    Dim Time As Double
    Dim ScanMode As String
    Dim newStgPos As Vector
    Dim FileName As String
    Dim FilePath As String
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    
    Success = False
    

    If AutofocusForm.Controls(JobName + "Active") Then
        DisplayProgress "Job " & JobName & ", Row " & Grids.thisRow(GridName) & ", Col " & Grids.thisColumn(GridName) & vbCrLf & _
        "subRow " & Grids.thisSubRow(GridName) & ", subCol " & Grids.thisSubColumn(GridName) & ", Rep " & Reps.thisIndex(GridName), RGB(&HC0, &HC0, 0)

        ScanMode = Jobs.GetScanMode(JobName)
        If ScanMode = "ZScan" Or ScanMode = "Line" Then
            AutofocusForm.Controls(JobName & "TrackXY").Value = False
        End If
        FileName = FileNameFromGrid(GridName, JobName)
        FilePath = parentPath & FilePathSuffix(GridName, JobName) & "\"
        If JobName <> "Autofocus" Then
            StgPos.Z = StgPos.Z + AutofocusForm.Controls(JobName + "ZOffset").Value
        End If
        
        If AutofocusForm.Controls(JobName + "OiaActive") And AutofocusForm.Controls(JobName + "OiaSequential") Then
            OiaSettings.writeKeyToRegistry "codeMic", "wait"
        End If
        
        If Not ExecuteJob(JobName, GlobalRecordingDoc, FilePath, FileName, StgPos) Then
            ScanStop = True
            Exit Function
        End If
        'do any recquired computation
        Time = Timer
        StgPos = TrackOffLine(JobName, GlobalRecordingDoc, StgPos)
        Debug.Print "Time to TrackOffLine " & Timer - Time
        If AutofocusForm.Controls(JobName + "OiaActive") And AutofocusForm.Controls(JobName + "OiaSequential") Then
            OiaSettings.writeKeyToRegistry "codeOia", "newImage"
            newStgPos = ComputeJobSequential(JobName, GridName, StgPos, FilePath, FileName, GlobalRecordingDoc)
            If Not checkForMaximalDisplacement(JobName, StgPos, newStgPos) Then
                newStgPos = StgPos
            End If
                
            Debug.Print "X =" & StgPos.x & ", " & newStgPos.x & ", " & StgPos.y & ", " & newStgPos.y & ", " & StgPos.Z & ", " & newStgPos.Z
            StgPos = TrackJob(JobName, StgPos, newStgPos)
        End If
        
        If JobName <> "Autofocus" Then
            StgPos.Z = StgPos.Z - AutofocusForm.Controls(JobName + "ZOffset").Value
        End If
    
    End If
    Success = True
    Exit Function
ErrorHandle:
    ErrorLog.UpdateLog "Error in ExecuteJobAndTrack " + GridName + " " + JobName + " " + parentPath + " " + Err.Description
    ScanStop = True
End Function



''''''
'   StartAcquisition(BleachingActivated)
'   Perform many things (TODO: write more). Pretty much the whole macro runs through here
''''''
Public Sub StartJobOnGrid(GridName As String, JobName As String, parentPath As String)
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    Dim StgPos As Vector, newStgPos As Vector
    Dim Time As Double
    '''The name of jobs run for the global mode
    Dim JobNamesGlobal(2) As String
    Dim iJobGlobal As Integer
    Dim iGuiDocument As Integer
    JobNamesGlobal(0) = "Autofocus"
    JobNamesGlobal(1) = "Acquisition"
    JobNamesGlobal(2) = "AlterAcquisition"
    
    Dim FileName As String
    Dim deltaZ As Integer
    deltaZ = -1
    Dim SuccessRecenter As Boolean
    Dim SuccessExecute As Boolean
    'Stop all running acquisitions (maybe to strong)
    AutofocusForm.StopScanCheck
    
    'coordinates
    Dim previousZ As Double   'remember position of previous position in Z
    
    
    
    'Coordinates
    Dim x As Double              ' x value where to move the stage (this is used as reference)
    Dim y As Double              ' y value where to move the stage
    Dim Z As Double              ' z value where to move the stage
    Dim Xold As Double
    Dim Yold As Double
    Dim Zold As Double
    Dim MaxMovementXY As Double
    Dim MaxMovementZ As Double
    
    'test variables
    Dim Success As Integer       ' Check if something was sucessfull
    
    'Recording stuff
    Dim FilePath As String   ' full path of file to save (changes through function)
    Dim Scancontroller As AimScanController ' the controller
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    
    OiaSettings.resetRegistry
    OiaSettings.readFromRegistry
    
    FileName = AutofocusForm.TextBoxFileName.Value & Grids.getName(JobName, 1, 1, 1, 1) & Grids.suffix(JobName, 1, 1, 1, 1) & Reps.suffix(JobName, 1)
    'create a new Gui document if recquired
    NewRecordGui GlobalRecordingDoc, FileName, ZEN, ZENv, 0
    
    CurrentJob = ""
    Running = True  'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.

     
    
    previousZ = Grids.getZ(JobName, 1, 1, 1, 1)
    Reps.resetIndex (JobName)
    '''
    ' Check if there are any valid positions
    ''''
    If Grids.getNrValidPts(GridName) = 0 Then
        DisplayProgress "Job " & JobName & ", on grid " & GridName & " has no valid positions !", RGB(&HC0, &HC0, 0)
        Sleep (500)
        Exit Sub
    End If
    
    While Reps.nextRep(GridName) ' cycle all repetitions
        Grids.setIndeces GridName, 1, 1, 1, 1
        Do ''Cycle all positions defined in grid
            If Grids.getThisValid(GridName) Then
               DisplayProgress "Job " & JobName & ", Row " & Grids.thisRow(GridName) & ", Col " & Grids.thisColumn(JobName) & vbCrLf & _
                "subRow " & Grids.thisSubRow(GridName) & ", subCol " & Grids.thisSubColumn(GridName) & ", Rep " & Reps.thisIndex(GridName), RGB(&HC0, &HC0, 0)

                'Do some positional Job
                StgPos.x = Grids.getThisX(GridName)
                StgPos.y = Grids.getThisY(GridName)
                StgPos.Z = Grids.getThisZ(GridName)
                
                If Reps.getIndex(GridName) = 1 And AutofocusForm.GridScanActive Then
                    StgPos.Z = previousZ
                End If

                ' Recenter and move where it should be. Job global is a series of jobs
                ' TODO move into one single function per task
                If JobName = "Global" Then
                    For iJobGlobal = 0 To UBound(JobNamesGlobal)
                        ' run subJobs for global setting
                         StgPos = ExecuteJobAndTrack(GridName, JobNamesGlobal(iJobGlobal), parentPath, StgPos, SuccessExecute)
                        If Not SuccessExecute Then
                            GoTo StopAcquisition
                        End If
                    Next iJobGlobal
                Else
                    StgPos = ExecuteJobAndTrack(GridName, JobName, parentPath, StgPos, SuccessExecute)
                    If Not SuccessExecute Then
                        GoTo StopAcquisition
                    End If
                End If
                
                Grids.setThisX GridName, StgPos.x
                Grids.setThisY GridName, StgPos.y
                Grids.setThisZ GridName, StgPos.Z
                previousZ = Grids.getThisZ(GridName)
            End If
            If ScanPause = True Then
                If Not AutofocusForm.Pause Then ' Pause is true if Resume
                    GoTo StopAcquisition
                    Exit Sub
                End If
            End If
        Loop While Grids.nextGridPt(JobName)
        ''Wait till next repetition
        Reps.updateTimeStart (JobName)
        
        If Reps.wait(JobName) > 0 Then
            DisplayProgress "Waiting " & CStr(CInt(Reps.wait(JobName))) & " s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
            DoEvents
        End If
        
        While ((Reps.wait(JobName) > 0) And (Reps.getIndex(JobName) < Reps.getRepetitionNumber(JobName)))
            Sleep (100)
            DoEvents
            If ScanPause = True Then
                If Not AutofocusForm.Pause Then ' Pause is true if Resume
                    GoTo StopAcquisition
                    Exit Sub
                End If
            End If
            If ScanStop Then
                GoTo StopAcquisition
            End If
            DisplayProgress "Waiting " & CStr(CInt(Reps.wait(JobName))) & " s before scanning repetition  " & Reps.getIndex(JobName) + 1, RGB(&HC0, &HC0, 0)
        Wend
        Sleep (100)
        DoEvents
        If ScanPause = True Then
            If Not AutofocusForm.Pause Then ' Pause is true is Resume
                GoTo StopAcquisition
            End If
        End If
        If ScanStop Then
            GoTo StopAcquisition
        End If
    Wend
    Exit Sub
StopAcquisition:
    ScanStop = True
    AutofocusForm.StopAcquisition
    Exit Sub
ErrorHandle1:
    ScanStop = True
    MsgBox "Error StartJobOnGrid for Job " + JobNamesGlobal(iJobGlobal) + " on Grid " + GridName + " " + Err.Description
    AutofocusForm.StopAcquisition
    Exit Sub
ErrorHandle2:
    ScanStop = True
    MsgBox "Error StartJobOnGrid for Job " + JobName + " on Grid " + GridName + " " + Err.Description
    AutofocusForm.StopAcquisition
    Exit Sub
End Sub

'''
' Derive filename from Grid and repetition
'''
Private Function FileNameFromGrid(GridName As String, JobName As String) As String
On Error GoTo ErrorHandle:
     FileNameFromGrid = AutofocusForm.TextBoxFileName.Value & Grids.getThisName(GridName) & JobShortNames(JobName) & "_" & Grids.thisSuffix(GridName) & Reps.thisSuffix(GridName)
     Exit Function
ErrorHandle:
    MsgBox "Error in FileNameOnGrid " + Err.Description
End Function

''''
' check  that newPos is not further away than the size of the image. In fact it should be half the image
''''
Public Function checkForMaximalDisplacement(JobName As String, currentPos As Vector, newPos As Vector) As Boolean
    Dim MaxMovementXY As Double
    Dim MaxMovementZ As Double
    
    MaxMovementXY = Max(Jobs.GetSamplesPerLine(JobName), Jobs.GetLinesPerFrame(JobName)) * Jobs.GetSampleSpacing(JobName)
    MaxMovementZ = Jobs.GetFramesPerStack(JobName) * Jobs.GetFrameSpacing(JobName)
                                
    If Abs(newPos.x - currentPos.x) > MaxMovementXY Or Abs(newPos.y - currentPos.y) > MaxMovementXY Or Abs(newPos.Z - currentPos.Z) > MaxMovementZ Then
        ErrorLog.UpdateLog "Job " & JobName & " online image analysis returned a too large displacement/focus " & _
        "dX, dY, dZ = " & Abs(newPos.x - currentPos.x) & ", " & Abs(newPos.y - currentPos.y) & ", " & Abs(newPos.Z - currentPos.Z) & vbCrLf & _
        "accepted dX, dY, dZ = " & MaxMovementXY & ", " & MaxMovementXY & ", " & MaxMovementZ
        Exit Function
    End If
    checkForMaximalDisplacement = True
End Function


''''
' check  that newPos is not further away than the size of the image. In fact it should be half the image
''''
Private Function checkForMaximalDisplacementVec(JobName As String, currentPos As Vector, newPos() As Vector) As Boolean
    Dim MaxMovementXY As Double
    Dim MaxMovementZ As Double
    Dim i As Integer
    MaxMovementXY = Jobs.GetFrameSize(JobName) * Jobs.GetSampleSpacing(JobName)
    MaxMovementZ = Jobs.GetFramesPerStack(JobName) * Jobs.GetFrameSpacing(JobName)
    For i = 0 To UBound(newPos)
        If Abs(newPos(i).x - currentPos.x) > MaxMovementXY Or Abs(newPos(i).y - currentPos.y) > MaxMovementXY Or Abs(newPos(i).Z - currentPos.Z) > MaxMovementZ Then
            ErrorLog.UpdateLog "Job " & JobName & " online image analysis returned a too large displacement/focus " & _
            "dX, dY, dZ = " & Abs(newPos(i).x - currentPos.x) & ", " & Abs(newPos(i).y - currentPos.y) & ", " & Abs(newPos(i).Z - currentPos.Z) & vbCrLf & _
            "accepted dX, dY, dZ = " & MaxMovementXY & ", " & MaxMovementXY & ", " & MaxMovementZ
            Exit Function
        End If
    Next i
    checkForMaximalDisplacementVec = True
End Function

'''
' Derive filepath Suffix from Grid and repetition
'''
Private Function FilePathSuffix(GridName As String, JobName As String) As String
    FilePathSuffix = AutofocusForm.TextBoxFileName.Value & Grids.getThisName(GridName) & JobShortNames(JobName)
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


'''
'   Update the settings of the corresponding Formpage from the Job
'''
Public Sub UpdateFormFromJob(Jobs As ImagingJobs, JobName As String)
    
    'update form for any new tracks
    'AutofocusForm.AutoFindTracks
    Dim i As Integer
    Dim Record As DsRecording
    Dim jobDescriptor() As String
    Set Record = Jobs.GetRecording(JobName)
    
    For i = 0 To TrackNumber - 1
       AutofocusForm.Controls(JobName + "Track" + CStr(i + 1)).Value = Jobs.GetAcquireTrack(JobName, i)
    Next i
         
    jobDescriptor = Jobs.splittedJobDescriptor(JobName, 8)
    AutofocusForm.Controls(JobName + "Label1").Caption = jobDescriptor(0)
    If UBound(jobDescriptor) > 0 Then
        AutofocusForm.Controls(JobName + "Label2").Caption = jobDescriptor(1)
    End If
    
    If Jobs.GetScanMode(JobName) = "ZScan" Or Jobs.GetScanMode(JobName) = "Line" Then
        AutofocusForm.Controls(JobName + "TrackXY").Value = False
        AutofocusForm.Controls(JobName + "TrackXY").Enabled = False
    Else
        AutofocusForm.Controls(JobName + "TrackXY").Enabled = AutofocusForm.Controls(JobName + "Active")
    End If
    AutofocusForm.FillTrackingChannelList JobName
End Sub

'''
'   Update the settings of Job with JobName from corresponding Formpage
'''
Public Sub UpdateJobFromForm(Jobs As ImagingJobs, JobName As String)
    Dim i As Integer
    
    For i = 0 To TrackNumber - 1
       Jobs.SetAcquireTrack JobName, i, AutofocusForm.Controls(JobName + "Track" + CStr(i + 1)).Value
    Next i
    AutofocusForm.UpdateRepetitionTimes
End Sub

'''
'   Updates the Gui AcquisitionMode from the Job
'''
Public Sub UpdateGuiFromJob(Jobs As ImagingJobs, JobName As String) ', ZEN As Object)
    Dim iTrack As Integer
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
    Dim ScanMode As String
    If ZEN Is Nothing Then
        Exit Sub
    End If
    ScanMode = Jobs.GetScanMode(JobName)
    ZEN.gui.Acquisition.AcquisitionMode.FrameSizeX.Value = Jobs.GetSamplesPerLine(JobName)
    ZEN.gui.Acquisition.AcquisitionMode.FrameSizeY.Value = Jobs.GetSamplesPerLine(JobName)

    If ScanMode = "ZScan" Or ScanMode = "Line" Then
        ZEN.gui.Acquisition.AcquisitionMode.ScanMode.ByName = "Line"
    End If
    
    If ScanMode = "Stack" Or ScanMode = "Plane" Then
        ZEN.gui.Acquisition.AcquisitionMode.ScanMode.ByName = "Frame"
    End If
    
    If ScanMode = "Point" Then
        ZEN.gui.Acquisition.AcquisitionMode.ScanMode.ByName = "Point"
    End If
    
    ZEN.gui.Acquisition.AcquisitionMode.ScanArea.Zoom.Value = Jobs.GetRecording(JobName).ZoomX
    ZEN.SetListEntrySelected "Scan.Mode.DirectionX", Jobs.GetRecording(JobName).ScanDirection
    
    'ZEN.GUI.Acquisition.AcquisitionMode.BitDepth.ByIndex = 1
    'This unfortunately does not update the GUI
'    For iTrack = 0 To Jobs.TrackNumber(JobName) - 1
'        ZEN.gui.Acquisition.Channels.Track.ByIndex = iTrack '(it does not display properly anyway)
'        ZEN.gui.Acquisition.Channels.Track.Acquire.Value = Jobs.GetAcquireTrack(JobName, iTrack)
'    Next iTrack

    
     
End Sub





'''''''
'   computeShiftedCoordinates(offsetPosition As Vector, currentPosition As Vector) As Vector
'   given offsetPosition with (0,0,0) center of image central slice (in um)
'   the function compute absolute stage/focus coordinates from currentPosition
''''''
Public Function computeShiftedCoordinates(offsetPosition As Vector, currentPosition As Vector) As Vector
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
        computeShiftedCoordinates.x = currentPosition.x + Xpre * offsetPosition.y
        computeShiftedCoordinates.y = currentPosition.y + Ypre * offsetPosition.x
    Else
        computeShiftedCoordinates.x = currentPosition.x + Xpre * offsetPosition.x
        computeShiftedCoordinates.y = currentPosition.y + Ypre * offsetPosition.y
    End If
      
    computeShiftedCoordinates.Z = currentPosition.Z + offsetPosition.Z
    
    computeShiftedCoordinates.x = Round(computeShiftedCoordinates.x, PrecXY)
    computeShiftedCoordinates.y = Round(computeShiftedCoordinates.y, PrecXY)
    computeShiftedCoordinates.Z = Round(computeShiftedCoordinates.Z, PrecZ)
End Function


''''
' compute offset coordinates for imaging from pixel coordinates
' newPosition() As Vector
' the values are returned in um!!
''''
Public Function computeCoordinatesImaging(JobName As String, currentPosition As Vector, newPosition() As Vector) As Vector()
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim imageSize As Integer
    Dim framesPerStack As Integer
    Dim i As Integer
    Dim position() As Vector
    position = newPosition
    'pixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing 'This is in meter!!!
    pixelSize = Jobs.GetSampleSpacing(JobName) ' this is in um
    'compute difference with respect to center
    imageSize = Jobs.GetFrameSize(JobName)
    framesPerStack = Jobs.GetFramesPerStack(JobName)
    frameSpacing = Jobs.GetFrameSpacing(JobName)
    For i = 0 To UBound(newPosition)
        position(i).x = (position(i).x - (imageSize - 1) / 2) * pixelSize
        position(i).y = (position(i).y - (imageSize - 1) / 2) * pixelSize
        If Jobs.isZStack(JobName) Then
            position(i).Z = (position(i).Z - (framesPerStack - 1) / 2) * frameSpacing
        Else
            position(i).Z = 0
        End If
        position(i) = computeShiftedCoordinates(position(i), currentPosition)
    Next i
    computeCoordinatesImaging = position
End Function


''''
'   compute coordinates for fcs from pixel coordinates
'   newPosition() assumes origin is (0,0), central slice is 0
'   the value are returned in meter!!
''''
Public Function computeCoordinatesFcs(JobName As String, newPosition() As Vector) As Vector()
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim imageSize As Integer
    Dim i As Integer
    Dim position() As Vector
    position = newPosition
    'convert in um
    pixelSize = Jobs.GetSampleSpacing(JobName)
    'compute difference with respect to center
    imageSize = Jobs.GetFrameSize(JobName)
    frameSpacing = Jobs.GetFrameSpacing(JobName)
    For i = 0 To UBound(newPosition)
        position(i).x = position(i).x * pixelSize
        position(i).y = position(i).y * pixelSize
        position(i).Z = position(i).Z * frameSpacing
    Next i
    computeCoordinatesFcs = position
End Function

Private Function MinInt(value1 As Integer, value2 As Integer) As Integer
    If value1 <= value2 Then
        MinInt = value1
    Else
        MinInt = value2
    End If
End Function


''''
' create and update a subgrid and eventually decide whether to run Job
''''
Public Function runSubImagingJob(GridName As String, JobName As String, newPositions() As Vector) As Boolean
    Dim i As Integer
    Dim ptNumber As Integer ' number of pts for the grid
    Dim maxWait As Double   ' maximal time to wait for the grid
    Dim GridLowBound As Integer
       
    If AutofocusForm.Controls(JobName + "OptimalPtNumber").Value <> "" Then
        ptNumber = CInt(AutofocusForm.Controls(JobName + "OptimalPtNumber").Value)
    Else
        ptNumber = 0
    End If
    
    If AutofocusForm.Controls(JobName + "maxWait").Value <> "" Then
        maxWait = CDbl(AutofocusForm.Controls(JobName + "maxWait").Value)
    Else
        maxWait = 0
    End If
    
    ''createnew grid if recquired
    If Not Grids.checkGridName(GridName) Then
        Grids.AddGrid (GridName)
    End If

    '' change size of grid
    If Grids.isGridEmpty(GridName) Then
        ''start counter for gridcreation!!!
        Grids.updateGridSize GridName, 1, 1, 1, UBound(newPositions) + 1
        GridLowBound = 1
        If TimersGridCreation Is Nothing Then
            Set TimersGridCreation = New Timers
        End If
        TimersGridCreation.addTimer GridName
        TimersGridCreation.updateTimeStart GridName
    Else
        GridLowBound = Grids.numColSub(GridName) + 1
        Grids.updateGridSizePreserve GridName, 1, 1, 1, UBound(newPositions) + GridLowBound
    End If
    
    ''' input grid positions
    For i = 0 To UBound(newPositions)
            Grids.setPt GridName, newPositions(i), True, 1, 1, 1, i + GridLowBound
    Next i
    
    If ptNumber = 0 Or maxWait = 0 Then
        runSubImagingJob = True
        Exit Function
    End If
        
    If AutofocusForm.Controls(JobName + "OptimalPtNumber").Value = "" And AutofocusForm.Controls(JobName + "maxWait").Value = "" Then
        ' if the value is empty we image whatever has been found
        runSubImagingJob = True
        Exit Function
    End If
    
    If AutofocusForm.Controls(JobName + "OptimalPtNumber").Value = "" Then
        If TimersGridCreation.wait(GridName, CDbl(AutofocusForm.Controls(JobName + "maxWait").Value)) < 0 Then
            runSubImagingJob = True
            Exit Function
        End If
    End If
    
        
    If AutofocusForm.Controls(JobName + "maxWait").Value = "" Then
        If Grids.getNrPts(GridName) >= ptNumber Then
            'trim grid
            Grids.updateGridSizePreserve GridName, 1, 1, 1, AutofocusForm.Controls(JobName + "OptimalPtNumber").Value
            runSubImagingJob = True
            Exit Function
        End If
    End If
    
    'both are unequal 0. you chose which occurs first
    If AutofocusForm.Controls(JobName + "OptimalPtNumber").Value <> "" Then
        If Grids.getNrPts(GridName) >= AutofocusForm.Controls(JobName + "OptimalPtNumber").Value Then
            'trim grid
            Grids.updateGridSizePreserve GridName, 1, 1, 1, AutofocusForm.Controls(JobName + "OptimalPtNumber").Value
            runSubImagingJob = True
            Exit Function
        End If
        
        If TimersGridCreation.wait(GridName, CDbl(AutofocusForm.Controls(JobName + "maxWait").Value)) < 0 Then
            runSubImagingJob = True
            Exit Function
        End If
    End If
    
    If AutofocusForm.Controls(JobName + "OptimalPtNumber").Value <> "" And AutofocusForm.Controls(JobName + "OptimalPtNumber").Value = "" Then
        If AutofocusForm.Controls(JobName + "OptimalPtNumber").Value >= Grids.getNrPts(GridName) Then
            'trim grid
            Grids.updateGridSizePreserve GridName, 1, 1, 1, AutofocusForm.Controls(JobName + "OptimalPtNumber").Value
            runSubImagingJob = True
            Exit Function
        End If
    End If
    
End Function

'''
'   Perform computation for tracking and wait for Onlineimageanalysis if sequential Oia is on
'   Create/update grid for OiaJobs
''''
Public Function ComputeJobSequential(parentJob As String, parentGrid As String, parentPosition As Vector, parentPath As String, parentFile As String, RecordingDoc As DsRecordingDoc, _
 Optional deltaZ As Integer = -1) As Vector
    Dim imageSize As Integer
    Dim newPositions() As Vector
    Dim codeMic As String
    Dim JobName As String 'local convenience variable
    
    Dim codeMicToJobName As Dictionary 'use to convert codes of regisrty into Jobnames as used in the code
    Set codeMicToJobName = New Dictionary
    codeMicToJobName.Add "trigger1", "Trigger1"
    codeMicToJobName.Add "trigger2", "Trigger2"
    
    Dim OiaSettings As OnlineIASettings
    Set OiaSettings = New OnlineIASettings
    codeMic = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="codeMic")
    
    
    Dim TimeWait, TimeStart, MaxTimeWait As Double
    
    MaxTimeWait = 100
    
    'default return value is currentPosition
    ComputeJobSequential = parentPosition
    
    Select Case codeMic
        Case "wait":
            'Wait for image analysis to finish
            DisplayProgress "Waiting for image analysis...", RGB(0, &HC0, 0)
            TimeStart = CDbl(GetTickCount) * 0.001
            Do While ((TimeWait < MaxTimeWait) And (codeMic = "wait"))
                Sleep (50)
                TimeWait = CDbl(GetTickCount) * 0.001 - TimeStart
                codeMic = GetSetting(appname:="OnlineImageAnalysis", section:="macro", _
                          Key:="codeMic")
                DoEvents
                If ScanStop Then
                    Exit Function
                End If
            Loop

            If TimeWait > MaxTimeWait Then
                codeMic = "timeExpired"
                SaveSetting "OnlineImageAnalysis", "macro", "codeMic", codeMic
                SaveSetting "OnlineImageAnalysis", "macro", "codeOia", "nothing"
            End If
    End Select

    ''Read all settings at once
    OiaSettings.readFromRegistry
    
    'read coordinates
    If Not OiaSettings.checkCodeMic(OiaSettings.getSettings("codeMic")) Then
        ScanStop = True
        Exit Function
    End If
    
    ComputeJobSequential = parentPosition
    Select Case codeMic
        Case "nothing", "": 'Nothing to do
            ComputeJobSequential = parentPosition
            
        Case "error":
            OiaSettings.writeKeyToRegistry "codeMic", "nothing"
            ErrorLog.Show
            ErrorLog.UpdateLog ("cedeMic error. Online image analysis for job " + parentJob + " file " + OiaSettings.getSettings("filePath") + " failed")
            
        Case "timeExpired":
            OiaSettings.writeKeyToRegistry "codeMic", "nothing"
            ErrorLog.Show
            ErrorLog.UpdateLog ("codeMic timeExpired. Online image analysis for job " + parentJob + " file " + OiaSettings.getSettings("filePath") + " took more then " & MaxTimeWait & " sec")
            
        Case "focus":
            OiaSettings.writeKeyToRegistry "codeMic", "nothing"
            If OiaSettings.getPositions(newPositions, parentPosition) Then
                newPositions = computeCoordinatesImaging(parentJob, parentPosition, newPositions)
            Else
                ErrorLog.Show
                ErrorLog.UpdateLog ("ComputeJobSequential: No position/wrong position for Job focus. " & vbCrLf & _
                "Specify one position in X, Y, Z of registry (in pixels, (X,Y) = (0,0) upper left corner image, Z = 0 -> central slice of current stack)!")
                Exit Function
            End If
            If UBound(newPositions) > 0 Then
                ErrorLog.Show
                ErrorLog.UpdateLog "ComputeJobSequential: for Job focus. Pass only one point instead of " & UBound(newPositions) + 1
            End If
            ComputeJobSequential = newPositions(0)
            
        Case "trigger1", "trigger2": 'store positions for later processing or direct imaging depending on settings
            OiaSettings.writeKeyToRegistry "codeMic", "nothing"
            JobName = codeMicToJobName.Item(codeMic)
            DisplayProgress "Registry codeMic " & codeMic & ": store positions and eventually image job" & JobName & "...", RGB(0, &HC0, 0)
            If Not AutofocusForm.Controls(JobName + "Active") Then
                ErrorLog.UpdateLog "ComputeJobSequential: job " & JobName & " is not active"
                Exit Function
            End If
            If OiaSettings.getPositions(newPositions, parentPosition) Then
                newPositions = computeCoordinatesImaging(parentJob, parentPosition, newPositions)
                ' if displacement are above the possible displacement estimated from current image then abort
                If Not checkForMaximalDisplacementVec(parentJob, parentPosition, newPositions) Then
                    ScanStop = True
                    Exit Function
                End If
            Else
                ErrorLog.UpdateLog "ComputeJobSequential: No position for Job " & JobName & " (key = " & codeMic & ") has been specified!"
                ScanStop = True
                Exit Function
            End If
            
            ''' if we run a subjob the grid and counter is reset
            If runSubImagingJob(JobName, JobName, newPositions) Then
                'remove positions from parent grid to avoid revisiting the position
                'Grids.setThisValid parentGrid, False
                'start acquisition of Job on grid named JobName
                StartJobOnGrid JobName, JobName, parentPath & parentFile & "\"
                'set all run positions to notValid
                Grids.setAllValid JobName, False
            End If
            
        Case "6", "fcs":
'            SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
'            DisplayProgress "Registry Code 6 (fcs): peform a fcs measurment ...", RGB(0, &HC0, 0)
'            'create empty arrays
'            Erase locX
'            Erase locY
'            Erase locZ
'            Erase locDeltaZ
'            DisplayProgress "Registry Code 6 (fcs): perform FCS measurement ...", RGB(0, &HC0, 0)
'            'StorePositionsFromRegistry Xref, Yref, Zref, locX, locY, locZ, locDeltaZ
'            DisplayProgress "Registry Code 6 (fcs): perform FCS measurement ...", RGB(0, &HC0, 0)
'            'SubImagingWorkFlowFcs FcsData, locX, locY, locZ, GlobalDataBaseName & BackSlash, AutofocusForm.TextBoxFileName.Value & UnderScore & FileNameID, _
'            RecordingDoc.Recording.SampleSpacing , RecordingDoc.Recording.frameSpacing
'
        Case "7", "Trigger2", "trigger2":
            'This only specify position of ROIs the stage is not moved
'            SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
'            DisplayProgress "Registry Code 7 (bleach): peform a bleach measurment ...", RGB(0, &HC0, 0)
'            ' read potitions from
'            ReDim locX(0)
'            ReDim locY(0)
'            ReDim locZ(0)
'            ReDim locDeltaZ(0)
'            locX(0) = Xref
'            locY(0) = Yref
'            locZ(0) = Zref + AutofocusForm.Trigger2ZOffset.Value
'            locDeltaZ(0) = -1
'
'            Repetitions.number = 1
'            Repetitions.time = 0
'            Repetitions.interval = True
'            DisplayProgress "Registry Code 8 (bleach): perform Bleaching ...", RGB(0, &HC0, 0)
'            'Delete All ROI's
'            Dim AcquisitionController As AimAcquisitionController40.AimScanController
'            Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
'            Dim vo As AimImageVectorOverlay
'            Set vo = AcquisitionController.AcquisitionRegions
'            vo.Cleanup
'            CreateRoisFromRegistry
'
'            'StorePositionsFromRegistry Xref, Yref, Zref, HighResArrayX, HighResArrayY, HighResArrayZ, HighResArrayDeltaZ
'            DisplayProgress "Registry Code 8 (bleach): perform Bleaching ...", RGB(0, &HC0, 0)
'            SubImagingWorkFlow RecordingDoc, GlobalBleachRecording, "Bleach", AutofocusForm.Trigger1Autofocus, AutofocusForm.Trigger1ZOffset, Repetitions, _
'            locX, locY, locZ, locDeltaZ, GridPos, FileNameID, iRepStart
'            Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
'            Track.UseBleachParameters = False  'switch off the bleaching

        Case Else
            MsgBox ("Invalid OnlineImageAnalysis codeMic = " & codeMic)
            AutofocusForm.StopAcquisition
            Exit Function
    End Select

End Function




Public Function ComputeJobParallel(JobName As String, Recording As DsRecording, FilePath As String, FileName As String, x As Double, _
y As Double, Z As Double, Optional deltaZ As Integer = -1) As Double()
    Dim OiaSettings As Dictionary
    Dim NewCoord() As Double
    ReDim NewCoord(3)
    'Defaults we dont change anything
    deltaZ = -1
    NewCoord(0) = x
    NewCoord(1) = y
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
