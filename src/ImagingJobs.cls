VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagingJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
' A Class to manage several AJob: imagingJob with different Settings and Tracks
'''
Option Explicit

'''' We can use number or a dictionary to find a Job in Jobs
Private Jobs() As AJob
Private JobNames As Dictionary
Private ZENv As Integer
Private ZoomPixelSlice(1 To 9, 1 To 3) As Double

Private Sub MsgNoJob(Name As String)
    MsgBox "ImagingJobs: Job " & Name & " has not been defined"
End Sub


'''
'   Initialize a series of Jobs with given name
'''
Public Sub initialize(Names() As String, Recording As DsRecording, ZEN As Object)
    Dim i As Integer
   ' Set Jobs = New AJob
    For i = 0 To UBound(Names)
        AddJob CStr(Names(i)), Lsm5.DsRecording, ZEN
    Next i
End Sub

Public Function getZENv() As Integer
    getZENv = ZENv
End Function

Public Function setZENv(Value As Integer)
    ZENv = Value
End Function

Public Sub resetJobs()
    Erase Jobs
    Set JobNames = New Dictionary
End Sub

'''
' AddJob
'   Add Job and increase number of Jobs
'''
Public Sub AddJob(Name As String, Recording As DsRecording, ZEN As Object)
    If JobNames Is Nothing Then
        Set JobNames = New Dictionary
    End If
    If Not JobNames.Exists(Name) Then
        If isArrayEmpty(Jobs) Then
            ReDim Jobs(0)
        Else
            ReDim Preserve Jobs(0 To UBound(Jobs) + 1)
        End If
        JobNames.Add Name, UBound(Jobs)
    End If
    Dim i As Integer
    Dim iJob As Integer
    iJob = JobNames.Item(Name)
    Set Jobs(iJob) = New AJob
    Jobs(iJob).Name = Name
    Jobs(iJob).setJob Lsm5.DsRecording, ZEN
End Sub


'''
'   Set the job with Name from ZEN using a given recording
'''
Public Function setJob(Name As String, Recording As DsRecording, ZEN As Object) As Boolean
    On Error GoTo ErrorHandle:
    Dim iJob As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If Jobs(JobNames.Item(Name)).setJob(Recording, ZEN) Then
        setJob = True
    End If
    Exit Function
ErrorHandle:
    MsgBox "Error in ImagingJobs.setJob " + Name + " " + Err.Description
End Function


Public Function GetRecording(Name As String) As DsRecording
    If JobNames.Exists(Name) Then
        Set GetRecording = Jobs(JobNames.Item(Name)).GetRecording
    Else
       MsgNoJob Name
    End If
End Function

Private Function checkJobName(Name As String) As Boolean
    If JobNames.Exists(Name) Then
        checkJobName = True
    Else
        MsgNoJob Name
    End If
End Function

'''
'   Put the job with Name from Macro into ZEN
'''
Public Function putJob(Name As String, ZEN As Object) As Boolean
    On Error GoTo ErrorHandle:
    Dim iJob As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If Jobs(JobNames.Item(Name)).putJob(ZEN) Then
        putJob = True
    End If
    Exit Function
ErrorHandle:
    MsgBox "Error in ImagingJobs.putJob " + Name + " " + Err.Description
 End Function


''''
'Roi Handler. This manages vectorElements in ZEN. Every Job can have some Rois associated to
''''
Public Sub setRoiMaxNr(Name As String, roiMaxNr As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).setRoiMaxNr roiMaxNr
End Sub

Public Sub setRoi(Name As String, ARoi As Roi, roiNr As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).setRoi ARoi, roiNr
End Sub

Public Sub setRois(Name As String, Rois() As Roi)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).setRois Rois
End Sub

Public Function getRoi(Name As String, roiNr As Integer) As Roi
    If Not checkJobName(Name) Then
        Exit Function
    End If
    Set getRoi = Jobs(JobNames.Item(Name)).getRoi(roiNr)
End Function


'''
' create ROIs in ZEN
'''
Public Sub createRois(Name As String)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).createRois
End Sub


Public Sub clearRois(Name As String)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).clearRois
End Sub


''''''
Public Function TrackNumber(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    TrackNumber = Jobs(JobNames.Item(Name)).TrackNumber
End Function


''''''
'    UpdateJobTimeSeries: if True the Job also has time Series
''''''
Public Sub setTimeSeries(Name As String, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.TimeSeries = Value
End Sub

Public Function getTimeSeries(Name As String) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getTimeSeries = Jobs(JobNames.Item(Name)).Recording.TimeSeries
End Function


'''
' X-Lines
'''
Public Function getSamplesPerLine(Name As String) As Long
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getSamplesPerLine = Jobs(JobNames.Item(Name)).Recording.SamplesPerLine
End Function

'''
' Y-Lines
''''
Public Function getLinesPerFrame(Name As String) As Long
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getLinesPerFrame = Jobs(JobNames.Item(Name)).Recording.LinesPerFrame
End Function

'''
' returns position of central point for a Job
''''
Public Function getCentralPointPx(Name As String) As Vector
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getCentralPointPx.X = (Jobs(JobNames.Item(Name)).Recording.SamplesPerLine - 1) / 2
    getCentralPointPx.Y = (Jobs(JobNames.Item(Name)).Recording.LinesPerFrame - 1) / 2
    If isZStack(Name) Then
        getCentralPointPx.Z = (Jobs(JobNames.Item(Name)).Recording.framesPerStack - 1) / 2
    Else
        getCentralPointPx.Z = 0
    End If
End Function

'''
' Interval for time series
'''
Public Sub setTimeBetweenStacks(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).TimeBetweenStacks = Value
End Sub

Public Function getTimeBetweenStacks(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getTimeBetweenStacks = Jobs(JobNames.Item(Name)).TimeBetweenStacks
End Function

Public Sub setTrackNumber(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).TrackNumber = Value
End Sub


Public Sub setAcquireTrack(Name As String, iTrack As Integer, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).ReDimAcquireTrack (Jobs(JobNames.Item(Name)).TrackNumber - 1)
    Jobs(JobNames.Item(Name)).setAcquireTrack iTrack, Value
End Sub

Public Function getAcquireTrack(Name As String, iTrack As Integer) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If iTrack < Jobs(JobNames.Item(Name)).TrackNumber Then
        getAcquireTrack = Jobs(JobNames.Item(Name)).getAcquireTrack(iTrack)
    End If
End Function

Public Function isAcquiring(Name As String) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If Jobs(JobNames.Item(Name)).isAcquiring Then
        isAcquiring = True
    End If
End Function

Public Sub setBleachTrack(Name As String, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Dim Success As Integer
    Jobs(JobNames.Item(Name)).Recording.TrackObjectBleach(Success).Acquire = Value
    Jobs(JobNames.Item(Name)).Recording.TrackObjectBleach(Success).UseBleachParameters = Value
End Sub

''''
'   True if track is bleaching
''''
Public Function getBleachTrack(Name As String) As Boolean
    Dim Success As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getBleachTrack = Jobs(JobNames.Item(Name)).Recording.TrackObjectBleach(Success).Acquire And Jobs(JobNames.Item(Name)).Recording.TrackObjectBleach(Success).UseBleachParameters
End Function


''''
'   This is the number of Z-stacks
''''
Public Sub setFramesPerStack(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.framesPerStack = Value
End Sub

Public Function getFramesPerStack(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getFramesPerStack = Jobs(JobNames.Item(Name)).Recording.framesPerStack
End Function

''''
'   This is dZ (in um)
''''
Public Sub setFrameSpacing(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.frameSpacing = Value
End Sub

Public Function getFrameSpacing(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getFrameSpacing = Jobs(JobNames.Item(Name)).Recording.frameSpacing
End Function

'''
'   This is dXdY in um
'''
Public Function getSampleSpacing(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    'Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing 'this is in meter different specification depending on the variable. What a mess!!!
    getSampleSpacing = Jobs(JobNames.Item(Name)).Recording.SampleSpacing
End Function


Public Sub setZoom(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.ZoomX = Value
    Jobs(JobNames.Item(Name)).Recording.ZoomY = Value
End Sub

Public Function getZoom(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getZoom = Jobs(JobNames.Item(Name)).Recording.ZoomX
End Function

Public Function getCorrX(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getCorrX = Jobs(JobNames.Item(Name)).CorrX
End Function


Public Function getCorrY(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getCorrY = Jobs(JobNames.Item(Name)).CorrY
End Function

Public Sub setCorrX(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).CorrX = Value
End Sub


Public Sub setCorrY(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).CorrY = Value
End Sub

Public Function getCentralPtPx(Name As String) As Vector
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getCentralPtPx.X = (Jobs(JobNames.Item(Name)).Recording.SamplesPerLine - 1) / 2
    getCentralPtPx.Y = (Jobs(JobNames.Item(Name)).Recording.LinesPerFrame - 1) / 2
    If Jobs(JobNames.Item(Name)).isZStack Then
        getCentralPtPx.Z = (Jobs(JobNames.Item(Name)).Recording.framesPerStack - 1) / 2
    Else
        getCentralPtPx.Z = 0
    End If
End Function


''''
'   Number of time points
'''
Public Sub setStacksPerRecord(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.StacksPerRecord = Value
End Sub


Public Function getStacksPerRecord(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getStacksPerRecord = Jobs(JobNames.Item(Name)).Recording.StacksPerRecord
End Function


'''''
' Sets usage of piezo or no
'   ZScan: X-Z Scan
'   Stack: Frame mode with Zstacks
'   Plane: just one plane
'   Line: Just a Line
'''''
Public Sub SetScanMode(Name As String, Value As String)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.ScanMode = Value
End Sub

Public Function GetScanMode(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetScanMode = Jobs(JobNames.Item(Name)).Recording.ScanMode
End Function

'true if record is a ZStack
Public Function isZStack(Name As String) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    isZStack = Jobs(JobNames.Item(Name)).isZStack
End Function

'''''
' Sets usage of piezo or no
'   ZScanner: Piezo
'   FocusStep: Focuswheel
'   OnTheFly: Focuswheel that is not syncronized with imaging
'''''
Public Sub SetSpecialScanMode(Name As String, Value As String)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.SpecialScanMode = Value
End Sub

Public Function GetSpecialScanMode(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetSpecialScanMode = Jobs(JobNames.Item(Name)).Recording.SpecialScanMode
End Function

'''''
' Sets Unidirectional or bidirectional scan
'    0: Unidirectional
'    1: bidirectional scan
'''''
Public Sub setScanDirection(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.ScanDirection = Value
End Sub

Public Function getScanDirection(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getScanDirection = Jobs(JobNames.Item(Name)).Recording.ScanDirection
End Function

'''
' Set and Get the usage of ROIs
'''
Public Function setUseRoi(Name As String, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Function
    End If
    Jobs(JobNames.Item(Name)).UseRoi = Value
End Function

Public Function getUseRoi(Name As String) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getUseRoi = Jobs(JobNames.Item(Name)).UseRoi
End Function

'''
' Set and Get the pixelDwellTime of Job
'''
Public Sub setPixelDwell(Name As String, pixelDwell As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).pixelDwell = pixelDwell
End Sub

Public Function getPixelDwell(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getPixelDwell = Jobs(JobNames.Item(Name)).pixelDwell
End Function



''''
' Create a string with the description of a Job
''''
Public Function jobDescriptor(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    jobDescriptor = Jobs(JobNames.Item(Name)).jobDescriptor
End Function

''''
' Create a string with the description of a Job to be used as settings
''''
Public Function jobDescriptorSettings(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    jobDescriptorSettings = Jobs(JobNames.Item(Name)).jobDescriptorSettings
End Function

''''
' Create a string with the description of a Job
''''
Public Function splittedJobDescriptor(Name As String, MaxLines As Integer) As String()
    If Not checkJobName(Name) Then
        Exit Function
    End If
    splittedJobDescriptor = Jobs(JobNames.Item(Name)).splittedJobDescriptor(MaxLines, jobDescriptor(Name))
End Function


Public Function changeJobFromDescriptor(Name As String, Descriptor As String, Value As Variant)
    If Not checkJobName(Name) Then
        Exit Function
    End If
    changeJobFromDescriptor = Jobs(JobNames.Item(Name)).changeJobFromDescriptor(Descriptor, Value)
End Function


