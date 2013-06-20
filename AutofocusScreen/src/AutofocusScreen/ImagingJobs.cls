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
' A Class to manage and imagingJob with different Settings and Tracks
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
Public Sub initialize(Names() As String, Recording As DsRecording)
    Dim i As Integer
   ' Set Jobs = New AJob
    For i = 0 To UBound(Names)
        AddJob CStr(Names(i)), Lsm5.DsRecording
    Next i
    
    
    ' This may be different depending on the setting and microscope
    ' define here the Zoom pixelSize slice relation specificed for 256x1 line. pixeldWell rescales with 256/FrameSize
    ZoomPixelSlice(1, 1) = 5
    ZoomPixelSlice(2, 1) = 3.1
    ZoomPixelSlice(3, 1) = 2
    ZoomPixelSlice(4, 1) = 1.2
    ZoomPixelSlice(5, 1) = 0.8
    ZoomPixelSlice(6, 1) = 0
    ZoomPixelSlice(7, 1) = 0
    ZoomPixelSlice(8, 1) = 0
    ZoomPixelSlice(9, 1) = 0
    
    'pixel dwell
    ZoomPixelSlice(1, 1) = 0.00000128 '1.28 us
    ZoomPixelSlice(2, 2) = 0.0000016  '1.6
    ZoomPixelSlice(3, 2) = 0.00000192 '1.92
    ZoomPixelSlice(4, 2) = 0.00000256 '2.56
    ZoomPixelSlice(5, 2) = 0.0000032  '3.2
    ZoomPixelSlice(6, 2) = 0.00000512 '5.12
    ZoomPixelSlice(7, 2) = 0.0000064  '6.4
    ZoomPixelSlice(8, 2) = 0.0000128  '12.8
    ZoomPixelSlice(9, 2) = 0.0000256  '25.6
    
    'slice size
    ZoomPixelSlice(1, 3) = 0.08
    ZoomPixelSlice(2, 3) = 0.1
    ZoomPixelSlice(3, 3) = 0.12
    ZoomPixelSlice(4, 3) = 0.15
    ZoomPixelSlice(5, 3) = 0.19
    ZoomPixelSlice(6, 3) = 0.31
    ZoomPixelSlice(7, 3) = 0.38
    ZoomPixelSlice(8, 3) = 0.77
    ZoomPixelSlice(9, 3) = 1.54
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
Public Sub AddJob(Name As String, Recording As DsRecording)
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
    Jobs(iJob).SetRecording Lsm5.DsRecording

    Dim Aq() As Boolean
    ReDim Aq(Jobs(iJob).TrackNumber - 1)
    For i = 0 To Jobs(iJob).TrackNumber - 1
        Aq(i) = Recording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i
    Jobs(iJob).AcquireTrack = Aq
    Jobs(iJob).TimeBetweenStacks = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).TimeBetweenStacks
    Jobs(iJob).pixelDwell = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
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

Public Function getRoi(Name As String, roiNr As Integer) As Roi
    If Not checkJobName(Name) Then
        Exit Function
    End If
    Set getRoi = Jobs(JobNames.Item(Name)).getRoi(roiNr)
End Function

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
Public Sub SetTimeSeries(Name As String, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.TimeSeries = Value
End Sub

Public Function GetTimeSeries(Name As String) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetTimeSeries = Jobs(JobNames.Item(Name)).Recording.TimeSeries
End Function

Public Sub SetFrameSize(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.LinesPerFrame = CLng(Value)  'This is Y (check)
    Jobs(JobNames.Item(Name)).Recording.SamplesPerLine = CLng(Value) 'This is X (check)
End Sub

Public Function GetFrameSize(Name As String) As Long
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetFrameSize = Jobs(JobNames.Item(Name)).Recording.SamplesPerLine
End Function

'''
' X-Lines
'''
Public Function GetSamplesPerLine(Name As String) As Long
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetSamplesPerLine = Jobs(JobNames.Item(Name)).Recording.SamplesPerLine
End Function

'''
' Y-Lines
''''
Public Function GetLinesPerFrame(Name As String) As Long
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetLinesPerFrame = Jobs(JobNames.Item(Name)).Recording.LinesPerFrame
End Function


'''
' Interval for time series
'''
Public Sub SetTimeBetweenStacks(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).TimeBetweenStacks = Value
End Sub

Public Function GetTimeBetweenStacks(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetTimeBetweenStacks = Jobs(JobNames.Item(Name)).TimeBetweenStacks
End Function

Public Sub SetTrackNumber(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).TrackNumber = Value
End Sub


Public Sub SetAcquireTrack(Name As String, iTrack As Integer, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).ReDimAcquireTrack (Jobs(JobNames.Item(Name)).TrackNumber - 1)
    Jobs(JobNames.Item(Name)).SetAcquireTrack iTrack, Value
End Sub

Public Function GetAcquireTrack(Name As String, iTrack As Integer) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If iTrack < Jobs(JobNames.Item(Name)).TrackNumber Then
        GetAcquireTrack = Jobs(JobNames.Item(Name)).GetAcquireTrack(iTrack)
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

Public Sub SetBleachTrack(Name As String, Value As Boolean)
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
Public Function GetBleachTrack(Name As String) As Boolean
    Dim Success As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetBleachTrack = Jobs(JobNames.Item(Name)).Recording.TrackObjectBleach(Success).Acquire And Jobs(JobNames.Item(Name)).Recording.TrackObjectBleach(Success).UseBleachParameters
End Function


''''
'   This is the number of Z-stacks
''''
Public Sub SetFramesPerStack(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.FramesPerStack = Value
End Sub

Public Function GetFramesPerStack(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetFramesPerStack = Jobs(JobNames.Item(Name)).Recording.FramesPerStack
End Function

''''
'   This is dZ (in um)
''''
Public Sub SetFrameSpacing(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.frameSpacing = Value
End Sub

Public Function GetFrameSpacing(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetFrameSpacing = Jobs(JobNames.Item(Name)).Recording.frameSpacing
End Function

'''
'   This is dXdY in meter
'''
Public Function GetSampleSpacing(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetSampleSpacing = Jobs(JobNames.Item(Name)).Recording.SampleSpacing
End Function


Public Sub SetZoom(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.ZoomX = Value
    Jobs(JobNames.Item(Name)).Recording.ZoomY = Value
End Sub

Public Function GetZoom(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetZoom = Jobs(JobNames.Item(Name)).Recording.ZoomX
End Function

Public Function GetCorrX(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetCorrX = Jobs(JobNames.Item(Name)).CorrX
End Function


Public Function GetCorrY(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetCorrY = Jobs(JobNames.Item(Name)).CorrY
End Function

Public Sub SetCorrX(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).CorrX = Value
End Sub


Public Sub SetCorrY(Name As String, Value As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).CorrY = Value
End Sub

''''
'   Number of time points
'''
Public Sub SetStacksPerRecord(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.StacksPerRecord = Value
End Sub


Public Function GetStacksPerRecord(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetStacksPerRecord = Jobs(JobNames.Item(Name)).Recording.StacksPerRecord
End Function


'''''
' Sets usage of piezo or no
'   ZScan: X-Z Scan
'   Stack: Frame mode
'   Plane: just one plane
'   Line: A Line
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
Public Sub SetScanDirection(Name As String, Value As Integer)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).Recording.ScanDirection = Value
End Sub

Public Function GetScanDirection(Name As String) As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetScanDirection = Jobs(JobNames.Item(Name)).Recording.ScanDirection
End Function

'''
' Set and Get the usage of ROIs
'''
Public Function SetUseRoi(Name As String, Value As Boolean)
    If Not checkJobName(Name) Then
        Exit Function
    End If
    Jobs(JobNames.Item(Name)).UseRoi = Value
End Function

Public Function GetUseRoi(Name As String) As Boolean
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetUseRoi = Jobs(JobNames.Item(Name)).UseRoi
End Function

'''
' Set and Get the pixelDwellTime of Job
'''
Public Sub SetPixelDwell(Name As String, pixelDwell As Double)
    If Not checkJobName(Name) Then
        Exit Sub
    End If
    Jobs(JobNames.Item(Name)).pixelDwell = pixelDwell
End Sub

Public Function GetPixelDwell(Name As String) As Double
    If Not checkJobName(Name) Then
        Exit Function
    End If
    GetPixelDwell = Jobs(JobNames.Item(Name)).pixelDwell
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

Public Function jobDescriptorSettings(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    jobDescriptorSettings
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


