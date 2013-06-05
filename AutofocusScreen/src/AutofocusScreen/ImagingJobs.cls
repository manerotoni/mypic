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


'''' We can use number or a dictionary to find a Job in Jobs
Private Jobs As Collection
Private ZENv As Integer
Private ZoomPixelSlice(1 To 9, 1 To 3) As Double

Private Sub MsgNoJob(Name As String)
    MsgBox "ImagingJobs: Job " & Name & " has not been defined"
End Sub


'''
'   Initialize a series of Jobs with given name
'''
Public Sub initialize(Names() As String, Recording As DsRecording)
    Dim Name As Variant
    Set Jobs = New Collection
    For Each Name In Names
        AddJob CStr(Name), Lsm5.DsRecording
    Next Name
    
    
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



'''
' AddJob
'   Add Job to Jobs Collection with key Name and use Recording
'''
Public Sub AddJob(Name As String, Recording As DsRecording)
    Dim Job As AJob
    Dim i As Integer
    Dim pixelDwell As Double
    Set Job = New AJob
    Job.Name = Name
    Job.SetRecording Lsm5.DsRecording

    Dim Aq() As Boolean
    ReDim Aq(Job.TrackNumber - 1)
    For i = 0 To Job.TrackNumber - 1
        Aq(i) = Recording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i
    Job.AcquireTrack = Aq
    Job.TimeBetweenStacks = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).TimeBetweenStacks
    Jobs.Add Job, Name
    SetPixelDwell Name, Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
End Sub


'''
'   Set the job with Name from ZEN using a given recording
'''
Public Sub SetJob(Name As String, Recording As DsRecording, ZEN As Object)
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    Dim AcquisitionParameters As AimAcquisitionController40.AimAcquisitionParameters
    Set AcquisitionParameters = AcquisitionController.AcquisitionParameters
    If InCollection(Jobs, Name) Then
        Jobs(Name).SetRecording Recording
        Jobs(Name).TrackNumber = Jobs(Name).Recording.GetNormalTrackCount
        Jobs(Name).TimeBetweenStacks = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).TimeBetweenStacks
        If ZENv > 2010 Then
            If GetScanDirection(Name) = 1 Then
                SetCorrX Name, ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrX.Value
                SetCorrY Name, ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrY.Value
            End If
        End If
        SetPixelDwell Name, Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
        Dim Aq() As Boolean
        ReDim Aq(Jobs(Name).TrackNumber - 1)
        For i = 0 To Jobs(Name).TrackNumber - 1
            Aq(i) = Recording.TrackObjectByMultiplexOrder(i, 1).Acquire
        Next i
        Jobs(Name).AcquireTrack = Aq
        Dim bo As Boolean
        bo = AcquisitionParameters.UseROIs
        SetUseRoi Name, AcquisitionParameters.UseROIs
    Else
        MsgNoJob Name
    End If
    
End Sub


Public Function GetRecording(Name As String) As DsRecording
    Set GetRecording = Jobs(Name).GetRecording
End Function
'''
'   Put the job with Name from Macro into ZEN
'''
Public Sub PutJob(Name As String, ZEN As Object)
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    Dim AcquisitionParameters As AimAcquisitionController40.AimAcquisitionParameters
    Set AcquisitionParameters = AcquisitionController.AcquisitionParameters
    Dim iTrack As Integer
    If InCollection(Jobs, Name) Then
        If Not Jobs(Name).Recording Is Nothing Then
            Lsm5.DsRecording.Copy Jobs(Name).Recording
        End If
        'at the moment we just set the active/inactive track and TimeBetweenStacks
        For iTrack = 0 To Jobs(Name).TrackNumber - 1
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).Acquire = Jobs(Name).GetAcquireTrack(iTrack)
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).TimeBetweenStacks = GetTimeBetweenStacks(Name)
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).SampleObservationTime = GetPixelDwell(Name)
        Next iTrack
        AcquisitionParameters.UseROIs = GetUseRoi(Name)
        'AcquisitionParameters.Bidirectional = True
        
    End If
    
    If ZENv > 2010 Then
        ZEN.gui.Acquisition.AcquisitionMode.ScanArea.Zoom.Value = GetZoom(Name)
        ZEN.SetListEntrySelected "Scan.Mode.DirectionX", GetScanDirection(Name)
        If GetScanDirection(Name) = 1 Then
            ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrX.Value = GetCorrX(Name)
            ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrY.Value = GetCorrY(Name)
        End If
  
    End If
 End Sub

Public Function TrackNumber(Name As String)
    TrackNumber = Jobs(Name).TrackNumber
End Function

''''''
'    UpdateJobTimeSeries: if True the Job also has time Series
''''''
Public Sub SetTimeSeries(Name As String, Value As Boolean)
    Jobs(Name).Recording.TimeSeries = Value
End Sub

Public Function GetTimeSeries(Name As String) As Boolean
    GetTimeSeries = Jobs(Name).Recording.TimeSeries
End Function

Public Sub SetFrameSize(Name As String, Value As Integer)
    Jobs(Name).Recording.LinesPerFrame = CLng(Value)  'This is Y (check)
    Jobs(Name).Recording.SamplesPerLine = CLng(Value) 'This is X (check)
End Sub

Public Function GetFrameSize(Name As String)
    GetFrameSize = Jobs(Name).Recording.SamplesPerLine
End Function

'''
' X-Lines
'''
Public Function GetSamplesPerLine(Name As String)
    GetSamplesPerLine = Jobs(Name).Recording.SamplesPerLine
End Function

'''
' Y-Lines
''''
Public Function GetLinesPerFrame(Name As String)
    GetLinesPerFrame = Jobs(Name).Recording.LinesPerFrame
End Function


'''
' Interval for time series
'''
Public Sub SetTimeBetweenStacks(Name As String, Value As Double)
    Jobs(Name).TimeBetweenStacks = Value
End Sub

Public Function GetTimeBetweenStacks(Name As String) As Double
    GetTimeBetweenStacks = Jobs(Name).TimeBetweenStacks
End Function

Public Sub SetTrackNumber(Name As String, Value As Integer)
    Jobs(Name).TrackNumber = Value
End Sub


Public Sub SetAcquireTrack(Name As String, iTrack As Integer, Value As Boolean)
    Jobs(Name).ReDimAcquireTrack (Jobs(Name).TrackNumber - 1)
    Jobs(Name).SetAcquireTrack iTrack, Value
End Sub

Public Function GetAcquireTrack(Name As String, iTrack As Integer) As Boolean
    If iTrack < Jobs(Name).TrackNumber Then
        GetAcquireTrack = Jobs(Name).GetAcquireTrack(iTrack)
    End If
End Function

Public Sub SetBleachTrack(Name As String, Value As Boolean)
    Dim Success As Integer
    Jobs(Name).Recording.TrackObjectBleach(Success).Acquire = Value
    Jobs(Name).Recording.TrackObjectBleach(Success).UseBleachParameters = Value
End Sub

''''
'   True if track is bleaching
''''
Public Function GetBleachTrack(Name As String) As Boolean
    Dim Success As Integer
    GetBleachTrack = Jobs(Name).Recording.TrackObjectBleach(Success).Acquire And Jobs(Name).Recording.TrackObjectBleach(Success).UseBleachParameters
End Function

Public Sub SetStacks(JobName As String, ZRange As Double, ZStep As Double)
    SetFramesPerStack JobName, CLng(ZRange / ZStep) + 1
    SetFrameSpacing JobName, ZStep
End Sub

''''
'   This is the number of Z-stacks
''''
Public Sub SetFramesPerStack(Name As String, Value As Integer)
    Jobs(Name).Recording.FramesPerStack = Value
End Sub

Public Function GetFramesPerStack(Name As String) As Integer
    GetFramesPerStack = Jobs(Name).Recording.FramesPerStack
End Function

''''
'   This is dZ (in um)
''''
Public Sub SetFrameSpacing(Name As String, Value As Double)
    Jobs(Name).Recording.FrameSpacing = Value
End Sub

Public Function GetFrameSpacing(Name As String) As Double
    GetFrameSpacing = Jobs(Name).Recording.FrameSpacing
End Function

'''
'   This is dXdY in meter
'''
Public Function GetSampleSpacing(Name As String) As Double
    GetSampleSpacing = Jobs(Name).Recording.SampleSpacing
End Function


Public Sub SetZoom(Name As String, Value As Double)
    Jobs(Name).Recording.ZoomX = Value
    Jobs(Name).Recording.ZoomY = Value
End Sub

Public Function GetZoom(Name As String) As Double
    GetZoom = Jobs(Name).Recording.ZoomX
End Function

Public Function GetCorrX(Name As String) As Double
    GetCorrX = Jobs(Name).CorrX
End Function


Public Function GetCorrY(Name As String) As Double
    GetCorrY = Jobs(Name).CorrY
End Function

Public Sub SetCorrX(Name As String, Value As Double)
    Jobs(Name).CorrX = Value
End Sub


Public Sub SetCorrY(Name As String, Value As Double)
    Jobs(Name).CorrY = Value
End Sub

''''
'   Number of time points
'''
Public Sub SetStacksPerRecord(Name As String, Value As Integer)
    Jobs(Name).Recording.StacksPerRecord = Value
End Sub


Public Function GetStacksPerRecord(Name As String) As Integer
    GetStacksPerRecord = Jobs(Name).Recording.StacksPerRecord
End Function


'''''
' Sets usage of piezo or no
'   ZScan: X-Z Scan
'   Stack: Frame mode
'   Plane: just one plane
'   Line: A Line
'''''
Public Sub SetScanMode(Name As String, Value As String)
    Jobs(Name).Recording.ScanMode = Value
End Sub

Public Function GetScanMode(Name As String) As String
    GetScanMode = Jobs(Name).Recording.ScanMode
End Function


'''''
' Sets usage of piezo or no
'   ZScanner: Piezo
'   FocusStep: Focuswheel
'   OnTheFly: Focuswheel that is not syncronized with imaging
'''''
Public Sub SetSpecialScanMode(Name As String, Value As String)
     Jobs(Name).Recording.SpecialScanMode = Value
End Sub

Public Function GetSpecialScanMode(Name As String) As String
    GetSpecialScanMode = Jobs(Name).Recording.SpecialScanMode
End Function
'''''
' Sets Unidirectional or bidirectional scan
'    0: Unidirectional
'    1: bidirectional scan
'''''
Public Sub SetScanDirection(Name As String, Value As Integer)
    Jobs(Name).Recording.ScanDirection = Value
End Sub

Public Function GetScanDirection(Name As String) As Integer
    GetScanDirection = Jobs(Name).Recording.ScanDirection
End Function

'''
' Set and Get the usage of ROIs
'''
Public Function SetUseRoi(Name As String, Value As Boolean)
    Jobs(Name).UseRoi = Value
End Function

Public Function GetUseRoi(Name As String) As Boolean
    GetUseRoi = Jobs(Name).UseRoi
End Function

'''
' Set and Get the pixelDwellTime of Job
'''
Public Sub SetPixelDwell(Name As String, pixelDwell As Double)
    Jobs(Name).pixelDwell = pixelDwell
End Sub

Public Function GetPixelDwell(Name As String) As Double
    GetPixelDwell = Jobs(Name).pixelDwell
End Function



''''
' Create a string with the description of a Job
''''
Public Function JobDescriptor1(JobName As String) As String
    Dim Recording As DsRecording
    Dim Description As String
    Dim iTrack As Integer
    If (GetScanMode(JobName) = "ZScan") Or (GetScanMode(JobName) = "Line") Then
        Description = Description & "ScanMode = Line" & vbCrLf
        Description = Description & "       X = " & GetSamplesPerLine(JobName) & ", Y = " & 1 & vbCrLf
    Else
        Description = Description & "ScanMode = Frame" & vbCrLf
        Description = Description & "       X = " & GetSamplesPerLine(JobName) & ", Y = " & GetLinesPerFrame(JobName) & vbCrLf
    End If
    
    Description = Description & "Bleaching = " & GetBleachTrack(JobName) & vbCrLf
    Description = Description & "Regions = " & GetUseRoi(JobName) & vbCrLf
    
    
    Description = Description & "PixelDwell = " & Round(GetPixelDwell(JobName) * 1000000#, 2) & " us" & vbCrLf
    Description = Description & "Zoom = " & GetZoom(JobName) & vbCrLf
    
    Description = Description & "Piezo = " & (GetSpecialScanMode(JobName) = "ZScanner") & vbCrLf
    Description = Description & "Fast-Z-Line = " & (GetSpecialScanMode(JobName) = "OnTheFly") & vbCrLf

    
        
    JobDescriptor1 = Description
End Function

''''
' Create a string with the description of a Job
''''
Public Function JobDescriptor2(JobName As String) As String
    Dim Description As String
    Dim iTrack As Integer
    Description = "Z-Stack = " & (GetScanMode(JobName) = "Stack" Or GetScanMode(JobName) = "ZScan") & vbCrLf
    If GetScanMode(JobName) = "Stack" Or GetScanMode(JobName) = "ZScan" Then
        Description = Description & "       Z-Slices    = " & GetFramesPerStack(JobName) & vbCrLf
        Description = Description & "       Z-Interval  = " & Round(GetFrameSpacing(JobName), 2) & " um" & vbCrLf
        Description = Description & "       Z-Range   = " & Round(GetFrameSpacing(JobName), 2) * (GetFramesPerStack(JobName) - 1) & " um" & vbCrLf
    End If

    Description = Description & "TimeSeries = " & GetTimeSeries(JobName) & vbCrLf
    If GetTimeSeries(JobName) Then
        Description = Description & "       Cycles     = " & GetStacksPerRecord(JobName) & vbCrLf
        Description = Description & "       Interval    = " & GetTimeBetweenStacks(JobName) & " sec" & vbCrLf
    End If


    Description = Description & vbCrLf
    
'    For iTrack = 0 To Jobs(JobName).TrackNumber - 1
'            Description = Description & Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).Name & " = " & Jobs(JobName).GetAcquireTrack(iTrack)
'    Next iTrack

    
    JobDescriptor2 = Description
End Function
