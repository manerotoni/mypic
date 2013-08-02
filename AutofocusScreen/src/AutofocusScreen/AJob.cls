VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Recording As DsRecording
Private pAcquireTrack() As Boolean 'must be private for no obvious reasons of VBA
Public pixelDwell As Double      'pixelDwell of acquisition
Public TrackNumber As Integer
Public TimeBetweenStacks As Double
Public UseRoi As Boolean
Public CorrX  As Double ' corretion for double directional scanning
Public CorrY As Double  ' corretion for double directional scanning
Private LaserOnOff() As Boolean
Const SwitchLaserLines = False ' this is not clear weather it works or not properly. Sometimes it is fast sometimes slow?


Private Rois() As Roi
Private pRoiMaxNr As Integer

Public Property Get AcquireTrack() As Boolean()
    AcquireTrack = pAcquireTrack
End Property

Public Property Let AcquireTrack(Value() As Boolean)
    Dim i As Integer
    ReDim pAcquireTrack(UBound(Value))
    For i = 0 To UBound(Value)
        pAcquireTrack(i) = Value(i)
    Next i
End Property

''''
' setJob load settings from recording/ZEN into macro
''''
Public Function setJob(RecordingLoc As DsRecording, ZEN As Object) As Boolean
    On Error GoTo ErrorHandle:
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    Dim AcquisitionParameters As AimAcquisitionController40.AimAcquisitionParameters
    Set AcquisitionParameters = AcquisitionController.AcquisitionParameters
    Dim iTrack As Integer
    Dim iLaser As Integer
    SetRecording RecordingLoc
    TrackNumber = RecordingLoc.GetNormalTrackCount
    TimeBetweenStacks = RecordingLoc.TrackObjectByMultiplexOrder(0, 1).TimeBetweenStacks
    
    If ZENv > 2010 Then
        Set ZEN = Application.ApplicationInterface
        If Recording.ScanDirection = 1 Then
            CorrX = ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrX.Value
            CorrY = ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrY.Value
        End If
        'Debug.Print "nr Lasers " & ZEN.gui.Acquisition.Channels.Track.LaserLine.ItemCount
        ReDim LaserOnOff(0 To TrackNumber - 1, 0 To ZEN.gui.Acquisition.Channels.Track.LaserLine.ItemCount - 1)
    End If
    
    pixelDwell = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    ReDim pAcquireTrack(TrackNumber - 1)
    For iTrack = 0 To TrackNumber - 1
        pAcquireTrack(iTrack) = Recording.TrackObjectByMultiplexOrder(iTrack, 1).Acquire
        If ZENv > 2010 Then
            If SwitchLaserLines Then
                ZEN.gui.Acquisition.Channels.Track.ByIndex = iTrack
                For iLaser = 0 To UBound(LaserOnOff, 2)
                    ZEN.gui.Acquisition.Channels.Track.LaserLine.ByIndex = iLaser
                    LaserOnOff(iTrack, iLaser) = ZEN.gui.Acquisition.Channels.Track.LaserLine.Acquire.Value
                    'Debug.Print " iTrack " & iTrack & " Laser " & iLaser & " status " & ZEN.gui.Acquisition.Channels.Track.LaserLine.Acquire.Value
                Next iLaser
            End If
        End If
    Next iTrack
    UseRoi = AcquisitionParameters.UseROIs
    setRoiFromZen
    setJob = True
    Exit Function
ErrorHandle:
    MsgBox "Error in AJob.setJob " + Err.Description
End Function

''''
' putJob put settings from Me.recording into ZEN
''''
Public Function putJob(ZEN As Object)
    On Error GoTo ErrorHandle:
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    Dim AcquisitionParameters As AimAcquisitionController40.AimAcquisitionParameters
    Set AcquisitionParameters = AcquisitionController.AcquisitionParameters
    Dim iTrack As Integer
    Dim iLaser As Integer
    Dim Time As Double
    If TrackNumber <> Lsm5.DsRecording.GetNormalTrackCount Then
        MsgBox "Nr of Tracks of  Job " + Name + " is " + CStr(TrackNumber) + ". Available tracks " + CStr(Lsm5.DsRecording.GetNormalTrackCount) + vbCrLf + "This may cause problems"
        
        'Exit Function
    End If
    
    Time = Timer
    If Not Recording Is Nothing Then
        Lsm5.DsRecording.Copy Recording
    Else
        MsgBox "No Recording initialized for Job " + Name
        Exit Function
    End If
    Debug.Print "Time put Recording " & Round(Timer - Time, 3)
    'at the moment we just set the active/inactive track and TimeBetweenStacks
    'consistency check for TrackNumber. May be pass error only to log?
    'this seems to be not recquired in Zen 2012
    
     Time = Timer
    For iTrack = 0 To TrackNumber - 1
        If ZENv > 2010 Then
            'ZEN.gui.Acquisition.Channels.Track.ByIndex = iTrack '(it does not display properly anyway)
            'ZEN.gui.Acquisition.Channels.Track.Acquire.Value = pAcquireTrack(iTrack)
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).Acquire = pAcquireTrack(iTrack)
            If Not isArrayEmpty(LaserOnOff) And pAcquireTrack(iTrack) And SwitchLaserLines Then
                For iLaser = 0 To UBound(LaserOnOff, 2)
                    ZEN.gui.Acquisition.Channels.Track.LaserLine.ByIndex = iLaser
                    If LaserOnOff(iTrack, iLaser) <> ZEN.gui.Acquisition.Channels.Track.LaserLine.Acquire.Value Then
                        ZEN.gui.Acquisition.Channels.Track.LaserLine.Acquire.Value = LaserOnOff(iTrack, iLaser)
                    End If
                Next iLaser
            End If
        Else
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).Acquire = pAcquireTrack(iTrack)
        End If
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).TimeBetweenStacks = TimeBetweenStacks
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).SampleObservationTime = pixelDwell
    Next iTrack
    Debug.Print "Times set tracks " & Round(Timer - Time, 3)
    
    AcquisitionParameters.UseROIs = UseRoi
    Time = Timer
  
   If ZENv > 2010 Then
      'this is slow ~1 sec (and can be only used for acquisition only
       'ZEN.gui.Acquisition.AcquisitionMode.ScanArea.Zoom.Value = Recording.ZoomX
       'ZEN.SetListEntrySelected "Scan.Mode.DirectionX", Recording.ScanDirection
        If Recording.ScanDirection Then
            If ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrX.Value <> CorrX Then
               ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrX.Value = CorrX
            End If
            If ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrY.Value <> CorrY Then
            ZEN.gui.Acquisition.AcquisitionMode.Bidirectional.CorrY.Value = CorrY
            End If
        End If
    End If
    Debug.Print "Time update Corr.X " & Round(Timer - Time, 3)
    createRois
    'testDescriptor
    putJob = True
    Exit Function
ErrorHandle:
    MsgBox "Error in AJob.putJob " + Name + " " + Err.Description
End Function

Public Sub SetAcquireTrack(iTrack As Integer, Value As Boolean)
    If UBound(pAcquireTrack) >= iTrack Then
        pAcquireTrack(iTrack) = Value
    End If
End Sub

Public Function isAcquiring() As Boolean
    Dim i As Integer
    For i = 0 To UBound(pAcquireTrack)
        If pAcquireTrack(i) = True Then
            isAcquiring = True
            Exit Function
        End If
    Next i
End Function

Public Function GetAcquireTrack(iTrack As Integer) As Boolean
    If UBound(pAcquireTrack) >= iTrack Then
        GetAcquireTrack = pAcquireTrack(iTrack)
    End If
End Function

Public Sub ReDimAcquireTrack(Value As Integer)
    ReDim Preserve pAcquireTrack(Value)
End Sub


Public Function GetRecording() As DsRecording
    Set GetRecording = Recording
End Function

Public Sub SetRecording(Value As DsRecording)
   Set Recording = Lsm5.CreateBackupRecording
   Recording.Copy Value
   TrackNumber = Recording.GetNormalTrackCount
End Sub

Public Sub setRoiMaxNr(roiMaxNr As Integer)
    Dim i As Integer
    pRoiMaxNr = roiMaxNr
    If pRoiMaxNr = 0 Then
        Erase Rois
    Else
        ReDim Rois(0 To pRoiMaxNr - 1)
        For i = 0 To pRoiMaxNr - 1
            Set Rois(i) = New Roi
        Next i
    End If
End Sub

Public Function getRoiMaxNr() As Integer
    getRoiMaxNr = pRoiMaxNr
End Function

Public Sub clearRois()
    ClearVectorElements
    pRoiMaxNr = 0
    Erase Rois
End Sub

''
' load Roi from ARoi inti class
''
Public Sub setRoi(ARoi As Roi, roiNr As Integer)
    If UBound(Rois) <= roiNr Then
        Rois(roiNr).setRoi ARoi.roiType, ARoi.roiAim, ARoi.getX, ARoi.getY
    Else
        MsgBox (" No Roi of nr" + roiNr + " initialize Rois first ")
    End If
End Sub

'''
' getRoi: returns a Roi
'''
Public Function getRoi(roiNr As Integer) As Roi
    If UBound(Rois) <= roiNr Then
        Set getRoi = Rois(roiNr)
    Else
        MsgBox (" No Roi of nr" + roiNr + " initialize Rois first ")
    End If
End Function

'''
' gets Roi directly from ZEN and save it. This is only used at the beginning
'''
Public Function setRoiFromZen() As Boolean
    On Error GoTo ErrorHandle:
    Dim i As Integer
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.Scancontroller.AcquisitionRegions
    'there are Rois already defined
    'Debug.Print "nr Elements " & vo.GetNumberElements
    setRoiMaxNr vo.GetNumberElements
    If vo.GetNumberElements > 0 Then
        For i = 0 To vo.GetNumberElements - 1
            Rois(i).setRoiFomZen i
        Next i
    End If
    setRoiFromZen = True
    Exit Function
ErrorHandle:
    MsgBox "Error AJob.setRoiFromZen " & Err.Description
    
End Function



Public Function createRois() As Boolean
    Dim ARoi As Roi
    Dim i As Integer
    ClearVectorElements
    For i = 0 To pRoiMaxNr - 1
        If Not Rois(i).MakeVectorElement Then
            Exit Function
        End If
    Next i
    createRois = True
End Function

Public Function isBleaching() As Boolean
    isBleaching = Recording.TrackObjectBleach(Success).Acquire And Recording.TrackObjectBleach(Success).UseBleachParameters
End Function

''''
' Create a string with the description of a Job
''''
Public Function jobDescriptor() As String
    Dim ABoolean As Boolean
    Dim Description As String
    Dim iTrack As Integer
    If (Recording.ScanMode = "ZScan") Or (Recording.ScanMode = "Line") Then
        Description = Description & "ScanMode = Line" & vbCrLf
        Description = Description & "       X = " & Recording.SamplesPerLine & ", Y = " & 1 & vbCrLf
    Else
        Description = Description & "ScanMode = Frame" & vbCrLf
        Description = Description & "       X = " & Recording.SamplesPerLine & ", Y = " & Recording.LinesPerFrame & vbCrLf
    End If
    
    Description = Description & "Bleaching = " & isBleaching & vbCrLf
    Description = Description & "Regions = " & Me.UseRoi & vbCrLf
    Description = Description & "PixelDwell = " & Round(Me.pixelDwell * 1000000#, 2) & " us" & vbCrLf
    Description = Description & "Zoom = " & Recording.ZoomX & vbCrLf
    Description = Description & "pixelSize = " & Round(Recording.SampleSpacing, 3) & " um" & vbCrLf
    If Recording.SpecialScanMode = "ZScanner" Then
        Description = Description & "Piezo = True" & vbCrLf
    Else
        Description = Description & "Piezo = False" & vbCrLf
    End If
    
    If Recording.SpecialScanMode = "OnTheFly" Then
        Description = Description & "Fast-Z-Line = True" & vbCrLf
    Else
        Description = Description & "Fast-Z-Line = False" & vbCrLf
    End If
    
    If Recording.ScanMode = "Stack" Or Recording.ScanMode = "ZScan" Then
        Description = Description & "Z-Stack = True" & vbCrLf
        Description = Description & "       Z-Slices    = " & Recording.framesPerStack & vbCrLf
        Description = Description & "       Z-Interval  = " & Round(Recording.frameSpacing, 2) & " um" & vbCrLf
        Description = Description & "       Z-Range   = " & Round(Recording.frameSpacing, 2) * (Recording.framesPerStack - 1) & " um" & vbCrLf
    Else
        Description = Description & "Z-Stack = False" & vbCrLf
    End If

     If Recording.TimeSeries Then
        Description = Description & "TimeSeries = " & Recording.TimeSeries & vbCrLf
        Description = Description & "       Cycles     = " & Recording.StacksPerRecord & vbCrLf
        Description = Description & "       Interval    = " & Me.TimeBetweenStacks & " sec"
    Else
        Description = Description & "TimeSeries = " & Recording.TimeSeries
    End If

        
    jobDescriptor = Description
End Function

Public Function splittedJobDescriptor(MaxLines As Integer, jobDescription As String) As String()
    Dim i As Integer
    Dim Label1() As String
    Dim Label2() As String
    Dim Output() As String
    ReDim Output(1)
    Label1 = Split(jobDescription, vbCrLf)
    If UBound(Label1) > MaxLines - 1 Then
        ReDim Label2(0 To UBound(Label1) - MaxLines)
        For i = 0 To UBound(Label1) - MaxLines
            Label2(i) = Label1(MaxLines + i)
        Next i
        ReDim Preserve Label1(0 To MaxLines - 1)
    Else
        ReDim Label2(0)
    End If
    Output(0) = Join(Label1, vbCrLf)
    Output(1) = Join(Label2, vbCrLf)
    splittedJobDescriptor = Output
End Function

''
' create a String to save of the "complete" job desriptin
''
Public Function jobDescriptorSettings() As String
    Dim Descriptor As String
    Descriptor = "JobName " & Me.Name & vbCrLf
    Descriptor = Descriptor & "pixelDwell " & Me.pixelDwell & vbCrLf
    Descriptor = Descriptor & "ScanMode " & Recording.ScanMode & vbCrLf
    Descriptor = Descriptor & "FramesPerStack " & Recording.framesPerStack & vbCrLf
    Descriptor = Descriptor & "FrameSpacing " & Recording.frameSpacing & vbCrLf
    Descriptor = Descriptor & "SamplesPerLine " & Recording.SamplesPerLine & vbCrLf
    Descriptor = Descriptor & "LinesPerFrame " & Recording.LinesPerFrame & vbCrLf
    Descriptor = Descriptor & "Bleaching " & isBleaching & vbCrLf
    Descriptor = Descriptor & "UseRoi " & Me.UseRoi & vbCrLf
    Descriptor = Descriptor & "Zoom " & Recording.ZoomX & vbCrLf
    Descriptor = Descriptor & "SpecialScanMode " & Recording.SpecialScanMode & vbCrLf
    Descriptor = Descriptor & "TimeSeries " & Recording.TimeSeries & vbCrLf
    Descriptor = Descriptor & "StacksPerRecord " & Recording.StacksPerRecord & vbCrLf
    Descriptor = Descriptor & "TimeBetweenStacks " & Me.TimeBetweenStacks & vbCrLf
    Descriptor = Descriptor & "EndJobDef " & vbCrLf
    jobDescriptorSettings = Descriptor

End Function

Public Sub testDescriptor()
    consistencyDescriptor (jobDescriptorSettings)
End Sub

Private Function consistencyDescriptor(Descriptor As String)
    Dim i As Integer
    Dim DerivedDescr() As String
    Dim KeyItem() As String
    DerivedDescr = Split(Descriptor, vbCrLf)
    For i = 0 To UBound(DerivedDescr)
        'Debug.Print DerivedDescr(i)
        Sleep (100)
        KeyItem = Split(DerivedDescr(i), " ")
        If UBound(KeyItem) = 1 Then
            If Not changeJobFromDescriptor(KeyItem(0), KeyItem(1)) Then
                Exit Function
            End If
        End If
    Next i
    consistencyDescriptor = True
End Function

''''
' change a job from a descriptor
'''
Public Function changeJobFromDescriptor(Descriptor As String, Value As Variant)
    If Recording Is Nothing Then
        MsgBox "AJob.changeJobFromDescriptor Recording has not been initiated!"
    End If
    
    Select Case Descriptor
        Case "JobName":
            Name = CStr(Value)
        Case "pixelDwell":
            pixelDwell = CDbl(Value)
        Case "ScanMode":
            Recording.ScanMode = CStr(Value)
        Case "FramesPerStack":
            Recording.framesPerStack = CLng(Value)
        Case "FrameSpacing":
            Recording.frameSpacing = CDbl(Value)
        Case "SamplesPerLine":
            Recording.SamplesPerLine = CLng(Value)
        Case "LinesPerFrame":
            Recording.LinesPerFrame = CLng(Value)
        Case "Bleaching":
            If CBool(Value) Then
                Recording.TrackObjectBleach(1).Acquire = CBool(Value)
                Recording.TrackObjectBleach(Success).UseBleachParameters = CBool(Value)
            End If
        Case "UseRoi":
            Me.UseRoi = CBool(Value)
        Case "Zoom":
            Recording.ZoomX = CDbl(Value)
            Recording.ZoomY = CDbl(Value)
        Case "SpecialScanMode":
            Recording.SpecialScanMode = CStr(Value)
        Case "TimeSeries":
            Recording.TimeSeries = CBool(Value)
        Case "StacksPerRecord":
            Recording.StacksPerRecord = CLng(Value)
        Case "TimeBetweenStacks":
            Me.TimeBetweenStacks = CDbl(Value)
        Case Else:
            MsgBox " No Item " & Descriptor
            Exit Function
    End Select
    changeJobFromDescriptor = True
End Function
