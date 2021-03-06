VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''
'   AJob stores the parameters for a imaging job
'''''

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
Public SamplingNumber As Long
Public SamplingMode As Long
Public SamplingMethod As Long
Public BitDepth As Long
Private ptimeToAcquire As Double 'time to acquire a job in seconds
'''a Job can also have Rois associated to it
Private Rois() As roi
''Number of Rois associated to the Job
Private pRoiMaxNr As Integer




Public Property Get timeToAcquire() As Double
    timeToAcquire = ptimeToAcquire
End Property
    

Public Property Let timeToAcquire(value As Double)
    ptimeToAcquire = value
End Property
    
Public Property Get AcquireTrack() As Boolean()
    AcquireTrack = pAcquireTrack
End Property

Public Property Let AcquireTrack(value() As Boolean)
    Dim i As Integer
    ReDim pAcquireTrack(UBound(value))
    For i = 0 To UBound(value)
        pAcquireTrack(i) = value(i)
    Next i
End Property

''''
' setJob load settings from recording/ZEN into macro
''''
Public Function SetJob(RecordingLoc As DsRecording, ZEN As Object) As Boolean

    Dim AcquisitionController As AimAcquisitionController40.AimScanController
On Error GoTo SetJob_Error

    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
    Dim AcquisitionParameters As AimAcquisitionController40.AimAcquisitionParameters
    Set AcquisitionParameters = AcquisitionController.AcquisitionParameters
    Dim iTrack As Integer
    Dim iLaser As Integer
    SetRecording RecordingLoc
    TrackNumber = RecordingLoc.GetNormalTrackCount
    TimeBetweenStacks = RecordingLoc.TrackObjectByMultiplexOrder(0, 1).TimeBetweenStacks
    
    If ZenV > 2010 Then
        If Recording.ScanDirection = 1 Then
            CorrX = ZEN.GUI.Acquisition.AcquisitionMode.Bidirectional.CorrX.value
            CorrY = ZEN.GUI.Acquisition.AcquisitionMode.Bidirectional.CorrY.value
        End If
        'Debug.Print "nr Lasers " & ZEN.gui.Acquisition.Channels.Track.LaserLine.ItemCount
        
    End If
    Recording.Sample0Z = Recording.frameSpacing * (Recording.framesPerStack - 1) / 2
    Me.pixelDwell = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    Me.SamplingNumber = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SamplingNumber
    Me.SamplingMethod = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SamplingMethod
    Me.SamplingMode = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SamplingMode
    Me.BitDepth = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).DataChannelObjectByIndex(0, 1).BitsPerSample
    Me.timeToAcquire = 0
    ReDim pAcquireTrack(TrackNumber - 1)
    For iTrack = 0 To TrackNumber - 1
        pAcquireTrack(iTrack) = Recording.TrackObjectByMultiplexOrder(iTrack, 1).Acquire
    Next iTrack
    UseRoi = AcquisitionParameters.UseROIs
    setRoiFromZen
    SetJob = True
    If Recording.TileAcquisition Then
        If Recording.TileAcquisitionStitchingThreshold < 1 Then
            MsgBox "Online image analysis may not work properly if Online Stitching Threshold  < 1.  ", VbExclamation, "Loading Imaging Job"
        End If
    End If
   On Error GoTo 0
   Exit Function

SetJob_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SetJob of Class Module AJob at line " & Erl & " "

End Function

''''
' putJob put settings from Me.recording into ZEN
''''
Public Function PutJob(ZEN As Object, Optional Extensive As Boolean)
    On Error GoTo ErrorHandle:
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
    Dim AcquisitionParameters As AimAcquisitionController40.AimAcquisitionParameters
    Set AcquisitionParameters = AcquisitionController.AcquisitionParameters
    Dim iTrack As Integer
    Dim iLaser As Integer
    Dim Time As Double
    
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
    Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).TimeBetweenStacks = TimeBetweenStacks
    Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = Me.pixelDwell
    If Extensive Then
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SamplingNumber = Me.SamplingNumber
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SamplingMode = Me.SamplingMode
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SamplingMethod = Me.SamplingMethod
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).DataChannelObjectByIndex(0, 1).BitsPerSample = Me.BitDepth
    End If
        
    For iTrack = 0 To TrackNumber - 1
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, 1).Acquire = pAcquireTrack(iTrack)
    Next iTrack
    Debug.Print "Times set tracks " & Round(Timer - Time, 3)
    
    AcquisitionParameters.UseROIs = UseRoi
    Time = Timer
 
'   the correction is not saved in the DsRecording
    If ZenV > 2010 Then
         If Recording.ScanDirection Then
             If ZEN.GUI.Acquisition.AcquisitionMode.Bidirectional.CorrX.value <> CorrX Then
                ZEN.GUI.Acquisition.AcquisitionMode.Bidirectional.CorrX.value = CorrX
             End If
             If ZEN.GUI.Acquisition.AcquisitionMode.Bidirectional.CorrY.value <> CorrY Then
                 ZEN.GUI.Acquisition.AcquisitionMode.Bidirectional.CorrY.value = CorrY
             End If
         End If
     End If
'    Debug.Print "Time update Corr.X " & Round(Timer - Time, 3)
    createRois
    PutJob = True
    Exit Function
ErrorHandle:
    MsgBox "Error in AJob.putJob " + Name + " " + Err.Description
End Function

Public Sub setAcquireTrack(iTrack As Integer, value As Boolean)
    If UBound(pAcquireTrack) >= iTrack Then
        pAcquireTrack(iTrack) = value
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

Public Function getAcquireTrack(iTrack As Integer) As Boolean
    If UBound(pAcquireTrack) >= iTrack Then
        getAcquireTrack = pAcquireTrack(iTrack)
    End If
End Function

Public Sub ReDimAcquireTrack(value As Integer)
    ReDim Preserve pAcquireTrack(value)
End Sub
 

Public Function GetRecording() As DsRecording
    Set GetRecording = Recording
End Function

Public Sub SetRecording(value As DsRecording)
   Set Recording = Lsm5.CreateBackupRecording
   Recording.Copy value
   TrackNumber = Recording.GetNormalTrackCount
End Sub

''
' returns size of image in pixels base 0
''''
Public Function imageSizePx() As Vector
    Dim TX As Integer
    Dim TY As Integer
    TX = 1
    TY = 1
    If Recording.TileAcquisition Then
        TX = Recording.TilesX
        TY = Recording.TilesY
    End If
    imageSizePx.X = Round(Recording.SamplesPerLine * (TX - (TX - 1) * Recording.TileAcquisitionOverlap))
    
    If Recording.ScanMode = "ZScan" Then
        imageSizePx.Y = 1
    Else
        imageSizePx.Y = Round(Recording.LinesPerFrame * (TY - (TY - 1) * Recording.TileAcquisitionOverlap))
    End If
    
    If isZStack Then
        imageSizePx.Z = Recording.framesPerStack
    Else
        imageSizePx.Z = 1
    End If
End Function


''
' returns size of image in um
''''
Public Function imageSize() As Vector
    Dim imgSizePx As Vector
    imgSizePx = imageSizePx
    imageSize.X = imgSizePx.X * Recording.SampleSpacing
    imageSize.Y = imgSizePx.Y * Recording.SampleSpacing
    imageSize.Z = imgSizePx.Z * Recording.frameSpacing
End Function

''
' returns position of central point for a Job (base 0)
''''
Public Function getCentralPointPx() As Vector
    Dim imgSizePx As Vector
    imgSizePx = imageSizePx
    getCentralPointPx.X = (imgSizePx.X - 1) / 2
    getCentralPointPx.Y = (imgSizePx.Y - 1) / 2
    getCentralPointPx.Z = (imgSizePx.Z - 1) / 2
End Function

Public Function isBleaching() As Boolean
    'On Error GoTo ExitFun
    Dim Success As Integer
    isBleaching = Recording.TrackObjectBleach(Success).Acquire And Recording.TrackObjectBleach(Success).UseBleachParameters
ExitFun:
End Function

Public Function isZStack() As Boolean
    Dim ScanMode As String
    ScanMode = Recording.ScanMode
    If ScanMode = "ZScan" Or ScanMode = "Stack" Then
        isZStack = True
    Else
        isZStack = False
    End If
End Function

Public Function isLineScan() As Boolean
    Dim ScanMode As String
    ScanMode = Recording.ScanMode
    If ScanMode = "ZScan" Or ScanMode = "Line" Then
        isLineScan = True
    Else
        isLineScan = False
    End If
End Function



Public Sub setRoiMaxNr(roiMaxNr As Integer)
    Dim i As Integer
    pRoiMaxNr = roiMaxNr
    If pRoiMaxNr = 0 Then
        Erase Rois
    Else
        ReDim Rois(0 To pRoiMaxNr - 1)
        For i = 0 To pRoiMaxNr - 1
            Set Rois(i) = New roi
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
Public Sub setRoi(ARoi As roi, roiNr As Integer)
    If UBound(Rois) <= roiNr Then
        Rois(roiNr).setRoi ARoi.roiType, ARoi.roiAim, ARoi.getX, ARoi.getY
    Else
        MsgBox (" No Roi of nr" + roiNr + " initialize Rois first ")
    End If
End Sub

''
' load Rois into class
''
Public Sub setRois(RoisIn() As roi)
    If isArrayEmpty(RoisIn) Then
         LogManager.UpdateErrorLog "AJob.setRois tried to initialize Rois with empty vector"
    Else
        Rois = RoisIn
        pRoiMaxNr = UBound(Rois) + 1
    End If
End Sub

'''
' getRoi: returns a Roi
'''
Public Function getRoi(roiNr As Integer) As roi
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
    Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions
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
    Dim ARoi As roi
    Dim i As Integer
    ClearVectorElements
    For i = 0 To pRoiMaxNr - 1
        If Not Rois(i).MakeVectorElement Then
            Exit Function
        End If
    Next i
    createRois = True
End Function



''''
' Create a string with the description of a Job
''''
Public Function jobDescriptor() As String
    Dim ABoolean As Boolean
    Dim Description As String
    Dim iTrack As Integer
    Dim ScanMode As String
    ScanMode = Recording.ScanMode
    If isLineScan Then
        'Description = Description & "ScanMode = Line" & vbCrLf
        Description = Description & "X = " & Recording.SamplesPerLine & ", Y = 1" & vbCrLf
        Description = Description & "LineStep = " & Recording.LineStepNumber & vbCrLf
    End If
    If ScanMode = "Frame" Or ScanMode = "Stack" Or ScanMode = "Plane" Then
        'Description = Description & "ScanMode = Frame" & vbCrLf
        Description = Description & "X = " & Recording.SamplesPerLine & ", Y = " & Recording.LinesPerFrame & vbCrLf
        Description = Description & "LineStep = " & Recording.LineStepNumber & vbCrLf
    End If
    If ScanMode = "Point" Then
        Description = Description & "X =  1, Y = 1 " & vbCrLf
        Description = Description & "LineStep = " & Recording.LineStepNumber & vbCrLf
    End If
    If isBleaching Then
        Description = Description & "Bleaching = " & isBleaching & vbCrLf
    End If
    If Me.UseRoi Then
        Description = Description & "Regions = " & Me.UseRoi & vbCrLf
    End If
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
    End If
    
    If isZStack Then
        Description = Description & "Z-Slices    = " & Recording.framesPerStack & vbCrLf
        Description = Description & "Z-Interval  = " & Round(Recording.frameSpacing, 2) & " um" & vbCrLf
        Description = Description & "Z-Range     = " & Round(Recording.frameSpacing, 2) * (Recording.framesPerStack - 1) & " um" & vbCrLf
    Else
        Description = Description & "Z-Slices = " & 1 & vbCrLf
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

'''
' split a String into two different Strigs depending on their maximal length
''''
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
' create a String to save of the "complete" job desription
''
Public Function jobDescriptorSettings() As String
    Dim Descriptor As String
    Descriptor = "JobName " & Me.Name & vbCrLf
    Descriptor = Descriptor & "ScanMode " & Recording.ScanMode & vbCrLf
    Descriptor = Descriptor & "SamplesPerLine " & Recording.SamplesPerLine & vbCrLf ' FrameSizeX
    Descriptor = Descriptor & "LinesPerFrame " & Recording.LinesPerFrame & vbCrLf   ' FrameSizeY
    Descriptor = Descriptor & "LineStepNumber " & Recording.LineStepNumber & vbCrLf
    Descriptor = Descriptor & "pixelDwell " & Me.pixelDwell & vbCrLf
    Descriptor = Descriptor & "SamplingNumber " & Me.SamplingNumber & vbCrLf
    Descriptor = Descriptor & "SamplingMode " & Me.SamplingMode & vbCrLf
    Descriptor = Descriptor & "SamplingMethod " & Me.SamplingMethod & vbCrLf
    Descriptor = Descriptor & "BitDepth " & Me.BitDepth & vbCrLf
    Descriptor = Descriptor & "ScanDirection " & Recording.ScanDirection & vbCrLf
    Descriptor = Descriptor & "CorrX " & Me.CorrX & vbCrLf
    Descriptor = Descriptor & "CorrY " & Me.CorrY & vbCrLf
    Descriptor = Descriptor & "FramesPerStack " & Recording.framesPerStack & vbCrLf
    Descriptor = Descriptor & "FrameSpacing " & Recording.frameSpacing & vbCrLf
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
Public Function changeJobFromDescriptor(Descriptor As String, value As Variant)
    Dim Success As Integer
    If Recording Is Nothing Then
        MsgBox "AJob.changeJobFromDescriptor Recording has not been initiated!"
    End If
    
    Select Case Descriptor
        Case "JobName":
            Name = CStr(value)
        Case "pixelDwell":
            pixelDwell = CDbl(value)
        Case "ScanMode":
            Recording.ScanMode = CStr(value)
        Case "SamplesPerLine":
            Recording.SamplesPerLine = CLng(value)
        Case "LinesPerFrame":
            Recording.LinesPerFrame = CLng(value)
        Case "LineStepNumber":
            Recording.LineStepNumber = CLng(value)
        Case "SamplingNumber":
            Debug.Print value
            Me.SamplingNumber = CLng(value)
        Case "SamplingMode":
            Me.SamplingMode = CLng(value)
        Case "SamplingMethod":
            Me.SamplingMethod = CLng(value)
        Case "BitDepth":
            Me.BitDepth = CLng(value)
        Case "ScanDirection":
            Recording.ScanDirection = CLng(value)
        Case "CorrX":
            Me.CorrX = CDbl(value)
        Case "CorrY":
            Me.CorrY = CDbl(value)
        Case "FramesPerStack":
            Recording.framesPerStack = CLng(value)
        Case "FrameSpacing":
            Recording.frameSpacing = CDbl(value)
        Case "Bleaching":
            If CBool(value) Then
                Recording.TrackObjectBleach(Success).Acquire = CBool(value)
                Recording.TrackObjectBleach(Success).UseBleachParameters = CBool(value)
            End If
        Case "UseRoi":
            Me.UseRoi = CBool(value)
        Case "Zoom":
            Recording.ZoomX = CDbl(value)
            Recording.ZoomY = CDbl(value)
        Case "SpecialScanMode":
            Recording.SpecialScanMode = CStr(value)
        Case "TimeSeries":
            Recording.TimeSeries = CBool(value)
        Case "StacksPerRecord":
            Recording.StacksPerRecord = CLng(value)
        Case "TimeBetweenStacks":
            Me.TimeBetweenStacks = CDbl(value)
        Case "EndJobDef":
            
        Case Else:
            MsgBox " No Item " & Descriptor
            Exit Function
    End Select
    changeJobFromDescriptor = True
End Function
