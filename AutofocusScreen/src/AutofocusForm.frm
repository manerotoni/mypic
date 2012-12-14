VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutofocusForm 
   Caption         =   "AutofocusScreen for ZEN"
   ClientHeight    =   13530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   OleObjectBlob   =   "AutofocusForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "AutofocusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force to declare all variables

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Version Description''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AutofocusScreen_ZEN_v2.0.1
'''''''''''''''''''''End: Version Description'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const Version = " v2.0.1"







Private Sub CommandButton1_Click()

End Sub





Private Sub SliderZStepLabel_Click()

End Sub

''''''
' UserForm_Initialize()
'   Function called from e.g. AutoFocusForm.Show
'   Load and initialize form
''
Private Sub UserForm_Initialize()

    Me.Caption = Me.Caption + Version
    FormatUserForm (Me.Caption) ' make minimizing button available
    Re_Start                    ' Initialize some of the variables
    
End Sub


''''
' Re_Start()
' Initializations that need to be performed only at the first start of the Macro
''''
Private Sub Re_Start()
    Dim delay As Single
    Dim standType As String
    Dim count As Long
    Dim ImageDatabase As DsGuidedModeDatabase
    Dim i As Long
    Dim MruList As DsMruList
    Dim cnt As Long
    Dim lpReOpenBuff As OFSTRUCT
    Dim wStyle As Long
    Dim lpRootPathName As String
    Dim lpSectorsPerCluster As Long
    Dim lpBytesPerSector As Long
    Dim lpNumberOfFreeClusters As Long
    Dim lpTotalNumberOfClusters As Long
    Dim lSpace As Long
    Dim lFreeSpace As Double
    Dim fSize As Double
    Dim hFile As Long
    Dim bLSM As Boolean
    Dim bLIVE As Boolean
    Dim bCamera As Boolean

    
    Set tools = Lsm5.tools
    GlobalMacroKey = "Autofocus"
    
    flgUserChange = True
    delay = 1
    flgEvent = 7
    flg = 0
    Lsm5.StopScan
    Wait (delay)
    TimerUnit = 1
    CommandTimeSec.BackColor = &HFF8080
    BlockRepetitions = 1
    ReDim Preserve GlobalImageIndex(BlockRepetitions)
    ScanLineToggle.Value = True
    LocationTextLabel.Caption = ""
    GlobalProject = "AutofocusScreen2.1"
    GlobalProjectName = GlobalProject + ".lvb"
    HelpNamePDF = "AutofocusScreen_help.pdf"
    UsedDevices40 bLSM, bLIVE, bCamera
    
    'Set standard values for Autofocus
    ' blSM is a flag to decide weather systen is LSM (ZEN is LSM for instance). LIVE is 5Live not anymore in use?
    If bLSM Then
        SystemName = "LSM"
        CheckBoxHighSpeed.Value = True
        BSliderFrameSize.Min = 16
        BSliderFrameSize.Max = 1024
        BSliderFrameSize.Step = 8
        BSliderFrameSize.StepSmall = 4
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
        DoEvents
    ElseIf bLIVE Then
        SystemName = "LIVE"
        BSliderFrameSize.Min = 128
        BSliderFrameSize.Max = 1024
        BSliderFrameSize.Step = 128
        BSliderFrameSize.StepSmall = 128
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
        DoEvents
    ElseIf bCamera Then
        SystemName = "Camera"
    End If
    
    'Check if GUI is available (ZEN2011 onward)

    
    ScanLineToggle.Value = True
    BSliderZOffset.Value = 0
    BSliderZRange.Value = 80
    BSliderZStep.Value = 0.1
    CheckBoxLowZoom.Value = False
    CheckBoxActiveAutofocus.Value = True
    
    'Set standard values for Post-Acquisition tracking
    TrackingToggle.Value = False
    SwitchEnableTrackingToggle (False)
 
    
    'Set standard values for Looping
    BSliderRepetitions = 300
    BSliderTime = 1
    
    'Set standard values for Micropilot
    CheckBoxActiveOnlineImageAnalysis.Value = False
    SwitchEnableOnlineImageAnalysisPage (False)
    CheckBoxZoomAutofocus.Value = False
    SwitchEnableZoomAutofocus (False)
    
    'Set standard values for Gridscan
    CheckBoxActiveGridScan.Value = False
    SwitchEnableGridScanPage (False)
    
    'Set standard values for Additional Acquisition
    CheckBoxAlterImage.Value = False
    SwitschEnableAlterImagePage (False)
    
    'Set Database name
    DatabaseTextbox.Value = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="OutputFolder")
    
    'Set repetition and locations
    RepetitionNumber = 1
    locationNumber = 1
    Re_Initialize
    
 
End Sub

''''''
'   CheckBoxActiveAutofocus_Click()
'       Activates Autofocus. If not toggled only Acquisition track is used
'''''
Private Sub CheckBoxActiveAutofocus_Click()
                                                    
    If CheckBoxActiveAutofocus.Value Then
        SwitchEnableAutofocusPage (True)
        CheckBoxTrackZ.Value = False
    Else
        SwitchEnableAutofocusPage (False)
    End If
    
End Sub

''''''
'   SwitchEnableAutofocusPage(Enable As Boolean)
'   Disable or enable all buttons and slider
'       [Enable] In - Sets the mini page enable status
''''''
Private Sub SwitchEnableAutofocusPage(Enable As Boolean)

    CheckBoxHighSpeed.Enabled = Enable
    CheckBoxHRZ.Enabled = Enable
    CheckBoxLowZoom.Enabled = Enable
    AFeveryNth.Enabled = Enable
    AFeveryNthLabel.Enabled = Enable
    AFeveryNthLabel2.Enabled = Enable
    ScanLineToggle.Enabled = Enable
    ScanFrameToggle.Enabled = Enable
    FrameSizeLabel.Enabled = Enable
    BSliderFrameSize.Enabled = Enable
    BSliderZOffset.Enabled = Enable
    SliderZOffsetLabel.Enabled = Enable
    BSliderZRange.Enabled = Enable
    SliderZRangeLabel.Enabled = Enable
    BSliderZStep.Enabled = Enable
    SliderZStepLabel.Enabled = Enable
    OptionButtonTrack1.Enabled = Enable
    OptionButtonTrack2.Enabled = Enable
    OptionButtonTrack3.Enabled = Enable
    OptionButtonTrack4.Enabled = Enable

End Sub

''''''
'   CheckBoxActiveOnlineImageAnalysis_Click()
'   Activate online image analysis, micropilot. Also enable the complete micropilot page
''''''
Private Sub CheckBoxActiveOnlineImageAnalysis_Click()

    SwitchEnableOnlineImageAnalysisPage (CheckBoxActiveOnlineImageAnalysis.Value)
    
End Sub

''''''
'   SwitchEnableOnlineImageAnalysisPage(Enable As Boolean)
'   Disable or enable all buttons and slider (aka Micropilot)
'       [Enable] In -  Sets the mini page enable status
''''''
Private Sub SwitchEnableOnlineImageAnalysisPage(Enable As Boolean)
    CheckBoxZoomTrack1.Enabled = Enable
    CheckBoxZoomTrack2.Enabled = Enable
    CheckBoxZoomTrack3.Enabled = Enable
    CheckBoxZoomTrack4.Enabled = Enable
    LabelZoom.Enabled = Enable
    TextBoxZoom.Enabled = Enable
    ZoomNumSlicesLabel.Enabled = Enable
    TextBoxZoomNumSlices.Enabled = Enable
    ZoomIntervalLabel.Enabled = Enable
    TextBoxZoomInterval.Enabled = Enable
    TextBoxZoomAutofocusZOffset.Enabled = Enable
    CheckBoxZoomAutofocus.Enabled = Enable
    ZoomFrameSizeLabel.Enabled = Enable
    TextBoxZoomFrameSize.Enabled = Enable
    ZoomCyclesLabel.Enabled = Enable
    TextBoxZoomCycles.Enabled = Enable
    ZoomCycleDelayLabel.Enabled = Enable
    TextBoxZoomCycleDelay.Enabled = Enable
    SwitchEnableGridScanPage (CheckBoxActiveGridScan.Value)
End Sub

''''''
'   CheckBoxZoomAutofocus_Click()
'   Activate extra autofocus for image analysis. Enable Z-offset box to be viewed
''''''
Private Sub CheckBoxZoomAutofocus_Click()
    
    SwitchEnableZoomAutofocus (CheckBoxZoomAutofocus.Value) 'Show Zoffset only when extra autofocus is clicked
    
End Sub

''''''
'   SwitchEnableZoomAutofocus(Enable As Boolean)
'   Enable/disable Z-offset form for Micropilot minipage
'       [Enable] In - Sets the visibility of box
''''''
Private Sub SwitchEnableZoomAutofocus(Enable As Boolean)

    TextBoxZoomAutofocusZOffset.Visible = Enable
    ZoomAutofocusZOffsetLabel.Visible = Enable
    
End Sub

''''''
'   CheckBoxAlterImage_Click()
'   Activate additional image that is acquired only from time to time
''''''
Private Sub CheckBoxAlterImage_Click()

    SwitschEnableAlterImagePage (CheckBoxAlterImage.Value)
    
End Sub

''''''
'   SwitschEnableAlterImagePage(Enable As Boolean)
'   Enable/disable Additional acquisition page
'       [Enable] In - Sets the enable Enable of minpage
''''''
Private Sub SwitschEnableAlterImagePage(Enable As Boolean)

    CheckBox2ndTrack1.Enabled = Enable
    CheckBox2ndTrack2.Enabled = Enable
    CheckBox2ndTrack3.Enabled = Enable
    CheckBox2ndTrack4.Enabled = Enable
    AlterZoomLabel.Enabled = Enable
    TextBoxAlterZoom.Enabled = Enable
    AlterNumSlicesLabel.Enabled = Enable
    TextBoxAlterNumSlices.Enabled = Enable
    AlterIntervalLabel.Enabled = Enable
    TextBoxAlterInterval.Enabled = Enable
    RoundAlterTrackLabel.Enabled = Enable
    TextBox_RoundAlterTrack.Enabled = Enable
    AlterTrackDescriptionLabel.Enabled = Enable
    
End Sub

''''
' CheckBoxActiveGridScan_Click()
'   Set the grid scan on or off. Changes also
''
Private Sub CheckBoxActiveGridScan_Click()

    SwitchEnableGridScanPage (CheckBoxActiveGridScan.Value)
    If CheckBoxActiveGridScan.Value Then
        TrackingToggle.Value = False
        MultipleLocationToggle.Value = False
    End If
    
End Sub

''''
'   SwitchEnableGridScanPage(Enable As Boolean)
'   Disable or enable all buttons and slider
'       [Enable] In - Sets the mini page enable status
''''
Private Sub SwitchEnableGridScanPage(Enable As Boolean)

    CheckBoxGridScan_Initialise.Enabled = Enable
    If CheckBoxActiveOnlineImageAnalysis.Value Then
        CheckBoxGridScan_FindGoodPositions.Enabled = Enable
    Else
        CheckBoxGridScan_FindGoodPositions.Enabled = False
    End If
    GridScan_posLabel.Enabled = Enable
    GridScan_nLabel.Enabled = Enable
    GridScan_nXLabel.Enabled = Enable
    GridScan_nYLabel.Enabled = Enable
    GridScan_nX.Enabled = Enable
    GridScan_nY.Enabled = Enable
    GridScan_dLabel.Enabled = Enable
    GridScan_dXLabel.Enabled = Enable
    GridScan_dYLabel.Enabled = Enable
    GridScan_dX.Enabled = Enable
    GridScan_dY.Enabled = Enable
    GridScan_subLabel.Enabled = Enable
    GridScan_nsubLabel.Enabled = Enable
    GridScan_nXsubLabel.Enabled = Enable
    GridScan_nYsubLabel.Enabled = Enable
    GridScan_nXsub.Enabled = Enable
    GridScan_nYsub.Enabled = Enable
    GridScan_dsubLabel.Enabled = Enable
    GridScan_dXsubLabel.Enabled = Enable
    GridScan_dYsubLabel.Enabled = Enable
    GridScan_dXsub.Enabled = Enable
    GridScan_dYsub.Enabled = Enable
    GridScanDescriptionLabel.Enabled = Enable
    
End Sub

'''''
' Where is this used? This should be present only when Frame tracking
' Not active anymore
'''''
Private Sub CheckBoxTrackXY_Click()
    
End Sub

''''''''
'   CommandButtonHelp_Click()
'   Look for Help file
'   TODO: Test
''''''''
Private Sub CommandButtonHelp_Click()

    Dim dblTask As Double
    Dim MacroPath As String
    Dim Mypath As String
    Dim MyPathPDF As String
    
    Dim bslash As String
    Dim Success As Integer
    Dim pos As Integer
    Dim Start As Integer
    Dim count As Long
    Dim ProjName As String
    Dim indx As Integer
    Dim AcrobatObject As Object
    Dim AcrobatViewer As Object
    Dim OK As Boolean
    Dim StrPath As String
    Dim ExecName As String
        
    count = ProjectCount()
    For indx = 0 To count - 1
        MacroPath = ProjectPath(indx, Success)
        ProjName = ProjectTitle(indx, Success)
        If StrComp(ProjName, GlobalProjectName, vbTextCompare) = 0 Then
            Start = 1
            bslash = "\"
            pos = Start
            Do While pos > 0
                pos = InStr(Start, MacroPath, bslash)
                If pos > 0 Then
                    Start = pos + 1
                End If
            Loop
            Mypath = Strings.Left(MacroPath, Start - 1)
            MyPathPDF = Mypath + HelpNamePDF

            OK = False
            On Error GoTo RTFhelp
            OK = FServerFromDescription("AcroExch.Document", StrPath, ExecName)
            dblTask = Shell(ExecName + " " + MyPathPDF, vbNormalFocus)
            
RTFhelp:
            If Not OK Then
                MsgBox "Install Acrobat Viewer!"
            End If
            Exit For
        End If
    Next indx
End Sub

'''''''''
'   StopButton_Click()
'   ScanStop is used to tell different functions to stop execution and acquisition
'   A second routine is called to stop the processes
'       [ScanStop] Global/Out - Set to true
'''''''
Private Sub StopButton_Click()
    ScanStop = True
End Sub

''''''''
'   StopAcquisition()
'   Stop scan and reset buttons of the form
''''''''
Public Sub StopAcquisition()
    If ScanStop Then
        Lsm5.StopScan
        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
        RestoreAcquisitionParameters
        Sleep (1000)
        DisplayProgress "Restore Settings", RGB(&HC0, 0, 0)
        DoEvents
        ' reset the buttons
        Dim FileName As String
        Running = False
        ScanStop = False
        ScanPause = False
        PauseButton.Caption = "Pause"
        PauseButton.BackColor = &H8000000F
        ExtraBleach = False
        ExtraBleachButton.Caption = "Bleach"
        ExtraBleachButton.BackColor = &H8000000F
        ReDim BleachTable(BlockRepetitions)
        ReDim BleachStartTable(BlockRepetitions)
        ReDim BleachStopTable(BlockRepetitions)
        BleachingActivated = False
        '
    '   If TrackingToggle Or FrameAutofocussing Then
    '   what is this?
    '        For i = 1 To PositionData.Sheets.count
    '            PositionData.Sheets.Item(i).Select
    '            Cells.Select
    '            Selection.Columns.AutoFit
    '        Next i
    '        FileName = Left(DataBaseLabel, Len(DataBaseLabel) - 4) & ".xls"
    '        PositionData.SaveAs FileName:=FileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    '        PositionData.Close
    '        Excel.Application.Quit
    '    End If
        ' TODO: How to check that the paramters has been restored ?
        Sleep (1000)
        DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
        ChangeButtonStatus True
    End If
End Sub

'''''
'   CommandButtonNewDataBase_Click()
'   Assigns output folder where files are stored in ZEN software
'   TODO: Change name of function
'''''
Private Sub CommandButtonNewDataBase_Click()

    GlobalDataBaseName = DatabaseTextbox.Value
    If Not GlobalDataBaseName = "" Then
        DatabaseLable.Caption = GlobalDataBaseName
        SaveSetting "OnlineImageAnalysis", "macro", "OutputFolder", GlobalDataBaseName
    End If
    
End Sub



''''''
'   BleachRegion(XShift As Double, YShift As Double)
'       [XShift] In - Shifts origin of x
'       [YShift] In - Shifts origin of y
'   Todo: function is never been used and does not belong to form or being called. Check it
''''''
Private Sub BleachRegion(XShift As Double, YShift As Double)
    Dim RecordingDoc As DsRecordingDoc
    Dim Recording As DsRecording
    Dim Track As DsTrack
    Dim Laser As DsLaser
    Dim DetectionChannel As DsDetectionChannel
    Dim IlluminationChannel As DsIlluminationChannel
    Dim DataChannel As DsDataChannel
    Dim BeamSplitter As DsBeamSplitter
    Dim Timers As DsTimers
    Dim Markers As DsMarkers
    Dim Success As Integer
    Set Recording = Lsm5.DsRecording
    Dim SampleObservationTime As Double
    Dim SampleOX As Double
    Dim SampleOY As Double
    
    
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
     
    SampleOX = Recording.Sample0X
    SampleOY = Recording.Sample0Y
    Recording.Sample0X = XShift
    Recording.Sample0Y = YShift
    'x = Lsm5.Hardware.CpStages.PositionX - XShift
    'y = Lsm5.Hardware.CpStages.PositionY - YShift
    'Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
    ' maybe wait here till it is finished moving
    Recording.SpecialScanMode = "NoSpecialMode"
    Recording.ScanMode = "Point"
    Recording.TimeSeries = True
    Recording.FramesPerStack = 1
    Recording.StacksPerRecord = 50  ' timepoints x 1000
    SampleObservationTime = Track.SampleObservationTime
    MsgBox "SampleObservationTime = " + CStr(SampleObservationTime)
    Track.SampleObservationTime = 0.0001 ' pixel-dwell time in seconds
    Track.TimeBetweenStacks = 0#
    'Timers.TimeInterval = 0#
    
    TakeImage
    
    Recording.Sample0X = SampleOX
    Recording.Sample0Y = SampleOY
    'x = Lsm5.Hardware.CpStages.PositionX + XShift
    'y = Lsm5.Hardware.CpStages.PositionY + YShift
    'Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
    ' maybe wait here till it is finished moving
    Recording.SpecialScanMode = "NoSpecialMode"
    Recording.ScanMode = "Frame"
    Recording.TimeSeries = False
    Recording.FramesPerStack = 1
    Recording.StacksPerRecord = 1  ' timepoints x 1000
    Track.SampleObservationTime = SampleObservationTime ' pixel-dwell time in seconds
    MsgBox "SampleObservationTime = " + CStr(SampleObservationTime)
    
 
    'Recording.ScanMode = "Plane"
    'Recording.FrameSpacing = 0.636243
       
        
End Sub

Public Function ComputeCenterAndAxis(dX As Double, dY As Double)

    Dim i, j, iFrame, channel, ni, bitDepth As Long
    Dim nj As Long
    
    Dim ic, jc, di, dj, PixelSize As Double
    Dim tot As Double
    
    Dim th As Double
    th = 20
    
    
    'Dim ColMax As Integer
    'Dim iRow As Integer
    'Dim nRow As Integer
    'Dim iFrame As Integer
    'Dim gvRow As Variant  ' gv = gray value
    'Dim iCol As Long
    'Dim nCol As Long
    'Dim bitDepth As Long
    'Dim channel As Integer
    'Dim gvMax As Double
    'Dim gvMaxBitRange As Double
    'Dim nSaturatedPixels As Long
    'Dim maxGV_nSat(2) As Double
    
    
    'DisplayProgress "Measuring Exposure...", RGB(0, &HC0, 0)
  
    'ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
    
    'nRow = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
    'MsgBox "nRow = " + CStr(nRow)
    
'        ElseIf SystemName = "LSM" Then
'            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
'            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
'        Else
'            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
''            Exit Sub
 '       End If
 '   End If
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    'ReDim Intline(nLines) As Long
    
    'iFrame = 0
    'gvMax = -1
        
    'iRow = 0
    'channel = 0
    'bitDepth = 0 ' leaves the internal bit depth
    'gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
    
    
    
    ni = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
    'nCol = 0
    nj = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
    
    'Dim image(,) As Variant
    
    'Dim replyCounts(,,) As Short = New Short(2, 1, 2) {}
    
    Dim srcline As Variant
    
    Dim image() As Long
    ReDim image(ni, nj)
    
    
    'Dim x(1 To ni, 1 To 4) As Variant

    'MsgBox "ni = " + CStr(ni) + " nj = " + CStr(nj)
    
   ' image = GetSubRegion(channel, xs, ys, zs, ts
    
    
    'Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
        
    PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
        
        
    ' get the image  (put into a subprocedure)
    iFrame = 0
    channel = 0
    bitDepth = 0 ' leaves the internal bit depth
    For i = 0 To ni - 1
        srcline = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, i, nj, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
        For j = 0 To nj - 1
            image(i, j) = srcline(j)
        Next j
    Next i
    'MsgBox "im = " + CStr(image(100, 100))
        
    ' computer center of mass
    ic = 0
    jc = 0
    tot = 0
    For i = 0 To ni - 1
        For j = 0 To nj - 1
            If (image(i, j) > th) Then
                ic = ic + image(i, j) * i
                jc = jc + image(i, j) * j
                tot = tot + image(i, j)
            End If
        Next j
    Next i
    
    ic = ic / tot
    jc = jc / tot
    'MsgBox "ic = " + CStr(ic) + " jc = " + CStr(jc) + " tot = " + CStr(tot)
    
    dX = (ic - ni / 2) * PixelSize
    dY = (jc - nj / 2) * PixelSize
    
    ' compute displacement vector
    di = 0
    dj = 0
    
    For i = 0 To ni - 1
        For j = 0 To nj - 1
            If (image(i, j) > th) Then
                di = di + image(i, j) * (i - ic) * Sgn(i - ic)
                dj = dj + image(i, j) * (j - jc) * Sgn(i - ic)
            End If
        Next j
    Next i
    
    di = di / tot
    dj = dj / tot
    'MsgBox "di = " + CStr(di) + " dj = " + CStr(dj) + " tot = " + CStr(tot)
        
        
    'PixelSize
        
        
        
    '    For iCol = 0 To nCol - 1            'Now I'm scanning all the pixels in the line
            
     '       If (gvRow(iCol) > gvMax) Then
      '          gvMax = gvRow(iCol)
       '     End If

    
    
    'iFrame = 0
    'gvMax = -1
    'iRow = 0
    'Channel = 0
    'bitDepth = 0 ' leaves the internal bit depth
    'gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(Channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
    'MsgBox "nCol = " + CStr(nCol)
    'MsgBox "bytes per pixel = " + CStr(bitDepth)

    ' todo: is there another function to find this out??
    'If (bitDepth = 1) Then
    '    gvMaxBitRange = 255
    'ElseIf (bitDepth = 2) Then
    '    gvMaxBitRange = 65536
    'End If
    
    'nSaturatedPixels = 0
 
End Function


Private Sub CommonDialog_Enter()

End Sub

''''''
'   TakeImage()
'   Acquire an image. Allow to stop acquisition and displaqy progress
'''''''
Private Sub TakeImage()

    Dim ScanImage As DsRecordingDoc
    
    Set ScanImage = Lsm5.StartScan

    DisplayProgress "Taking Image.......", RGB(0, &HC0, 0)
    Do While ScanImage.IsBusy ' Waiting until the image acquisition is done
        Sleep (100)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StopAcquisition
                Exit Sub
            End If
        End If
    Loop
    DisplayProgress "Taking Image...DONE.", RGB(0, &HC0, 0)
End Sub




Public Function SetGetLaserPower(power As Double)
    
    Dim Recording As DsRecording
    Dim Track As DsTrack
    Dim IlluminationChannel As DsIlluminationChannel
    
    Set Recording = Lsm5.DsRecording
    Set Track = Recording.TrackObjectByIndex(0, Success)
    Set IlluminationChannel = Track.IlluminationObjectByIndex(0, Success)

    If (power > 0) Then
        IlluminationChannel.power = power
    End If
    
    power = IlluminationChannel.power
       
End Function
 

Public Function MeasureExposure(fractionMax As Double, fractionSat As Double)
   
'    Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
 '   DoEvents
    
    'Dim ColMax As Integer
    Dim iRow As Integer
    Dim nRow As Integer
    Dim iFrame As Integer
    Dim gvRow As Variant  ' gv = gray value
    Dim iCol As Long
    Dim nCol As Long
    Dim bitDepth As Long
    Dim channel As Integer
    Dim gvMax As Double
    Dim gvMaxBitRange As Double
    Dim nSaturatedPixels As Long
    Dim maxGV_nSat(2) As Double
    
    
    DisplayProgress "Measuring Exposure...", RGB(0, &HC0, 0)
  
    'ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
    
    nRow = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
    'MsgBox "nRow = " + CStr(nRow)
    
'        ElseIf SystemName = "LSM" Then
'            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
'            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
'        Else
'            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
''            Exit Sub
 '       End If
 '   End If
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    'ReDim Intline(nLines) As Long
    
    iFrame = 0
    gvMax = -1
        
    iRow = 0
    channel = 0
    bitDepth = 0 ' leaves the internal bit depth
    gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
    'MsgBox "nCol = " + CStr(nCol)
    'MsgBox "bytes per pixel = " + CStr(bitDepth)

    ' todo: is there another function to find this out??
    If (bitDepth = 1) Then
        gvMaxBitRange = 255
    ElseIf (bitDepth = 2) Then
        gvMaxBitRange = 65536
    End If
    
    nSaturatedPixels = 0
    
    For iRow = 0 To nRow - 1
        gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
        For iCol = 0 To nCol - 1            'Now I'm scanning all the pixels in the line
            
            If (gvRow(iCol) > gvMax) Then
                gvMax = gvRow(iCol)
            End If
            
            If (gvRow(iCol) = gvMaxBitRange) Then
                nSaturatedPixels = nSaturatedPixels + 1
                ' TODO: measure neighbouring saturated pixels
            End If

        Next iCol
    Next iRow
        
    fractionMax = gvMax / gvMaxBitRange
    fractionSat = nSaturatedPixels / (nRow * nCol)
        
    'MsgBox "maximal gray value in image = " + CStr(gvMax)
    'MsgBox "fractional brightness of maximal gray value in image = " + CStr(fractionMax)
    'MsgBox "number of saturated pixles = " + CStr(nSaturatedPixels)
    'MsgBox "fraction of saturated pixles = " + CStr(fractionSat)
      
    DisplayProgress "Measuring Exposure...DONE", RGB(0, &HC0, 0)
  
End Function


Private Sub ScanLineToggle_Click()
    ScanFrameToggle.Value = Not ScanLineToggle.Value 'if ScanFrame is true ScanLine is false (you can only chose one of them)
    FrameAutofocussing = ScanFrameToggle.Value 'if ScanFrame is true than FrameAutofocusing (boolean variable) will be set true as well
    FrameSizeLabel.Visible = ScanFrameToggle.Value 'FrameSize Label is only displayed if ScanFrame is activated
    BSliderFrameSize.Visible = ScanFrameToggle.Value 'FrameSize Slider is only displayed if ScanFrame is activated
'    BSliderScanSpeed.Visible = ScanLineToggle.Value
'    ScanSpeedLabel.Visible = ScanLineToggle.Value
End Sub

Private Sub ScanFrameToggle_Click()
    ScanLineToggle.Value = Not ScanFrameToggle.Value 'if ScanLine is chosen, ScanFrame will be unchecked
    
    FrameAutofocussing = ScanFrameToggle.Value 'if ScanFrame is true than FrameAutofocusing (boolean variable) will be set true
    FrameSizeLabel.Visible = ScanFrameToggle.Value
    BSliderFrameSize.Visible = ScanFrameToggle.Value
    CheckBoxAutofocusTrackXY.Visible = ScanFrameToggle.Value

'    ScanSpeedLabel.Visible = ScanLineToggle.Value

'         If SystemName = "LSM" Then
'
'            BSliderFrameSize.ValueEditable = True
'             BSliderFrameSize.Min = 16
'            BSliderFrameSize.Max = 1024
'            BSliderFrameSize.Step = 128
'            BSliderFrameSize.StepSmall = 4
'           BSliderFrameSize.ValueDisplay = True
'
'        ElseIf SystemName = "LIVE" Then
'
'
'            BSliderFrameSize.ValueEditable = False
'            BSliderFrameSize.Min = 128
'            BSliderFrameSize.Max = 1024
'            BSliderFrameSize.Step = 128
'            BSliderFrameSize.StepSmall = 128
'            BSliderFrameSize.Value = 128
'
'        End If
    
End Sub

Private Sub ScanSpeedLabel_Click()

End Sub

''''''
'   GetCurrentPositionOffsetButton_Click()
'       Read autofocus parameters BlockZRange, BlockZStep....
'       Performs the scan in Z (line or Frame), to find the offset value according to actual position
''''''
Private Sub GetCurrentPositionOffsetButton_Click()
    
    AutofocusForm.GetBlockValues  ' Update parameter                                 'Updates the parameters value for BlockZRange, BlockZStep....
    GetCurrentPositionOffset BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset   ' Performs scan  in Z (line or Frame, to find the offset value
    
End Sub

'''''''
'   AutofocusButton_Click()
'   Perform Autofocus if Track is selected. If LineScan only Z is changed
'   if FrameScan X and Y are changed
'   TODO: No Check that ZShift make sense (original CheckRefControl has been removed because not properly working)
'
''''''''
Private Sub AutofocusButton_Click()
    
    Dim AutofocusDoc As DsRecordingDoc
    Dim Success As Boolean
    Try = 1
    AutofocusForm.GetBlockValues 'Updates the parameters value for BlockZRange, BlockZStep..
    
    DisplayProgress "Autofocus 0", RGB(0, &HC0, 0)
    StopScanCheck
    StoreAcquisitionParameters 'stores Parameters in GlobalBackupRecording and BackupRecording
    
    'Acquire image and calculate center of mass. This is stored in ZShift, (XShift and YShift)
    Success = Autofocus_StackShift(BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset, AutofocusDoc)
    
    If Not Success Then
        StopAcquisition
        Exit Sub
    End If
    
    ' fine focus is done with focuswheel of microscope
    Autofocus_MoveAcquisition BlockZOffset
    
    If ScanStop Then
        StopAcquisition
        Exit Sub
    End If
    
    ActivateAcquisitionTrack
    If IsAcquisitionTrackSelected And IsAutofocusTrackSelected Then 'TODO why both conditions
        ScanToImageNew AutofocusDoc
    End If
    
    DisplayProgress "AF: Taking image at found position...", RGB(0, &HC0, 0)
    While AcquisitionController.IsGrabbing
        Sleep (100)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StopAcquisition
                Exit Sub
            End If
        End If
    Wend
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
End Sub


Private Sub StartBleachButton_Click()
    
    Dim Success As Integer
    Dim nt As Integer
    
    BleachingActivated = True
    AutomaticBleaching = False
    
    If TrackingToggle And TrackingChannelString = "" Then
        MsgBox ("Select a channel for tracking, or uncheck the tracking button")
        Exit Sub
    End If
    If MultipleLocationToggle.Value And Lsm5.Hardware.CpStages.Markcount < 1 Then
        MsgBox ("Select at least one location in the stage control window, or uncheck the multiple location button")
        Exit Sub
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No Output Folder selected ! Cannot start acquisition.")
        Exit Sub
    End If
    
    
    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
    
    If Success Then
        If Track.BleachPositionZ <> 0 Then
            MsgBox ("This macro does not enable to bleach at a different Z position. Please uncheck the corresponding box in the Bleach Control Window")
            Exit Sub
        End If
        
        If Lsm5.IsValidBleachRoi Then
            
            If CheckBoxActiveOnlineImageAnalysis Then
                nt = TextBoxZoomCycles
            Else
                nt = BlockRepetitions
            End If
                    
            If (Track.BleachScanNumber + 1) > nt Then
                MsgBox ("Not enough repetitions to bleach; either increase the Number of Acquisitions, or, when using MicroPilot, the Cycles")
                Exit Sub
            End If
            
            FillBleachTable
            AutomaticBleaching = True
           'Track.UseBleachParameters = True ' deleted 20100818 , can probably not work with Micropilot
        Else
            MsgBox ("A bleaching ROI needs to be defined to start the macro in the bleaching mode")
            Exit Sub
        End If
    Else
        MsgBox ("A bleach track needs to be defined to start the macro in the bleaching mode")
        Exit Sub
    End If
    
    StoreAcquisitionParameters
    
    StartAcquisition BleachingActivated
End Sub

Private Sub FillBleachTable()  'Fills a table for the macro to know when the bleaches have to occur. This works for FRAPs (and FLIPS if working with LSM 3.2)
    
    Dim i As Integer
    Dim nt As Integer
    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
    If Success Then
        
        If CheckBoxActiveOnlineImageAnalysis Then
            nt = TextBoxZoomCycles.Value
        Else
            nt = BlockRepetitions
        End If
            
        ReDim BleachTable(nt)               'The bleach table contains as many timepoints as blockrepetitions
        
        'When working with the Lsm 2.8, remove all this test, except the one indicated line
        If Track.EnableBleachRepeat Then
            For i = Track.BleachScanNumber + 1 To nt Step Track.BleachRepeat
                BleachTable(i) = True
            Next
        Else
        '    BleachTable(Track.BleachScanNumber + 1) = True  'This is the only line to be kept when working with the Lsm 2.8
        End If
        
    End If
End Sub


'''''
'   StartButton_Click()
'''''
Private Sub StartButton_Click()
    Dim Success As Boolean
    Success = StartSetting
    If Not Success Then
        ScanStop = True
        StopAcquisition
        Exit Sub
    End If
    
    'Set counters back to 1
    locationNumber = 1    ' first location
    RepetitionNumber = 1 ' first time point
    
    StartAcquisition BleachingActivated 'This is the main function of the macro
End Sub


Private Sub ContinueFromCurrentLocation_Click()
    Dim Success As Boolean
    Success = StartSetting
    If Not Success Then
        ScanStop = True
        StopAcquisition
        Exit Sub
    End If
    StartAcquisition BleachingActivated 'This is the main function of the macro
End Sub

Private Function StartSetting() As Boolean
    StartSetting = False
    BleachingActivated = False
    AutomaticBleaching = False                                  'We do not do FRAps or FLIPS in this case. Bleaches can still be done with the "ExtraBleach" button.
    If TrackingToggle And TrackingChannelString = "" Then
        MsgBox ("Select a channel for tracking, or uncheck the tracking button")
    
        Exit Function
    End If
    If MultipleLocationToggle.Value And Lsm5.Hardware.CpStages.Markcount < 1 Then
        MsgBox ("Select at least one location in the stage control window, or uncheck the multiple location button")
        Exit Function
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start acquisition.")
        Exit Function
    End If
    
    StoreAcquisitionParameters
    
    'As default we do not overwrite files
    OverwriteFiles = False
    
    ' load starting position from stage for GridScan
    If CheckBoxActiveGridScan Then
        If Lsm5.Hardware.CpStages.Markcount = 0 Then  ' No marked position
            MsgBox " GridScan: Use stage to Mark at the initial position "
            ScanStop = True
            StopAcquisition
            Exit Function
        End If
        ' Store starting position for later restart. This is the first marked point
        Lsm5.Hardware.CpStages.MarkGetZ 0, XStart, YStart, ZStart
    End If
       
    ' fill positions for MultipleLocations
    If MultipleLocationToggle Then
        Dim i As Integer
        If Lsm5.Hardware.CpStages.Markcount > 0 Then
            ReDim posMultiLocationX(1 To Lsm5.Hardware.CpStages.Markcount)
            ReDim posMultiLocationY(1 To Lsm5.Hardware.CpStages.Markcount)
            ReDim posMultiLocationZ(1 To Lsm5.Hardware.CpStages.Markcount)
            For i = 1 To Lsm5.Hardware.CpStages.Markcount
                Lsm5.Hardware.CpStages.MarkGetZ i - 1, posMultiLocationX(i), posMultiLocationY(i), _
                posMultiLocationZ(i)
            Next i
        End If
    End If
    StartSetting = True
End Function



''''''
'   StartAcquisition(BleachingActivated)
'   Perform many things (TODO: write more). Pretty much the whole macro runs through here
''''''
Private Sub StartAcquisition(BleachingActivated)
    'measure time required
    Dim rettime, difftime As Double
    Dim GlobalPrvTime As Double
    Dim StartTime As Double
    
    'Counters
    Dim Location As Long         ' Location counter
    Dim iLoc As Integer          ' second location counter (could eventually be removed)
    Dim MaxNrLocations As Long   ' Maximal number of locations
    Dim iPosition As Long        ' id of Well/position
    Dim iPositionMax As Long     ' Maximal number of Well/position
    Dim iSubposition As Long     ' id of subposition
    Dim iSubpositionMax As Long  ' Maximal number of subpositions per Position
    Dim HighResExperimentCounter As Integer
    Dim HighResCounter As Integer

    'Coordinates
    Dim x As Double              ' x value where to move the stage
    Dim y As Double              ' y value where to move the stage
    Dim z As Double              ' z value where to move the stage
    Dim XCor As Double           ' Shift in X calculated from Autofocus
    Dim YCor As Double           ' Shift in Y calculated from autofocus
    Dim ZCor As Double           ' Shift in Z calculated from autofocus
    
    'test variables
    Dim Success As Integer       ' Check if something was sucessfull
    Dim SuccessAF As Boolean     ' Check if AF was succesful
    Dim LocationSoFarBest As Integer
    Dim soFarBestGoodCellsPerImage As Integer
    
    'Recording stuff
    Dim FileNameId As String ' ID name of file (Well/Position, Subpositio, Timepoint)
    Dim filepath As String   ' full path of file to save (changes through function)
    Dim RecordingDoc As DsRecordingDoc  ' contains the images
    Dim Scancontroller As AimScanController ' the controller
  
    
    ' Set the offset in z-stack to 0; otherwise there can be errors...
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * Int(Lsm5.DsRecording.FramesPerStack / 2)
                       
    ' Store current settings
    CopyRecording BackupRecording, Lsm5.DsRecording
    
    
    ' set up the imaging
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
    ' set up RecordingDoc
    If RecordingDoc Is Nothing Then
        Set RecordingDoc = Lsm5.NewScanWindow
        While RecordingDoc.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
     
    ' CheckBoxActiveOnlineImageAnalysis  refers to the MicroPilot
    If CheckBoxActiveOnlineImageAnalysis Then
        
        Dim HighResArrayX() As Double ' this is an array of values why do you need to store values?
        Dim HighResArrayY() As Double
        Dim HighResArrayZ() As Double
        ReDim Preserve HighResArrayX(100) 'define 100 a priori (even if there are less)
        ReDim Preserve HighResArrayY(100)
        ReDim Preserve HighResArrayZ(100)
        HighResExperimentCounter = 0
        HighResCounter = 0
        SaveSetting "OnlineImageAnalysis", "macro", "code", 0
        SaveSetting "OnlineImageAnalysis", "macro", "offsetx", 0
        SaveSetting "OnlineImageAnalysis", "macro", "offsety", 0
        
    End If
    
    
    InitializeStageProperties
    SetStageSpeed 9, True
        
    
            
    Running = True  'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.
    ChangeButtonStatus False ' disable buttons
    MaxNrLocations = 1  'If using the single location you do not have to mark it in the stage control window.
    

    If MultipleLocationToggle.Value Then                    'Defines the Location Number parameter
        MaxNrLocations = Lsm5.Hardware.CpStages.Markcount       'Counts the locations stored in the Stage control window from the LSM
    End If
    
    '''''''''''''''''''''''
    '***Set up GridScan***'
    '''''''''''''''''''''''
    If CheckBoxActiveGridScan Then
        MaxNrLocations = GridScan_nX.Value * GridScan_nY.Value * GridScan_nXsub.Value * GridScan_nYsub.Value
        If MaxNrLocations > 10000 Then
            MsgBox "GridScan: Maximal number of locations is 10000. Please change Numbers  X and/or Y."
            ScanStop = True
            StopAcquisition
            Exit Sub
        End If
    End If
    
    If CheckBoxActiveGridScan Then
    
        Dim GridInit As Boolean 'initialize grid
        If CheckBoxGridScan_Initialise Then 'forced initialization
            GridInit = True
        ElseIf isArrayEmpty(posGridX) Then  'when empty grid
            GridInit = True
        ElseIf UBound(posGridX) < MaxNrLocations Then 'when change in number of grid points
            GridInit = True
        Else
            GridInit = False
        End If
        
        If GridInit Then
            ReDim posGridX(1 To MaxNrLocations)
            ReDim posGridY(1 To MaxNrLocations)
            ReDim posGridXY_valid(1 To MaxNrLocations)
            ReDim locationNumbersMainGrid(1 To MaxNrLocations)
            DisplayProgress "Initialize all grid positions....", RGB(0, &HC0, 0)
            Sleep (1000)
            MakeGrid posGridX, posGridY, posGridXY_valid, locationNumbersMainGrid
            DisplayProgress "Initialize all grid positions...DONE", RGB(0, &HC0, 0)
        End If
        
    End If
    '''''''''''''''''''''''''''
    '***End Set up GridScan***'
    '''''''''''''''''''''''''''

            
    If TrackingToggle Or FrameAutofocussing Then
        'Here you could add code for storing the XYZ position of the cells at each time point in Excel
        'code is in "unused code" ExcelXYZstoring
    End If
    
    
    Do While Running   'As long as the macro is running we're in this loop. At everystop one will save actual location, and repetition

        ' Todo: what is happening here?
        ' Todo: remember the last focus position for each location! (this automatically would create a ZMap)
        ' Tischi: i commented the following lines, because the z-positions for multiple location are updated already within the location loop..
        ' ..so i do not understand what is happening here
        'If Not (TrackingToggle Or FrameAutofocussing) Then
        '   UpdateZvalues Grid, MultipleLocationToggle.Value, z ' cleaned 2010.07.15
        'End If
        
        nGoodCellsPerWell = 0
        iPosition = 1 ' not consisten with name used before
        iSubposition = 1  ' this is the local position according to submask
        ' start counting how long it takes

        iPositionMax = MaxNrLocations / (GridScan_nXsub.Value * GridScan_nYsub.Value)
        iSubpositionMax = GridScan_nXsub.Value * GridScan_nYsub.Value
        GlobalPrvTime = CDbl(GetTickCount) * 0.001
        
        For Location = locationNumber To MaxNrLocations    'This loops all the locations (only one if Single location is selected)
                              
            ''''''''Start stage movement to a different position
            If MultipleLocationToggle.Value Then
                Success = FailSafeMoveStage(posMultiLocationX(Location), posMultiLocationY(Location), posMultiLocationZ(Location))
                LocationTextLabel.Caption = "Now at X= " & posMultiLocationX(Location) & ", Y = " & posMultiLocationY(Location) & ", Z = " & posMultiLocationZ(Location)
                iPosition = Location
            End If
            
            If CheckBoxActiveGridScan.Value Then
                ' TODO: check for good cells. The check is done in Micropilot but afterwards. Default? However this should be done in the workflow manager
                If CheckBoxGridScan_FindGoodPositions And (Location > 1) Then
                    If nGoodCellsPerWell >= minGoodCellsPerWell Then
                        MsgBox "Enough Cells Per Well " + CStr(nGoodCellsPerWell) + "/" + CStr(minGoodCellsPerWell) + ". Going to Next Well. "
                        If (iPosition + 1 > GridScan_nX.Value * GridScan_nY.Value) Then ' we are in the last well
                            ' set all remaining positions to 0
                            For iLoc = Location To MaxNrLocations
                                posGridXY_valid(iLoc) = 0
                            Next iLoc
                        Else
                            ' only set all positions till the next well to 0
                            For iLoc = Location To locationNumbersMainGrid(iPosition + 1) - 1
                                posGridXY_valid(iLoc) = 0
                            Next iLoc
                        End If
                        ' select next position/next Well
                        Location = locationNumbersMainGrid(iPosition + 1)
                        ' stop if done
                        If (Location > MaxNrLocations) Or (iPosition + 1 > GridScan_nX.Value * GridScan_nY.Value) Then
                            MsgBox "Done with the Location Checking."
                            GoTo DoneWithLocations
                        End If
                    End If
                End If 'CheckBoxGridScan_FindGoodPositions And (Location > 1)
                ' compute whether we are entering a new position (Well) and do iPosition + 1
                If ((Location - 1) Mod (GridScan_nXsub.Value * GridScan_nYsub.Value)) = 0 Then
                    If CheckBoxGridScan_FindGoodPositions And (Location > 1) Then
                        If nGoodCellsPerWell < minGoodCellsPerWell Then ' still the values for the last well
                            MsgBox "New Well: Not enough cells in last well, making valid position " + CStr(LocationSoFarBest) + " with " + CStr(soFarBestGoodCellsPerImage) + " cells."
                            posGridXY_valid(LocationSoFarBest) = 1  ' set the so far best position as valid
                        End If
                    End If
                    If CheckBoxGridScan_FindGoodPositions Then
                        ' init for the new well
                        nGoodCellsPerWell = 0
                        LocationSoFarBest = Location
                        soFarBestGoodCellsPerImage = 0
                    End If
                    If Location > 1 Then  ' iPosition is already initialised with 1
                        iPosition = iPosition + 1
                    End If
                End If '((Location - 1) Mod (GridScan_nXsub.Value * GridScan_nYsub.Value)) = 0
                
                iSubposition = Location - (iPosition - 1) * (GridScan_nXsub.Value * GridScan_nYsub.Value)
                ' setting value of x and y according to grid
                If posGridXY_valid(Location) Then
                    x = posGridX(Location)
                    y = posGridY(Location)
                Else
                    GoTo NextLocation ' skip this position
                End If
               
                '** Here we finally move to next Grid location**'
                Success = FailSafeMoveStage(x, y)
                If Not Success Then
                    ScanStop = True
                    StopAcquisition
                    Exit Sub
                End If
                
            End If 'CheckBoxActiveGridScan.Value
            ''''''''end stage movement to a different position
            
            ' At every positon and repetition  check if Autofocus needs to be required
            If (RepetitionNumber - 1) Mod AFeveryNth = 0 Then
                     
                If Not CheckBoxActiveAutofocus Then  ' Looking if needs to perform an autofocus
                     ZShift = 0
                Else ' perform AUTOFOCUS
                     AutofocusForm.GetBlockValues
                     DisplayProgress "Autofocus 0", RGB(0, &HC0, 0)
                     StopScanCheck 'stop any running jobs
                     RestoreAcquisitionParameters ' has to be there, because after hires mode settings would be wrong for autofocus
                     ' take a z-stack and finds the brightest plane:
                     SuccessAF = Autofocus_StackShift(BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset, RecordingDoc)
                     If Not SuccessAF Then
                        StopAcquisition
                        Exit Sub
                     End If
                     ' move the xyz to the right position
                     DisplayProgress "Autofocus move stage", RGB(0, &HC0, 0)
                     Autofocus_MoveAcquisition BlockZOffset
                End If
            End If '(RepetitionNumber - 1) Mod AFeveryNth = 0
 
            Lsm5.DsRecording.TimeSeries = True  ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
            Lsm5.DsRecording.StacksPerRecord = 1 ' This is time series stack!
            
            ' Set FileNameId. In case of no subpositions then there is also no well
            If GridScan_nXsub.Value * GridScan_nYsub.Value = 1 Then
                FileNameId = FileName(1, iPosition, RepetitionNumber)
            Else
                FileNameId = FileName(iPosition, iSubposition, RepetitionNumber)
            End If
            
            ''''''''''''''''''''''''''''''
            '*Begin Alternative imaging**'
            ''''''''''''''''''''''''''''''
            If CheckBoxAlterLocation.Value = True Then  'this is not in use at the moment? Would use a different alternative imaging
                If Location Mod TextBox_RoundAlterLocation = 0 Then
                    ActivateAlterAcquisitionTrack
                    DisplayProgress "using alternative tracks", RGB(0, 0, &HC0)
                End If
            End If
            
            If CheckBoxAlterImage.Value = True Then
                CopyRecording Lsm5.DsRecording, BackupRecording
                filepath = GlobalDataBaseName & "\" & GlobalFileName & "_" & FileNameId & "_Alt" & ".lsm" ' fullpath of alternative file
                StartAlternativeImaging RecordingDoc, StartTime, filepath, _
                GlobalFileName & "_" & FileNameId & "_Alt" & ".lsm"
            End If
            '****************************'
            
                        
            '''''''''''''''''''''''''''''''''''''
            '*Begin Normal acquisition imaging**'
            '''''''''''''''''''''''''''''''''''''
            CopyRecording Lsm5.DsRecording, BackupRecording  ' restore acquisition parameters
            Sleep (100)
            
            AutofocusForm.ActivateAcquisitionTrack           ' set the tracks to be imaged
            Sleep (100)
            
            If Not IsAcquisitionTrackSelected Then           'An additional control....
                MsgBox "No track selected for Acquisition! Cannot Acquire!"
                ScanStop = True
                StopAcquisition
                Exit Sub
            End If
            
            ScanToImageNew RecordingDoc                       ' **** HERE THE IMAGE IS ACQUIRED ****
            
            If GridScan_nXsub.Value * GridScan_nYsub.Value = 1 Then
                DisplayProgress "Acquiring Position " & iPosition & "(" & iPositionMax & "), Repetition " & RepetitionNumber _
                & "(" & BlockRepetitions & ")", RGB(&HC0, &HC0, 0)  'Now we're going to do the acquisition
            Else
                DisplayProgress "Acquiring Position " & iPosition & "(" & iPositionMax & "), Sub-position " & iSubposition & "(" & iSubpositionMax & ")," _
                & vbCrLf & "Repetition " & RepetitionNumber & "(" & BlockRepetitions & ")", RGB(&HC0, &HC0, 0)       'Now we're going to do the acquisition
            End If
            
            If RepetitionNumber = 1 Then
                StartTime = GetTickCount    'Get the time when the acquisition was started
            End If

            While AcquisitionController.IsGrabbing 'TODO: test function
                Sleep (100)
                If GetInputState() <> 0 Then
                    DoEvents
                    If ScanStop Then
                        StopAcquisition
                        locationNumber = Location
                        Exit Sub
                    End If
                End If
            Wend
            ' ************************************'
            
            ''''''''''''''''''''''''''
            '*** Store bleachTable ***'
            ''''''''''''''''''''''''''
            If BleachStartTable(RepetitionNumber) > 0 Then          'If a bleach was performed we add the information to the image metadata
                Lsm5.DsRecordingActiveDocObject.AddEvent (BleachStartTable(RepetitionNumber) - StartTime) / 1000, eEventTypeBleachStart, "Bleach Start"
                Lsm5.DsRecordingActiveDocObject.AddEvent (BleachStopTable(RepetitionNumber) - StartTime) / 1000, eEventTypeBleachStop, "Bleach End"
            End If
            
            
            ''''''''''''''''''''''''
            '*** Save Image *******'
            ''''''''''''''''''''''''
            RecordingDoc.SetTitle GlobalFileName & "_" & FileNameId
            'this is the name of the file to be saved
            filepath = GlobalDataBaseName & "\" & GlobalFileName & "_" & FileNameId & ".lsm"
            'Check existance of file and warn
            If Not OverwriteFiles Then
                If FileExist(filepath) Then
                    If MsgBox("File " & filepath & " exists. Do you want to overwrite this and subsequent files? ", VbYesNo) = vbYes Then
                        OverwriteFiles = True
                    Else
                        ScanStop = True
                        StopAcquisition
                        locationNumber = Location
                        Exit Sub
                    End If
                End If
            End If
           
            SaveDsRecordingDoc RecordingDoc, filepath  ' HERE THE IMAGE IS FINALLY SAVED
            
            If ScanStop Then    'TODO Check this!
                StopAcquisition
                locationNumber = Location
                Exit Sub
            End If
            ' *******************************
            
            

              
            If Not CheckBoxActiveOnlineImageAnalysis Then ' without MicroPilot
                
                If BleachTable(RepetitionNumber) = True Then   'Check if we're performing a bleach before image acquisition
                    
                    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
                    If Success Then
                        DisplayProgress "Bleaching...", &HFF00FF
                        DoEvents
                        Track.UseBleachParameters = True            'Bleach parameters are lasers lines, bleach iterations... stored in the bleach control window
    '                   BleachStartTable(RepetitionNumber) = Lsm5.ExternalCpObject.pHardwareObjects.pScanController.GetDspTime
                        BleachStartTable(RepetitionNumber) = GetTickCount      'Get the time right before bleach to store this in the image metadata
                        Lsm5.Bleach 0
                        Lsm5.tools.WaitForScanEnd False, 1                     'Waits for the end of the bleach during one second, I think this is not long enough
                        BleachStopTable(RepetitionNumber) = GetTickCount       'Get the time right after bleach to store this in the image metadata
    '                   BleachStopTable(RepetitionNumber) = Lsm5.ExternalCpObject.pHardwareObjects.pScanController.GetDspTime
                        Track.UseBleachParameters = False  'switch off the bleaching
                    Else
                        MsgBox ("Could not set bleach track. Did not bleach.")
                    End If
                    If Location = MaxNrLocations Then   'Alowas again to do an extrableach at the en
                        ExtraBleachButton.Caption = "Bleach"
                        ExtraBleachButton.BackColor = &H8000000F
                    End If
                
                End If
                
                ' todo:
                ' but where is the bleaching image stored ??
            End If
                        
                        
            If CheckBoxActiveOnlineImageAnalysis Then ' MicroPilot Active
                            
                SaveSetting "OnlineImageAnalysis", "macro", "filepath", filepath
                'TODO Check this!
                Do While RecordingDoc.IsBusy
                    Sleep (100)
                    If GetInputState() <> 0 Then
                        DoEvents
                        If ScanStop Then
                            StopAcquisition
                            locationNumber = Location
                            Exit Sub
                        End If
                    End If
                Loop
                
                SaveSetting "OnlineImageAnalysis", "macro", "Refresh", 0
                SaveSetting "OnlineImageAnalysis", "macro", "code", 1
    '            Sleep (600)
    '            SaveSetting "OnlineImageAnalysis", "Ainput", "Refresh", 0
            
            End If
               
                
            If TrackingToggle Or FrameAutofocussing Then
                'not used at the moment find code in unusedCode: ExcelXYZstoring II
            End If
                
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            '**** Updatepositions (x,y)z *********************'''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            If TrackingToggle Then 'This is if we're doing some postacquisition tracking
            
                DisplayProgress "Analysing the new position of location " & Location, &H80FF&
                DoEvents
                MassCenter ("Tracking")
                'If CheckBoxTrackXY Then
                
                If AreStageCoordinateExchanged Then  ' if X and Y are Swapped
                    XCor = YMass
                    YCor = XMass
                Else
                    XCor = XMass
                    YCor = YMass
                End If

                    
                If CheckBoxTrackZ.Value Then
                    ZCor = ZMass
                Else
                    If HRZ Then
                        ZCor = 0
                        'Success = Lsm5.Hardware.CpHrz.Leveling
                    Else
                        ZCor = 0
                    End If
                End If
                
                
            Else ' no location tracking
                
                ' Todo: find out what is happening here
                XCor = 0
                YCor = 0
                If HRZ Then
                    ZCor = 0
                    Success = Lsm5.Hardware.CpHrz.Leveling   'This I think puts the HRZ to its resting position, and moves the focuswheel correspondingly
                Else
                    ZCor = 0
                End If
            
            End If
                    
                    
            Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                Sleep (100)
                'TODO: Check this
                If GetInputState() <> 0 Then
                    DoEvents
                    If ScanStop Then
                        StopAcquisition
                        locationNumber = Location
                        Exit Sub
                    End If
                End If
            Loop
            'sets the new position
            x = Lsm5.Hardware.CpStages.PositionX + XCor                     'Records the current X,Y,Z positions
            y = Lsm5.Hardware.CpStages.PositionY - YCor
            z = Lsm5.Hardware.CpFocus.Position + ZCor   ' this is the current position, including the z-offset
            
            ' End: Defining new (x,y)z positions
            'If Not CheckBoxInactivateAutofocus Then
            '    z = z - BlockZOffset
            'End If
    
            ' Updating positions during tracking (x,y)z positions ***************************
            If MultipleLocationToggle.Value Then
            
'                Success = Lsm5.Hardware.CpStages.MarkClear(0)                   ' Deletes the first mark location in the stage control (the current one)
'                                                                                ' This deletion and new addition of the location
'                                                                                ' was necessary to change the X, Y and Z properties of that location.
'                                                                                ' I did not know how to do it otherwise
'                Lsm5.Hardware.CpStages.MinMarkDistance = 0.1                    ' Put a very small mark distance to make it possible to have two cells coming close together.
'                                                                                ' This parameter can be cahnged with the macro but is not accessible from the main software !
'                While Lsm5.Hardware.CpStages.MarkGetIndex(x, y) <> -1
'                    x = x + 0.1
'                    y = y + 0.1
'                Wend
'
'                ' update the stage positions (particularly important for Location Tracking)
'                Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ(x, y, z) 'Adds the location again,at the end of the list
'
'                Lsm5.Hardware.CpStages.MinMarkDistance = 10                     'Put back the minimal marking distance to its default value
'                'test if this is working
'                Do While Lsm5.Info.IsAnyHardwareBusy
'                    Sleep (20)
'                    DoEvents
'                Loop
                
            Else  ' In the single location case with postacquisition tracking one still has to move to the new focus before next acquisition
                
                Lsm5.Hardware.CpFocus.Position = z + ZBacklash
                Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                    Sleep (20)
                    DoEvents
                Loop
                Lsm5.Hardware.CpFocus.Position = z
                Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                    Sleep (20)
                    DoEvents
                Loop
                
                If TrackingToggle Then   ' In the single location case one also neess to correct for the XY movements if location tracking is activated
                    Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y) ' moves here
                    Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                         'check this
                         Sleep (100)
                         If GetInputState() <> 0 Then
                            DoEvents
                            If ScanStop Then
                                StopAcquisition
                                locationNumber = Location
                                Exit Sub
                            End If
                        End If
                    Loop
                End If
                
            End If
                
            ''  End: Setting new (x,y)z positions *******************************
             
             
             
            ' COMMUNICATION WITH MICROPILOT: START *****************
              
            If CheckBoxActiveOnlineImageAnalysis Then
                
                MicroscopePilot RecordingDoc, BleachingActivated, HighResExperimentCounter, HighResCounter, HighResArrayX, HighResArrayY, HighResArrayZ
            
            End If
            
            If CheckBoxGridScan_FindGoodPositions Then
                    
                'MsgBox "nGoodCells " + CStr(nGoodCells) + " minGoodCells " + CStr(minGoodCellsPerImage)
                
                ' compute whether we just entered a new well or whether we are in the very last Location
                
                
                If nGoodCells > soFarBestGoodCellsPerImage Then
                    LocationSoFarBest = Location
                    soFarBestGoodCellsPerImage = nGoodCells
                End If
                    
                
                If nGoodCells >= minGoodCellsPerImage Then
                    posGridXY_valid(Location) = 1 ' image this position
                    nGoodCellsPerWell = nGoodCellsPerWell + nGoodCells
                Else
                    MsgBox "not enough cells; remove this image from position list"
                    posGridXY_valid(Location) = 0 ' do not image this position
                End If
                
                
                If Location = MaxNrLocations Then ' we are at the last image, check whether this well has enough cells
                    If nGoodCellsPerWell < minGoodCellsPerWell Then ' still the values for the last well
                        MsgBox "Last image: Not enough cells in this well, making valid position " + CStr(LocationSoFarBest) + " with " + CStr(soFarBestGoodCellsPerImage) + " cells."
                        posGridXY_valid(LocationSoFarBest) = 1  ' set the so far best position as valid
                    End If
                End If
                
                    
                
                
            End If
                
            
            ' COMMUNICATION WITH MICROPILOT: END *****************
                 
                 
            ' the following is done here already, beacuse in case the imaging ends the
            ' zoom settings are still on, which would be annoying
            
            ' reset all acquistion parameters
            CopyRecording Lsm5.DsRecording, BackupRecording  ' destination <- source
            
            ' reset the imaging tracks
            ActivateAcquisitionTrack
             
             
NextLocation:
        
        
        Next Location
        'reset location to first location for a new round of repetition
        locationNumber = 1
            
DoneWithLocations:
            
        
        ' DONE WITH THE IMAGING....NOW POSTPROCESSING...
        
        If AutomaticBleaching Then
            FillBleachTable     ' Updating the bleaching table before the next acquisitions, just in case there were changes n the bleaching window
        End If
        
        
        If (RepetitionNumber < BlockRepetitions) Then
            
            If (CheckBoxInterval) Then
                ' do nothing => leave GlobalPrvTime as the time that set at the beginng of the position loop
            Else ' delay
                GlobalPrvTime = CDbl(GetTickCount) * 0.001    'Reset the time to NOW
            End If
            
            rettime = CDbl(GetTickCount) * 0.001
            difftime = rettime - GlobalPrvTime
            'TODO: Check this
            'This loops define the waiting delay before going back to the first location
            Do While (difftime <= BlockTimeDelay) And Not (BleachTable(RepetitionNumber + 1) = True)
                Sleep (100)
                If GetInputState() <> 0 Then
                    DoEvents
                    If ScanPause = True Then
                        Pause
                    End If
                    If ExtraBleach Then                                 'Modifies the bleaching table to do an Extrableach for al locatins at the next repetition
                        ExtraBleach = False
                        BleachTable(RepetitionNumber + 1) = True
                    End If
                    If ScanStop Then
                        StopAcquisition
                        locationNumber = Location
                        Exit Sub
                    End If
                End If
                DisplayProgress "Waiting " & CStr(CInt(BlockTimeDelay - difftime)) + " s before scanning repetition  " & (RepetitionNumber + 1), RGB(&HC0, &HC0, 0)
                rettime = CDbl(GetTickCount) * 0.001
                difftime = rettime - GlobalPrvTime
            Loop
            
        Else
            
            Running = False  ' done with everything done all repetitions
        
        End If
        
        RepetitionNumber = RepetitionNumber + 1
        ' TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
        
    
    Loop ' RepetitonLoop ; Do While Running
    
    ' set back the tracks to be imaged
    ActivateAcquisitionTrack
            
    StopAcquisition
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)


End Sub

'''''
'   MakeGrid( posGridX() As Double, posGridY() As Double, posGridXY_valid() )
'   Create a grid
'       [posGridX] In/Out - Array where X grid positions are stored
'       [posGridY] In/Out - Array where Y grid positions are stored
'       [posGridXY_valid] In/Out - Array that says if position is valid
'       [locationNumbersMainGrid] In/Out - location number on main grid
'''''
Private Sub MakeGrid(posGridX() As Double, posGridY() As Double, posGridXY_valid() As Integer _
, locationNumbersMainGrid() As Long)
    
        'Positions
        Dim tmpGridX As Double
        Dim tmpGridY As Double
        'subPosition
        Dim tmpGridXsub As Double
        Dim tmpGridYsub As Double
                
        'counters
        Dim iy As Long
        Dim ix As Long
        Dim iyy As Long
        Dim ixx As Long
        Dim iLoc As Long
        Dim iLocMainGrid As Long
        'for changing direction, Meander
        Dim xDirection As Integer
        Dim xxDirection As Integer
        
                
        tmpGridX = XStart
        tmpGridY = YStart
        
        iLoc = 1
        iLocMainGrid = 0
        xDirection = 1 ' meander
        
        For iy = 1 To GridScan_nY.Value
            
            For ix = 1 To GridScan_nX.Value
                If ix = 1 Then
                    tmpGridX = tmpGridX
                Else
                    tmpGridX = tmpGridX + xDirection * GridScan_dX.Value
                End If
                    
                iLocMainGrid = iLocMainGrid + 1
                locationNumbersMainGrid(iLocMainGrid) = iLoc  ' remember where the start position of sub is with respect to global position
                
                ' Sub-Positions: start
                tmpGridXsub = tmpGridX
                tmpGridYsub = tmpGridY
                
                xxDirection = 1 ' meander
                
                For iyy = 1 To GridScan_nYsub.Value
                    
                    For ixx = 1 To GridScan_nXsub.Value
                        
                        If ixx = 1 Then
                            tmpGridXsub = tmpGridXsub
                        Else
                            tmpGridXsub = tmpGridXsub + xxDirection * GridScan_dXsub.Value
                        End If
                            
                        posGridX(iLoc) = tmpGridXsub
                        posGridY(iLoc) = tmpGridYsub
                        posGridXY_valid(iLoc) = 1 ' image this position
                    
                        iLoc = iLoc + 1
                
                    Next ixx
                    
                    xxDirection = xxDirection * (-1) ' meander back and forth
                    tmpGridYsub = tmpGridYsub + GridScan_dYsub.Value
                
                Next iyy
                ' Sub-Positions: end
                
            Next ix
            xDirection = xDirection * (-1) ' meander
            tmpGridY = tmpGridY + GridScan_dY.Value ' update Y position
        Next iy
End Sub


''''''
'   MassCenter(Context As String)
'   TODO: No test of Goodness of Mass estimation. what is the exact algorithm?
''''''
Public Sub MassCenter(Context As String)
    Dim scrline As Variant
    Dim spl As Long
    Dim bpp As Long
    Dim ColMax As Long
    Dim LineMax As Long
    Dim FrameNumber As Integer
    Dim PixelSize As Double
    Dim FrameSpacing As Double
    Dim Intline() As Long
    Dim IntCol() As Long
    Dim IntFrame() As Long
    Dim channel As Integer
    Dim frame As Long
    Dim line As Long
    Dim Col As Long
    Dim MinColValue As Single
    Dim minLineValue As Single
    Dim minFrameValue As Single
    Dim MaxColValue As Single
    Dim MaxLineValue As Single
    Dim MaxframeValue As Single
    Dim LineSum As Double
    Dim LineWeight As Single
    Dim MidLine As Single
    Dim Threshold As Single
    Dim LineValue As Single
    Dim PosValue As Single
    Dim ColSum As Single
    Dim ColWeight As Single
    Dim MidCol As Single
    Dim ColValue As Single
    Dim FrameSum As Single
    Dim FrameWeight As Single
    Dim MidFrame As Single
    Dim FrameValue As Single
    
   
    'Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
    DoEvents
    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
    If FrameAutofocussing And SystemName = "LIVE" Then ' binning only with LIVE device
        ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
        LineMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionHeight
    Else
        If SystemName = "LIVE" Then
            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth
            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionHeight
        ElseIf SystemName = "LSM" Then
            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
        Else
            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
            Exit Sub
        End If
    End If
    If Lsm5.DsRecordingActiveDocObject.Recording.ScanMode = "ZScan" Or Lsm5.DsRecordingActiveDocObject.Recording.ScanMode = "Stack" Then
        FrameNumber = Lsm5.DsRecordingActiveDocObject.Recording.FramesPerStack
    Else
        FrameNumber = 1
    End If
    'Gets the pixel size
    PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
    'Gets the distance between frames in Z
    FrameSpacing = Lsm5.DsRecordingActiveDocObject.Recording.FrameSpacing
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    ReDim Intline(LineMax) As Long
    ReDim IntCol(ColMax) As Long
    ReDim IntFrame(FrameNumber) As Long

    'Select the image channel on which to do the calculations
    If Context = "Autofocus" Then       'Takes the first channel in the context of preacquisition focussing
        channel = 0
    ElseIf Context = "Tracking" Then    'Takes the channel selected in the pop-up menue when doing postacquisition tracking
        For channel = 0 To Lsm5.DsRecordingActiveDocObject.GetDimensionChannels - 1
            If Lsm5.DsRecordingActiveDocObject.ChannelName(channel) = Left(TrackingChannelString, 4) Then
                Exit For
            End If
        Next channel
    End If
    
    'Tracking is not the correct word. It just does center of mass on an additional channel

    'Reads the pixel values and fills the tables with the projected (integrated) pixels values in the three directions
    For frame = 1 To FrameNumber
        For line = 1 To LineMax
            bpp = 0
            'channel = 0: This will allow to do the tracking on a different channel
            scrline = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, frame - 1, line - 1, spl, bpp) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 2 To ColMax               'Now I'm scanning all the pixels in the line
                Intline(line - 1) = Intline(line - 1) + scrline(Col - 1)
                IntCol(Col - 1) = IntCol(Col - 1) + scrline(Col - 1)
                IntFrame(frame - 1) = IntFrame(frame - 1) + scrline(Col - 1)
            Next Col
        Next line
    Next frame
    
    'First it finds the minimum and maximum porjected (integrated) pixel values in the 3 dimensions
    MinColValue = 4095 * LineMax * FrameNumber          'The maximum values are initially set to the maximum possible value
    minLineValue = 4095 * ColMax * FrameNumber
    minFrameValue = 4095 * LineMax * ColMax
    MaxColValue = 0                                     'The maximun values are initialliy set to 0
    MaxLineValue = 0
    MaxframeValue = 0
    For line = 1 To LineMax
        If Intline(line - 1) < minLineValue Then
            minLineValue = Intline(line - 1)
        End If
        If Intline(line - 1) > MaxLineValue Then
            MaxLineValue = Intline(line - 1)
        End If
    Next line
    For Col = 1 To ColMax
        If IntCol(Col - 1) < MinColValue Then
            MinColValue = IntCol(Col - 1)
        End If
        If IntCol(Col - 1) > MaxColValue Then
            MaxColValue = IntCol(Col - 1)
        End If
    Next Col
    For frame = 1 To FrameNumber
        If IntFrame(frame - 1) < minFrameValue Then
            minFrameValue = IntFrame(frame - 1)
        End If
        If IntFrame(frame - 1) > MaxframeValue Then
            MaxframeValue = IntFrame(frame - 1)
        End If
    Next frame
    ' Why do you need to threshold the image? (this is probably to remove noise
    'Calculates the threshold values. It is set to an arbitrary value of the minimum projected value plus 20% of the difference between the minimum and the maximum projected value.
    'Then calculates the center of mass
    LineSum = 0
    LineWeight = 0
    MidLine = (LineMax + 1) / 2
    Threshold = minLineValue + (MaxLineValue - minLineValue) * 0.8         'Threshold calculation
    For line = 1 To LineMax
        LineValue = Intline(line - 1) - Threshold                           'Subtracs the threshold
        PosValue = LineValue + Abs(LineValue)                               'Makes sure that the value is positive or zero. If LineValue is negative, the Posvalue = 0; if Line value is positive, then Posvalue = 2*LineValue
        LineWeight = LineWeight + (PixelSize * (line - MidLine)) * PosValue 'Calculates the weight of the Thresholded projected pixel values according to their position relative to the center of the image and sums them up
        LineSum = LineSum + PosValue                                        'Calculates the sum of the thresholded pixel values
    Next line
    If LineSum = 0 Then
        YMass = 0
    Else
        YMass = LineWeight / LineSum                                       'Normalizes the weights to get the center of mass
    End If

    ColSum = 0
    ColWeight = 0
    MidCol = (ColMax + 1) / 2
    Threshold = MinColValue + (MaxColValue - MinColValue) * 0.8
    For Col = 1 To ColMax
        ColValue = IntCol(Col - 1) - Threshold
        PosValue = ColValue + Abs(ColValue)
        ColWeight = ColWeight + (PixelSize * (Col - MidCol)) * PosValue
        ColSum = ColSum + PosValue
    Next Col
    If ColSum = 0 Then
        XMass = 0
    Else
        XMass = ColWeight / ColSum
    End If

    FrameSum = 0
    FrameWeight = 0
    MidFrame = (FrameNumber + 1) / 2
    Threshold = minFrameValue + (MaxframeValue - minFrameValue) * 0.8
    For frame = 1 To FrameNumber
        FrameValue = IntFrame(frame - 1) - Threshold
        PosValue = FrameValue + Abs(FrameValue)
        FrameWeight = FrameWeight + (FrameSpacing * (frame - MidFrame)) * PosValue
        FrameSum = FrameSum + PosValue
    Next frame
    
    If FrameSum = 0 Then
        ZMass = 0
    Else
        ZMass = FrameWeight / FrameSum
    End If
        
End Sub

''''''
'   MassCenterF(Context As String)
'   TODO: Make a faster procedure here
''''''
Public Sub MassCenterF(Context As String)
    Dim scrline As Variant
    Dim spl As Long ' samples per line
    Dim bpp As Long ' bytes per pixel
    Dim ColMax As Long
    Dim LineMax As Long
    Dim FrameNumber As Integer
    Dim PixelSize As Double
    Dim FrameSpacing As Double
    Dim Intline() As Long
    Dim IntCol() As Long
    Dim IntFrame() As Long
    Dim channel As Integer
    Dim frame As Long
    Dim line As Long
    Dim Col As Long
    Dim MinColValue As Single
    Dim minLineValue As Single
    Dim minFrameValue As Single
    Dim MaxColValue As Single
    Dim MaxLineValue As Single
    Dim MaxframeValue As Single
    Dim LineSum As Double
    Dim LineWeight As Single
    Dim MidLine As Single
    Dim Threshold As Single
    Dim LineValue As Single
    Dim PosValue As Single
    Dim ColSum As Single
    Dim ColWeight As Single
    Dim MidCol As Single
    Dim ColValue As Single
    Dim FrameSum As Single
    Dim FrameWeight As Single
    Dim MidFrame As Single
    Dim FrameValue As Single
    
   
    'Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
    DoEvents
    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
    If FrameAutofocussing And SystemName = "LIVE" Then ' binning only with LIVE device
        ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
        LineMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionHeight
    Else
        If SystemName = "LIVE" Then
            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth
            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionHeight
        ElseIf SystemName = "LSM" Then
            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
        Else
            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
            Exit Sub
        End If
    End If
    If Lsm5.DsRecordingActiveDocObject.Recording.ScanMode = "ZScan" Or Lsm5.DsRecordingActiveDocObject.Recording.ScanMode = "Stack" Then
        FrameNumber = Lsm5.DsRecordingActiveDocObject.Recording.FramesPerStack
    Else
        FrameNumber = 1
    End If
    'Gets the pixel size
    PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
    'Gets the distance between frames in Z
    FrameSpacing = Lsm5.DsRecordingActiveDocObject.Recording.FrameSpacing
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    ReDim Intline(LineMax) As Long
    ReDim IntCol(ColMax) As Long
    ReDim IntFrame(FrameNumber) As Long

    'Select the image channel on which to do the calculations
    If Context = "Autofocus" Then       'Takes the first channel in the context of preacquisition focussing
        channel = 0
    ElseIf Context = "Tracking" Then    'Takes the channel selected in the pop-up menue when doing postacquisition tracking
        For channel = 0 To Lsm5.DsRecordingActiveDocObject.GetDimensionChannels - 1
            If Lsm5.DsRecordingActiveDocObject.ChannelName(channel) = Left(TrackingChannelString, 4) Then
                Exit For
            End If
        Next channel
    End If
    
    'Tracking is not the correct word. It just does center of mass on an additional channel

    'Reads the pixel values and fills the tables with the projected (integrated) pixels values in the three directions
    For frame = 1 To FrameNumber
        For line = 1 To LineMax
            bpp = 0 ' bytesperpixel
            'channel = 0: This will allow to do the tracking on a different channel
            scrline = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, frame - 1, line - 1, spl, bpp) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 2 To ColMax               'Now I'm scanning all the pixels in the line
                Intline(line - 1) = Intline(line - 1) + scrline(Col - 1)
                IntCol(Col - 1) = IntCol(Col - 1) + scrline(Col - 1)
                IntFrame(frame - 1) = IntFrame(frame - 1) + scrline(Col - 1)
            Next Col
        Next line
    Next frame
    
    'First it finds the minimum and maximum projected (integrated) pixel values in the 3 dimensions
    MinColValue = 4095 * LineMax * FrameNumber          'The maximum values are initially set to the maximum possible value
    minLineValue = 4095 * ColMax * FrameNumber
    minFrameValue = 4095 * LineMax * ColMax
    MaxColValue = 0                                     'The maximun values are initialliy set to 0
    MaxLineValue = 0
    MaxframeValue = 0
    For line = 1 To LineMax
        If Intline(line - 1) < minLineValue Then
            minLineValue = Intline(line - 1)
        End If
        If Intline(line - 1) > MaxLineValue Then
            MaxLineValue = Intline(line - 1)
        End If
    Next line
    For Col = 1 To ColMax
        If IntCol(Col - 1) < MinColValue Then
            MinColValue = IntCol(Col - 1)
        End If
        If IntCol(Col - 1) > MaxColValue Then
            MaxColValue = IntCol(Col - 1)
        End If
    Next Col
    For frame = 1 To FrameNumber
        If IntFrame(frame - 1) < minFrameValue Then
            minFrameValue = IntFrame(frame - 1)
        End If
        If IntFrame(frame - 1) > MaxframeValue Then
            MaxframeValue = IntFrame(frame - 1)
        End If
    Next frame
    ' Why do you need to threshold the image? (this is probably to remove noise
    'Calculates the threshold values. It is set to an arbitrary value of the minimum projected value plus 20% of the difference between the minimum and the maximum projected value.
    'Then calculates the center of mass
    LineSum = 0
    LineWeight = 0
    MidLine = (LineMax + 1) / 2
    Threshold = minLineValue + (MaxLineValue - minLineValue) * 0.8         'Threshold calculation
    For line = 1 To LineMax
        LineValue = Intline(line - 1) - Threshold                           'Subtracs the threshold
        PosValue = LineValue + Abs(LineValue)                               'Makes sure that the value is positive or zero. If LineValue is negative, the Posvalue = 0; if Line value is positive, then Posvalue = 2*LineValue
        LineWeight = LineWeight + (PixelSize * (line - MidLine)) * PosValue 'Calculates the weight of the Thresholded projected pixel values according to their position relative to the center of the image and sums them up
        LineSum = LineSum + PosValue                                        'Calculates the sum of the thresholded pixel values
    Next line
    If LineSum = 0 Then
        YMass = 0
    Else
        YMass = LineWeight / LineSum                                       'Normalizes the weights to get the center of mass
    End If

    ColSum = 0
    ColWeight = 0
    MidCol = (ColMax + 1) / 2
    Threshold = MinColValue + (MaxColValue - MinColValue) * 0.8
    For Col = 1 To ColMax
        ColValue = IntCol(Col - 1) - Threshold
        PosValue = ColValue + Abs(ColValue)
        ColWeight = ColWeight + (PixelSize * (Col - MidCol)) * PosValue
        ColSum = ColSum + PosValue
    Next Col
    If ColSum = 0 Then
        XMass = 0
    Else
        XMass = ColWeight / ColSum
    End If

    FrameSum = 0
    FrameWeight = 0
    MidFrame = (FrameNumber + 1) / 2
    Threshold = minFrameValue + (MaxframeValue - minFrameValue) * 0.8
    For frame = 1 To FrameNumber
        FrameValue = IntFrame(frame - 1) - Threshold
        PosValue = FrameValue + Abs(FrameValue)
        FrameWeight = FrameWeight + (FrameSpacing * (frame - MidFrame)) * PosValue
        FrameSum = FrameSum + PosValue
    Next frame
    
    If FrameSum = 0 Then
        ZMass = 0
    Else
        ZMass = FrameWeight / FrameSum
    End If
        
End Sub




Private Sub PauseButton_Click()
    If Running Then
        If ScanPause = False Then
            ScanPause = True
            PauseButton.Caption = "Resume"
            PauseButton.BackColor = 12648447
        Else
            ScanPause = False
            PauseButton.Caption = "Pause"
            PauseButton.BackColor = &H8000000F
        End If
    Else
        MsgBox "The acquisition has not started yet or is already finished. Cannot pause."
    End If
End Sub


'''''
'   Pause()
'   Function called when ScanPause = True
'   Checks state and wait for action in Form
'''''
Public Sub Pause()
    
    Dim rettime As Double
    Dim GlobalPrvTime As Double
    Dim difftime As Double
    
    GetCurrentPositionOffsetButton.Enabled = True
    AutofocusButton.Enabled = True
    GlobalPrvTime = CDbl(GetTickCount) * 0.001
    rettime = GlobalPrvTime
    difftime = rettime - GlobalPrvTime
    'TODO: test this function
    Do While True
        Sleep (100)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StopAcquisition
                Exit Sub
            End If
            If ScanPause = False Then
                GetCurrentPositionOffsetButton.Enabled = False
                AutofocusButton.Enabled = False
                Exit Sub
            End If
        End If
        DisplayProgress "Pause " & CStr(CInt(difftime)) & " s", RGB(&HC0, &HC0, 0)
        rettime = CDbl(GetTickCount) * 0.001
        difftime = rettime - GlobalPrvTime
    Loop
End Sub


Private Sub ExtraBleachButton_Click()
    
    If Running Then
        ExtraBleach = True
        ExtraBleachButton.Caption = "Will Bleach"
        ExtraBleachButton.BackColor = 12648447
    Else
        MsgBox "The acquisition has not started yet or is already finished. Cannot bleach."
    End If

End Sub

'''''''
'   MultipleLocationToggle_Change()
'   Activate MultipleLocation and deactivate SingleLocation
'''''''
Private Sub MultipleLocationToggle_Change()
        
    If MultipleLocationToggle.Value = True Then
        CheckBoxActiveGridScan.Value = False
        SetMultipleLocationToggle_True
    Else
        SingleLocationToggle.Value = True
    End If
    
End Sub


'''''''
'   SingleLocationToggle_Change()
'   Activate Singlelocation and deactivate MultipleLocation
'''''''
Private Sub SingleLocationToggle_Change()
    
    If SingleLocationToggle.Value = True Then
        SetSingleLocationToggle_True
    Else
        MultipleLocationToggle.Value = True
    End If

End Sub

Private Sub SetSingleLocationToggle_True()
                
        ' MsgBox "Setting Single Locations True"
        
        SingleLocationToggle.Value = True
        MultipleLocationToggle.Value = False
        LocationTextLabel.Caption = ""
        
        ' CheckBoxScannAll.Visible = False
        ' GridObjectsandVarialbles False
        ' StartBleachButton.Visible = True
        ' ExtraBleachButton.Visible = True
        ' Frame15.Visible = False
        ' If GridToggle.Value = True Then GridToggle.Value = Not SingleLocationToggle.Value

End Sub
  
Private Sub SetMultipleLocationToggle_True()
  
        ' MsgBox "Setting Multiple Locations True"
        
        SingleLocationToggle.Value = False
        MultipleLocationToggle.Value = True
        LocationTextLabel.Caption = "Define locations using the Stage (NOT the Positions) dialog !"
        
        CheckBoxActiveGridScan.Value = False ' currently not compatible
        
        
        ' CheckBoxScannAll.Visible = False
        ' GridObjectsandVarialbles False
        ' StartBleachButton.Visible = True
        ' ZMapButton.Left = 12
        ' ZMapButton.Top = 258
        ' CheckBoxZMap.Left = 80
        ' CheckBoxZMap.Top = 258
        'ZMapButton.Visible = True
        'CheckBoxZMap.Visible = True
        'ExtraBleachButton.Visible = True
        'Frame15.Visible = True
        'TextBoxTileX.Visible = True
        'TextBoxTileY.Visible = True
        'Tileframe.Visible = True
        'Label17.Visible = True
        'Label18.Visible = True
        'Label20.Visible = True
        'CreateLocationsButton.Visible = True
        'TextBoxOverlap.Visible = True
        'If GridToggle.Value = True Then GridToggle.Value = Not MultipleLocationToggle.Value
        
End Sub
  
  


Public Sub AutoFindTracks()

    Dim i, j As Integer
    Dim ChannelOK As Boolean
    Dim DataChannel As DsDataChannel
    Dim Color As Long
    Dim ConfiguredTracks As Integer
    Dim GoodTracks As Integer

    
    OptionButtonTrack1.Visible = False
    OptionButtonTrack1.Enabled = False
    OptionButtonTrack1.Value = False
    CheckBoxTrack1.Visible = False
    CheckBoxTrack1.Enabled = False
    CheckBoxTrack1.Value = False
    CheckBoxZoomTrack1.Visible = False
    CheckBoxZoomTrack1.Enabled = False
    CheckBoxZoomTrack1.Value = False
    CheckBox2ndTrack1.Visible = False
    CheckBox2ndTrack1.Enabled = False
    CheckBox2ndTrack1.Value = False
                         
    
    OptionButtonTrack2.Visible = False
    OptionButtonTrack2.Enabled = False
    OptionButtonTrack2.Value = False
    CheckBoxTrack2.Visible = False
    CheckBoxTrack2.Enabled = False
    CheckBoxTrack2.Value = False
    CheckBoxZoomTrack2.Visible = False
    CheckBoxZoomTrack2.Enabled = False
    CheckBoxZoomTrack2.Value = False
    CheckBox2ndTrack2.Visible = False
    CheckBox2ndTrack2.Enabled = False
    CheckBox2ndTrack2.Value = False
    
    OptionButtonTrack3.Visible = False
    OptionButtonTrack3.Enabled = False
    OptionButtonTrack3.Value = False
    CheckBoxTrack3.Visible = False
    CheckBoxTrack3.Enabled = False
    CheckBoxTrack3.Value = False
    CheckBoxZoomTrack3.Visible = False
    CheckBoxZoomTrack3.Enabled = False
    CheckBoxZoomTrack3.Value = False
    CheckBox2ndTrack3.Visible = False
    CheckBox2ndTrack3.Enabled = False
    CheckBox2ndTrack3.Value = False
   
    OptionButtonTrack4.Visible = False
    OptionButtonTrack4.Enabled = False
    OptionButtonTrack4.Value = False
    CheckBoxTrack4.Visible = False
    CheckBoxTrack4.Enabled = False
    CheckBoxTrack4.Value = False
    CheckBoxZoomTrack4.Visible = False
    CheckBoxZoomTrack4.Enabled = False
    CheckBoxZoomTrack4.Value = False
    CheckBox2ndTrack4.Visible = False
    CheckBox2ndTrack4.Enabled = False
    CheckBox2ndTrack4.Value = False
   

    ConfiguredTracks = Lsm5.DsRecording.TrackCount
    ChannelOK = False
    GoodTracks = 0
    
    'The next line and the following "if" should be removed when working with the LSM 2.8 software (where the lambda mode is not defined)
    Set Track = Lsm5.DsRecording.TrackObjectLambda(Success)
    If Success Then
        If Track.Acquire Then
            MsgBox ("This macro does not work in the Lambda Mode. Please switch to the Channel Mode and reinitialize the Macro.")
            Exit Sub
        End If
    End If
            
    For i = 1 To ConfiguredTracks
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        TrackName = Track.name
        j = 0
        'In the next line remove "Or Track.IslambdaTrack" when working with the LSM 2.8 software
        If Not (Track.IsBleachTrack Or Track.IsLambdaTrack) Then
            Do While (Not ChannelOK) And (j < Track.DataChannelCount)
                Set DataChannel = Track.DataChannelObjectByIndex(j, Success)
                If DataChannel.Acquire = True Then ChannelOK = True
                Color = DataChannel.ColorRef
                j = j + 1
            Loop
            If ChannelOK Then
                If Not Track.IsRatioTrack Then
                    GoodTracks = GoodTracks + 1
                    If GoodTracks = 5 Then
                        MsgBox ("This Macro only accepts 4 different tracks")
                    End If
                    If GoodTracks = 1 Then
                        OptionButtonTrack1.Visible = True
                        OptionButtonTrack1.Caption = TrackName
                        OptionButtonTrack1.Enabled = True
                        OptionButtonTrack1.BackColor = Color
                        
                        CheckBoxTrack1.Visible = True
                        CheckBoxTrack1.Caption = TrackName
                        CheckBoxTrack1.Enabled = True
                        CheckBoxTrack1.BackColor = Color
                        
                        CheckBoxZoomTrack1.Visible = True
                        CheckBoxZoomTrack1.Caption = TrackName
                        CheckBoxZoomTrack1.Enabled = CheckBoxActiveOnlineImageAnalysis.Value
                        CheckBoxZoomTrack1.BackColor = Color
                        
                        CheckBox2ndTrack1.Visible = True
                        CheckBox2ndTrack1.Caption = TrackName
                        CheckBox2ndTrack1.Enabled = CheckBoxAlterImage.Value
                        CheckBox2ndTrack1.BackColor = Color
                        
                    End If
                    If GoodTracks = 2 Then
                        OptionButtonTrack2.Visible = True
                        OptionButtonTrack2.Caption = TrackName
                        OptionButtonTrack2.Enabled = True
                        OptionButtonTrack2.BackColor = Color
                        CheckBoxTrack2.Visible = True
                        CheckBoxTrack2.Caption = TrackName
                        CheckBoxTrack2.Enabled = True
                        CheckBoxTrack2.BackColor = Color
                        
                        CheckBoxZoomTrack2.Visible = True
                        CheckBoxZoomTrack2.Caption = TrackName
                        CheckBoxZoomTrack2.Enabled = True
                        CheckBoxZoomTrack2.BackColor = Color
                        
                        CheckBox2ndTrack2.Visible = True
                        CheckBox2ndTrack2.Caption = TrackName
                        CheckBox2ndTrack2.Enabled = True
                        CheckBox2ndTrack2.BackColor = Color
                        
                    End If
                    If GoodTracks = 3 Then
                        OptionButtonTrack3.Visible = True
                        OptionButtonTrack3.Caption = TrackName
                        OptionButtonTrack3.Enabled = True
                        OptionButtonTrack3.BackColor = Color
                        
                        CheckBoxTrack3.Visible = True
                        CheckBoxTrack3.Caption = TrackName
                        CheckBoxTrack3.Enabled = True
                        CheckBoxTrack3.BackColor = Color
                        
                        CheckBoxZoomTrack3.Visible = True
                        CheckBoxZoomTrack3.Caption = TrackName
                        CheckBoxZoomTrack3.Enabled = True
                        CheckBoxZoomTrack3.BackColor = Color
                        
                        CheckBox2ndTrack3.Visible = True
                        CheckBox2ndTrack3.Caption = TrackName
                        CheckBox2ndTrack3.Enabled = True
                        CheckBox2ndTrack3.BackColor = Color
                        
                    End If
                    If GoodTracks = 4 Then
                        OptionButtonTrack4.Visible = True
                        OptionButtonTrack4.Caption = TrackName
                        OptionButtonTrack4.Enabled = True
                        OptionButtonTrack4.BackColor = Color
                        
                        CheckBoxTrack4.Visible = True
                        CheckBoxTrack4.Caption = TrackName
                        CheckBoxTrack4.Enabled = True
                        CheckBoxTrack4.BackColor = Color
                        
                        CheckBoxZoomTrack4.Visible = True
                        CheckBoxZoomTrack4.Caption = TrackName
                        CheckBoxZoomTrack4.Enabled = True
                        CheckBoxZoomTrack4.BackColor = Color
                        
                        CheckBox2ndTrack4.Visible = True
                        CheckBox2ndTrack4.Caption = TrackName
                        CheckBox2ndTrack4.Enabled = True
                        CheckBox2ndTrack4.BackColor = Color
                        
                    End If
                Else
                    MsgBox ("This macro does not allow to use a Ratio Channel. The Ratio Channel will thus be disabled.")
                    For j = 0 To Track.DataChannelCount - 1
                        Set DataChannel = Track.DataChannelObjectByIndex(j, Success)
                        DataChannel.Acquire = False
                    Next
                End If
                ChannelOK = False
            End If
        End If
    Next
    If GoodTracks < 4 Then
        TrackNumber = GoodTracks
    Else
        TrackNumber = 4
    End If
End Sub



Private Sub BSliderZOffset_Change()
    'Tests whether chosen Offset is less or equal to half of the working distance of the objective but why can't it be bigger??
'    Dim Position As Long 'gets the postion of the actual objective revolver by number
'    Dim Range As Double 'contains value of working distance in um
    If flgUserChange Then '??? What is the sense of flgUserChange
'        Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
'        If Position >= 0 Then ' ??? is it possible that Revolver Position has another value
'            Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000# ' ??? why is there a # behind that number if range is already defined as double
'                                                                                            ' in which unit is working distance read out and why multiplication with 1000
'        Else
'            Range = 0#
'        End If
'substituted29.06.2010 by Function Range
        If Abs(BSliderZOffset.Value) > Range * 0.9 Then
            BSliderZOffset.Value = 0
            MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um"
        End If
    End If
End Sub

Private Sub BSliderZRange_Change()    ' It should be possible to change the limit of the range to bigger values than half of the working distance
'    Dim Position As Long
'    Dim Range As Double
    If flgUserChange Then
'        Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
'        If Position >= 0 Then
'            Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
'        Else
'            Range = 0#
'        End If
'substituted29.06.2010 by Function Range
        If BSliderZRange.Value > Range * 0.9 Then
            BSliderZRange.Value = Range * 0.9
            MsgBox "ZRange has to be less or equal to the working distance of the objective: " + CStr(Range) + " um"
        End If
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub CloseButton_Click()
    AutoStore
'    Excel.Application.DisplayAlerts = False
'    Excel.Application.Quit
    End
End Sub

Private Sub ReInitializeButton_Click()
    Re_Initialize
End Sub


Private Sub TextBox1_Change()

End Sub







'''''
'   Re_Initialize()
'   Initializations that need to be performed only when clicking the "Reinitialize" button
'''''
Public Sub Re_Initialize()
    Dim delay As Single
    Dim standType As String
    Dim count As Long
    
    AutoFindTracks
  
    
    PubSearchScan = False
    NoReflectionSignal = False
    PubSentStageGrid = False
    
    '  AutofocusForm.Caption = GlobalProject + " for " + SystemName
    BleachingActivated = False
      
End Sub

Private Sub CreditButton_Click()
    CreditForm.Show
End Sub

''''''
'  TrackingToggle_Click()
'  Add extra tracking channel.
'  Tracking is the wrong word. It just uses an extra channel for the calculation of center of mass which is then used to move the stage
'''''
Private Sub TrackingToggle_Click()
    If TrackingToggle.Value Then
        CheckBoxActiveGridScan.Value = False
    End If
    SwitchEnableTrackingToggle (TrackingToggle.Value)
End Sub

'''''
'   SwitchEnableTrackingToggle(Enable As Boolean)
'   Changes Enable visibility of AcquiistionForm Tracking part
'       [Enable] In - Enable of tracking
'''''
Private Sub SwitchEnableTrackingToggle(Enable As Boolean)
    ComboBoxTrackingChannel.Visible = Enable
    If Enable Then
       FillTrackingChannelList
    End If
    CheckBoxPostTrackXY.Visible = Enable
    CheckBoxTrackZ.Visible = Enable
    If Lsm5.DsRecording.ScanMode = "Stack" Then
        CheckBoxTrackZ.Enabled = True
        PostAcquisitionLabel.Visible = Enable
    Else
        CheckBoxTrackZ.Enabled = False
        CheckBoxTrackZ.Value = False
        PostAcquisitionLabel.Visible = Enable
    End If
End Sub
    

'''''''
' CheckBoxTrackZ_Click()
'   Activate post-acquisition Z-tracking.
'   Inactivates Autofocus
'''''''
Private Sub CheckBoxTrackZ_Click()
    If CheckBoxTrackZ.Value = True Then
        CheckBoxActiveAutofocus.Value = False 'inactivate Autofocus
        SwitchEnableAutofocusPage (False)
    Else
        CheckBoxActiveAutofocus.Value = True
        SwitchEnableAutofocusPage (True)
    End If
End Sub

'fills popup menu for chosing a track for post-acquisition tracking in ScanLine mode
Private Sub FillTrackingChannelList()
    Dim t As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim channel As DsDetectionChannel

    ActivateAcquisitionTrack 'will set IsAcquisitionTrack selected true if a valid track is selected for acquisition, and "marks the track in the Zeiss config window
    
    ReDim ActiveChannels(Lsm5.Constants.MaxActiveChannels)  'ActiveChannels is a dynamic array (variable size), ReDim defines array size required next
                                                            'Array size is (MaxActiveChannels gets) the total max number of active channels in all tracks
    ComboBoxTrackingChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    
    If IsAcquisitionTrackSelected Then 'IsAcquisitionTrackSelected is True if one channel is activated in tracks 1-4
        For t = 1 To TrackNumber 'This loop goes through all tracks and will collect all activated channels to display them in popup menu
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(t - 1, Success) 'goes through all defined tracks
            If Track.Acquire Then 'if track is activated for acquisition
                For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
                    Set channel = Track.DetectionChannelObjectByIndex(c - 1, Success)
                    If channel.Acquire Then 'if channel is activated
                        ca = ca + 1 'counter for active channels will increase by one
                        ComboBoxTrackingChannel.AddItem Track.name & " " & channel.name 'entry is added to combo box to chose track for post-acquisition tracking
                        ActiveChannels(ca) = channel.name & "-T" & Track.MultiplexOrder + 1  'adds entry to ActiveChannel Array with name of channel + name of track
                    End If
                Next c
            End If
        Next t
        ComboBoxTrackingChannel.Value = ComboBoxTrackingChannel.List(0) 'initially displayed text in popup menu is a blank line (first channel is 1).
    End If
End Sub

Private Sub ComboBoxTrackingChannel_Change()        'Sets the name of the channel for PostAcquisition tracking.
    TrackingChannelString = ActiveChannels(ComboBoxTrackingChannel.ListIndex + 1)
End Sub



Private Sub CommandTimeMin_Click()
    TimerUnit = 60
    BSliderTime.Max = 60                        'When workings with minutes the maximum delay that can be set with the slider is 1 hour
    BSliderTime.Value = BlockTimeDelay / 60
    CommandTimeMin.BackColor = &HFF8080
    CommandTimeSec.BackColor = &H8000000F
End Sub

Private Sub CommandTimeSec_Click()
    TimerUnit = 1
    BSliderTime.Max = 180                       'When workings with seconds the maximum delay that can be set with the slider is 3 minutes
    BSliderTime.Value = BlockTimeDelay
    CommandTimeSec.BackColor = &HFF8080
    CommandTimeMin.BackColor = &H8000000F
End Sub

Private Sub BSliderTime_Click()
    BlockTimeDelay = BSliderTime.Value * TimerUnit                      'BlockTimedelay gets the value of the slider in seconds
End Sub

Private Sub BSliderRepetitions_Change()
    If Not Running Then
        BlockRepetitions = BSliderRepetitions.Value
    ElseIf Not (BSliderRepetitions.Value <= (RepetitionNumber + 1)) Then
        BlockRepetitions = BSliderRepetitions.Value
    Else
        BSliderRepetitions.Value = RepetitionNumber + 1
        BlockRepetitions = BSliderRepetitions.Value
    End If
    
    ReDim Preserve GlobalImageIndex(BlockRepetitions)           'The global image index I'm not sure how this is working.
    ReDim Preserve BleachTable(BlockRepetitions)                'BleachTable defines when bleaching will have to occur
    If AutomaticBleaching Then FillBleachTable                  'Reads the parameters defined in the Bleach control window of the main software
    ReDim Preserve BleachStartTable(BlockRepetitions)           'This is to store the timepoints when the bleaches started. Preserve is to keep the timepoints if the slider is moved during an experiment
    ReDim Preserve BleachStopTable(BlockRepetitions)            'This is to store the timepoints when the bleaches stopped. Preserve is to keep the timepoints if the slider is moved during an experiment
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub TextBoxFileName_Change()
    GlobalFileName = TextBoxFileName.Value
End Sub

 
'''''''
'   ActivateAutofocusTrack(HighSpeed As Boolean)
'   Check which track has been activated for Autofocus and set the track properties accordingly
'   TODO: Test
''''''
Public Sub ActivateAutofocusTrack(HighSpeed As Boolean)
    Dim i As Integer
    IsAutofocusTrackSelected = False
    ' Set all tracks to non-aquisition
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        Track.Acquire = 0
    Next i
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        Track.Acquire = 0
        If OptionButtonTrack1.Value = True And i = 1 Then
            IsAutofocusTrackSelected = True
            Exit For
        ElseIf OptionButtonTrack2.Value = True And i = 2 Then
            IsAutofocusTrackSelected = True
            Exit For
        ElseIf OptionButtonTrack3.Value = True And i = 3 Then
            IsAutofocusTrackSelected = True
            Exit For
        ElseIf OptionButtonTrack4.Value = True And i = 4 Then
            IsAutofocusTrackSelected = True
            Exit For
        End If
    Next i
    
    If IsAutofocusTrackSelected Then
        AutofocusTrack = i - 1
        Track.Acquire = 1 ' this basically sets the track belonging to DsRecording to acquire.
                          ' This can be cleaned up by creating a DsRecording for each operation
        If HighSpeed Then
            Track.SamplingNumber = 1
        End If
    End If
    
End Sub

'''''''
'   ActivateAlterAcquisitionTrack
'   Check which track has been activated and for AlternativeAcquisitionTrack set the track properties accordingly
'   TODO: Test
''''''
Public Sub ActivateAlterAcquisitionTrack()
    Dim i As Integer
    IsAcquisitionTrackSelected = False
    'Set all to zero
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        Track.Acquire = 0
    Next i
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        If CheckBox2ndTrack1.Value = True And i = 1 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        ElseIf CheckBox2ndTrack2.Value = True And i = 2 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        ElseIf CheckBox2ndTrack3.Value = True And i = 3 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        ElseIf CheckBox2ndTrack4.Value = True And i = 4 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        End If
    Next i

End Sub


'''''''''
' ActivateAcquisitionTrack()
' If any of the checkboxes in the AutoFocusForm Acquisition are checked
'        Track.Acquire = 1 and IsAcquisitionTrackSelected = True
' otherwise
'       Track.Acquire = 0 and IsAcquisitionTrackSelected = False
' This sets the Macro to perform acquisition after Autofocus
' TODO: Test
''''''''''
Public Sub ActivateAcquisitionTrack()
    Dim i As Integer
    IsAcquisitionTrackSelected = False
    'Set all track to zero
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        Track.Acquire = 0
    Next i
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        If CheckBoxTrack1.Value = True And i = 1 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        ElseIf CheckBoxTrack2.Value = True And i = 2 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        ElseIf CheckBoxTrack3.Value = True And i = 3 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        ElseIf CheckBoxTrack4.Value = True And i = 4 Then
            IsAcquisitionTrackSelected = True
            Track.Acquire = 1
        End If
    Next i
End Sub


'''''''''
' ActivateZoomTrack()
' Micropilotpage. This is extra track for micropilot
' TODO: Test and change name
''''''''''
Private Sub ActivateZoomTrack()
    Dim i As Integer
    IsAcquisitionTrackSelected = False
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, Success)
        If i = 1 Then
            If CheckBoxZoomTrack1.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 2 Then
            If CheckBoxZoomTrack2.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 3 Then
            If CheckBoxZoomTrack3.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 4 Then
            If CheckBoxZoomTrack4.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
    Next i
End Sub

Sub Wait(PauseTime As Single)
    Dim Start As Single
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
       DoEvents    ' Yield to other processes.
       'Lsm5.DsRecording.StartScanTriggerIn
    Loop
End Sub

''''''
'   GetCurrentPositionOffset(ZRange As Double, ZStep As Double, HighSpeed As Boolean, ZOffset As Double)
'   Calculates offset according to actual position of image
'       [ZRange] In - Range in um over which to perform the scan
'       [ZStep]  In - zStep size in um
'       [HighSpeed] In - Use Fast Z-line for LineScan
'       [ZOffset]   In/Out - Return calculated offset value
''''''
Public Sub GetCurrentPositionOffset(ZRange As Double, ZStep As Double, HighSpeed As Boolean, ZOffset As Double)
    Dim SpeedCopy As Double
    Dim ZoomXCopy As Double
    Dim ZoomYCopy As Double
    Dim SamplesPerLineCopy As Long
    Dim LinesPerFrameCopy As Long
    Dim ScanModeCopy As String
    Dim SpecialScanModeCopy As String

'    Dim Range As Double
'    Dim Position As Long
  
    Dim MyRecording As DsRecording

    Dim Tnum As Long
    Dim i As Long
    Dim Success As Integer
    Dim NewPicture As DsRecordingDoc
    Dim Pixel As Long
    Dim scrline As Variant
    Dim PxlArray() As Long
    Dim spl As Long
    Dim bpp As Long
    Dim IntensityStr As String
    Dim ChNumber As Long
    Dim channel As Long
    Dim LongRange As Long
    Dim PxlMax As Long
    Dim PxlTot As Long
    Dim LineMax As Long
    Dim StackSize As Double
    Dim SavedSampling As Long
    Dim key As String
    Dim line As Long
    Dim lT As Long
    'Dim NoFrames As Long MadePublic29.06.2010
    Dim SystemVersion As String
    Dim Speed As Long
    Dim MaxSpeed As Long
    
   
        
    Zbefore = Lsm5.Hardware.CpFocus.Position
    
    DisplayProgress "Get Offset Value", RGB(0, &HC0, 0)             'Gives information to the user
    StopScanCheck
    
    ' ZAuto = 0   removed29.07.2010                                                    'I do not know why is this Z Auto there. I believe it is obsolete
    ' ZBacklash = -50 'Has to do with the movements of the focus wheel that are "better" if they are long enough.
    
    StoreAcquisitionParameters
    
    
    ActivateAutofocusTrack HighSpeed                                'Sets the track for autofocussing (i.e. "selects" the track in the Zeiss config window )
    If Not IsAutofocusTrackSelected Then                                'The variable IsAutofocusTrackSelected has been updated in the ActivateAutofocausTrack function
        MsgBox "No track selected for Autofocus! Cannot Autofocus!"
        StopAcquisition
        Exit Sub
    End If
  
'    Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition       'Verifies that the working distnce is OK. Comes from the initial Zeiss autofocussing macro
'    If Position >= 0 Then
'        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
'    Else
'        Range = 0#
'    End If
'substituted29.06.2010 by Function Range
    
    'MsgBox "ZOffset = " + CStr(ZOffset) + "; Range = " + CStr(Range) + "; ZRange = " + CStr(ZRange)
    
    If Range = 0 Then
        MsgBox "Objective's working distance not defined! Cannot Autofocus!"
        Exit Sub
    End If
    If ZRange > Range * 0.9 Then
        ZRange = Range * 0.9
    End If
    If Abs(ZOffset) > Range * 0.9 Then                   'The offset has to be within half of the working distance. May want to change this when working with large samples in Z
        ZOffset = 0
    End If

    SystemVersionOffset
    
    AutofocusForm.AutofocusSetting HRZ, BlockHighSpeed, BlockZStep
    Lsm5.DsRecording.FrameSpacing = ZStep
    NoFrames = CLng(ZRange / ZStep) + 1                     'Calculates the number of frames per stack. Clng converts it to a long and rounds up the fraction
    Lsm5.DsRecording.FramesPerStack = NoFrames
    
    If NoFrames > 2048 Then                                 'overwrites the userdefined value if too many frames have been defined by the user
        NoFrames = 2048
    End If
    
    'If Not HRZ Then
        Lsm5.DsRecording.Sample0Z = ZStep * NoFrames / 2
    'End If                                                    'Distance of the actual focus to the first Z position of the image (or line) to acquire in the stack.
                                                            'I think this is only valid for the focus wheel and not the HRZ
    
    If ZOffset <= Range * 0.9 Then
       
       'MsgBox " Doing ZBacklash "
       
       Lsm5.Hardware.CpFocus.Position = Zbefore - ZOffset + GlobalCorrectionOffset + ZBacklash 'Move down 50um (=ZBacklash) below the position of the offset
       Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
            Sleep (20)  '20ms
            DoEvents
       Loop
       Lsm5.Hardware.CpFocus.Position = Zbefore - ZOffset + GlobalCorrectionOffset             'Moves up to the position of the offset
       Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
           Sleep (20)
           DoEvents
       Loop
    
    End If
    

    If Not FrameAutofocussing Then
        Lsm5.DsRecording.ScanMode = "ZScan"
        If Not HRZ Then
            Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
        End If
    End If
    
    Set NewPicture = Lsm5.StartScan                             'Starts the image acquisition for autofocussing
    'TODO: Test code
    Do While NewPicture.IsBusy                                  ' Waiting untill the image acquisition is done
        Sleep (100)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StopAcquisition
                Exit Sub
            End If
        End If
    Loop
    
    Lsm5.tools.WaitForScanEnd False, 40                        'TODO: redundancy? This looks redoundant with the previous, but I had trried to remove it and had problems. It's better to have 2 contols than none !
 
    AutofocusForm.MassCenter ("Autofocus")                     'Calculates the mass center in 3 dimensions
    XShift = XMass
    YShift = YMass
    ZShift = ZMass
    
        
    If ZOffset <= Range * 0.9 Then
       Lsm5.Hardware.CpFocus.Position = Zbefore + GlobalCorrectionOffset + ZBacklash  'Move down 50um (=ZBacklash) below the position of the offset
       Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
            Sleep (20)  '20ms
            DoEvents
       Loop
       Lsm5.Hardware.CpFocus.Position = Zbefore + GlobalCorrectionOffset             'Moves up to the position of the offset
       Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
           Sleep (20)
           DoEvents
       Loop
    End If

    If HRZ Then                             'The HRZ and the focus wheel are acquiring Z stacks in opposite directions. TODO: This is now the same. OK?
        ZOffset = -ZShift + ZOffset
    Else
        ZOffset = -ZShift + ZOffset
    End If
    BSliderZOffset.Value = ZOffset          'Update Box ZOffset in AutofocusForm
    
End Sub


Public Sub SetBlockValues()
'    Dim Position As Long
'    Dim Range As Double
 
    CheckBoxHighSpeed.Value = BlockHighSpeed
    CheckBoxLowZoom.Value = BlockLowZoom
    CheckBoxHRZ.Value = BlockHRZ
'    Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
'    If Position >= 0 Then
'        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
'    Else
'        Range = 0#
'    End If
'substituted29.06.2010 by Function Range
    If BlockZRange > Range * 0.9 Then
        BlockZRange = Range * 0.9
    End If
    If Abs(BlockZOffset) > Range * 0.9 Then
        BlockZOffset = 0
    End If
    BSliderZOffset.Value = BlockZOffset
    BSliderZRange.Value = BlockZRange
    BSliderZStep.Value = BlockZStep

End Sub


'''''
' TODO: All block values should use the checkboxes directly
'''''
Public Sub GetBlockValues()
   
    BlockHighSpeed = CheckBoxHighSpeed.Value
    BlockLowZoom = CheckBoxLowZoom.Value
    HRZ = CheckBoxHRZ.Value ' this is for the piezo
    BlockZOffset = BSliderZOffset.Value
    BlockZRange = BSliderZRange.Value
    BlockZStep = BSliderZStep.Value

End Sub



Private Function TimeDisplay(Value As Double) As String         'Calculates the String to display in a "user frindly format". Value is in seconds
    Dim Hour, Min As Integer
    Dim Sec As Double

    Hour = Int(Value / 3600)                                        'calculates number of full hours                           '
    Min = Int(Value / 60) - (60 * Hour)                             'calculates number of left minutes
    Sec = (Fix((Value - (60 * Min) - (3600 * Hour)) * 100)) / 100   'calculates the number of left seconds
    If (Hour = 0) And (Min = 0) Then                                'Defines a "user friendly" string to display the time
        TimeDisplay = Sec & " sec"
    ElseIf (Hour = 0) And (Sec = 0) Then
        TimeDisplay = Min & " min"
    ElseIf (Hour = 0) Then
        TimeDisplay = Min & " min " & Sec
    Else
        TimeDisplay = Hour & " h " & Min
    End If
End Function


Public Function AcquisitionTime() As Double
    Dim Track1Speed, Track2Speed, Track3Speed, Track4Speed As Double
    Dim Pixels As Long
    Dim FrameNumber As Integer
    Dim ScanDirection As Integer
    Dim i As Integer
   
    Track1Speed = 0
    Track2Speed = 0
    Track3Speed = 0
    Track4Speed = 0
    If CheckBoxTrack1.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, Success)
        Track1Speed = Track.SampleObservationTime
    End If
    If CheckBoxTrack2.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
        Track2Speed = Track.SampleObservationTime
    End If
    If CheckBoxTrack3.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(2, Success)
        Track3Speed = Track.SampleObservationTime
    End If
    If CheckBoxTrack4.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(3, Success)
        Track4Speed = Track.SampleObservationTime
    End If
    Pixels = Lsm5.DsRecording.LinesPerFrame * Lsm5.DsRecording.SamplesPerLine
    FrameNumber = Lsm5.DsRecording.FramesPerStack
    If Lsm5.DsRecording.ScanDirection = 0 Then
        ScanDirection = 1
    Else
        ScanDirection = 2
    End If
    AcquisitionTime = (Track1Speed + Track2Speed + Track3Speed + Track4Speed) * Pixels * FrameNumber / ScanDirection * 3.3485
End Function



Private Sub CheckBoxTrack1_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub

Private Sub CheckBoxTrack2_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub

Private Sub CheckBoxTrack3_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub

Private Sub CheckBoxTrack4_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub



Public Function AutofocusTime() As Double
    Dim Speed As Double
    Dim Pixels As Long
    Dim FrameNumber As Integer
    Dim ScanDirection As Integer
    Dim i As Integer

    Speed = 0
    If CheckBoxHighSpeed.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, Success)
        Speed = 1.76 * 10 ^ -6
    Else
        If OptionButtonTrack1.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
        If OptionButtonTrack2.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
        If OptionButtonTrack3.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
        If OptionButtonTrack4.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
    End If
    Pixels = 512
    AutofocusForm.GetBlockValues
    FrameNumber = CLng(BlockZRange / BlockZStep) + 1
    If Lsm5.DsRecording.ScanDirection = 0 Then
        ScanDirection = 1
    Else
        ScanDirection = 2
    End If
    If CheckBoxHRZ.Value = True Then
        AutofocusTime = Speed * Pixels * FrameNumber * 3.3485 + 4.9
    Else
        AutofocusTime = Speed * Pixels * FrameNumber / ScanDirection * 3.3485 + 4.9
    End If
End Function



Public Sub CheckBoxHRZ_Change() 'I was trying to display the time needed for autofocus, single image acquisition and total time of the experiments, but I gave and and commented out those functions
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Public Sub CheckBoxHighSpeed_Change()  'I was trying to display the time needed for autofocus, single image acquisition and total time of the experiments, but I gave and and commented out those functions
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub BSliderZStep_Change()  'I was trying to display the time needed for autofocus, single image acquisition and total time of the experiments, but I gave and and commented out those functions
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack1_Click()
    If OptionButtonTrack1.Value Then 'if track 1 checked others are not autofocus track but false
        OptionButtonTrack2.Value = Not OptionButtonTrack1.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack1.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack1.Value
        CheckAutofocusTrack (1) 'sets SelectedTrack to 1, see below
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack2_Click()
    If OptionButtonTrack2.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack2.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack2.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack2.Value
        CheckAutofocusTrack (2)
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack3_Click()
    If OptionButtonTrack3.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack3.Value
        OptionButtonTrack2.Value = Not OptionButtonTrack3.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack3.Value
        CheckAutofocusTrack (3)
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack4_Click()
    If OptionButtonTrack4.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack4.Value
        OptionButtonTrack2.Value = Not OptionButtonTrack4.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack4.Value
        CheckAutofocusTrack (4)
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

''''''
'    CheckAutofocusTrack( SelectedTrack As Integer )
'    Checks whether the track that was selected for autofocusing only contains a single channel (alternetivly defines one of the checked channels)
'    and finds the name of the autofocusing channel
'       [SelectedTrack] In - Number of selected track
''''''
Private Sub CheckAutofocusTrack(SelectedTrack As Integer)
    Dim Track As DsTrack 'a new track is defined
    Dim DataChannel As DsDataChannel    'a new interface to a data channel is defined
                                        'contains channel dependend parameters of the
                                        'scan memory/display/calculation of scan data during scan operation
    Dim ActiveChannelNumber As Integer
    Dim AutofocusChannel As String
    Dim j As Integer
    
    Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(SelectedTrack - 1, Success)
        'gets the track object by multiplexorder which starts with 0
        'since selected track starts with 1 (see CheckAutofocusTrack (n)), 1 has to be substracted
        
    'the following loop will count the number of activated channels in the track chosen for autofocusing
    ActiveChannelNumber = 0
    
    For j = 0 To Track.DataChannelCount - 1 'gets number of channels that are potentially activatable in track
        Set DataChannel = Track.DataChannelObjectByIndex(j, Success) 'data channel corresponding to loop index is analysed
        If DataChannel.Acquire = True Then  'checks whether the data channel corresponding to loop index is activated
            ActiveChannelNumber = ActiveChannelNumber + 1 'counts the number of activated channels
            If ActiveChannelNumber = 1 Then AutofocusChannel = DataChannel.name 'Gets the name of the first activated channel
        End If
    Next
    
    If ActiveChannelNumber > 1 Then 'if more than one channel is activated...
        MsgBox ("The track you selected has more than one active Channel. " & AutofocusChannel & " will be used to calculate autofocus parameters.")
    End If
End Sub


Public Function TotalTimeLeft() As Double
    Dim Speed As Double
    Dim Pixels As Long
    Dim ScanDirection As Integer
    Dim i As Integer
    TotalTimeLeft = (AcquisitionTime + AutofocusTime + BlockTimeDelay) * (BlockRepetitions - RepetitionNumber + 1) - BlockTimeDelay
End Function

'''''
' this should all move to a single
Public Sub AutofocusSetting(HRZ As Boolean, HighSpeed As Boolean, ZStep As Double)
    
    If BlockLowZoom Then                                         'Changes the zoom if necessary
        Lsm5.DsRecording.ZoomX = 1
        Lsm5.DsRecording.ZoomY = 1
    End If
        
    Lsm5.DsRecording.TimeSeries = False                     'Disable the timeseries, because autofocussing is juste one image at one timepoint.
    
    If FrameAutofocussing Then                              'Setting the way the Stage is going to move in Z, plus speed and number of pixels
        
        Lsm5.DsRecording.ScanMode = "Stack"                 'This is defining to acquire a Z stack of Z-Y images
        
        If HRZ Then
            
            Lsm5.DsRecording.SpecialScanMode = "ZScanner"
        
        Else
    
            ' !!!!!!!!!!!! potential error source  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            ' improvement
            If SystemName = "LSM" Then
                Lsm5.DsRecording.SpecialScanMode = "FocusStep"
        
                'Lsm5.DsRecording.FrameSpacing = ZStep
                '    NoFrames = CLng(ZRange / ZStep) + 1
                '    Lsm5.DsRecording.FramesPerStack = NoFrames
                '    If NoFrames > 2048 Then
                '        NoFrames = 2048
                '    End If
                Lsm5.DsRecording.Sample0Z = ZStep * NoFrames / 2
            Else
                Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
                Lsm5.DsRecording.FramesPerStack = 1201
                Lsm5.DsRecording.Sample0Z = Range / 2
                Lsm5.DsRecording.FrameSpacing = Range / 1200
                Sleep (100)
            End If
                
        End If
        
        
        If HighSpeed Then
            Lsm5.DsRecording.ScanDirection = 1                  'If Highspeed is selected it uses the bidirectionnal scanning
        End If
        If SystemName = "LIVE" Then
            Lsm5.DsRecording.RtRegionWidth = BSliderFrameSize.Value 'If doing frame autofocussing it uses the userdefined frame size
            Lsm5.DsRecording.RtBinning = 512 / BSliderFrameSize.Value
            Lsm5.DsRecording.RtRegionHeight = BSliderFrameSize.Value
        ElseIf SystemName = "LSM" Then
            Lsm5.DsRecording.SamplesPerLine = BSliderFrameSize.Value  'If doing frame autofocussing it uses the userdefined frame size
            'Lsm5.DsRecording.RtBinning = 4
            Lsm5.DsRecording.LinesPerFrame = BSliderFrameSize.Value
        Else
            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
        Exit Sub
        End If
    
    
    Else  ' Not FrameAutoFocussing
        
        Lsm5.DsRecording.ScanMode = "ZScan"                     'This is defining to acquire a single X-Z image, like what is done with the "Range" button in the LSM ScanControl window
        If HRZ Then
        
            Lsm5.DsRecording.SpecialScanMode = "ZScanner"
            If SystemName = "LIVE" Then
                Lsm5.DsRecording.RtLinePeriod = 1 / 1000 'BSliderScanSpeed.Value
                Lsm5.DsRecording.RtRegionWidth = 512
                Lsm5.DsRecording.RtRegionHeight = 1
            ElseIf SystemName = "LSM" Then
                'MsgBox "HRZ LSM 256"
                Lsm5.DsRecording.SamplesPerLine = 256
                Lsm5.DsRecording.LinesPerFrame = 1
            Else
                MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
                Exit Sub
            End If
           
        Else ' Not HRZ
        
            'Lsm5.DsRecording.SpecialScanMode = "FocusStep" ' I this does not work, use "FocusStep"
            Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
            
        End If
        
        
        If SystemName = "LIVE" Then
            'TODO: Legacy code
            Lsm5.DsRecording.RtLinePeriod = 1 / 1000 'BSliderScanSpeed.Value
            Lsm5.DsRecording.RtRegionWidth = 512
            Lsm5.DsRecording.RtRegionHeight = 1
            
       
        ElseIf SystemName = "LSM" Then
                Lsm5.DsRecording.SamplesPerLine = 256
                Lsm5.DsRecording.LinesPerFrame = 1
        Else
                MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
                Exit Sub
        End If
        
    End If
    
    
    Sleep (100)
    
End Sub





'''''
'   ChangeButtonStatus(Enable As Boolean)
'   Reset status of buttons on rightside of form
'''''
Private Sub ChangeButtonStatus(Enable As Boolean)
    StartButton.Enabled = Enable
    StartBleachButton.Enabled = Enable
    CloseButton.Enabled = Enable
    ReinitializeButton.Enabled = Enable
End Sub

'''''
'   FailSafeMoveStage(Optional Mark As Integer = 0)
'   Moves stage and wait till it is finished
'       [x] In - x-position
'       [y] In - y-position
'       [z] In - z-position (this is optional)
'''''
Private Function FailSafeMoveStage(x As Double, y As Double, Optional z As Double = -10000) As Boolean

    If z <> -10000 Then ' also sets first move down
         Lsm5.Hardware.CpFocus.Position = z + ZBacklash  ' move backward and then forward again (this should be apparently better than direct movement)
    End If
    'Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition x, y
    Lsm5.Hardware.CpStages.SetXYPosition x, y
    'TODO Check this
    Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
        Sleep (200)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                FailSafeMoveStage = False
                StopAcquisition
                Exit Function
            End If
        End If
    Loop
    
    If z <> -10000 Then ' move to actual position
        Lsm5.Hardware.CpFocus.Position = z
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy      'Waits that the objective movement is finished
           Sleep (20)
           DoEvents
        Loop
    End If
    FailSafeMoveStage = True
    
End Function


'''''
'   MoveToNextLocation(Optional Mark As Integer = 0)
'   Moves to next location as set in the stage (mark)
'   Default will cycle through all positions sequentially starting from actual position
'       [Mark] In - Number of position where to move.
'''''
Private Sub MoveToNextLocation(Optional Mark As Integer = 0)
        Dim Markcount As Long
        Dim count As Long
        Dim idx As Long
        Dim dX As Double
        Dim dY As Double
        Dim dZ As Double
        Dim i As Integer
        Lsm5.Hardware.CpStages.MarkMoveToZ (0)
        'Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ (0)  'old code Moves to the first location marked in the stage control. How to move to next point?
        ' the points were deleted and readded at the end of list in the Acquisition function
        'TODO: Check code
        Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
            Sleep (100)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    StopAcquisition
                    Exit Sub
                End If
            End If
        Loop
End Sub


'''''
' Sub StopScanCheck()
' This stop all running scans. Check is the wrong name
'''''
Private Sub StopScanCheck()
    Lsm5.StopScan
    DoEvents
End Sub

'''''
'   UpdateZvalues(Grid, MultipleLocation, z)
'   Adds a ZShift to all positions from MultiLocation (the ZShift is the one determined by the Autofocus)
'''''
Private Sub UpdateZvalues(Grid, MultipleLocation, z)
        
        
        Dim idpos As Integer
        Dim Sucess As Integer
   
        If Grid Or MultipleLocationToggle.Value Then
            
            Lsm5.Hardware.CpStages.MarkClearAll
            For idpos = 1 To GlobalPositionsStage
                GlobalZpos(idpos) = GlobalZpos(idpos) + ZShift
                Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
               
            Next idpos
    
        
        Else  ' Todo: what is this doing?

        '      GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
       '     For idpos = 1 To GlobalPositionsStage
        '        Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).GetMarkZ(0, x, y, z)
         '       Success = Lsm5.Hardware.CpStages.MarkClear(0)
                    
         '       z = z + ZShift
          '      Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ x, y, z
          '  Next idpos
             
         End If
    
End Sub



Private Sub CreateZoomDatabase(ZoomDatabaseName, HighResExperimentCounter, ZoomExpname)
            'Create ZoomDatabase
            Dim Start As Integer
            Dim bslash As String
            Dim pos As Long
            Dim NameLength As Long
            Dim Mypath As String
            
            Start = 1
            bslash = "\"
            pos = Start
            Do While pos > 0
                pos = InStr(Start, GlobalDataBaseName, bslash)
                If pos > 0 Then
                    Start = pos + 1
                End If
            Loop
            
            Mypath = GlobalDataBaseName + bslash
            NameLength = Len(GlobalDataBaseName)
            ZoomExpname = Strings.Right(GlobalDataBaseName, NameLength - Start + 1)
           ' NameLength = Len(Myname)
           ' Myname = Strings.Left(Myname, NameLength - 4)
            ZoomDatabaseName = Mypath & ZoomExpname & "_" & GlobalFileName & LocationName & "_R" & RepetitionNumber & "_Exp" & HighResExperimentCounter & "_zoom"
            ' Lsm5.NewDatabase (ZoomDatabaseName)
           ' ZoomDatabaseName = ZoomDatabaseName & "\" & Myname & "_zoom.mdb"
    
End Sub

Private Sub CreateAlterImageDatabase(AlterDatabaseName, Mypath)
        Dim Start As Integer
        Dim bslash As String
        Dim pos As Long
        Dim NameLength As Long
        Dim Myname As String

         Start = 1
         bslash = "\"
         pos = Start
         Do While pos > 0
             pos = InStr(Start, GlobalDataBaseName, bslash)
             If pos > 0 Then
                 Start = pos + 1
             End If
         Loop
         Mypath = Strings.Left(GlobalDataBaseName, Start - 1)
         NameLength = Len(GlobalDataBaseName)
         Myname = Strings.Right(GlobalDataBaseName, NameLength - Start + 1)
         NameLength = Len(Myname)
         ' Myname = Strings.Left(Myname, NameLength - 4)
         AlterDatabaseName = Mypath & Myname & "_additionalTracks"
        ' Lsm5.NewDatabase (AlterDatabaseName)
        '  AlterDatabaseName = AlterDatabaseName & "\" & Myname & "_additionalTracks"
         
End Sub

''''''
'   MicroscopePilot(RecordingDoc As DsRecordingDoc, BleachingActivated As Boolean, HighResExperimentCounter As Integer, HighResCounter As Integer _
'   HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double)
'   TODO: test stricter way of passing arguments
''''''
Private Sub MicroscopePilot(RecordingDoc As DsRecordingDoc, ByVal BleachingActivated As Boolean, HighResExperimentCounter As Integer, HighResCounter As Integer _
, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double)
    
    Dim ZoomNumber As Integer
    Dim code As String
    Dim codeArray() As String
        
    ' Get Code from Windows registry
    code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="code")

    Do While (code = "1" Or code = "0")
        ' TODO: Check Code
        DisplayProgress "Waiting for Micropilot...", RGB(0, &HC0, 0)
        Sleep (100)
        code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", _
                  key:="Code")
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StopAcquisition
                Exit Sub
            End If
        End If
    Loop
    
    'MsgBox ("Code = " + code)
    
    DisplayProgress "Received Code " + CStr(code), RGB(0, &HC0, 0)
    
    'TODO: create a better procedure to check for cells
    If (CheckBoxGridScan_FindGoodPositions) Then
        
        codeArray = Split(code, "_")
        
        nGoodCells = CInt(codeArray(1))
        minGoodCellsPerImage = CInt(codeArray(2))
        minGoodCellsPerWell = CInt(codeArray(3))
    
        'MsgBox "nGoodCellsPerWell = " + CStr(nGoodCellsPerWell)
    
        GoTo Mark
    
    End If
    

    If code = "2" Then   ' no interesting cell
    
        DisplayProgress "Micropilot Code 2", RGB(0, &HC0, 0)
        SaveSetting "OnlineImageAnalysis", "macro", "Refresh", 0
        'SaveSetting "OnlineImageAnalysis", "Cinput", "Code", 0
        'If RecordingDoc.IsValid Then   ' window is closed later anyway
        '    RecordingDoc.CloseAllWindows
        '    Set RecordingDoc = Nothing
        'End If
        GoTo Mark '(because Image does not show any interesting pheotype)
    
    ElseIf code = "4" Then   'store position in a list
    
        DisplayProgress "Micropilot Code 4", RGB(0, &HC0, 0)
        HighResCounter = HighResCounter + 1 ' Counts the postions, where Highres Imaging will be carried out
        ' store postion from windows registry in array
        StorePositioninHighResArray HighResArrayX, HighResArrayY, HighResArrayZ, HighResCounter
        
    ElseIf code = "5" Then ' start Highres Batch Imaging 1 to n postions
        
        DisplayProgress "Micropilot Code 5", RGB(0, &HC0, 0)
        HighResCounter = HighResCounter + 1
        ' store postion from windows registry in array
       
        StorePositioninHighResArray HighResArrayX, HighResArrayY, HighResArrayZ, HighResCounter
        ' BatchHighresImagingRoutine
        HighResExperimentCounter = HighResExperimentCounter + 1 ' counts the number of highres-multipositionexperiments (important for naming the datafolder)
        
        ' HERE THE IMAGES ARE ACQUIRED
        BatchHighresImagingRoutine RecordingDoc, HighResArrayX, HighResArrayY, HighResArrayZ, HighResCounter, HighResExperimentCounter
        'After the whole MultiposExperiment HighResCounter must be set to 0 again
        HighResCounter = 0
        ReDim HighResArrayX(100)
        ReDim HighResArrayY(100)
        ReDim HighResArrayZ(100)
    
    Else
        
        'Error Message "OnlineImageAnalysis Value = 'Code'"
        MsgBox ("Invalid OnlineImageAnalysis Code = " & code)
    
    End If
    
    'Reset Code to 0 in Windows Registry
    'SaveSetting "OnlineImageAnalysis", "Cinput", "Code", 0
      
Mark:
        
            
End Sub

'''''
'   Private Sub StartAlternativeImaging(RecordingDoc As DsRecordingDoc, StartTime As Double, _
'   AlterDatabaseName As String, name As String)
'   Alternative Acquisition in every .. round
'   TODO: What are all the parameters
'''''
Private Sub StartAlternativeImaging(RecordingDoc As DsRecordingDoc, StartTime As Double, _
filepath As String, name As String)
    If RepetitionNumber Mod TextBox_RoundAlterTrack = 0 Then
        Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
        If RecordingDoc Is Nothing Then
            Set RecordingDoc = Lsm5.NewScanWindow
            While RecordingDoc.IsBusy
                Sleep (20)
                DoEvents
            Wend
        End If
        DisplayProgress "Acquiring Additional Track...", RGB(0, &HC0, 0)
         
        ActivateAlterAcquisitionTrack
         Sleep (100)
              
         If Not IsAcquisitionTrackSelected Then      'An additional control....
             MsgBox "No track selected for Acquisition! Cannot Acquire!"
             StopAcquisition
             Exit Sub
         End If
                
         'MsgBox "Piezo Position = " + CStr(Lsm5.Hardware.CpHrz.Position)
         '= 0  ' Center Piezo
         'Sleep (100)
         
         
         ' get and set the values from the GUI
         Lsm5.DsRecording.ZoomX = TextBoxAlterZoom.Value
         Lsm5.DsRecording.ZoomY = TextBoxAlterZoom.Value
         Lsm5.DsRecording.FramesPerStack = TextBoxAlterNumSlices.Value
         Lsm5.DsRecording.FrameSpacing = TextBoxAlterInterval.Value
         If Lsm5.DsRecording.FramesPerStack > 1 Then
            'Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * Int(Lsm5.DsRecording.FramesPerStack / 2) ' maybe necessary for non-piezo
            Lsm5.DsRecording.SpecialScanMode = "ZScanner" ' this is a problem if people do not have a piezo
            Lsm5.DsRecording.ScanMode = "Stack"
         End If
         
         'MsgBox "all settings set   " + CStr(Lsm5.DsRecording.Sample0Z)
        
         ' take the image
         ScanToImageNew RecordingDoc
        'TODO Check this
         While AcquisitionController.IsGrabbing
            Sleep (100)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    StopAcquisition
                    Exit Sub
                End If
            End If
         Wend
         
         RecordingDoc.SetTitle name
        
         SaveDsRecordingDoc RecordingDoc, filepath
         
         'Lsm5.DsRecording.Sample0Z = SampleOZold
      End If
End Sub

'''
'   StorePositioninHighResArray(HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, HighResCounter As Integer)
'   TODO: Test stricter way of passing arguments
''''
Private Sub StorePositioninHighResArray(HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, HighResCounter As Integer)
    
    ' store postion from windows registry in array
    
    Dim zoomXoffset As Double
    Dim zoomYoffset As Double
    Dim x As Double
    Dim y As Double
    Dim PixelSize As Double

    'zoomXoffset = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="offsetx")
    'zoomYoffset = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="offsety")
    
    zoomXoffset = CDbl(GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="offsetx"))
    zoomYoffset = CDbl(GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="offsety"))
    
    
    'MsgBox ("zoomXoffset,zoomYoffset " + CStr(zoomXoffset) + "," + CStr(zoomYoffset))
    
    
    If HRZ Then
        Success = Lsm5.Hardware.CpHrz.Leveling   'This I think puts the HRZ to its resting position, and moves the focuswheel correspondingly
    Else
        ' do nothing
    End If
                
    Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
        Sleep (100)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StopAcquisition
                Exit Sub
            End If
        End If
    Loop

    'Move x,y,
     
    PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
    x = Lsm5.Hardware.CpStages.PositionX
    y = Lsm5.Hardware.CpStages.PositionY
    
    'MsgBox ("PixelSize " + CStr(PixelSize))
    'MsgBox ("zoomXoffset*ps,zoomYoffset*ps " + CStr(zoomXoffset * PixelSize) + "," + CStr(zoomYoffset * PixelSize))
    
    
    HighResArrayX(HighResCounter) = x - zoomXoffset * PixelSize
    HighResArrayY(HighResCounter) = y + zoomYoffset * PixelSize
    HighResArrayZ(HighResCounter) = Lsm5.Hardware.CpFocus.Position
   ' MsgBox "Current Z Position = " + CStr(Lsm5.Hardware.CpFocus.Position)
    DisplayProgress "Micropilot - Position stored", RGB(0, &HC0, 0)

End Sub


'''''
'   BatchHighresImagingRoutine(RecordingDoc As DsRecordingDoc, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, _
'   HighResCounter As Integer, HighResExperimentCounter As Integer)
'   TODO: Test stricter way of passing arguments
'''''
Private Sub BatchHighresImagingRoutine(RecordingDoc As DsRecordingDoc, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, _
HighResCounter As Integer, HighResExperimentCounter As Integer)
    
    Dim PixelSize As Double
    Dim Succes As Integer
    Dim ZoomExpname As String
    Dim ZoomImageIndex() As Long
    ReDim Preserve ZoomImageIndex(10000)
    Dim zoomname As String
    Dim ZoomDatabaseName As String
    'Timer and Looping Variables
    Dim highrespos As Integer
    Dim ZoomTimeDelay As Long  ' this seems to be an interval rather than delay
    Dim ZoomRepetitions As Integer
    Dim ZoomRepetitionNumber As Integer
    Dim ZoomRunning As Boolean
    Dim ZoomStartTime As Double
    Dim ZoomNewTime As Double
    Dim Zoomdifftime As Double
    
    Dim fullpathname As String
    
     
    ' set up the imaging
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    'Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
    
    If RecordingDoc Is Nothing Then
        Set RecordingDoc = Lsm5.NewScanWindow
        While RecordingDoc.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
    
    
    
    'Create Database ' own folder for each new BatchHighres Experiment !
    CreateZoomDatabase ZoomDatabaseName, HighResExperimentCounter, ZoomExpname
    
    ' Set parameters for time loop
    If BleachingActivated Then
        ZoomRepetitions = 1 ' do everything in one go
    Else
        ZoomRepetitions = TextBoxZoomCycles.Value
    End If
                  
    ZoomTimeDelay = TextBoxZoomCycleDelay.Value
    ZoomRepetitionNumber = 1
    ZoomRunning = True ' We are in this loop till all repetitions are done

    Do While ZoomRunning = True ' We are in this loop till all repetitions are done (timerepetitions loop)
        
        'MsgBox "HighResCounter " + CStr(HighResCounter)
        
        For highrespos = 1 To HighResCounter ' Postition loop
        
                ' Move to Positon in x,y
                DisplayProgress "Micropilot Code 5 - Move to Position", RGB(0, &HC0, 0)
                
                x = Lsm5.Hardware.CpStages.PositionX
                y = Lsm5.Hardware.CpStages.PositionY
                'MsgBox ("x,y " + CStr(x) + "," + CStr(y) + "dx,dy" + CStr(HighResArrayX(highrespos)) + "," + CStr(HighResArrayY(highrespos)))
                
                Succes = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(HighResArrayX(highrespos), HighResArrayY(highrespos))
                'TODO: Check
                Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                    Sleep (100)
                    If GetInputState() <> 0 Then
                        DoEvents
                        If ScanStop Then
                            StopAcquisition
                            Exit Sub
                        End If
                    End If
                Loop
        
                ' Move to Positon in z
                ' MsgBox "HighResArrayZ(highrespos) " + CStr(HighResArrayZ(highrespos))
                ' MsgBox "ZBacklash " + CStr(ZBacklash)
                
                Lsm5.Hardware.CpFocus.Position = HighResArrayZ(highrespos) + ZBacklash 'Move down 50um (=ZBacklash) below the position of the offset
                Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
                     Sleep (20)  '20ms
                     DoEvents
                Loop
                Lsm5.Hardware.CpFocus.Position = HighResArrayZ(highrespos)          'Moves up to the position of the offset
                Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
                    Sleep (20)
                    DoEvents
                Loop
                
                'Autofocus. This does an extra Autofocus also for the HighresImaging
                If CheckBoxZoomAutofocus.Value = True Then
                    BlockZOffset = TextBoxZoomAutofocusZOffset.Value
                    DisplayProgress "Micropilot Code 5 - Do Autofocus", RGB(0, &HC0, 0)
                    Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset, RecordingDoc
                    Autofocus_MoveAcquisition BlockZOffset
                End If
        
                ' Load AcquisitionSettings
                Lsm5.DsRecording.SamplesPerLine = TextBoxZoomFrameSize.Value
                Lsm5.DsRecording.LinesPerFrame = TextBoxZoomFrameSize.Value
                Sleep (100)
                ActivateZoomTrack
                Lsm5.DsRecording.ZoomX = TextBoxZoom.Value
                Lsm5.DsRecording.ZoomY = TextBoxZoom.Value
                
                If BleachingActivated Then
                                
                    DisplayProgress "Bleaching...", &HFF00FF
                        
                    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
                    If Success Then
                        Track.Acquire = True
                        Lsm5.DsRecording.TimeSeries = True
                        Lsm5.DsRecording.StacksPerRecord = TextBoxZoomCycles.Value
                        Track.TimeBetweenStacks = TextBoxZoomCycleDelay.Value
                        'MsgBox "Track.IsBleachTrack " + CStr(Track.IsBleachTrack)
                        'MsgBox "BleachScanNumber " + CStr(Track.BleachScanNumber)
                        DoEvents
                        Track.UseBleachParameters = True            'Bleach parameters are lasers lines, bleach iterations... stored in the bleach control window
                        'BleachStartTable(RepetitionNumber) = GetTickCount      'Get the time right before bleach to store this in the image metadata
                                                               
                        'ScanToImageNew RecordingDoc
    
                        'While AcquisitionController.IsGrabbing
                        '    Sleep (20)
                        '    If ScanStop Then
                        '        Lsm5.StopScan
                        '        'ScanStop = True
                        '        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        '        Exit Sub
                        '    End If
                        '    DoEvents
                        'Wend
                    
                        
                        Set RecordingDoc = Lsm5.StartScan
                        'TODO Check
                        Do While RecordingDoc.IsBusy
                            Sleep (100)
                            If GetInputState() <> 0 Then
                                DoEvents
                                If ScanStop Then
                                StopAcquisition
                                Exit Sub
                            End If
                        End If
                        Loop
                        
                        Lsm5.tools.WaitForScanEnd False, 10
                                                   
                        
                        Track.UseBleachParameters = False  'switch off the bleaching
                        Lsm5.DsRecording.TimeSeries = False
                        
                    Else
                    
                        MsgBox ("Could not set bleach track. Did not bleach.")
                    
                    End If
                
                                 
                    'Save Image  ' modified by Tischi
                    zoomname = GlobalFileName & LocationName & "_R" & RepetitionNumber & "_Exp_" & HighResExperimentCounter & "_MP" & highrespos & "_Bleach"
                    
        
                    fullpathname = ZoomDatabaseName & "\" & zoomname & ".lsm"
                    SaveDsRecordingDoc RecordingDoc, fullpathname
        
                    DisplayProgress "Micropilot Code 5 - SaveImage", RGB(0, &HC0, 0)
                    'If RecordingDocNew.IsValid Then
                    '    RecordingDocNew.CloseAllWindows
                    '    Set RecordingDoc = Nothing
                    'End If
                    
                    
                Else ' normal acquistion (non bleaching mode)
                    
                    Lsm5.DsRecording.FramesPerStack = TextBoxZoomNumSlices.Value
                    Lsm5.DsRecording.FrameSpacing = TextBoxZoomInterval.Value
                    
                    'preliminary take it out and make it better
                    Lsm5.DsRecording.ScanMode = "Stack"
                    Lsm5.DsRecording.SpecialScanMode = "ZScanner"
                
                    'Acquisition
                    DisplayProgress "Micropilot Code 5 - Start Scan", RGB(0, &HC0, 0)
                    If highrespos = 1 Then
                        ZoomStartTime = CDbl(GetTickCount) * 0.001
                    End If
                    DisplayProgress "Acquisition HighRes Position " & highrespos, RGB(&HC0, &HC0, 0)
                    
                    
                    ScanToImageNew RecordingDoc
                    'TODO Check
                    While AcquisitionController.IsGrabbing
                        Sleep (100)
                        If GetInputState() <> 0 Then
                            DoEvents
                            If ScanStop Then
                                StopAcquisition
                                Exit Sub
                            End If
                        End If
                    Wend
                                    
                    'Set RecordingDocNew = Lsm5.StartScan
                    'Do While RecordingDocNew.IsBusy
                    '   If ScanStop Then
                    '        Lsm5.StopScan
                    '        StopAcquisition
                    '        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                    '        Exit Sub
                    '    End If
                    '    DoEvents
                    '    Sleep (5)
                    'Loop
                    
                    Lsm5.tools.WaitForScanEnd False, 10
            
                    'Save Image ' Tischi: changed filename such that it can be traced back to the correspoding location
                    zoomname = GlobalFileName & LocationName & "_R" & RepetitionNumber & "_Exp_" & HighResExperimentCounter & "_MP" & highrespos & "_R" & ZoomRepetitionNumber
                    
                    fullpathname = ZoomDatabaseName & "\" & zoomname & ".lsm"
                    SaveDsRecordingDoc RecordingDoc, fullpathname
        
                    DisplayProgress "Micropilot Code 5 - SaveImage", RGB(0, &HC0, 0)
                    
                    
                    ' Tischi: Here the Location-tracking code needs be added!
                    ' and these variable need to be updated!
                    
                    'If LocationTracking_HighRes Then 'This is if we're doing some postacquisition tracking
                
                     '   DisplayProgress "Analysing the new position of location " & Location, &H80FF&
                     '   DoEvents
                     '   MassCenter ("Tracking")
                     '   XCor = XMass
                     '   YCor = YMass
                     '   If TrackZ Then
                     '       ZCor = ZMass
                     '   Else
                     '   If HRZ Then
                     '       ZCor = 0
    '                '        Success = Lsm5.Hardware.CpHrz.Leveling
                     '   Else
                     '       ZCor = 0
                     '   End If
                    'End If
                    '''''changed
                    'If AreStageCoordinateExchanged Then
                    '    XCor = YMass
                    '    YCor = XMass
                    'End If
                    '''changed
                
                    
                    'HighResArrayX (highrespos) = HighResArrayX (highrespos) + XCor
                    'HighResArrayY (highrespos) = HighResArrayY (highrespos) - YCor
                    'HighResArrayZ (highrespos) = HighResArrayZ (highrespos) - ZCor
                    
                    
                    ' LocationTracking HighRes End -----------
                    
                    
                    'If RecordingDocNew.IsValid Then
                    '    RecordingDocNew.CloseAllWindows
                    '    Set RecordingDoc = Nothing
                    'End If
                    
                End If ' Bleaching

                        
                        
                
        Next highrespos ' End of postions loop
        
        
    
        If ZoomRepetitionNumber < ZoomRepetitions Then
            ZoomNewTime = CDbl(GetTickCount) * 0.001
            Zoomdifftime = ZoomNewTime - ZoomStartTime
            'TODO Check
            Do While Zoomdifftime <= ZoomTimeDelay
                Sleep (10)
                If GetInputState() <> 0 Then
                    DoEvents
                    If ScanStop Then
                        StopAcquisition
                        Exit Sub
                    End If
                End If
                ZoomNewTime = CDbl(GetTickCount) * 0.001
                Zoomdifftime = ZoomNewTime - ZoomStartTime
                DisplayProgress "Waiting " & CStr(CInt(ZoomTimeDelay - Zoomdifftime)) + " s before scanning repetition  " & (ZoomRepetitionNumber + 1), RGB(&HC0, &HC0, 0)
            Loop
        Else
            ZoomRunning = False ' now all repetitions are done, so  we leave the do while zoomrunnning = true loop
        End If
        ZoomRepetitionNumber = ZoomRepetitionNumber + 1
    
       
    Loop  ' End of time repetition loop
    
    
    
End Sub



' Copied and adapted from MultiTimeSeries macro
Public Function SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String) As Boolean
    Dim Export As AimImageExport
    Dim image As AimImageMemory
    Dim Error As AimError
    Dim Planes As Long
    Dim Plane As Long
    Dim Horizontal As enumAimImportExportCoordinate
    Dim Vertical As enumAimImportExportCoordinate

    On Error GoTo Done

    'Set Image = EngelImageToHechtImage(Document).Image(0, True)
    If Not Document Is Nothing Then
        Set image = Document.RecordingDocument.image(0, True)
    End If
    
    Set Export = Lsm5.CreateObject("AimImageImportExport.Export.4.5")
'    Set Export = New AimImageExport
    Export.FileName = FileName
    Export.Format = eAimExportFormatLsm5
    Export.StartExport image, image
    Set Error = Export
    Error.LastErrorMessage
    
    Planes = 1
    Export.GetPlaneDimensions Horizontal, Vertical
    
    Select Case Vertical
        Case eAimImportExportCoordinateY:
             Planes = image.GetDimensionZ * image.GetDimensionT
        Case eAimImportExportCoordinateZ:
            Planes = image.GetDimensionT
    End Select
    
    'TODO check. what happens here with Export.ExportPlane Nothing why Nothing (thumbnails)
    For Plane = 0 To Planes - 1
        If GetInputState() <> 0 Then
            DoEvents
             If ScanStop Then
                SaveDsRecordingDoc = False
                Export.FinishExport
                StopAcquisition
                Exit Function
            End If
        End If
        Export.ExportPlane Nothing
    Next Plane
    Export.FinishExport
    SaveDsRecordingDoc = True
    Exit Function
    
Done:
    MsgBox "Check Temporary Files Folder! Cannot Save Temporary File(s)!"
    ScanStop = True
    SaveDsRecordingDoc = False
    Export.FinishExport
    StopAcquisition
    
End Function


