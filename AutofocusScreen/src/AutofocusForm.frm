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
' AutofocusScreen_ZEN_v2.0.3
'''''''''''''''''''''End: Version Description'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const Version = " v2.0.3"
Public LogFile As TextStream 'This is the file where a log of the procedure is saved
Public LogFileName As String
Public FileSystem As FileSystemObject
Public Log     As Boolean          'If true we log data during the macro
Public posTempZ  As Double                  'This is position at start after pushing AutofocusButton
Public CntAutofocusButtonTest As Integer
Public DebugCode As Boolean




''''''
' UserForm_Initialize()
'   Function called from e.g. AutoFocusForm.Show
'   Load and initialize form
''
Private Sub UserForm_Initialize()
    'Setting of some global variables
    LogFileName = "Z:\AntonioP\Code\AutomatedMicroscopy\ZeissMacro\AutofocusScreen\AutofocusLog"  'you can't write on root directory!
    Log = True
    DebugCode = True
    
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
    Dim bLSM As Boolean
    Dim bLIVE As Boolean
    Dim bCamera As Boolean

    
    Set tools = Lsm5.tools
    GlobalMacroKey = "Autofocus"
    
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
    
    ' Set standard values for Autofocus
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
    
    'TODO: Check if GUI is available (ZEN2011 onward). How do you do this!!

    'Set default value
    ScanLineToggle.Value = True
    BSliderZOffset.Value = 0
    BSliderZRange.Value = 80
    BSliderZStep.Value = 0.1
    CheckBoxLowZoom.Value = False
    CheckBoxActiveAutofocus.Value = True
    
    'Set standard values for Post-Acquisition tracking
    TrackingToggle.Value = False
    SwitchEnableTrackingToggle (False)
    TrackingToggle.Enabled = False
    
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
    DatabaseTextbox.Value = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="OutputFolder")
    
    'Set repetition and locations
    RepetitionNumber = 1
    locationNumber = 1
    
    'If we log a new logfile is created
    If Log Then
        Set FileSystem = New FileSystemObject
        Dim i As Integer
        i = 0
        While FileSystem.FileExists(LogFileName & i & ".txt")
            i = i + 1
        Wend
        LogFileName = LogFileName & i & ".txt"
        Set LogFile = FileSystem.OpenTextFile(LogFileName, 8, True)
        LogFile.Close
    End If
    If DebugCode Then
        RunTests.Visible = True
    Else
        RunTests.Visible = False
    End If
    Re_Initialize
    
    
 
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
    
    'This sets standard values for all task we want to do. This will be changed by the macro
    
    If CheckBoxHRZ Then
        Lsm5.Hardware.CpHrz.Leveling
        While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
    posTempZ = Lsm5.Hardware.CpFocus.Position


    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2
    
    Set GlobalAutoFocusRecording = Lsm5.CreateBackupRecording
    Set GlobalAcquisitionRecording = Lsm5.CreateBackupRecording
    Set GlobalZoomRecording = Lsm5.CreateBackupRecording
    Set GlobalAltRecording = Lsm5.CreateBackupRecording
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    GlobalAutoFocusRecording.Copy Lsm5.DsRecording
    GlobalAcquisitionRecording.Copy Lsm5.DsRecording
    GlobalZoomRecording.Copy Lsm5.DsRecording
    GlobalAltRecording.Copy Lsm5.DsRecording
    GlobalBackupRecording.Copy Lsm5.DsRecording ' this will not be changed remains always the same
    GlobalBackupSampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    Dim i As Long
    Dim NrTracks As Long
    ReDim GlobalBackupActiveTracks(Lsm5.DsRecording.TrackCount)
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       GlobalBackupActiveTracks(i) = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i
    
    'set a reference Z position

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
    CheckBoxAutofocusTrackZ.Enabled = Enable
    CheckBoxAutofocusTrackXY.Enabled = Enable
End Sub

''''''
'   BSliderZOffset_Change()
'   BSliderZOffset is the offset added after AF
''''''
Private Sub BSliderZOffset_Change()
    'make range checks
     If Abs(BSliderZOffset.Value) > Range() * 0.9 Then
            BSliderZOffset.Value = 0
            MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um"
    End If
End Sub

''''''
'   BSliderZRange_Change()
'   Set the range in um during AF
''''''
Private Sub BSliderZRange_Change()    ' It should be possible to change the limit of the range to bigger values than half of the working distance
    If BSliderZRange.Value > Range * 0.9 Then 'make range checks
            BSliderZRange.Value = Range * 0.9
            MsgBox "ZRange has to be less or equal to the working distance of the objective: " + CStr(Range) + " um"
    End If
End Sub


'''''
'   CheckZRanges()
'   Check if Z movements are in agreement with range of microscope
'''''
Public Function CheckZRanges() As Boolean
    If Range() = 0 Then
        MsgBox "Objective's working distance not defined! Cannot Autofocus!"
        CheckZRanges = False
        Exit Function
    Else
        CheckZRanges = True
    End If
    
    If BSliderZRange.Value > Range() * 0.9 Then 'this is already tested in the slider could be removed
        AutofocusForm.BSliderZRange.Value = Range() * 0.9
        MsgBox "Autofocus range is too large! Has been reduced to " + Str(AutofocusForm.BSliderZRange.Value)
    End If
    If Abs(BSliderZOffset.Value) > Range() * 0.9 Then 'this is already tested in the slider could be removed
        AutofocusForm.BSliderZOffset = 0
        MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um. Has been put back to " + Str(AutofocusForm.BSliderZOffset)
    End If
    
End Function
  
''''''
'   ************************
'   The tracks for Autofocus
''''''
Private Sub OptionButtonTrack1_Click()
    If OptionButtonTrack1.Value Then 'if track 1 checked others are not autofocus track but false
        OptionButtonTrack2.Value = Not OptionButtonTrack1.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack1.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack1.Value
        CheckAutofocusTrack (1) 'sets SelectedTrack to 1, see below
    End If
End Sub

Private Sub OptionButtonTrack2_Click()
    If OptionButtonTrack2.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack2.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack2.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack2.Value
        CheckAutofocusTrack (2)
    End If
End Sub

Private Sub OptionButtonTrack3_Click()
    If OptionButtonTrack3.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack3.Value
        OptionButtonTrack2.Value = Not OptionButtonTrack3.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack3.Value
        CheckAutofocusTrack (3)
    End If
End Sub

Private Sub OptionButtonTrack4_Click()
    If OptionButtonTrack4.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack4.Value
        OptionButtonTrack2.Value = Not OptionButtonTrack4.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack4.Value
        CheckAutofocusTrack (4)
    End If
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
    TextBoxZoomAutofocusZOffset.Value = BSliderZOffset.Value
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

    If CheckBoxActiveOnlineImageAnalysis.Value Then
        CheckBoxGridScan_FindGoodPositions.Enabled = Enable
    Else
        CheckBoxGridScan_FindGoodPositions.Enabled = False
    End If
    GridScan_posLabel.Enabled = Enable
    GridScan_nLabel.Enabled = Enable
    GridScan_nColumnLabel.Enabled = Enable
    GridScan_nRowLabel.Enabled = Enable
    GridScan_nColumn.Enabled = Enable
    GridScan_nRow.Enabled = Enable
    GridScan_dLabel.Enabled = Enable
    GridScan_dColumnLabel.Enabled = Enable
    GridScan_dRowLabel.Enabled = Enable
    GridScan_dColumn.Enabled = Enable
    GridScan_dRow.Enabled = Enable
    GridScan_subLabel.Enabled = Enable
    GridScan_nsubLabel.Enabled = Enable
    GridScan_nColumnsubLabel.Enabled = Enable
    GridScan_nRowsubLabel.Enabled = Enable
    GridScan_nColumnsub.Enabled = Enable
    GridScan_nRowsub.Enabled = Enable
    GridScan_dsubLabel.Enabled = Enable
    GridScan_dColumnsubLabel.Enabled = Enable
    GridScan_dRowsubLabel.Enabled = Enable
    GridScan_dColumnsub.Enabled = Enable
    GridScan_dRowsub.Enabled = Enable
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
    Dim Pos As Integer
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
            Pos = Start
            Do While Pos > 0
                Pos = InStr(Start, MacroPath, bslash)
                If Pos > 0 Then
                    Start = Pos + 1
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
        DisplayProgress "Restore Settings", RGB(&HC0, &HC0, 0)
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
        Sleep (1000)
        DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
        ChangeButtonStatus True
    Else
        DisplayProgress "Restore Settings", RGB(&HC0, &HC0, 0)
        RestoreAcquisitionParameters
        DoEvents
        Sleep (200)
        ' reset the buttons
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
        Sleep (200)
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


'''
' This button runs several tests of multiple acquisitions in different modes. The test will be run in the working directory
'''
Private Sub RunTests_Click()
    Dim RecordingDoc As DsRecordingDoc
    Dim FilePath As String
    Dim MaxTestRepeats As Integer
    Dim TestNr As Integer
    Dim pixelDwell As Double
    MaxTestRepeats = 20
    
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Sub
    End If
    
    'Setup a single recording doc
    If RecordingDoc Is Nothing Then
        Set RecordingDoc = Lsm5.NewScanWindow
        While RecordingDoc.IsBusy
            Sleep (100)
            DoEvents
        Wend
    End If
    
    If Not CheckDir(GlobalDataBaseName) Then
        StopAcquisition
        Exit Sub
    End If
    
    Log = True
    
    '''''''
    ' No Z-Tracking, No Acquistion after Autofocus
    '''''''
    CheckBoxAutofocusTrackZ.Value = False
    CheckBoxTrack1.Value = False
    CheckBoxTrack2.Value = False
    CheckBoxTrack3.Value = False
    CheckBoxTrack4.Value = False

    TestNr = 1
    CheckBoxHighSpeed.Value = False
    CheckBoxHRZ.Value = False
    CheckBoxLowZoom.Value = False
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    TestNr = 2
    CheckBoxHighSpeed.Value = True
    CheckBoxHRZ.Value = False
    CheckBoxLowZoom.Value = True
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    TestNr = 3
    CheckBoxHighSpeed.Value = True
    CheckBoxHRZ.Value = True
    CheckBoxLowZoom.Value = True
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    ''''''
    ' No Z-Tracking, Acquistion after Autofocus
    ''''''
    CheckBoxAutofocusTrackZ.Value = False
    CheckBoxTrack1.Value = OptionButtonTrack1.Value
    CheckBoxTrack2.Value = OptionButtonTrack2.Value
    CheckBoxTrack3.Value = OptionButtonTrack3.Value
    CheckBoxTrack4.Value = OptionButtonTrack4.Value

    TestNr = 4
    CheckBoxHighSpeed.Value = False
    CheckBoxHRZ.Value = False
    CheckBoxLowZoom.Value = False
    ActivateAutofocusTrack GlobalAcquisitionRecording, Lsm5.Hardware.CpFocus.Position, pixelDwell
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    TestNr = 5
    CheckBoxHighSpeed.Value = True
    CheckBoxHRZ.Value = False
    CheckBoxLowZoom.Value = True
    ActivateAutofocusTrack GlobalAcquisitionRecording, Lsm5.Hardware.CpFocus.Position, pixelDwell
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    TestNr = 6
    CheckBoxHighSpeed.Value = True
    CheckBoxHRZ.Value = True
    CheckBoxLowZoom.Value = True
    ActivateAutofocusTrack GlobalAcquisitionRecording, Lsm5.Hardware.CpFocus.Position, pixelDwell
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    CheckBoxAutofocusTrackZ.Value = True
    CheckBoxTrack1.Value = OptionButtonTrack1.Value
    CheckBoxTrack2.Value = OptionButtonTrack2.Value
    CheckBoxTrack3.Value = OptionButtonTrack3.Value
    CheckBoxTrack4.Value = OptionButtonTrack4.Value

    TestNr = 7
    CheckBoxHighSpeed.Value = False
    CheckBoxHRZ.Value = False
    CheckBoxLowZoom.Value = False
    ActivateAutofocusTrack GlobalAcquisitionRecording, Lsm5.Hardware.CpFocus.Position, pixelDwell
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    TestNr = 8
    CheckBoxHighSpeed.Value = True
    CheckBoxHRZ.Value = False
    CheckBoxLowZoom.Value = True
    ActivateAutofocusTrack GlobalAcquisitionRecording, Lsm5.Hardware.CpFocus.Position, pixelDwell
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If

    TestNr = 9
    CheckBoxHighSpeed.Value = True
    CheckBoxHRZ.Value = True
    CheckBoxLowZoom.Value = True
    ActivateAutofocusTrack GlobalAcquisitionRecording, Lsm5.Hardware.CpFocus.Position, pixelDwell
    If Not RunTestAutofocusButton(RecordingDoc, TestNr, MaxTestRepeats) Then
        StopAcquisition
        Exit Sub
    End If


End Sub


''''
'   RunTestAutofocusButton(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer) As Boolean
'   Using the actual setting for autofocusing function runs AutofocusButton. Save images and logfile on the GlobalDataBaseName directory
'       [RecordingDoc] - A recording where images are overwritten
'       [TestNr]       - Number of the test, this sets the name of the image files and logfiles.
'       [MaxTestRepeats] - Maximal number of tests for each repeat
''''
Private Function RunTestAutofocusButton(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer) As Boolean

    Dim FilePath As String
    Dim TestRepeats As Integer
    TestRepeats = 1
    LogFileName = GlobalDataBaseName & "\AutofocusLogTest" & TestNr & ".txt"
    
    If Log Then
        Set LogFile = FileSystem.OpenTextFile(LogFileName, 8, True)
        LogFile.WriteLine "% Autofocus Test " & TestNr & ". Repeated AutofocusButton executions. "
        LogFile.WriteLine "% MaxSpeed " & CheckBoxHighSpeed.Value & ", Zoom1 " & CheckBoxLowZoom.Value & ", Piezo " & CheckBoxHRZ.Value & ", AFTrackZ " & CheckBoxAutofocusTrackZ.Value & _
        ", AFTrackXY " & CheckBoxAutofocusTrackXY.Value
        LogFile.Close
    End If
    
    While TestRepeats < MaxTestRepeats + 1
        DisplayProgress "Running Test " & TestNr & ". Repeat " & TestRepeats & "/" & MaxTestRepeats & ".......", RGB(0, &HC0, 0)
        FilePath = GlobalDataBaseName & "\Test" & TestNr & "_" & TestRepeats & ".lsm"
        If Log Then
            Set LogFile = FileSystem.OpenTextFile(LogFileName, 8, True)
            LogFile.WriteLine "% Save image in file " & FilePath
            LogFile.Close
        End If
        DoEvents
        Sleep (5000)
        DoEvents
        If Not AutofocusButtonRun(RecordingDoc) Then
            Exit Function
        End If
        'save file
        SaveDsRecordingDoc RecordingDoc, FilePath
        TestRepeats = TestRepeats + 1
        If ScanStop Then
            Exit Function
        End If
    Wend
    RunTestAutofocusButton = True
End Function

    

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



''''''
'   RestoreAcquisitionParameters()
'   Restores the image acquisition recording parameters from GlobalBackupRecording
'   recenter acquisition
'   Lsm5.DsRecording Out - Recording settings
''''''
Public Sub RestoreAcquisitionParameters()
    Dim i As Integer
    Dim Pos As Double

    Lsm5.DsRecording.Copy GlobalBackupRecording
    Lsm5.DsRecording.FrameSpacing = GlobalBackupRecording.FrameSpacing
    Lsm5.DsRecording.FramesPerStack = GlobalBackupRecording.FramesPerStack
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire = GlobalBackupActiveTracks(i)
    Next i
    'move to start pos
    
    i = 1
    While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + Lsm5.Hardware.CpFocus.Position - posTempZ, 1) And i < 10
        Sleep (400)
        i = i + 1
        DoEvents
    Wend
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

'''
'   ScanLineToggle_Click()
'   Switch setting for Line Autofocus on and FrameAutofocus off
'''
Private Sub ScanLineToggle_Click()
    ScanFrameToggle.Value = Not ScanLineToggle.Value 'if ScanFrame is true ScanLine is false (you can only chose one of them)
    FrameSizeLabel.Visible = ScanFrameToggle.Value   'FrameSize Label is only displayed if ScanFrame is activated
    BSliderFrameSize.Visible = ScanFrameToggle.Value 'FrameSize Slider is only displayed if ScanFrame is activated
    CheckBoxAutofocusTrackXY.Visible = ScanFrameToggle.Value
End Sub

'''
'   ScanFrameToggle_Click()
'   Switch setting for FrameAutofocus on and LineAutofocus off
'''
Private Sub ScanFrameToggle_Click()
    ScanLineToggle.Value = Not ScanFrameToggle.Value 'if ScanLine is chosen, ScanFrame will be unchecked
    FrameSizeLabel.Visible = ScanFrameToggle.Value
    BSliderFrameSize.Visible = ScanFrameToggle.Value
    CheckBoxAutofocusTrackXY.Visible = ScanFrameToggle.Value
End Sub


''''''
'   GetCurrentPositionOffsetButton_Click()
'       Performs Autofocus and update ZOffset according to ZShift
''''''
Private Sub GetCurrentPositionOffsetButton_Click()
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim NewPicture As DsRecordingDoc
    Set NewPicture = Lsm5.NewScanWindow
    While NewPicture.IsBusy
        Sleep (100)
        DoEvents
    Wend
    
    DisplayProgress "Get Current Position Offset - Autofocus", RGB(0, &HC0, 0)             'Gives information to the user
    StopScanCheck
    
    posTempZ = Lsm5.Hardware.CpFocus.Position

    If Not (newMacros.Autofocus_StackShift(NewPicture)) Then
        StopAcquisition
        Exit Sub
    End If
    
    ComputeShiftedCoordinates XMass, YMass, ZMass, X, Y, Z
    
    BSliderZOffset.Value = BSliderZOffset.Value - Round(ZMass, 1)
    
    RestoreAcquisitionParameters
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
End Sub

'''''''
'   AutofocusButton_Click()
'   calls AutofocusButtonRun
''''''''
Public Sub AutofocusButton_Click()
    Dim RecordingDoc As DsRecordingDoc
    If Not AutofocusButtonRun(RecordingDoc) Then
        StopAcquisition
    End If
End Sub

'''''''
'   AutofocusButtonRunn(Optional AutofocusDoc As DsRecordingDoc = Nothing) As Boolean
'   Runs a Z-stacks, compute center of mass, if selected acquire an image at computed position + ZOffset
'   If CheckBoxAutofocusTrackZ : position is updated to computed position from autofocus (without ZOffset!)
'   If CheckBoxAutofocusTrackXY and FrameToggle: position of X and Y are changed
'   Function uses a posTempZ to remember starting position
'       [AutofocusDoc] - A recording Doc. If = Nothing then it will create a new recording
'
'   Additional comments: The function works best with piezo. With Fast-Zline (Onthefly) acquisition is less precise
'                        Lots of test to check that focus returned to workingposition. Lsm5.Hardware.CpFocus.Position
'                        does not give actual position when stage is moving after acquisition.
'                        Lsm5.DsRecording.Sample0Z provides the actual shift to the central slice
''''''''
Private Function AutofocusButtonRun(Optional AutofocusDoc As DsRecordingDoc = Nothing) As Boolean
    Dim Time As Double
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim Success As Boolean
    Try = 1

    AutofocusForm.GetBlockValues 'Updates the parameters value for BlockZRange, BlockZStep..
    
    DisplayProgress "Autofocus Setup", RGB(0, &HC0, 0)
    
    StopScanCheck
    ' really move where it should
    Z = posTempZ
    FailSafeMoveStageZ (posTempZ)
    X = Lsm5.Hardware.CpStages.PositionX
    Y = Lsm5.Hardware.CpStages.PositionY
    ' check that 0 slice is correct
    Dim Cnt As Integer
    Cnt = 1
   
    ' Wait up to 4 sec for centering
    ' Note pculiarity of centering
    ' position central slice is Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position (or the real actual position)
    ' this waits for central slice at posTempZ
    While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
    Lsm5.Hardware.CpFocus.Position - posTempZ, 1) And Cnt < 10
        Sleep (400)
        If ScanStop Then
            Exit Function
        End If
        DoEvents
        Cnt = Cnt + 1
    Wend
    
    If CheckBoxActiveAutofocus Then
        ' Acquire image and calculate center of mass stored in XMass, YMass and ZMass
        If Not newMacros.Autofocus_StackShift(AutofocusDoc) Then
            ScanStop = True
            StopAcquisition
            Exit Function
        End If

        'recenter: This is essential to obtain correct slice when acquiring new stack
        DisplayProgress "AF: Recenter ...", RGB(0, &HC0, 0)

        Cnt = 1
        Time = Timer
        ' this waits for central slice at posTempZ
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
        Lsm5.Hardware.CpFocus.Position - posTempZ, 1) And Cnt < 10
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend

        'compute new coordinates
        ComputeShiftedCoordinates XMass, YMass, ZMass, X, Y, Z
        'round up position
        Z = Round(Z, 1)

        If Log Then
            Set LogFile = FileSystem.OpenTextFile(LogFileName, 8, True)
            LogFile.WriteLine ("% AutofocusButton: Current position " & posTempZ & "; Autofocus computed position " & Z & "; Center of mass " & ZMass)
            LogFile.WriteLine ("% AutofocusButton: Time recentering " & Timer - Time & ", waiting repeats (max 9) " & Cnt)
        End If

        'move X and Y if tracking is on
        If ScanFrameToggle And CheckBoxAutofocusTrackXY Then
            If Not FailSafeMoveStageXY(X, Y) Then
                Exit Function
            End If
        End If
    End If
    
    If CheckBoxHRZ Then
        Lsm5.Hardware.CpHrz.Position = 0
    End If

    ' Set the acquisitiontrack and record
    If ActivateAcquisitionTrack(GlobalAcquisitionRecording) Then
        Dim Offset As Double
        If CheckBoxActiveAutofocus Then
            Offset = BSliderZOffset
        Else
            Offset = 0
        End If

        DisplayProgress "AF: Taking image at ZOffset position...", RGB(0, &HC0, 0)
        'center the slide
        If Lsm5.DsRecording.ScanMode = "Stack" Or Lsm5.DsRecording.ScanMode = "ZScan" Then
            'central slide is at Z + Offset
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 _
            + Lsm5.Hardware.CpFocus.Position - Z - Offset
        Else
            If Not FailSafeMoveStageZ(Z + Offset) Then
                Exit Function
            End If
        End If

        If Log Then
            Sleep (2000)
            LogFile.WriteLine ("% Target Centralslide " & Z + BSliderZOffset & "; obtained Centralslide " & _
            Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position)
        End If

        If Not ScanToImageNew(AutofocusDoc) Then
            Exit Function
        End If
        'wait that slice recentered after acquisition
        Cnt = 1
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * _
        (Lsm5.DsRecording.FramesPerStack - 1) / 2 + Lsm5.Hardware.CpFocus.Position - Z - Offset, 1) And Cnt < 10
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
    End If

    If Log Then
        LogFile.Close
    End If



    '''Update position. Central slide is without offset!!
    If CheckBoxAutofocusTrackZ Then
        posTempZ = Z
        Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + Lsm5.Hardware.CpFocus.Position - Z
        If Not FailSafeMoveStageZ(Z) Then
            Exit Function
        End If
        Cnt = 1
        'check that Z-Stack is centered again
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2, 1) And Cnt < 10
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
    Else
        'just recenter and move to original position
        Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + Lsm5.Hardware.CpFocus.Position - posTempZ
        If Not FailSafeMoveStageZ(posTempZ) Then
            Exit Function
        End If
    End If
    StopAcquisition
    AutofocusButtonRun = True
    
End Function


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

    If Not StartSetting() Then
        ScanStop = True
        StopAcquisition
        Exit Sub
    End If
    
    'Set counters back to 1
    RepetitionNumber = 1 ' first time point
    
    StartAcquisition BleachingActivated 'This is the main function of the macro
End Sub


Private Sub ContinueFromCurrentLocation_Click()
    If Not StartSetting Then
        ScanStop = True
        StopAcquisition
        Exit Sub
    End If
    StartAcquisition BleachingActivated 'This is the main function of the macro
End Sub


''''''
'   StartSetting()
'   Setups and controls before start of experiment
'       Create list of positions for Grid or Multiposition
''''''
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
    
    'As default we do not overwrite files
    OverwriteFiles = False
    

       
    '''''''''''''''''''''''
    '***Set up GridScan***'
    '''''''''''''''''''''''
    If CheckBoxActiveGridScan Then
        'Load starting position from stage
        If Lsm5.Hardware.CpStages.Markcount = 0 Then  ' No marked position
            MsgBox " GridScan: Use stage to Mark the initial position "
            ScanStop = True
            StopAcquisition
            Exit Function
        End If
        MsgBox " GridScan: Uses as initial position the first Marked point on stage "
        ' Store starting position for later restart. This is the first marked point
        Lsm5.Hardware.CpStages.MarkGetZ 0, XStart, YStart, ZStart
 
        If GridScan_nColumn.Value * GridScan_nRow.Value * GridScan_nColumnsub.Value * GridScan_nRowsub.Value > 10000 Then
            MsgBox "GridScan: Maximal number of locations is 10000. Please change Numbers  X and/or Y."
            ScanStop = True
            StopAcquisition
            Exit Function
        End If
        
        ReDim posGridX(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value)
        ReDim posGridY(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value)
        ReDim posGridZ(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value)
        ReDim posGridXY_valid(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value) ' A well may be active or not
        ReDim posGridXsub(1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value)
        ReDim posGridYsub(1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value)
        ReDim posGridZsub(1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value)
        ReDim posGridXYsub_valid(1 To GridScan_nColumnsub.Value, 1 To GridScan_nRowsub.Value) ' A subposition may be active or not
        DisplayProgress "Initialize main grid positions....", RGB(0, &HC0, 0)
        Sleep (1000)
        MakeGrid posGridX, posGridY, posGridZ, posGridXY_valid, XStart, YStart, ZStart, GridScan_dColumn.Value, GridScan_dRow.Value, True
        DisplayProgress "Initialize all grid positions...DONE", RGB(0, &HC0, 0)
    End If
    '''''''''''''''''''''''''''
    '***End Set up GridScan***'
    '''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''
    '***Set up MultiLocationScan***'
    ''''''''''''''''''''''''''''''''
    If MultipleLocationToggle Then
        Dim i As Integer
        If Lsm5.Hardware.CpStages.Markcount > 0 Then
            ReDim posGridX(1 To 1, 1 To Lsm5.Hardware.CpStages.Markcount)
            ReDim posGridY(1 To 1, 1 To Lsm5.Hardware.CpStages.Markcount)
            ReDim posGridZ(1 To 1, 1 To Lsm5.Hardware.CpStages.Markcount)
            ReDim posGridXY_valid(1 To 1, 1 To Lsm5.Hardware.CpStages.Markcount) ' A well may be active or not
            ReDim posGridXsub(1 To 1, 1 To 1)
            ReDim posGridYsub(1 To 1, 1 To 1)
            ReDim posGridZsub(1 To 1, 1 To 1)
            ReDim posGridXYsub_valid(1, 1)
            For i = 1 To Lsm5.Hardware.CpStages.Markcount
                Lsm5.Hardware.CpStages.MarkGetZ i - 1, posGridX(1, i), posGridY(1, i), _
                posGridZ(1, i)
                posGridXY_valid(1, i) = True
            Next i
        End If
    End If
    
  
    If SingleLocationToggle And Not CheckBoxActiveGridScan Then
            ReDim posGridX(1 To 1, 1 To 1)
            ReDim posGridY(1 To 1, 1 To 1)
            ReDim posGridZ(1 To 1, 1 To 1)
            ReDim posGridXY_valid(1 To 1, 1 To 1) 'A well may be active or not
            ReDim posGridXsub(1 To 1, 1 To 1)
            ReDim posGridYsub(1 To 1, 1 To 1)
            ReDim posGridZsub(1 To 1, 1 To 1)
            ReDim posGridXYsub_valid(1 To 1, 1 To 1)
            Lsm5.Hardware.CpStages.GetXYPosition posGridX(1, 1), posGridY(1, 1)
            posGridZ(1, 1) = Lsm5.Hardware.CpFocus.Position
            posGridXY_valid(1, 1) = 1
    End If
    
    StartSetting = True
End Function



''''''
'   StartAcquisition(BleachingActivated)
'   Perform many things (TODO: write more). Pretty much the whole macro runs through here
''''''
Private Sub StartAcquisition(BleachingActivated As Boolean)
    
    'measure time required
    Dim rettime, difftime As Double
    Dim GlobalPrvTime As Double
    Dim StartTime As Double
    
    'Counters
    Dim iRow As Long
    Dim iCol As Long
    Dim iRowSub As Long
    Dim iColSub As Long
    Dim RowMax As Long
    Dim ColMax  As Long
    Dim RowSubMax As Long
    Dim ColSubMax As Long
    Dim StartCol As Long
    Dim StartColSub As Long
    Dim EndCol As Long
    Dim EndColSub As Long
    Dim StepCol As Integer
    Dim StepColSub As Integer
    Dim Cnt As Integer ' a local counter
    
    HighResExperimentCounter = 1
    HighResCounter = 1
    ' CheckBoxActiveOnlineImageAnalysis  refers to the MicroPilot
    If CheckBoxActiveOnlineImageAnalysis Then
        ReDim Preserve HighResArrayX(100) 'define 100 a priori (even if there are less)
        ReDim Preserve HighResArrayY(100)
        ReDim Preserve HighResArrayZ(100)
        SaveSetting "OnlineImageAnalysis", "macro", "code", 0
        SaveSetting "OnlineImageAnalysis", "macro", "offsetx", 0
        SaveSetting "OnlineImageAnalysis", "macro", "offsety", 0
    End If
  
    'Coordinates
    Dim X As Double              ' x value where to move the stage (this is used as reference)
    Dim Y As Double              ' y value where to move the stage
    Dim Z As Double              ' z value where to move the stage

    
    'test variables
    Dim Success As Integer       ' Check if something was sucessfull
    Dim SuccessAF As Boolean     ' Check if AF was succesful
    Dim LocationSoFarBest As Integer
    Dim soFarBestGoodCellsPerImage As Integer
    
    'Recording stuff
    Dim FileNameID As String ' ID name of file (Well/Position, Subpositio, Timepoint)
    Dim FilePath As String   ' full path of file to save (changes through function)
    Dim RecordingDoc As DsRecordingDoc  ' contains the images
    Dim Scancontroller As AimScanController ' the controller
  
        'do once leveling
    If CheckBoxHRZ Then
        Lsm5.Hardware.CpHrz.Leveling ' not sure if this is needed
        Sleep (1000)
    End If
    
    
    ' Set the offset in z-stack to 0; otherwise there can be errors...
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2
                           
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

    InitializeStageProperties
    SetStageSpeed 9, True

    Running = True  'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.
    ChangeButtonStatus False ' disable buttons
    
    Do While Running   'As long as the macro is running we're in this loop. At everystop one will save actual location, and repetition
                
        RowMax = UBound(posGridX, 1)
        ColMax = UBound(posGridX, 2)
        
        RowSubMax = UBound(posGridXsub, 1)
        ColSubMax = UBound(posGridXsub, 2)
        
        GlobalPrvTime = CDbl(GetTickCount) * 0.001
        iRow = 1
        For iRow = 1 To RowMax
            'Meander
            If iRow Mod 2 = 0 Then
                StartCol = ColMax
                EndCol = 1
                StepCol = -1
            Else
                StartCol = 1
                EndCol = ColMax
                StepCol = 1
            End If
            iCol = StartCol
            For iCol = StartCol To EndCol Step StepCol
                ' Create Sub positions and loop through them
                If CheckBoxActiveGridScan Then
                    'Create a sub mask (with no submask Grid and Gridsub are identical)
                    MakeGrid posGridXsub, posGridYsub, posGridZsub, posGridXYsub_valid, posGridX(iRow, iCol), _
                    posGridY(iRow, iCol), posGridZ(iRow, iCol), GridScan_dColumnsub, GridScan_dRowsub, posGridXY_valid(iRow, iCol)
                Else
                    ' for one position or multiposition the Gridsub = Grid
                    posGridXsub(1, 1) = posGridX(iRow, iCol)
                    posGridYsub(1, 1) = posGridY(iRow, iCol)
                    posGridZsub(1, 1) = posGridZ(iRow, iCol)
                    posGridXYsub_valid(1, 1) = posGridXY_valid(iRow, iCol)
                End If
                iRowSub = 1
                ' Move in the subGrid
                For iRowSub = 1 To RowSubMax
                    'Meander in the subgrid
                    If iRowSub Mod 2 = 0 Then
                        StartColSub = ColSubMax
                        EndColSub = 1
                        StepColSub = -1
                    Else
                        StartColSub = 1
                        EndColSub = ColSubMax
                        StepColSub = 1
                    End If
                    iColSub = StartColSub
                    For iColSub = StartColSub To EndColSub Step StepColSub
                        ' Here comes the check for good or bad location ...
                        If posGridXYsub_valid(iRowSub, iColSub) Then
                            'define actual positions and move there
                            X = posGridXsub(iRowSub, iColSub)
                            Y = posGridYsub(iRowSub, iColSub)
                            Z = posGridZsub(iRowSub, iColSub)
                            'move Z if required
                            If Round(Z, 1) <> Round(Lsm5.Hardware.CpFocus.Position, 1) Then
                                If Not FailSafeMoveStageZ(Z) Then
                                    StopAcquisition
                                    Exit Sub
                                End If
                            End If
                            If Not FailSafeMoveStageXY(X, Y) Then
                                StopAcquisition
                                Exit Sub
                            End If
                            'recenter the imaging to be sure and wait
                            Cnt = 1
                            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2
                            Sleep (400)
                            While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2, 1) And Cnt < 10
                                Sleep (400)
                                DoEvents
                                Cnt = Cnt + 1
                            Wend
                        Else ' jump to next location
                            GoTo NextLocation
                        End If
                        
                        ' Show position of stage
                        If SingleLocationToggle Then
                            LocationTextLabel.Caption = " X= " & posGridXsub(iRowSub, iColSub) & ",  Y = " & posGridYsub(iRowSub, iColSub) & ", Z = " & posGridZsub(iRowSub, iColSub)
                        End If
                        
                        If MultipleLocationToggle Then
                            LocationTextLabel.Caption = "Marked Position: " & iCol & vbCrLf & _
                                                        " X= " & posGridXsub(iRowSub, iColSub) & ",  Y = " & posGridYsub(iRowSub, iColSub) & ", Z = " & posGridZsub(iRowSub, iColSub)
                        End If
                        If CheckBoxActiveGridScan Then
                            LocationTextLabel.Caption = "Well/Position Row: " & iRow & ", Column: " & iCol & vbCrLf & _
                                                        "subposition   Row: " & iRowSub & ", Column: " & iColSub & vbCrLf & _
                                                        " X= " & posGridXsub(iRowSub, iColSub) & ",  Y = " & posGridYsub(iRowSub, iColSub) & ", Z = " & posGridZsub(iRowSub, iColSub)
                        End If
                        
                        If ScanPause Then
                            If Not Pause Then ' Pause is true is Resume
                                ScanStop = True
                                StopAcquisition
                                Exit Sub
                            End If
                        End If
                        If RepetitionNumber = 1 Then
                            StartTime = GetTickCount    'Get the time when the acquisition was started
                        End If
                        
                        'Do the imaging
                        If Not ImagingWorkFlow(RecordingDoc, StartTime, iRow, iCol, iColSub, iRowSub) Then
                            ' Return Z to its original position
                            FailSafeMoveStageZ posGridZsub(iRowSub, iColSub)
                            StopAcquisition
                            Exit Sub
                        End If
        
NextLocation:
                        If ScanPause Then
                            If Not Pause Then ' Pause is true is Resume
                                ScanStop = True
                                StopAcquisition
                                Exit Sub
                            End If
                        End If
                    
                Next iColSub
            Next iRowSub
        Next iCol
    Next iRow

       
    ' DONE WITH THE IMAGING....NOW POSTPROCESSING...
    
    If AutomaticBleaching Then
        FillBleachTable     ' Updating the bleaching table before the next acquisitions, just in case there were changes n the bleaching window
    End If
    
        
    If (RepetitionNumber < BSliderRepetitions.Value) Then
        
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
                    If Not Pause Then ' Pause is true is Resume
                        ScanStop = True
                        StopAcquisition
                        Exit Sub
                    End If
                End If
                If ExtraBleach Then                                 'Modifies the bleaching table to do an Extrableach for al locatins at the next repetition
                    ExtraBleach = False
                    BleachTable(RepetitionNumber + 1) = True
                End If
                If ScanStop Then
                    StopAcquisition
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
    
    DoEvents
    
    Loop ' RepetitonLoop ; Do While Running
    
    
    StopAcquisition
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)

End Sub


'''''
'   Contains all the jobs performed at one position
'   It will check for Autofocus, Additional image acquisition, normal acquisitions, Micropilot acquisition
'       [Row] In - Actual Row of Well/Position
'       [Col] In - Actual Column of Well/Position
'       [RowSub] In - Row of subpositions grid
'       [ColSub] In - Column of subpositions grid
'''''
Private Function ImagingWorkFlow(RecordingDoc As DsRecordingDoc, StartTime As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    
    ImagingWorkFlow = False
    Dim Xnew As Double
    Dim Ynew As Double
    Dim Znew As Double
    Dim Xold As Double
    Dim Yold As Double
    Dim Zold As Double
    Dim FileNameID As String
    Dim FilePath As String
    Dim Cnt As Integer
    Dim Time As Double
    Xnew = Lsm5.Hardware.CpStages.PositionX
    Ynew = Lsm5.Hardware.CpStages.PositionY
    Znew = Lsm5.Hardware.CpFocus.Position
    Xold = Xnew
    Yold = Ynew
    Zold = Znew
    ' Set FileNameId. W....P....T....
    FileNameID = FileName((Row - 1) * UBound(posGridX, 1) + Col, (RowSub - 1) * UBound(posGridXsub, 1) + ColSub, RepetitionNumber)
    FilePath = DatabaseTextbox.Value & "\" & TextBoxFileName.Value & FileNameID & ".lsm"
    
    If Log Then
        Set LogFile = FileSystem.OpenTextFile(LogFileName, 8, True)
        LogFile.WriteLine ("% StartButton: Acquire image " & FilePath)
    End If
    
    ' At every positon and repetition  check if Autofocus needs to be required. Update of positions in Z is only done at the end of acquisition
    If (RepetitionNumber - 1) Mod AFeveryNth = 0 Then
        If CheckBoxActiveAutofocus Then  ' Perform Autofocus
            DisplayProgress "Autofocus", RGB(0, &HC0, 0)
            StopScanCheck 'stop any running jobs
            ' take a z-stack and finds the brightest plane:
            If Not Autofocus_StackShift(RecordingDoc) Then
               Exit Function
            End If
            ' move the xyz to the right position
            DisplayProgress "Autofocus move stage", RGB(0, &HC0, 0)
            ComputeShiftedCoordinates XMass, YMass, ZMass, Xnew, Ynew, Znew
            If CheckBoxAutofocusTrackXY.Value Then
                If Not FailSafeMoveStageXY(Xnew, Ynew) Then
                    Exit Function
                End If
                posGridX(Row, Col) = Xnew
                posGridY(Row, Col) = Ynew
            End If
            If Log Then
                LogFile.WriteLine ("% StartButton: Autofocus. Current position XYZ" & posGridXsub(RowSub, ColSub) & ", " & posGridYsub(RowSub, ColSub) & ", " & posGridZsub(RowSub, ColSub) & _
                ". AF computed position XYZ " & Xnew & ", " & Ynew & ", " & Znew & ". Center of mass XYZ" & XMass & ", " & YMass & ", " & ZMass)
                LogFile.WriteLine ("% StartButton: Autofocus. Current position XYZ" & posGridXsub(RowSub, ColSub) & ", " & posGridYsub(RowSub, ColSub) & ", " & posGridZsub(RowSub, ColSub) & _
                ". AF computed position XYZ " & Xnew & ", " & Ynew & ", " & Znew & ". Center of mass XYZ" & XMass & ", " & YMass & ", " & ZMass)
            End If
        End If
    End If '(RepetitionNumber - 1) Mod AFeveryNth = 0

    FileNameID = FileName((Row - 1) * UBound(posGridX, 1) + Col, (RowSub - 1) * UBound(posGridXsub, 1) + ColSub, RepetitionNumber)
          
    'recenter
    DisplayProgress "Autofocus: Wait for recenter", RGB(0, &HC0, 0)
    Cnt = 1
    Time = Timer
    While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
        Lsm5.Hardware.CpFocus.Position - posGridZsub(RowSub, ColSub), 1) And Cnt < 10
        Sleep (400)
        DoEvents
        Cnt = Cnt + 1
        If ScanStop Then
            Exit Function
        End If
    Wend
    
    If Log Then
        LogFile.WriteLine ("% Recenter time: " & Timer - Time)
    End If
       
    If ScanPause Then
        If Not Pause Then ' Pause is true if Resume
            Exit Function
        End If
    End If
    
    ''''''''''''''''''''''''''''''
    '*Begin Alternative imaging**'
    ''''''''''''''''''''''''''''''
    If CheckBoxAlterImage.Value Then
        If RepetitionNumber Mod TextBox_RoundAlterTrack = 0 Then
            ' use subgrid
            If GridScan_nColumnsub.Value * GridScan_nRowsub.Value > 1 And ((RowSub - 1) * UBound(posGridXsub, 1) + ColSub) Mod TextBox_RoundAlterLocation = 0 Then
                FilePath = DatabaseTextbox.Value & "\" & TextBoxFileName.Value & "--Alt" & FileNameID & ".lsm" ' fullpath of alternative file
                If Not StartAlternativeImaging(RecordingDoc, FilePath, _
                    TextBoxFileName.Value & "--Alt" & FileNameID & ".lsm") Then
                    Exit Function
                End If
            ElseIf ((Row - 1) * UBound(posGridX, 1) + Col) Mod TextBox_RoundAlterLocation = 0 Then
                FilePath = DatabaseTextbox.Value & "\" & TextBoxFileName.Value & "--Alt" & FileNameID & ".lsm" ' fullpath of alternative file
                If Not StartAlternativeImaging(RecordingDoc, FilePath, _
                    TextBoxFileName.Value & "--Alt" & FileNameID & ".lsm") Then
                    Exit Function
                End If
            End If
        End If
    End If

    If ScanPause Then
        If Not Pause Then ' Pause is true if Resume
            Exit Function
        End If
    End If

    '''''''''''''''''''''''''''''''''''''
    '*Begin Normal acquisition imaging**'
    '''''''''''''''''''''''''''''''''''''
    DisplayProgress "Acquiring at current location" & vbCrLf & _
                    "Repetition: " & RepetitionNumber, RGB(&HC0, &HC0, 0)
    If Not ActivateAcquisitionTrack(GlobalAcquisitionRecording) Then           'An additional control....
        MsgBox "No track selected for Acquisition! Cannot Acquire!"
        ScanStop = True
        StopAcquisition
        Exit Function
    End If
    'center the slide
    If Lsm5.DsRecording.ScanMode = "Stack" Or Lsm5.DsRecording.ScanMode = "ZScan" Then
        'central slide is at Z + BSliderZOffset
        Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 _
        + Lsm5.Hardware.CpFocus.Position - Znew - BSliderZOffset
    Else
        If Not FailSafeMoveStageZ(Znew + BSliderZOffset) Then
            Exit Function
        End If
    End If
    
    If Log Then
        Sleep (2000)
        LogFile.WriteLine ("% Target central slide " & Znew + BSliderZOffset & "; obtained Central slide " & _
        Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position)
    End If
    
    If Not ScanToImageNew(RecordingDoc) Then
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''
    '*** Save Acquisition Image *******'
    ''''''''''''''''''''''''''''''''''''
    RecordingDoc.SetTitle TextBoxFileName.Value & FileNameID
    'this is the name of the file to be saved
    FilePath = DatabaseTextbox.Value & "\" & TextBoxFileName.Value & FileNameID & ".lsm"
    'Check existance of file and warn
    If Not OverwriteFiles Then
        If FileExist(FilePath) Then
            If MsgBox("File " & FilePath & " exists. Do you want to overwrite this and subsequent files? ", VbYesNo) = vbYes Then
                OverwriteFiles = True
            Else
                ScanStop = True
                Exit Function
            End If
        End If
    End If

    If Not SaveDsRecordingDoc(RecordingDoc, FilePath) Then    ' HERE THE IMAGE IS FINALLY SAVED
        Exit Function
    End If

    'wait that slice recentered after acquisition
    Cnt = 1
    While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * _
    (Lsm5.DsRecording.FramesPerStack - 1) / 2 + Lsm5.Hardware.CpFocus.Position - Znew - BSliderZOffset, 1) And Cnt < 10
        Sleep (400)
        DoEvents
        Cnt = Cnt + 1
        If ScanStop Then
            Exit Function
        End If
    Wend

    If ScanPause Then
        If Not Pause Then ' Pause is true is Resume
            Exit Function
        End If
    End If

    ''''''''''''''''''''''''''
    '*** Store bleachTable ***'
    ''''''''''''''''''''''''''
    If BleachStartTable(RepetitionNumber) > 0 Then          'If a bleach was performed we add the information to the image metadata
        Lsm5.DsRecordingActiveDocObject.AddEvent (BleachStartTable(RepetitionNumber) - StartTime) / 1000, eEventTypeBleachStart, "Bleach Start"
        Lsm5.DsRecordingActiveDocObject.AddEvent (BleachStopTable(RepetitionNumber) - StartTime) / 1000, eEventTypeBleachStop, "Bleach End"
    End If
    
    If ScanPause Then
        If Not Pause Then ' Pause is true if Resume
            ScanStop = True
            Exit Function
        End If
    End If

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
            If Row = UBound(posGridX, 1) And Col = UBound(posGridX, 2) Then
                If RowSub = UBound(posGridXsub, 1) And Col = UBound(posGridXsub, 2) Then  'Allows again to do an extrableach at the en
                    ExtraBleachButton.Caption = "Bleach"
                    ExtraBleachButton.BackColor = &H8000000F
                End If
            End If
        
        End If
        ' todo:
        ' but where is the bleaching image stored ??
    End If
                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**** Updatepositions (x,y)z: Tracking ***********'''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TrackingToggle And Not CheckBoxActiveGridScan Then 'This is if we're doing some postacquisition tracking (not possible with Grid) this is done before Micropilot analysis
        Xnew = Xold
        Ynew = Yold
        Znew = Zold
        DisplayProgress "Tracking and computing new coordinates of " & vbCrLf & _
                "Well/Position Row: " & Row & ", Column: " & Col & vbCrLf & _
                "subposition   Row: " & RowSub & ", Column: " & ColSub & vbCrLf, RGB(&HC0, &HC0, 0)
        DoEvents
        MassCenter ("Tracking")
        'compute XYZShift from XYZMass
        ComputeShiftedCoordinates XMass, YMass, ZMass, Xnew, Ynew, Znew
    End If



    ' COMMUNICATION WITH MICROPILOT: START *****************
      
    If CheckBoxActiveOnlineImageAnalysis Then
        SaveSetting "OnlineImageAnalysis", "macro", "filepath", FilePath
        'Wait for anything to stop
        Do While RecordingDoc.IsBusy
            Sleep (100)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    StopAcquisition
                    Exit Function
                End If
            End If
        Loop
        
        SaveSetting "OnlineImageAnalysis", "macro", "Refresh", 0
        SaveSetting "OnlineImageAnalysis", "macro", "code", 1

        If Not MicroscopePilot(RecordingDoc, BleachingActivated, HighResExperimentCounter, HighResCounter, HighResArrayX, HighResArrayY, HighResArrayZ, _
        Row, Col, RowSub, ColSub) Then
            Exit Function
        End If
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**** Updatepositions (x,y)z: Tracking ***********'''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CheckBoxAutofocusTrackZ Or CheckBoxTrackZ Then
        If Not FailSafeMoveStageZ(Znew) Then
            Exit Function
        End If
        posGridZ(Row, Col) = Znew
    End If
    
    If TrackingToggle And Not CheckBoxActiveGridScan Then 'This is if we're doing some postacquisition tracking (not possible with Grid)
        
        DisplayProgress "Tracking and computing new coordinates of " & vbCrLf & _
                "Well/Position Row: " & Row & ", Column: " & Col & vbCrLf & _
                "subposition   Row: " & RowSub & ", Column: " & ColSub & vbCrLf, RGB(&HC0, &HC0, 0)

        DoEvents
        MassCenter ("Tracking")
    
        'compute XYZShift from XYZMass
        ComputeShiftedCoordinates XMass, YMass, ZMass, Xnew, Ynew, Znew
        
        'move to new position
        If CheckBoxTrackZ.Value And CheckBoxPostTrackXY.Value Then
            If Not FailSafeMoveStageXY(Xnew, Ynew) Or Not FailSafeMoveStageZ(Znew) Then
                Exit Function
            End If
        ElseIf CheckBoxPostTrackXY.Value Then
            If Not FailSafeMoveStageXY(Xnew, Ynew) Then
                Exit Function
            End If
        ElseIf CheckBoxTrackZ.Value Then
            If Not FailSafeMoveStageZ(Znew) Then
                Exit Function
            End If
        End If
        
        If Not Success Then
            Exit Function
        End If
        
        'update positionList x, y, z (Only if in multilocation mode)
        If MultipleLocationToggle.Value Then
            posGridX(Row, Col) = Lsm5.Hardware.CpStages.PositionX
            posGridY(Row, Col) = Lsm5.Hardware.CpStages.PositionY
            posGridZ(Row, Col) = Lsm5.Hardware.CpFocus.Position
        End If
        
    Else ' no location tracking
        Lsm5.Hardware.CpHrz.Leveling   'This I think puts the HRZ to its resting position, and moves the focuswheel correspondingly. Do we need this?
    End If
    ''  End: Setting new (x,y)z positions *******************************

    ImagingWorkFlow = True
    'one could monitor weather this position was any good at all here. Goodpositions
    ' COMMUNICATION WITH MICROPILOT: END *****************
End Function
    
                            
 


'''''
'   MakeGrid( posGridX() As Double, posGridY() As Double, posGridXY_valid() )
'   Create a grid
'       [posGridX] In/Out - Array where X grid positions are stored
'       [posGridY] In/Out - Array where Y grid positions are stored
'       [posGridXY_valid] In/Out - Array that says if position is valid
'       [locationNumbersMainGrid] In/Out - location number on main grid
'''''
Private Sub MakeGrid(posGridX() As Double, posGridY() As Double, posGridZ() As Double, posGridXY_valid() As Boolean _
, XStart As Double, YStart As Double, ZStart As Double, dX As Double, dY As Double, Valid As Boolean)
        ' A row correspond to Y movement and Column to X shift
        ' entries are posGridX(row, column)!! this what is
        'counters
        Dim iRow As Long
        Dim iCol As Long
        iRow = UBound(posGridX, 1)
        For iRow = 1 To UBound(posGridX, 1)
            For iCol = 1 To UBound(posGridX, 2)
                posGridX(iRow, iCol) = XStart + (iCol - 1) * dX
                posGridY(iRow, iCol) = YStart + (iRow - 1) * dY
                posGridZ(iRow, iCol) = ZStart
                posGridXY_valid(iRow, iCol) = Valid
            Next iCol
        Next iRow
End Sub


''''''
'   MassCenter(Context As String)
'   TODO: No test of Goodness of Mass estimation. Very slow function
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
    If ScanFrameToggle And SystemName = "LIVE" Then ' binning only with LIVE device
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
    ElseIf Context = "Tracking" Then    'Takes the channel selected in the pop-up menu when doing postacquisition tracking
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
    If ScanFrameToggle And SystemName = "LIVE" Then ' binning only with LIVE device
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
            PauseButton.Caption = "RESUME"
            PauseButton.BackColor = 12648447
        Else
            ScanPause = False
            PauseButton.Caption = "PAUSE"
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
Public Function Pause() As Boolean
    
    Dim rettime As Double
    Dim GlobalPrvTime As Double
    Dim difftime As Double
    
    GetCurrentPositionOffsetButton.Enabled = True
    AutofocusButton.Enabled = True
    GlobalPrvTime = CDbl(GetTickCount) * 0.001
    rettime = GlobalPrvTime
    difftime = rettime - GlobalPrvTime
    'TODO: test this function
    DoEvents
    Do While True
        Sleep (100)
        DoEvents
        If ScanStop Then
            StopAcquisition
            Pause = False
            Exit Function
        End If
        If ScanPause = False Then
            GetCurrentPositionOffsetButton.Enabled = False
            AutofocusButton.Enabled = False
            Pause = True
            Exit Function
        End If

        DisplayProgress "Pause " & CStr(CInt(difftime)) & " s", RGB(&HC0, &HC0, 0)
        rettime = CDbl(GetTickCount) * 0.001
        difftime = rettime - GlobalPrvTime
    Loop
End Function


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





Private Sub CloseButton_Click()
    RestoreAcquisitionParameters
    Sleep (1000)
    AutoStore
    
'    Excel.Application.DisplayAlerts = False
'    Excel.Application.Quit
    End
End Sub

Private Sub ReInitializeButton_Click()
    Re_Initialize
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

'fills popup menu for chosing a track for post-acquisition tracking
' TODO: move in form
Private Sub FillTrackingChannelList()
    Dim t As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim channel As DsDetectionChannel
    Dim Track As DsTrack
    
    ReDim ActiveChannels(Lsm5.Constants.MaxActiveChannels)  'ActiveChannels is a dynamic array (variable size), ReDim defines array size required next
                                                            'Array size is (MaxActiveChannels gets) the total max number of active channels in all tracks
    ComboBoxTrackingChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    
    If ActivateAcquisitionTrack(GlobalAcquisitionRecording) Then
        For t = 1 To TrackNumber 'This loop goes through all tracks and will collect all activated channels to display them in popup menu
            Set Track = GlobalAcquisitionRecording.TrackObjectByMultiplexOrder(t - 1, Success)
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
'   ActivateAutofocusTrack(Recording As DsRecording, posZ As Double, pixelDwell As Double)
'   Check which Track should be used for Autofocus and update passed DsRercoding.
'   posZ is position with respect to which calculate central slice (often the CpFocus still moves after acquisition)
'   This sets also the Z acquisition parameters for Acquisition document. For this one uses parameters of the AutofocusForm
'       [Recording] In/Out - a DsRecording
'   TODO: test
''''''
Public Function ActivateAutofocusTrack(Recording As DsRecording, posZ As Double, pixelDwell As Double) As Boolean
    Dim i As Integer
    Dim TrackSuccess As Integer
    Dim FunSuccess As Boolean
    
    FunSuccess = False
    ' Set all tracks to non-acquisition first
    For i = 1 To TrackNumber
       Recording.TrackObjectByMultiplexOrder(i - 1, TrackSuccess).Acquire = 0
    Next i
    
    For i = 1 To TrackNumber
        If OptionButtonTrack1.Value = True And i = 1 Then
            FunSuccess = True
            Exit For
        ElseIf OptionButtonTrack2.Value = True And i = 2 Then
            FunSuccess = True
            Exit For
        ElseIf OptionButtonTrack3.Value = True And i = 3 Then
            FunSuccess = True
            Exit For
        ElseIf OptionButtonTrack4.Value = True And i = 4 Then
            FunSuccess = True
            Exit For
        End If
    Next i
    
    If FunSuccess Then
        AutofocusTrack = i - 1
        Recording.TrackObjectByMultiplexOrder(AutofocusTrack, Success).Acquire = True
        If CheckBoxHighSpeed.Value Then
           Recording.TrackObjectByMultiplexOrder(AutofocusTrack, Success).SamplingNumber = 1  'TODO what happens here
        End If
    Else
        Exit Function
    End If
    
    If Not (SystemName = "LSM" Or SystemName = "LIVE") Then
        MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
        ActivateAutofocusTrack = False
        Exit Function
    End If
    
    If CheckBoxLowZoom.Value Then                   ' Changes the zoom if necessary
        Recording.ZoomX = 1
        Recording.ZoomY = 1
    Else                                            ' Use AcquisitionRecording as default
        Recording.ZoomX = GlobalBackupRecording.ZoomX
        Recording.ZoomY = GlobalBackupRecording.ZoomY
        Recording.ZoomZ = GlobalBackupRecording.ZoomZ
    End If
        
    Recording.TimeSeries = False                     'Disable the timeseries, because autofocussing is just one image at one timepoint.
    
    ''''''''''''''''''''''''''''
    '*Setting for LSM system***'
    ''''''''''''''''''''''''''''
    If SystemName = "LSM" Then
        
        '''How to do the Z-stacks
        If CheckBoxHRZ.Value Then                'Piezo
            Recording.SpecialScanMode = "ZScanner"
        Else
            Recording.SpecialScanMode = "FocusStep"
        End If
                    
        '''''''''''''''''''''''''''
        '**Setting for line scan**'
        '''''''''''''''''''''''''''
        If ScanLineToggle.Value Then
        
            Recording.ScanMode = "ZScan"             'This acquires  single X-Z image, like with "Range Select" button Z-stack Window.
            Recording.LinesPerFrame = 1
     
            If CheckBoxHighSpeed.Value Then           'For Highspeed we change the settings otherwise use the standard settings as for Acquisition
                Recording.SamplesPerLine = 256
                pixelDwell = 0.00000256
                If Not CheckBoxHRZ.Value Then         'OnTheFly speeds need to be adapted
                    Recording.SpecialScanMode = "OnTheFly" 'aka: Fast Z-line in Z-Stack menu
                    If BSliderZStep.Value < 0.15 Then
                        pixelDwell = 0.00000256
                    ElseIf BSliderZStep.Value < 0.19 Then
                        pixelDwell = 0.0000032
                    ElseIf BSliderZStep.Value < 0.31 Then
                        pixelDwell = 0.00000512
                    ElseIf BSliderZStep.Value < 0.38 Then
                        pixelDwell = 0.0000064
                    ElseIf BSliderZStep.Value < 0.77 Then
                        pixelDwell = 0.0000128
                    ElseIf BSliderZStep.Value < 1.54 Then
                        pixelDwell = 0.0000256
                    Else
                        pixelDwell = 0.00000256
                        Recording.SpecialScanMode = "FocusStep"
                        DisplayProgress "Highest Z Step of 1.54 um with no piezo and Fast Z line has been reached. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
                    End If
                End If
            Else ' Use GlobalAcquisitionsTrack as default for pixel dwell
                Recording.SpecialScanMode = "FocusStep"
                Recording.SamplesPerLine = GlobalBackupRecording.SamplesPerLine        'TODO: Check if a value is always given also in frame mode
                pixelDwell = GlobalBackupSampleObservationTime
                
            End If
            
        End If
        
        ''''''''''''''''''''''''''''
        '**Setting for frame scan**'
        ''''''''''''''''''''''''''''
        If ScanFrameToggle.Value Then
        
            Recording.ScanMode = "Stack"                       'This is defining to acquire a Z stack of Z-Y images
            Recording.SamplesPerLine = BSliderFrameSize.Value  'If doing frame autofocussing it uses the userdefined frame size
            Recording.LinesPerFrame = BSliderFrameSize.Value
            
            If CheckBoxHighSpeed.Value Then
               Recording.ScanDirection = 1                     'If Highspeed is selected it uses the bidirectionnal scanning
               pixelDwell = (256 / BSliderFrameSize.Value) * 0.00000256
            Else                                               ' Default is GlobalAcquisitionTrack
               Recording.ScanDirection = GlobalBackupRecording.ScanDirection
               pixelDwell = GlobalBackupSampleObservationTime
            End If
            
        End If
    
    End If  ' If SystemName = "LSM"
    
    '''''''''''''''''''''''''''''
    '*Setting for LIVE system***'
    ' TODO: Legacy Code         '
    '''''''''''''''''''''''''''''
    If SystemName = "LIVE" Then
       If ScanLineToggle.Value Then
           Recording.ScanMode = "ZScan"
           Recording.RtLinePeriod = 1 / 1000 'BSliderScanSpeed.Value
           Recording.RtRegionWidth = 512
           Recording.RtRegionHeight = 1
           If CheckBoxHRZ.Value Then
               Recording.SpecialScanMode = "ZScanner"
           Else ' Not HRZ
               Recording.SpecialScanMode = "OnTheFly"
           End If
       End If
       
       If ScanFrameToggle.Value Then
           Recording.ScanMode = "Stack"
           If CheckBoxHRZ.Value Then                           'piezo
               Recording.SpecialScanMode = "ZScanner"
           Else
               Recording.SpecialScanMode = "OnTheFly"                   ' TODO is OnTheFly possible for frame mode?
               Recording.FramesPerStack = 1201
               Recording.Sample0Z = Range() / 2
               Recording.FrameSpacing = Range() / 1200
           End If
           If CheckBoxHighSpeed.Value Then
               Recording.ScanDirection = 1                  'If Highspeed is selected it uses the bidirectionnal scanning
           End If
           Recording.RtRegionWidth = BSliderFrameSize.Value 'If doing frame autofocussing it uses the userdefined frame size
           Recording.RtBinning = 512 / BSliderFrameSize.Value
           Recording.RtRegionHeight = BSliderFrameSize.Value
       End If
    End If
    
    Sleep (100)
    ' set the pixelDwellTime globally
    NoFrames = CLng(BSliderZRange.Value / BSliderZStep.Value) + 1   'Calculates the number of frames per stack. Clng converts it to a long and rounds up the fraction
    If NoFrames > 2048 Then                                         'overwrites the userdefined value if too many frames have been defined by the user
        NoFrames = 2048
    End If
    Recording.FrameSpacing = BSliderZStep.Value
    Recording.FramesPerStack = NoFrames
    Recording.TrackObjectByMultiplexOrder(AutofocusTrack, 1).SampleObservationTime = pixelDwell
    Lsm5.DsRecording.Copy Recording
    ' need to do it twice:  set new pixelDwell and FrameSpacing (in case it hase not been set on the first go)
    Lsm5.DsRecording.TrackObjectByMultiplexOrder(AutofocusTrack, 1).SampleObservationTime = pixelDwell
    Lsm5.DsRecording.FrameSpacing = BSliderZStep.Value
    ActivateAutofocusTrack = FunSuccess
    
End Function

'''''''''
' ActivateAcquisitionTrack()
' If any of the checkboxes in the AutoFocusForm Acquisition are checked activates themin DsRecording
'   [Recording] In/Out - a DsRecording
' TODO: Test
''''''''''
Public Function ActivateAcquisitionTrack(Recording As DsRecording) As Boolean
    Dim i As Integer
    Dim TrackSuccess As Integer
    Dim FunSuccess As Boolean
    Dim Activate As Boolean
    
    FunSuccess = False

    For i = 1 To TrackNumber
        Activate = False
        If CheckBoxTrack1.Value = True And i = 1 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBoxTrack2.Value = True And i = 2 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBoxTrack3.Value = True And i = 3 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBoxTrack4.Value = True And i = 4 Then
            Activate = True
            FunSuccess = True
        End If
        Recording.TrackObjectByMultiplexOrder(i - 1, Success).Acquire = Activate ' this is not a property specific to this recording
    Next i


    Recording.TimeSeries = True   ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
    Recording.StacksPerRecord = 1
    'can't put Lsm5.DsRecording here. as it is not followed. Why?
    Lsm5.DsRecording.Copy Recording
    Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
    ActivateAcquisitionTrack = FunSuccess
    
End Function



'''''''
'   ActivateAlterAcquisitionTrack
'   Check which track has been activated and for AlternativeAcquisitionTrack set the track properties accordingly
'   TODO: Test
''''''
Public Function ActivateAlterAcquisitionTrack(Recording As DsRecording) As Boolean
    Dim i As Integer
    Dim FunSuccess As Boolean
    Dim Activate As Boolean
    
    FunSuccess = False
    ' Set all tracks to non-acquisition first

    For i = 1 To TrackNumber
        Activate = False
        If CheckBox2ndTrack1.Value = True And i = 1 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBox2ndTrack2.Value = True And i = 2 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBox2ndTrack3.Value = True And i = 3 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBox2ndTrack4.Value = True And i = 4 Then
            Activate = True
            FunSuccess = True
        End If
        Recording.TrackObjectByMultiplexOrder(i - 1, Success).Acquire = Activate
    Next i
    
    Recording.TimeSeries = True  ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
    Recording.StacksPerRecord = 1 ' This is time series stack!
    Recording.Sample0Z = Recording.FrameSpacing * (Recording.FramesPerStack - 1) / 2 ' center the recording
    ActivateAlterAcquisitionTrack = FunSuccess
    
End Function



'''''''''
' ActivateZoomTrack()
' Micropilotpage. This is extra track for micropilot
' TODO: Test and change name
''''''''''
Private Function ActivateZoomTrack(Recording As DsRecording) As Boolean
    Dim i As Integer
    Dim FunSuccess As Boolean
    Dim Activate As Boolean
    
    FunSuccess = False
    ' Set all tracks to non-acquisition first

    For i = 1 To TrackNumber
        Activate = False
        If CheckBoxZoomTrack1.Value = True And i = 1 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBoxZoomTrack2.Value = True And i = 2 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBoxZoomTrack3.Value = True And i = 3 Then
            Activate = True
            FunSuccess = True
        ElseIf CheckBoxZoomTrack4.Value = True And i = 4 Then
            Activate = True
            FunSuccess = True
        End If
        Recording.TrackObjectByMultiplexOrder(i - 1, Success).Acquire = Activate
    Next i
    'Recording.TimeSeries = True  ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
    'Recording.StacksPerRecord = 1 ' This is time series stack!
    Recording.SamplesPerLine = TextBoxZoomFrameSize.Value
    Recording.LinesPerFrame = TextBoxZoomFrameSize.Value
    Recording.ZoomX = TextBoxZoom.Value
    Recording.ZoomY = TextBoxZoom.Value
    Lsm5.DsRecording.TimeSeries = True
    Lsm5.DsRecording.StacksPerRecord = TextBoxZoomCycles.Value
    Recording.Sample0Z = Recording.FrameSpacing * (Recording.FramesPerStack - 1) / 2 ' center the recording

    ActivateZoomTrack = FunSuccess
    
End Function

' TODO a long does it wait
'Wait time in sec?
Sub Wait(PauseTime As Single)
    Dim Start As Single
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
       DoEvents    ' Yield to other processes.
       'Lsm5.DsRecording.StartScanTriggerIn
    Loop
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
    If BSliderZRange.Value > Range() * 0.9 Then
        BSliderZRange.Value = Range() * 0.9
    End If
    If Abs(BSliderZOffset.Value) > Range() * 0.9 Then
        BSliderZOffset.Value = 0
    End If
    BSliderZOffset.Value = BSliderZOffset.Value
    BSliderZRange.Value = BSliderZRange.Value
    BSliderZStep.Value = BlockZStep

End Sub


'''''
' TODO: All block values should use the checkboxes directly
'''''
Public Sub GetBlockValues()
   
    BlockHighSpeed = CheckBoxHighSpeed.Value
    BlockLowZoom = CheckBoxLowZoom.Value
    HRZ = CheckBoxHRZ.Value  ' this is for the piezo
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
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Sub CheckBoxTrack2_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Sub CheckBoxTrack3_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Sub CheckBoxTrack4_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Function AcquisitionTracksOn() As Boolean
    If CheckBoxTrack1 Then
        AcquisitionTracksOn = True
    End If
    If CheckBoxTrack2 Then
        AcquisitionTracksOn = True
    End If
    If CheckBoxTrack3 Then
        AcquisitionTracksOn = True
    End If
    If CheckBoxTrack4 Then
        AcquisitionTracksOn = True
    End If
End Function

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
    FrameNumber = CLng(BSliderZRange.Value / BSliderZStep.Value) + 1
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
Private Sub UpdateZvalues(Grid, MultipleLocation, Z)
        
        
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


''''
' Not anymore in use
''''
Private Sub CreateZoomDatabase(ZoomDatabaseName, HighResExperimentCounter, ZoomExpname)
            'Create ZoomDatabase
            Dim Start As Integer
            Dim bslash As String
            Dim Pos As Long
            Dim NameLength As Long
            Dim Mypath As String
            
            Start = 1
            bslash = "\"
            Pos = Start
            Do While Pos > 0
                Pos = InStr(Start, DatabaseTextbox.Value, bslash)
                If Pos > 0 Then
                    Start = Pos + 1
                End If
            Loop
            
            Mypath = DatabaseTextbox.Value + bslash
            NameLength = Len(DatabaseTextbox.Value)
            ZoomExpname = Strings.Right(DatabaseTextbox.Value, NameLength - Start + 1)
           ' NameLength = Len(Myname)
           ' Myname = Strings.Left(Myname, NameLength - 4)
            ZoomDatabaseName = Mypath & ZoomExpname & "_" & TextBoxFileName.Value & LocationName & "_R" & RepetitionNumber & "_Exp" & HighResExperimentCounter & "_zoom"
            ' Lsm5.NewDatabase (ZoomDatabaseName)
           ' ZoomDatabaseName = ZoomDatabaseName & "\" & Myname & "_zoom.mdb"
    
End Sub

Private Sub CreateAlterImageDatabase(AlterDatabaseName, Mypath)
        Dim Start As Integer
        Dim bslash As String
        Dim Pos As Long
        Dim NameLength As Long
        Dim Myname As String

         Start = 1
         bslash = "\"
         Pos = Start
         Do While Pos > 0
             Pos = InStr(Start, DatabaseTextbox.Value, bslash)
             If Pos > 0 Then
                 Start = Pos + 1
             End If
         Loop
         Mypath = Strings.Left(DatabaseTextbox.Value, Start - 1)
         NameLength = Len(DatabaseTextbox.Value)
         Myname = Strings.Right(DatabaseTextbox.Value, NameLength - Start + 1)
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
Private Function MicroscopePilot(RecordingDoc As DsRecordingDoc, ByVal BleachingActivated As Boolean, HighResExperimentCounter As Integer, HighResCounter As Integer _
, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    
    Dim ZoomNumber As Integer
    Dim code As String
    Dim codeArray() As String
        
    ' Get Code from Windows registry
    code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="code")
    DisplayProgress "Waiting for Micropilot...", RGB(0, &HC0, 0)
    DoEvents
    Do While (code = "1" Or code = "0")
        ' TODO: Check Code
        Sleep (100)
        code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", _
                  Key:="Code")
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                MicroscopePilot = False
                Exit Function
            End If
        End If
    Loop
    
    DisplayProgress "Received Code " + CStr(code), RGB(0, &HC0, 0)
    
    'TODO: create a better procedure to check for cells
'    If (CheckBoxGridScan_FindGoodPositions) Then
'
'        codeArray = Split(code, "_")
'
'        nGoodCells = CInt(codeArray(1))
'        minGoodCellsPerImage = CInt(codeArray(2))
'        minGoodCellsPerWell = CInt(codeArray(3))
'
'        GoTo Mark
'
'    End If
'

    If code = "2" Then   ' no interesting cell
    
        DisplayProgress "Micropilot Code 2", RGB(0, &HC0, 0)
        SaveSetting "OnlineImageAnalysis", "macro", "Refresh", 0
        GoTo Mark '(because Image does not show any interesting pheotype)
    
    ElseIf code = "4" Then   'store position in a list
    
        DisplayProgress "Micropilot Code 4", RGB(0, &HC0, 0)
        HighResCounter = HighResCounter + 1 ' Counts the postions, where Highres Imaging will be carried out
        ' store postion from windows registry in array
        StorePositioninHighResArray HighResArrayX, HighResArrayY, HighResArrayZ, HighResCounter
        
    ElseIf code = "5" Then  ' start Highres Batch Imaging 1 to n postions
        
        DisplayProgress "Micropilot Code 5", RGB(0, &HC0, 0)
        
        ' store postion from windows registry in array
       
        StorePositioninHighResArray HighResArrayX, HighResArrayY, HighResArrayZ, HighResCounter
        ' BatchHighresImagingRoutine
        ' HERE THE IMAGES ARE ACQUIRED
        BatchHighresImagingRoutine RecordingDoc, HighResArrayX, HighResArrayY, HighResArrayZ, HighResCounter, HighResExperimentCounter, Row, Col, RowSub, ColSub
        HighResExperimentCounter = HighResExperimentCounter + 1 ' counts the number of highres-multipositionexperiments (important for naming the datafolder)

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
    
    MicroscopePilot = True
            
End Function

'''''
'   Private Sub StartAlternativeImaging(RecordingDoc As DsRecordingDoc, StartTime As Double, _
'   AlterDatabaseName As String, name As String)
'   Alternative Acquisition in every .. round
'   TODO: Bring it up to normal setting for all
'''''
Private Function StartAlternativeImaging(RecordingDoc As DsRecordingDoc, _
FilePath As String, name As String) As Boolean
    
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    If RecordingDoc Is Nothing Then
        Set RecordingDoc = Lsm5.NewScanWindow
        While RecordingDoc.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
    DisplayProgress "Acquiring Additional Track...", RGB(0, &HC0, 0)
    
    If Not ActivateAlterAcquisitionTrack(GlobalAltRecording) Then           'An additional control....
        MsgBox "No track selected for Additional Acquisition! Cannot Acquire!"
        StartAlternativeImaging = False
        Exit Function
    End If

    
    ' get and set the values from the Form
    GlobalAltRecording.ZoomX = TextBoxAlterZoom.Value
    GlobalAltRecording.ZoomY = TextBoxAlterZoom.Value
    GlobalAltRecording.FramesPerStack = TextBoxAlterNumSlices.Value
    GlobalAltRecording.FrameSpacing = TextBoxAlterInterval.Value

     If GlobalAltRecording.FramesPerStack > 1 Then
        GlobalAltRecording.ScanMode = "Stack"
        If CheckBoxHRZ.Value Then
            GlobalAltRecording.SpecialScanMode = "ZScanner" ' this is a problem if people do not have a piezo
        Else
            GlobalAltRecording.SpecialScanMode = "FocusStep"
        End If
     End If
     
     Lsm5.DsRecording.Copy GlobalAltRecording
     Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
     
     ' take the image
     ScanToImageNew RecordingDoc
     ' TODO Check this
     While AcquisitionController.IsGrabbing
        Sleep (100)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                StartAlternativeImaging = False
                Exit Function
            End If
        End If
     Wend
     
     RecordingDoc.SetTitle name
    
     SaveDsRecordingDoc RecordingDoc, FilePath
     StartAlternativeImaging = True
        
End Function

'''
'   StorePositioninHighResArray(HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, HighResCounter As Integer)
'   TODO: Test stricter way of passing arguments
''''
Private Sub StorePositioninHighResArray(HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, HighResCounter As Integer)
    
    ' store postion from windows registry in array
    
    Dim zoomXoffset As Double
    Dim zoomYoffset As Double
    Dim X As Double
    Dim Y As Double
    Dim PixelSize As Double

    'zoomXoffset = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="offsetx")
    'zoomYoffset = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="offsety")
    
    zoomXoffset = CDbl(GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="offsetx"))
    zoomYoffset = CDbl(GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="offsety"))
    
    
    'MsgBox ("zoomXoffset,zoomYoffset " + CStr(zoomXoffset) + "," + CStr(zoomYoffset))
    
    
    If CheckBoxHRZ.Value Then
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
    X = Lsm5.Hardware.CpStages.PositionX
    Y = Lsm5.Hardware.CpStages.PositionY
    
    'MsgBox ("PixelSize " + CStr(PixelSize))
    'MsgBox ("zoomXoffset*ps,zoomYoffset*ps " + CStr(zoomXoffset * PixelSize) + "," + CStr(zoomYoffset * PixelSize))
    
    
    HighResArrayX(HighResCounter) = X - zoomXoffset * PixelSize
    HighResArrayY(HighResCounter) = Y + zoomYoffset * PixelSize
    HighResArrayZ(HighResCounter) = Lsm5.Hardware.CpFocus.Position
   ' MsgBox "Current Z Position = " + CStr(Lsm5.Hardware.CpFocus.Position)
    DisplayProgress "Micropilot - Position stored", RGB(0, &HC0, 0)

End Sub


'''''
'   BatchHighresImagingRoutine(RecordingDoc As DsRecordingDoc, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, _
'   HighResCounter As Integer, HighResExperimentCounter As Integer)
'   TODO: Test stricter way of passing arguments
'''''
Private Function BatchHighresImagingRoutine(RecordingDoc As DsRecordingDoc, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, _
HighResCounter As Integer, HighResExperimentCounter As Integer, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    
    Dim FileNameID  As String
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
    
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
     
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
    'CreateZoomDatabase ZoomDatabaseName, HighResExperimentCounter, ZoomExpname
    
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
        
                ' Move to Positon in x, y, z for Highresscan
                DisplayProgress "Micropilot Code 5 - Move to Position", RGB(0, &HC0, 0)

                If Not FailSafeMoveStageXY(HighResArrayX(highrespos), HighResArrayY(highrespos)) Then
                    Exit Function
                End If
                
                If Not FailSafeMoveStageZ(HighResArrayZ(highrespos)) Then
                    Exit Function
                End If
                
                'Autofocus. This does an extra Autofocus also for the HighresImaging with the same parameters as Autofocus only Zoffset chnages
                If CheckBoxZoomAutofocus.Value = True Then
                    DisplayProgress "Micropilot Code 5 - Autofocus acquire", RGB(0, &HC0, 0)
                    If Not Autofocus_StackShift(RecordingDoc) Then
                        BatchHighresImagingRoutine = False
                        Exit Function
                    End If
                    ' move the xyz to the right position
                    DisplayProgress "Micropilot Code 5 - Autofocus move stage", RGB(0, &HC0, 0)
                    ComputeShiftedCoordinates XMass, YMass, ZMass + TextBoxZoomAutofocusZOffset.Value, X, Y, Z
                    If Not FailSafeMoveStageXY(X, Y) Then
                        BatchHighresImagingRoutine = False
                        Exit Function
                    End If
                End If
        
                ' Load AcquisitionSettings
                If Not ActivateZoomTrack(GlobalZoomRecording) Then
                    MsgBox " No Track selected for Micropilot! Macro stops here"
                    ScanStop = True
                    StopAcquisition
                    Exit Function
                End If
                    
                
                Lsm5.DsRecording.Copy GlobalZoomRecording
                'set the correct dwelltime
                Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
                
                If BleachingActivated Then
                                
                    DisplayProgress "Bleaching...", &HFF00FF
                        
                    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
                    If Success Then ' do only one stack
                        Track.Acquire = True
                        Lsm5.DsRecording.TimeSeries = True
                        Lsm5.DsRecording.StacksPerRecord = TextBoxZoomCycles.Value
                        Lsm5.DsRecording.FramesPerStack = 1
                        Lsm5.DsRecording.Sample0Z = 0
                        Track.TimeBetweenStacks = TextBoxZoomCycleDelay.Value
                        'MsgBox "Track.IsBleachTrack " + CStr(Track.IsBleachTrack)
                        'MsgBox "BleachScanNumber " + CStr(Track.BleachScanNumber)
                        DoEvents
                        Track.UseBleachParameters = True            'Bleach parameters are lasers lines, bleach iterations... stored in the bleach control window
                        'BleachStartTable(RepetitionNumber) = GetTickCount      'Get the time right before bleach to store this in the image metadata
                        Set RecordingDoc = Lsm5.StartScan
                        'TODO Check
                        Do While RecordingDoc.IsBusy
                            Sleep (100)
                            If GetInputState() <> 0 Then
                                DoEvents
                                If ScanStop Then
                                BatchHighresImagingRoutine = False
                                Exit Function
                            End If
                        End If
                        Loop
                        
                        Track.UseBleachParameters = False  'switch off the bleaching
                        Lsm5.DsRecording.TimeSeries = False
                        
                    Else
                    
                        MsgBox ("Could not set bleach track. Did not bleach.")
                    
                    End If
                
                                 
                    'Save Image  ' modified by Tischi
                    FileNameID = FileName((Row - 1) * UBound(posGridX, 1) + Col, (RowSub - 1) * UBound(posGridXsub, 1) + ColSub, -1)
                    ' e.g. name--Bleach--HRExp00001--HRPos00001--W00001--P00001.lsm
                    fullpathname = DatabaseTextbox.Value & "\" & TextBoxFileName.Value & "--Bleach" & "--HRExp" & ZeroString(5 - Len(CStr(HighResExperimentCounter))) & _
                    HighResExperimentCounter & "--HRPos" & ZeroString(5 - Len(CStr(highrespos))) _
                    & highrespos & FileNameID & ".lsm"
                    SaveDsRecordingDoc RecordingDoc, fullpathname
                    DisplayProgress "Micropilot Code 5 - SaveImage", RGB(0, &HC0, 0)
                    
                Else ' normal acquistion (non bleaching mode)
                    
                    Lsm5.DsRecording.ScanMode = "Stack"
                    If CheckBoxHRZ.Value Then
                        Lsm5.DsRecording.SpecialScanMode = "ZScanner"
                    Else
                        Lsm5.DsRecording.SpecialScanMode = "FocusStep"
                    End If
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
                                BatchHighresImagingRoutine = False
                                Exit Function
                            End If
                        End If
                    Wend
                                    
                    FileNameID = FileName((Row - 1) * UBound(posGridX, 1) + Col, (RowSub - 1) * UBound(posGridXsub, 1) + ColSub, RepetitionNumber)
                    fullpathname = DatabaseTextbox.Value & "\" & TextBoxFileName.Value & "--HRExp" & FileNameID & "\"
                    FileNameID = FileName((Row - 1) * UBound(posGridX, 1) + Col, (RowSub - 1) * UBound(posGridXsub, 1) + ColSub, ZoomRepetitionNumber)
                    
                    ' e.g. name--HRExp00001--HRPos00001--W00001--P00001--T00001.lsm where T is for ZoomRepetitionNumber
                    fullpathname = fullpathname & TextBoxFileName.Value & "--HRExp" & ZeroString(5 - Len(CStr(HighResExperimentCounter))) & _
                    HighResExperimentCounter & "--HRPos" & ZeroString(5 - Len(CStr(highrespos))) _
                    & highrespos & FileNameID & ".lsm"

                    SaveDsRecordingDoc RecordingDoc, fullpathname
        
                    DisplayProgress "Micropilot Code 5 - SaveImage", RGB(0, &HC0, 0)
                        
                    ' Tischi: Here the Location-tracking in High-resmode code needs be added!

                End If ' BleachingActivated
          
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
                        BatchHighresImagingRoutine = False
                        Exit Function
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
    
    BatchHighresImagingRoutine = True
End Function



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
                Export.FinishExport
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


