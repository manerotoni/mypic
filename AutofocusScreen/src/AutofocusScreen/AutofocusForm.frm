VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutofocusForm 
   Caption         =   "AutofocusScreen"
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
'Option Explicit 'force to declare all variables

Private shlShell As Shell32.Shell
Private shlFolder As Shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Version Description''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AutofocusScreen_ZEN_v2.1.3.10
'''''''''''''''''''''End: Version Description'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const Version = " v2.1.3.12"
Public posTempZ  As Double                  'This is position at start after pushing AutofocusButton
Private Const DebugCode = False             'sets key to run tests visible or not
Private Const ReleaseName = True            'this adds the ZEN version
Private Const LogCode = True                'sets key to run tests visible or not

Private AlterImageInitialize As Boolean ' first time aternative image is activated values from acquisition are loaded. Then variable is ste to false
Private ZoomImageInitialize As Boolean  ' first time ZoomImage/Micropilot is activated values from acquisition are loaded



Private Sub AutofocusAlgorithm_Change()
    If AutofocusAlgorithm.Value = "external" Then
        SaveAFImage.Value = True
    End If
End Sub

''''''
' UserForm_Initialize()
'   Function called from e.g. AutoFocusForm.Show
'   Load and initialize form
'''''
Public Sub UserForm_Initialize()
    ZEN = "2011"
    'Setting of some global variables
    LogFileNameBase = ""
    Log = LogCode
    
    Me.Caption = Me.Caption + Version + " for ZEN "
    
    If ReleaseName Then
        Me.Caption = Me.Caption + ZEN
    End If
    FormatUserForm (Me.Caption) ' make minimizing button available
    AutofocusForm.Show
    
    StageSettings MirrorX, MirrorY, ExchangeXY
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
    LoopingTimerUnit = 1
    CommandTimeSec.BackColor = &HFF8080
    BlockRepetitions = 1
    ReDim Preserve GlobalImageIndex(BlockRepetitions)
    ScanLineToggle.Value = True
    LocationTextLabel.Caption = ""
    GlobalProject = "AutofocusScreen2.1"
    GlobalProjectName = GlobalProject + ".lvb"
    HelpNamePDF = "AutofocusScreen_help.pdf"
    UsedDevices40 bLSM, bLIVE, bCamera
    SystemVersionOffset         ' extra offset depending on macroscope

    ' Set standard values for Autofocus
    ' blSM is a flag to decide weather systen is LSM (ZEN is LSM for instance). LIVE is 5Live not anymore in use?
    If bLSM Then
        SystemName = "LSM"
        AutofocusMaxSpeed.Value = True
        AutofocusFrameSize.Min = 16
        AutofocusFrameSize.Max = 1024
        AutofocusLineSize.Min = 16
        AutofocusLineSize.Max = 1024
        AutofocusFrameSize.Step = 8
        AutofocusLineSize.Step = 8
        AutofocusFrameSize.StepSmall = 4
        AutofocusLineSize.StepSmall = 4
        'Lsm5Vba.Application.ThrowEvent eRootReuse, 0 ' not clear what it does. On ZEN2012 the marked positions are reset
        DoEvents
    ElseIf bLIVE Then
        SystemName = "LIVE"
        AutofocusFrameSize.Min = 128
        AutofocusFrameSize.Max = 1024
        AutofocusFrameSize.Step = 128
        AutofocusFrameSize.StepSmall = 128
        AutofocusLineSize.Min = 128
        AutofocusLineSize.Max = 1024
        AutofocusLineSize.Step = 128
        AutofocusLineSize.StepSmall = 128
        'Lsm5Vba.Application.ThrowEvent eRootReuse, 0
        DoEvents
    ElseIf bCamera Then
        SystemName = "Camera"
    End If
    
    'TODO: Check if GUI is available (ZEN2011 onward). How do you do this!!

    'Set default value
    ScanLineToggle.Value = True
    AutofocusZOffset.Value = 0
    AutofocusZRange.Value = 80
    AutofocusZStep.Value = 0.1
    AutofocusLowZoom.Value = False
    ActiveAutofocus.Value = True
    FillAutofocusAlgorithm
    
    'Set standard values for Post-Acquisition tracking
    TrackingToggle.Value = False
    SwitchEnableTrackingToggle (False)
    TrackingToggle.Enabled = False
    
    'Set standard values for Looping
    LoopingRepetitions = 300
    LoopingRepetitionTime = 1
    
    'Set standard values for Micropilot
    ActiveMicropilot.Value = False
    SwitchEnableMicropilotPage (False)
    MicropilotAutofocus.Value = False
    SwitchEnableZoomAutofocus (False)
    
    'Set standard values for Gridscan
    ActiveGridScan.Value = False
    SwitchEnableGridScanPage (False)
    
    'Set standard values for Additional Acquisition
    ActiveAlterImage.Value = False
    SwitchEnableAlterImagePage (False)
    
    'Set default for bleach
    ActiveBleach.Value = False
    SwitchEnableBleachPage (False)
    
    
    'Set Database name
    DatabaseTextbox.Value = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="OutputFolder")
    
    'Set repetition and locations
    RepetitionNumber = 1
    locationNumber = 1
    Set FileSystem = New FileSystemObject
    'If we log a new logfile is created
    If LogCode And LogFileNameBase <> "" Then
        LogFileName = LogFileNameBase
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        LogFile.Close
        Log = True
    Else
        Log = False
    End If

    AutofocusTrackZ.Visible = DebugCode
    MultiPage1.Pages("TestsPage").Visible = DebugCode
    
    
    AlterImageInitialize = True
    ZoomImageInitialize = True
    
    If ZEN = "2010" Then
        ZBacklash = 0.5
    ElseIf ZEN = "2011" Then
        ZBacklash = 0.5
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
    Dim SuccessRecenter As Boolean
    AutoFindTracks
    SwitchEnableAutofocusPage ActiveAutofocus
    SwitchEnableAlterImagePage ActiveAlterImage
    SwitchEnableMicropilotPage ActiveMicropilot
    
    PubSearchScan = False
    NoReflectionSignal = False
    PubSentStageGrid = False
    
    '  AutofocusForm.Caption = GlobalProject + " for " + SystemName
    BleachingActivated = False
    FocusMapPresent = False
    'This sets standard values for all task we want to do. This will be changed by the macro
    
    If AutofocusHRZ Then
        Lsm5.Hardware.CpHrz.Leveling
        While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
    
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Recenter_pre posTempZ, SuccessRecenter, ZEN
    
    Set GlobalAutoFocusRecording = Lsm5.CreateBackupRecording
    Set GlobalAcquisitionRecording = Lsm5.CreateBackupRecording
    Set GlobalMicropilotRecording = Lsm5.CreateBackupRecording
    Set GlobalBleachRecording = Lsm5.CreateBackupRecording
    Set GlobalAltRecording = Lsm5.CreateBackupRecording
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    GlobalAutoFocusRecording.Copy Lsm5.DsRecording
    GlobalAcquisitionRecording.Copy Lsm5.DsRecording
    GlobalMicropilotRecording.Copy Lsm5.DsRecording
    GlobalBleachRecording.Copy Lsm5.DsRecording
    GlobalAltRecording.Copy Lsm5.DsRecording
    GlobalBackupRecording.Copy Lsm5.DsRecording ' this will not be changed remains always the same
    GlobalBackupSampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    Dim i As Long
    Dim NrTracks As Long
    ReDim GlobalBackupActiveTracks(Lsm5.DsRecording.TrackCount)
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       GlobalBackupActiveTracks(i) = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i
    If Not Recenter_post(posTempZ, SuccessRecenter, ZEN) Then
        Exit Sub
    End If
    Set FileSystem = New FileSystemObject
    'If we log a new logfile is created
    If LogCode And LogFileNameBase <> "" Then
        LogFileName = LogFileNameBase
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        LogFile.Close
        Log = True
    Else
        Log = False
    End If
End Sub


'''''
'   SaveSettings(FileName As String)
'   SaveSettings of the UserForm AutofocusForm in file name FileName.
'   Name should correspond exactly to name used in Form
'''''
Private Sub SaveSettings(fileName As String)
    Dim iFileNum As Integer
    Close
    On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open fileName For Output As iFileNum
    
    Print #iFileNum, "% Settings for AutofocusMacro for ZEN " & ZEN & "  " & Version

    'Single MultipelocationToggle
    Print #iFileNum, "% Single Multiple "
    Print #iFileNum, "MultipleLocationToggle " & MultipleLocationToggle.Value
    Print #iFileNum, "SingleLocationToggle " & SingleLocationToggle.Value
    
    'Autofocus
    Print #iFileNum, "% Autofocus "
    Print #iFileNum, "ActiveAutofocus " & ActiveAutofocus.Value
    Print #iFileNum, "AutofocusTrack1 " & AutofocusTrack1.Value
    Print #iFileNum, "AutofocusTrack2 " & AutofocusTrack2.Value
    Print #iFileNum, "AutofocusTrack3 " & AutofocusTrack3.Value
    Print #iFileNum, "AutofocusTrack4 " & AutofocusTrack4.Value
    Print #iFileNum, "AutofocusMaxSpeed " & AutofocusMaxSpeed.Value
    Print #iFileNum, "AutofocusLowZoom " & AutofocusLowZoom.Value
    Print #iFileNum, "AutofocusHRZ " & AutofocusHRZ.Value
    Print #iFileNum, "AutofocusFastZline " & AutofocusFastZline.Value
    Print #iFileNum, "AFeveryNth " & AFeveryNth.Value
    Print #iFileNum, "AutofocusTrackZ " & AutofocusTrackZ.Value
    Print #iFileNum, "AutofocusTrackXY " & AutofocusTrackXY.Value
    Print #iFileNum, "ScanLineToggle " & ScanLineToggle.Value
    Print #iFileNum, "ScanFrameToggle " & ScanFrameToggle.Value
    Print #iFileNum, "AutofocusLineSize " & AutofocusLineSize.Value
    Print #iFileNum, "AutofocusFrameSize " & AutofocusFrameSize.Value
    Print #iFileNum, "AutofocusZOffset " & AutofocusZOffset.Value
    Print #iFileNum, "AutofocusZRange " & AutofocusZRange.Value
    Print #iFileNum, "AutofocusZStep " & AutofocusZStep.Value
    Print #iFileNum, "SaveAFImage " & SaveAFImage.Value
    Print #iFileNum, "AutofocusAlgorithm " & AutofocusAlgorithm.Value
    
    'Acquisition
    Print #iFileNum, "% Acquisition "
    Print #iFileNum, "AcquisitionTrack1 " & AcquisitionTrack1.Value
    Print #iFileNum, "AcquisitionTrack2 " & AcquisitionTrack2.Value
    Print #iFileNum, "AcquisitionTrack3 " & AcquisitionTrack3.Value
    Print #iFileNum, "AcquisitionTrack4 " & AcquisitionTrack4.Value

    
    'PostAcquisitionTracking
    Print #iFileNum, "% PostAcquisitionTracking "
    Print #iFileNum, "TrackingToggle " & TrackingToggle.Value
    Print #iFileNum, "ComboBoxTrackingChannel " & ComboBoxTrackingChannel.Value
    Print #iFileNum, "PostTrackXY " & PostTrackXY.Value
    Print #iFileNum, "PostTrackZ " & PostTrackZ.Value
    
    'Looping
    Print #iFileNum, "% Looping "
    Print #iFileNum, "LoopingTimerUnit " & LoopingTimerUnit
    Print #iFileNum, "LoopingRepetitionTime " & LoopingRepetitionTime.Value
    Print #iFileNum, "LoopingInterval " & LoopingInterval.Value
    Print #iFileNum, "LoopingRepetitions " & LoopingRepetitions.Value
    
    'Output
    Print #iFileNum, "% Output "
    Print #iFileNum, "DatabaseTextbox " & DatabaseTextbox.Value
    Print #iFileNum, "TextBoxFileName " & TextBoxFileName.Value
    
    'Micropilot
    Print #iFileNum, "% MicroPilot "
    Print #iFileNum, "ActiveMicropilot " & ActiveMicropilot.Value
    Print #iFileNum, "MicropilotTrack1 " & MicropilotTrack1.Value
    Print #iFileNum, "MicropilotTrack2 " & MicropilotTrack2.Value
    Print #iFileNum, "MicropilotTrack3 " & MicropilotTrack3.Value
    Print #iFileNum, "MicropilotTrack4 " & MicropilotTrack4.Value
    Print #iFileNum, "MicropilotRepetitions " & MicropilotRepetitions.Value
    Print #iFileNum, "MicropilotRepetitionTime " & MicropilotRepetitionTime.Value
    Print #iFileNum, "MicropilotMaxPositions " & MicropilotMaxPositions.Value
    Print #iFileNum, "MicropilotFrameSize " & MicropilotFrameSize.Value
    Print #iFileNum, "MicropilotZOffset " & MicropilotZOffset.Value
    Print #iFileNum, "MicropilotZSlices " & MicropilotZSlices.Value
    Print #iFileNum, "MicropilotZStep " & MicropilotZStep.Value
    Print #iFileNum, "MicropilotZoom " & MicropilotZoom.Value
    Print #iFileNum, "MicropilotAutofocus " & MicropilotAutofocus.Value
    
    'Additional Acquisition
    Print #iFileNum, "% Additional Acquisition "
    Print #iFileNum, "ActiveAlterImage " & ActiveAlterImage.Value
    Print #iFileNum, "AlterTrack1 " & AlterTrack1.Value
    Print #iFileNum, "AlterTrack2 " & AlterTrack2.Value
    Print #iFileNum, "AlterTrack3 " & AlterTrack3.Value
    Print #iFileNum, "AlterTrack4 " & AlterTrack4.Value
    Print #iFileNum, "RoundAlterTrack " & RoundAlterTrack.Value
    Print #iFileNum, "RoundAlterLocation " & RoundAlterLocation.Value
    Print #iFileNum, "AlterFrameSize " & AlterFrameSize.Value
    Print #iFileNum, "AlterZOffset " & AlterZOffset.Value
    Print #iFileNum, "AlterZSlices " & AlterZSlices.Value
    Print #iFileNum, "AlterZStep " & AlterZStep.Value
    Print #iFileNum, "AlterZoom " & AlterZoom.Value

    'Grid Acquisition
    Print #iFileNum, "% Additional Acquisition "
    Print #iFileNum, "ActiveGridScan " & ActiveGridScan.Value
    Print #iFileNum, "GridScan_validGridDefault " & GridScan_validGridDefault.Value
    Print #iFileNum, "GridScan_nRow " & GridScan_nRow.Value
    Print #iFileNum, "GridScan_nColumn " & GridScan_nColumn.Value
    Print #iFileNum, "GridScan_dRow " & GridScan_dRow.Value
    Print #iFileNum, "GridScan_dColumn " & GridScan_dColumn.Value
    Print #iFileNum, "GridScan_refRow " & GridScan_refRow.Value
    Print #iFileNum, "GridScan_refColumn " & GridScan_refColumn.Value
    Print #iFileNum, "GridScan_nRowsub " & GridScan_nRowsub.Value
    Print #iFileNum, "GridScan_nColumnsub " & GridScan_nColumnsub.Value
    Print #iFileNum, "GridScan_dRowsub " & GridScan_dRowsub.Value
    Print #iFileNum, "GridScan_dColumnsub " & GridScan_dColumnsub.Value

       'Additional Acquisition
    Print #iFileNum, "% Bleach Acquisition "
    Print #iFileNum, "ActiveBleach " & ActiveBleach.Value
    Print #iFileNum, "BleachTrack1 " & BleachTrack1.Value
    Print #iFileNum, "BleachTrack2 " & BleachTrack2.Value
    Print #iFileNum, "BleachTrack3 " & BleachTrack3.Value
    Print #iFileNum, "BleachTrack4 " & BleachTrack4.Value
    Print #iFileNum, "BleachRepetitions " & BleachRepetitions.Value
    Print #iFileNum, "BleachRepetitionTime " & BleachRepetitionTime.Value
    Print #iFileNum, "BleachFrameSize " & BleachFrameSize.Value
    Print #iFileNum, "BleachZOffset " & BleachZOffset.Value
    Print #iFileNum, "BleachZSlices " & BleachZSlices.Value
    Print #iFileNum, "BleachZStep " & BleachZStep.Value
    Print #iFileNum, "BleachZoom " & BleachZoom.Value

    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to open " & fileName & " for saving settings"
End Sub

''''
'   LoadSettings(FileName As String)
'   LoadSettings of Form from FileName
''''
Private Sub LoadSettings(fileName As String)
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    Dim Entries() As String
    Close
    On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open fileName For Input As iFileNum
    Do While Not EOF(iFileNum)
  
            Line Input #iFileNum, Fields
            While Left(Fields, 1) = "%"
                Line Input #iFileNum, Fields
            Wend
            FieldEntries = Split(Fields, " ", 2)
            If FieldEntries(0) = "LoopingTimerUnit" Then
                LoopingTimerUnit = CDbl(FieldEntries(1))
                If LoopingTimerUnit = 60 Then
                    CommandTimeMin_Click
                Else
                    CommandTimeSec_Click
                End If
            Else
                On Error GoTo nextLine
                Me.Controls(FieldEntries(0)).Value = FieldEntries(1)
nextLine:
            End If
    Loop
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & fileName & " for AutofocusScreen settings"
End Sub

''''
'   ButtonSaveSettings_Click()
'   Open a dialog to save setting of the macro
''''
Private Sub ButtonSaveSettings_Click()
    Dim Filter As String, fileName As String
    Dim Flags As Long
  
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST
    Filter$ = "Settings (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
            
    
    fileName = CommonDialogAPI.ShowSave(Filter, Flags, "", DatabaseTextbox.Value, "Save AutofocusScreen settings")
    
    If fileName <> "" Then
        If Right(fileName, 4) <> ".ini" Then
            fileName = fileName & ".ini"
        End If
        SaveSettings fileName
    End If
    
End Sub

''''
'   ButtonSaveSettings_Click()
'   Open a dialog to save setting of the macro
''''
Private Sub ButtonLoadSettings_Click()
    Dim Filter As String, fileName As String
    Dim Flags As Long
  
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST
    Filter$ = "Settings (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
            
    'Filter = "ini file (*.ini) |*.ini"
    
    fileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DatabaseTextbox.Value, "Load AutofocusScreen settings")
    
    If fileName <> "" Then
        LoadSettings fileName
    End If
End Sub

''''
'   FocusMap_Click()
'   create a focusMap using teh Autofocus Channel
''''
Private Sub FocusMap_Click()
    ' This will run just in the AutofocusMode all the AcquisitionTracks are set off
    SetDatabase
    SaveSettings GlobalDataBaseName & "\tmpSettings.ini"
    AcquisitionTracksSetOff
    'change values
    LoopingRepetitions.Value = 1
    BlockTimeDelay = 0
    CommandTimeSec_Click
    ActiveMicropilot.Value = False
    ActiveAlterImage.Value = False
    StartButton_Click
    WritePosFile GlobalDataBaseName & "\" & TextBoxFileName.Value & "positionsGrid.csv", posGridX, posGridY, posGridZ
    'Return to original values for the
    LoadSettings GlobalDataBaseName & "\tmpSettings.ini"
End Sub


Private Sub AutofocusTrackXY_Click()
    If AutofocusTrackXY Then
        PostTrackXY.Value = Not AutofocusTrackXY
    End If
End Sub

Private Sub AutofocusFastZline_Click()
    If AutofocusFastZline Then
        LocationTextLabel.Caption = "WARNING: " & vbCrLf & _
        "ScanLine with FastZLine is fast but can have low reliability." & vbCrLf & _
        "Please test reproducibility of LineScan with FastZLine before using AutofocusScreen. Otherwise use piezo or normal mode with large Z Step and smaller Z Range."
        LocationTextLabel.BackColor = &H80FF&
    Else
        LocationTextLabel.Caption = " "
        LocationTextLabel.BackColor = &HFFFF&
    End If
        
End Sub

Private Sub AutofocusHRZ_Click()
    AutofocusFastZline.Enabled = Not AutofocusHRZ
    If AutofocusFastZline And Not AutofocusHRZ Then
        LocationTextLabel.Caption = "WARNING: " & vbCrLf & _
        "ScanLine with FastZLine is fast but can have low reliability." & vbCrLf & _
        "Please test reproducibility of LineScan with FastZLine before using AutofocusScreen. Otherwise use piezo or normal mode with large Z Step and smaller Z Range."
        LocationTextLabel.BackColor = &H80FF&
    Else
        LocationTextLabel.Caption = " "
        LocationTextLabel.BackColor = &HFFFF&
    End If
End Sub


Private Sub PostTrackXY_Click()
    If PostTrackXY Then
        AutofocusTrackXY.Value = Not PostTrackXY
    End If
End Sub



Private Sub StopAfterRepetition_Click()
    If Not Running Then
        StopAfterRepetition.Value = False
        StopAfterRepetition.BackColor = &H8000000F
    Else
        If StopAfterRepetition.Value Then
            StopAfterRepetition.BackColor = 12648447
        Else
            StopAfterRepetition.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub PauseButton_Click()
    If Not Running Then
        ScanPause = False
        PauseButton.Value = False
        PauseButton.Caption = "PAUSE"
        PauseButton.BackColor = &H8000000F
    Else
        If PauseButton.Value Then
            ScanPause = True
            PauseButton.Caption = "RESUME"
            PauseButton.BackColor = 12648447
        Else
            ScanPause = False
            PauseButton.Caption = "PAUSE"
            PauseButton.BackColor = &H8000000F
        End If
    End If
End Sub


Private Sub MicropilotZSlices_Change()
    MicropilotZSlices.Value = Round(MicropilotZSlices.Value)
End Sub


''''''
'   ActiveAutofocus_Click()
'       Activates Autofocus. If not toggled only Acquisition track is used
'''''
Private Sub ActiveAutofocus_Click()
    If ActiveAutofocus.Value Then
        SwitchEnableAutofocusPage (True)
        AutofocusTrackZ.Value = False
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
    AutofocusMaxSpeed.Enabled = Enable
    AutofocusHRZ.Enabled = Enable
    AutofocusLowZoom.Enabled = Enable
    AFeveryNth.Enabled = Enable
    AFeveryNthLabel.Enabled = Enable
    AFeveryNthLabel2.Enabled = Enable
    ScanLineToggle.Enabled = Enable
    ScanFrameToggle.Enabled = Enable
    FrameSizeLabel.Enabled = Enable
    AutofocusFrameSize.Enabled = Enable
    AutofocusLineSize.Enabled = Enable
    AutofocusZOffset.Enabled = Enable
    SliderZOffsetLabel.Enabled = Enable
    AutofocusZRange.Enabled = Enable
    SliderZRangeLabel.Enabled = Enable
    AutofocusZStep.Enabled = Enable
    SliderZStepLabel.Enabled = Enable
    AutofocusTrack1.Enabled = Enable
    AutofocusTrack2.Enabled = Enable
    AutofocusTrack3.Enabled = Enable
    AutofocusTrack4.Enabled = Enable
    AutofocusTrackZ.Enabled = Enable
    AutofocusTrackXY.Enabled = Enable
    AutofocusAlgorithmLabel.Enabled = Enable
    AutofocusAlgorithmLabel.Enabled = Enable
    
    AutofocusFastZline.Enabled = Enable
    SaveAFImage.Enabled = Enable
    
End Sub

'''
' fills popup menu for UseAF compo box
'''
Private Sub FillAutofocusAlgorithm()
    AutofocusAlgorithm.AddItem "internal"
    AutofocusAlgorithm.AddItem "external"
    AutofocusAlgorithm.Value = AutofocusAlgorithm.List(0)
End Sub


''''''
'   AutofocusZOffset_Change()
'   AutofocusZOffset is the offset added after AF
''''''
Private Sub AutofocusZOffset_Change()
    'make range checks
     If Abs(AutofocusZOffset.Value) > Range() * 0.9 Then
            AutofocusZOffset.Value = 0
            MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um"
    End If
End Sub

''''''
'   AutofocusZRange_Change()
'   Set the range in um during AF
''''''
Private Sub AutofocusZRange_Change()    ' It should be possible to change the limit of the range to bigger values than half of the working distance
    If AutofocusZRange.Value > Range * 0.9 Then 'make range checks
            AutofocusZRange.Value = Range * 0.9
            MsgBox "ZRange has to be less or equal to the working distance of the objective: " + CStr(Range) + " um"
    End If
End Sub


'''''
'   CheckZRanges()
'   Check if Z movements are in agreement with range of microscope
'''''
Public Function CheckZRanges() As Boolean
    If ScanStop Then
        Exit Function
    End If
    If Range() = 0 Then
        MsgBox "Objective's working distance not defined! Cannot Autofocus!"
        CheckZRanges = False
        Exit Function
    Else
        CheckZRanges = True
    End If
    
    If AutofocusZRange.Value > Range() * 0.9 Then 'this is already tested in the slider could be removed
        AutofocusForm.AutofocusZRange.Value = Range() * 0.9
        MsgBox "Autofocus range is too large! Has been reduced to " + Str(AutofocusForm.AutofocusZRange.Value)
    End If
    
    If Abs(AutofocusZOffset.Value) > Range() * 0.9 Then 'this is already tested in the slider could be removed
        AutofocusForm.AutofocusZOffset = 0
        MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um. Has been put back to " + Str(AutofocusForm.AutofocusZOffset)
    End If
    
End Function
  
''''''
'   The tracks for Autofocus
''''''
Private Sub AutofocusTrack1_Click()
    If AutofocusTrack1.Value Then 'if track 1 checked others are not autofocus track but false
        AutofocusTrack2.Value = Not AutofocusTrack1.Value
        AutofocusTrack3.Value = Not AutofocusTrack1.Value
        AutofocusTrack4.Value = Not AutofocusTrack1.Value
        CheckAutofocusTrack (1) 'sets SelectedTrack to 1, see below
    End If
End Sub

Private Sub AutofocusTrack2_Click()
    If AutofocusTrack2.Value Then
        AutofocusTrack1.Value = Not AutofocusTrack2.Value
        AutofocusTrack3.Value = Not AutofocusTrack2.Value
        AutofocusTrack4.Value = Not AutofocusTrack2.Value
        CheckAutofocusTrack (2)
    End If
End Sub

Private Sub AutofocusTrack3_Click()
    If AutofocusTrack3.Value Then
        AutofocusTrack1.Value = Not AutofocusTrack3.Value
        AutofocusTrack2.Value = Not AutofocusTrack3.Value
        AutofocusTrack4.Value = Not AutofocusTrack3.Value
        CheckAutofocusTrack (3)
    End If
End Sub

Private Sub AutofocusTrack4_Click()
    If AutofocusTrack4.Value Then
        AutofocusTrack1.Value = Not AutofocusTrack4.Value
        AutofocusTrack2.Value = Not AutofocusTrack4.Value
        AutofocusTrack3.Value = Not AutofocusTrack4.Value
        CheckAutofocusTrack (4)
    End If
End Sub


''''''
'   ActiveMicropilot_Click()
'   Activate online image analysis, micropilot. Also enable the complete micropilot page
''''''
Private Sub ActiveMicropilot_Click()

    SwitchEnableMicropilotPage (ActiveMicropilot.Value)
    If ActiveMicropilot.Value And ZoomImageInitialize Then
        MicropilotZOffset.Value = AutofocusZOffset.Value
        MicropilotZSlices.Value = GlobalAcquisitionRecording.FramesPerStack
        MicropilotFrameSize.Value = GlobalAcquisitionRecording.SamplesPerLine
        MicropilotZoom.Value = GlobalAcquisitionRecording.ZoomX
        MicropilotZStep.Value = GlobalAcquisitionRecording.FrameSpacing
        ZoomImageInitialize = False
    End If
        
End Sub

''''''
'   SwitchEnableMicropilotPage(Enable As Boolean)
'   Disable or enable all buttons and slider (aka Micropilot)
'       [Enable] In -  Sets the mini page enable status
''''''
Private Sub SwitchEnableMicropilotPage(Enable As Boolean)
    MicropilotTrack1.Enabled = Enable
    MicropilotTrack2.Enabled = Enable
    MicropilotTrack3.Enabled = Enable
    MicropilotTrack4.Enabled = Enable
    LabelZoom.Enabled = Enable
    MicropilotZoom.Enabled = Enable
    ZoomNumSlicesLabel.Enabled = Enable
    MicropilotZSlices.Enabled = Enable
    ZoomIntervalLabel.Enabled = Enable
    MicropilotZStep.Enabled = Enable
    MicropilotZOffset.Enabled = Enable
    MicropilotAutofocus.Enabled = Enable
    ZoomFrameSizeLabel.Enabled = Enable
    MicropilotFrameSize.Enabled = Enable
    ZoomCyclesLabel.Enabled = Enable
    MicropilotRepetitions.Enabled = Enable
    ZoomCycleDelayLabel.Enabled = Enable
    MicropilotRepetitionTime.Enabled = Enable
    SwitchEnableGridScanPage (ActiveGridScan.Value)
    MicropilotZOffset.Enabled = Enable
    ZoomAutofocusZOffsetLabel.Enabled = Enable
    MicropilotZOffset.Value = AutofocusZOffset.Value
    MicropilotMaxPositions.Enabled = Enable
    MicropilotMaxPositionsLabel1.Enabled = Enable
    MicroPilotMaxPositionsLabel2.Enabled = Enable
    
End Sub

''''''
'   MicropilotAutofocus_Click()
'   Activate extra autofocus for image analysis. Enable Z-offset box to be viewed
''''''
Private Sub MicropilotAutofocus_Click()
    
    SwitchEnableZoomAutofocus (MicropilotAutofocus.Value) 'Show Zoffset only when extra autofocus is clicked
    
End Sub

''''''
'   SwitchEnableZoomAutofocus(Enable As Boolean)
'   Enable/disable Z-offset form for Micropilot minipage
'       [Enable] In - Sets the visibility of box
''''''
Private Sub SwitchEnableZoomAutofocus(Enable As Boolean)
'    MicropilotZOffset.Visible = Enable
'    ZoomAutofocusZOffsetLabel.Visible = Enable
'    MicropilotZOffset.Value = AutofocusZOffset.Value
End Sub

''''''
'   ActiveAlterImage_Click()
'   Activate additional image that is acquired only from time to time
''''''
Private Sub ActiveAlterImage_Click()

    SwitchEnableAlterImagePage (ActiveAlterImage.Value)
    If ActiveAlterImage.Value And AlterImageInitialize Then
        AlterFrameSize.Value = GlobalAcquisitionRecording.SamplesPerLine
        AlterZoom.Value = GlobalAcquisitionRecording.ZoomX
        AlterZOffset.Value = AutofocusZOffset.Value
        AlterZStep.Value = GlobalAcquisitionRecording.FrameSpacing
        AlterZSlices.Value = GlobalAcquisitionRecording.FramesPerStack
        AlterImageInitialize = False
    End If
End Sub

''''''
'   SwitschEnableAlterImagePage(Enable As Boolean)
'   Enable/disable Additional acquisition page
'       [Enable] In - Sets the enable Enable of minpage
''''''
Private Sub SwitchEnableAlterImagePage(Enable As Boolean)

    AlterTrack1.Enabled = Enable
    AlterTrack2.Enabled = Enable
    AlterTrack3.Enabled = Enable
    AlterTrack4.Enabled = Enable
    AlterFrameSizeLabel.Enabled = Enable
    AlterFrameSize.Enabled = Enable
    AlterZoomLabel.Enabled = Enable
    AlterZoom.Enabled = Enable
    AlterNumSlicesLabel.Enabled = Enable
    AlterZSlices.Enabled = Enable
    AlterIntervalLabel.Enabled = Enable
    AlterZStep.Enabled = Enable
    RoundAlterTrackLabel1.Enabled = Enable
    RoundAlterTrackLabel2.Enabled = Enable
    RoundAlterTrack.Enabled = Enable
    RoundAlterLocationLabel1.Enabled = Enable
    RoundAlterLocationLabel2.Enabled = Enable
    AlterZOffset.Enabled = Enable
    AlterZOffsetLabel.Enabled = Enable
    
End Sub

Private Sub ActiveBleach_Click()
      SwitchEnableBleachPage (ActiveBleach.Value)
End Sub



Private Sub SwitchEnableBleachPage(Enable As Boolean)
    BleachTrack1.Enabled = Enable
    BleachTrack2.Enabled = Enable
    BleachTrack3.Enabled = Enable
    BleachTrack4.Enabled = Enable
    BleachFrameSizeLabel.Enabled = Enable
    BleachFrameSize.Enabled = Enable
    BleachZoomLabel.Enabled = Enable
    BleachZoom.Enabled = Enable
    BleachZSlicesLabel.Enabled = Enable
    BleachZSlices.Enabled = Enable
    BleachZStepLabel.Enabled = Enable
    BleachZStep.Enabled = Enable
    BleachZOffset.Enabled = Enable
    BleachZOffsetLabel.Enabled = Enable
    BleachRepetitionsLabel.Enabled = Enable
    BleachRepetitions.Enabled = Enable
    BleachRepetitionTimeLabel.Enabled = Enable
    BleachRepetitionTime.Enabled = Enable

End Sub

''''
' ActiveGridScan_Click()
'   Set the grid scan on or off. Changes also
''
Private Sub ActiveGridScan_Click()
    SwitchEnableGridScanPage (ActiveGridScan.Value)
    If ActiveGridScan.Value Then
        MultipleLocationToggle.Value = False
    End If
End Sub

''''
'   SwitchEnableGridScanPage(Enable As Boolean)
'   Disable or enable all buttons and slider
'       [Enable] In - Sets the mini page enable status
''''
Private Sub SwitchEnableGridScanPage(Enable As Boolean)

    GridScan_validGridDefault.Enabled = Enable
    GridScan_posLabel.Enabled = Enable
    GridScan_nColumnLabel.Enabled = Enable
    GridScan_nRowLabel.Enabled = Enable
    GridScan_nColumn.Enabled = Enable
    GridScan_nRow.Enabled = Enable
    GridScan_dColumnLabel.Enabled = Enable
    GridScan_dRowLabel.Enabled = Enable
    GridScan_dColumn.Enabled = Enable
    GridScan_dRow.Enabled = Enable
    GridScan_refColumn.Enabled = Enable
    GridScan_refRow.Enabled = Enable
    GridScan_refColumnLabel.Enabled = Enable
    GridScan_refRowLabel.Enabled = Enable
    GridScan_subLabel.Enabled = Enable
    GridScan_nColumnsub.Enabled = Enable
    GridScan_nRowsub.Enabled = Enable
    GridScan_nColumnsubLabel.Enabled = Enable
    GridScan_nRowsubLabel.Enabled = Enable
    
    GridScan_dColumnsubLabel.Enabled = Enable
    GridScan_dRowsubLabel.Enabled = Enable
    GridScan_dColumnsub.Enabled = Enable
    GridScan_dRowsub.Enabled = Enable
    GridScanDescriptionLabel.Enabled = Enable
    
End Sub





''''''''
'   CommandButtonHelp_Click()
'   Look for Help file
'   TODO: Test
''''''''
Private Sub CommandButtonHelp_Click()

    Dim dblTask As Double
    Dim MacroPath As String
    Dim MyPath As String
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
            MyPath = Strings.Left(MacroPath, Start - 1)
            MyPathPDF = MyPath + HelpNamePDF

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
    If Not Running Then
        StopButton.Value = False
        StopButton.BackColor = &H8000000F
        ScanStop = False
    Else
        If StopButton.Value Then
            StopButton.BackColor = 12648447
            ScanStop = True
        Else
            StopButton.BackColor = &H8000000F
            ScanStop = False
        End If
    End If

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
        DoEvents
    Else
        DisplayProgress "Restore Settings", RGB(&HC0, &HC0, 0)
        RestoreAcquisitionParameters
        DoEvents
    End If
    
    ReDim BleachTable(BlockRepetitions)
    ReDim BleachStartTable(BlockRepetitions)
    ReDim BleachStopTable(BlockRepetitions)
    ChangeButtonStatus True
    Running = False
    ScanStop = False
    ScanPause = False
    PauseButton.Value = False
    PauseButton.Caption = "PAUSE"
    PauseButton.BackColor = &H8000000F
    ExtraBleach = False
    ExtraBleachButton.Caption = "Bleach"
    ExtraBleachButton.BackColor = &H8000000F
    StopAfterRepetition.Value = False
    StopAfterRepetition.BackColor = &H8000000F
    StopButton.BackColor = &H8000000F
    StopButton.Value = False
    BleachingActivated = False
    LocationTextLabel.Caption = ""
    Sleep (1000)
    If Log Then
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        If Not LogFile Is Nothing Then
            LogFile.Close
        End If
    End If
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)

End Sub

'''''
'   CommandButtonNewDataBase_Click()
'   Open a Dialog to set output folder where to save the results. then cal SetDatabase to set global variables
'''''
Private Sub CommandButtonNewDataBase_Click()
    Dim Filter As String, fileName As String
    Dim Flags As Long
  
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
            
    Filter = "Alle Dateien (*.*)" & Chr$(0) & "*.*"
    
    fileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", "", "Select output folder")
    
    If Len(fileName) > 3 Then
        fileName = Left(fileName, Len(fileName) - 3)
        DatabaseTextbox.Value = fileName
        SetDatabase
    End If
    
End Sub

'''''
'   DatabaseTextbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'   Only update the outputfolder when enter is pressed. This avoids creating a folded at every keystroke
'''''
Private Sub DatabaseTextbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'this is the enter key
        SetDatabase
    End If
End Sub

'''''
'   SetDatabase()
'       [GlobalDataBaseName] Out/Global - The name of Outputfolder
'       [LogFileNameBase]    Out/Global - The name of the LogfileName
'       [Log]                Out/Global - If yes results are logged
'       Set global variables and check if we can create Outputfolder
'''''
Private Sub SetDatabase()
    GlobalDataBaseName = DatabaseTextbox.Value
    If GlobalDataBaseName = "" Then
        DatabaseLable.Caption = "No output folder"
    End If

    If Not GlobalDataBaseName = "" Then
        On Error GoTo ErrorHandleDataBase
        If Not CheckDir(GlobalDataBaseName) Then
            Exit Sub
        End If
        DatabaseLable.Caption = GlobalDataBaseName
        SaveSetting "OnlineImageAnalysis", "macro", "OutputFolder", GlobalDataBaseName
        LogFileNameBase = GlobalDataBaseName & "\AutofocusScreen.log"
        If Right(GlobalDataBaseName, 1) = "\" Then
            BackSlash = ""
        Else
            BackSlash = "\"
        End If
    End If

    If LogCode And LogFileNameBase <> "" Then
        On Error GoTo ErrorHandleLogFile
        'Set FileSystem = New FileSystemObject
        LogFileName = LogFileNameBase
        'SafeOpenTextFile LogFileName, LogFile, FileSystem
        'LogFile.Close
        Log = True
    Else
        Log = False
    End If
    Exit Sub
ErrorHandleDataBase:
    MsgBox "Could not create output Directory " & GlobalDataBaseName
    Exit Sub
ErrorHandleLogFile:
    MsgBox "Could not create LogFile " & LogFileName
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
    Dim pos As Double
    Dim Time As Double
    Dim LogMsg As String
    Dim SuccessRecenter As Boolean
    
    Time = Timer
    Lsm5.DsRecording.Copy GlobalBackupRecording
    Lsm5.DsRecording.FrameSpacing = GlobalBackupRecording.FrameSpacing
    Lsm5.DsRecording.FramesPerStack = GlobalBackupRecording.FramesPerStack
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire = GlobalBackupActiveTracks(i)
    Next i
    Time = Round(Timer - Time, 2)
    LogMsg = "% Restore settings: time to return to backuprecording " & Time
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    
    Sleep (1000)
 
    Time = Timer
    Recenter_pre posTempZ, SuccessRecenter, ZEN
    pos = Lsm5.Hardware.CpFocus.Position
    'move to posTempZ
    If ZEN = "2011" Or ZEN = "2010" Then
        If Round(pos, PrecZ) <> Round(posTempZ, PrecZ) Then
            If Not FailSafeMoveStageZ(posTempZ) Then
                Exit Sub
            End If
        End If
        Recenter_post posTempZ, SuccessRecenter, ZEN
    End If
    Time = Round(Timer - Time, 2)
    LogMsg = "% Restore settings: recenter Z " & posTempZ & ", Time required " & Time & ", success within rep. " & SuccessRecenter & vbCrLf
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    
    ''' close LogFile
    If Log Then
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        If LogFile Is Nothing Then
            Exit Sub
        Else
            LogFile.Close
        End If
    End If
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
    FrameSizeLabel.Visible = ScanLineToggle.Value   'FrameSize Label is only displayed if ScanFrame is activated
    AutofocusFrameSize.Visible = ScanFrameToggle.Value 'FrameSize Slider is only displayed if ScanFrame is activated
    AutofocusTrackXY.Visible = ScanFrameToggle.Value
    AutofocusLineSize.Visible = ScanLineToggle.Value 'LineSize is only displayed if ScanFrame is activated
    If ScanLineToggle.Value Then
        FrameSizeLabel.Caption = "LineSize (px)"
    End If
    AutofocusFastZline.Enabled = ScanLineToggle And Not AutofocusHRZ
End Sub

'''
'   ScanFrameToggle_Click()
'   Switch setting for FrameAutofocus on and LineAutofocus off
'''
Private Sub ScanFrameToggle_Click()
    ScanLineToggle.Value = Not ScanFrameToggle.Value 'if ScanLine is chosen, ScanFrame will be unchecked
    FrameSizeLabel.Visible = ScanFrameToggle.Value
    AutofocusFrameSize.Visible = ScanFrameToggle.Value
    AutofocusTrackXY.Visible = ScanFrameToggle.Value
    If ScanFrameToggle.Value Then
        FrameSizeLabel.Caption = "FrameSize (px)"
    End If
    AutofocusFastZline.Enabled = Not ScanFrameToggle.Value And Not AutofocusHRZ
End Sub


''''''
'   GetCurrentPositionOffsetButton_Click()
'       Performs Autofocus and update ZOffset according to ZShift
''''''
Private Sub GetCurrentPositionOffsetButton_Click()
    If Not GetCurrentPositionOffsetButtonRun Then
        ScanStop = True
        StopAcquisition
    Else
        StopAcquisition
    End If
End Sub

Private Function GetCurrentPositionOffsetButtonRun() As Boolean
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim DeltaZ As Double
    Dim GridPos As GridPosType
    Dim Time As Double
    Dim pos As Double
    Dim LogMsg As String
    Dim SuccessRecenter As Boolean
    Running = True
    Dim NewPicture As DsRecordingDoc
    DisplayProgress "Get Current Position Offset - Autofocus", RGB(0, &HC0, 0)             'Gives information to the user
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Z = posTempZ
    X = Lsm5.Hardware.CpStages.PositionX
    Y = Lsm5.Hardware.CpStages.PositionY

    'recenter only after activation of new track
    If ActiveAutofocus Then
        StopScanCheck
        If AutofocusHRZ Then
            Lsm5.Hardware.CpHrz.Leveling
        End If
       'FailSafeMoveStageZ (posTempZ) 'move at position
        ' Acquire image and calculate center of mass stored in XMass, YMass and ZMass
        DisplayProgress "Autofocus Activate Tracks", RGB(0, &HC0, 0)
        Time = Timer
        If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
            MsgBox "No track selected for Autofocus! Cannot Autofocus!"
            Exit Function
        End If
        
        LogMessage "% Get current position: time activate AF track " & Round(Timer - Time), Log, LogFileName, LogFile, FileSystem
        
        'DoEvents
        'Sample0Z = Lsm5.DsRecording.Sample0Z
        DisplayProgress "Autofocus: Recenter prior AF acquisition.... ", RGB(0, &HC0, 0)
        DoEvents
        Time = Timer
        If Not Recenter_pre(posTempZ, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        pos = Lsm5.Hardware.CpFocus.Position
        Time = Round(Timer - Time, 2)
        LogMsg = "% Get current position: center Z (pre AFimg) " & posTempZ & ", time required" & Time & ", succes within rep. " & SuccessRecenter
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        'Use internal agorithm to compute Xmass etc.
        If Not MicroscopeIO.Autofocus_StackShift(NewPicture) Then
                Exit Function
        End If
        
        DisplayProgress "Autofocus compute", RGB(0, &HC0, 0)
        
        If Not ComputeNewCoordinatesAfterAF(NewPicture, X, Y, Z, DeltaZ, "Autofocus") Then
            Exit Function
        End If
        AutofocusZOffset.Value = posTempZ - Z
                    
        DisplayProgress "Autofocus: Recenter after AF acquisition...", RGB(0, &HC0, 0)
        
        Time = Timer
        If Not Recenter_post(posTempZ, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        Time = Round(Timer - Time, 2)
        LogMsg = "% Get current position: recenter Z (post AFImg) " & posTempZ
        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
                LogMsg = LogMsg & "; obtained central slide " & pos & "; position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        Else
            LogMsg = LogMsg & "; obtained central slide " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 _
            - Lsm5.DsRecording.Sample0Z + pos & "; position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        End If
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        posTempZ = Z
        Time = Timer
        If Not Recenter_pre(posTempZ, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        Time = Round(Timer - Time, 2)
        LogMsg = "% Get current position: center Z (end) " & posTempZ & ", time required" & Time & ", success" & SuccessRecenter
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    End If
    GetCurrentPositionOffsetButtonRun = True
End Function

'''''''
'   AutofocusButton_Click()
'   calls AutofocusButtonRun
''''''''
Public Sub AutofocusButton_Click()
    Dim RecordingDoc As DsRecordingDoc
    Dim SuccessRecenter As Boolean
    Running = True
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Recenter_pre posTempZ, SuccessRecenter, ZEN
    Set GlobalAutoFocusRecording = Lsm5.CreateBackupRecording
    Set GlobalAcquisitionRecording = Lsm5.CreateBackupRecording
    Set GlobalMicropilotRecording = Lsm5.CreateBackupRecording
    Set GlobalBleachRecording = Lsm5.CreateBackupRecording
    Set GlobalAltRecording = Lsm5.CreateBackupRecording
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    GlobalAutoFocusRecording.Copy Lsm5.DsRecording
    GlobalAcquisitionRecording.Copy Lsm5.DsRecording
    GlobalMicropilotRecording.Copy Lsm5.DsRecording
    GlobalBleachRecording.Copy Lsm5.DsRecording
    GlobalAltRecording.Copy Lsm5.DsRecording
    GlobalBackupRecording.Copy Lsm5.DsRecording ' this will not be changed remains always the same
    GlobalBackupSampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    Dim i As Long
    Dim NrTracks As Long
    ReDim GlobalBackupActiveTracks(Lsm5.DsRecording.TrackCount)
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       GlobalBackupActiveTracks(i) = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i
    AutofocusButtonRun RecordingDoc, GlobalDataBaseName
    StopAcquisition
End Sub

'''''''
'   AutofocusButtonRun (Optional AutofocusDoc As DsRecordingDoc = Nothing) As Boolean
'   Runs a Z-stacks, compute center of mass, if selected acquire an image at computed position + ZOffset
'   If AutofocusTrackZ : position is updated to computed position from autofocus (without ZOffset!)
'   If AutofocusTrackXY and FrameToggle: position of X and Y are changed
'   Function uses a posTempZ to remember starting position
'       [AutofocusDoc] - A recording Doc. If = Nothing then it will create a new recording
'
'   Additional comments: The function works best with piezo. With Fast-Zline (Onthefly) acquisition is less precise
'                        Lots of test to check that focus returned to workingposition. Lsm5.Hardware.CpFocus.Position
'                        does not give actual position when stage is moving after acquisition.
'                        Lsm5.DsRecording.Sample0Z provides the actual shift to the central slice
''''''''
Private Function AutofocusButtonRun(Optional AutofocusDoc As DsRecordingDoc = Nothing, Optional FilePath As String = "") As Boolean
    Running = True
    Dim Time As Double
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim DeltaZ As Double
    Dim GridPos As GridPosType
    Dim Sample0Z As Double ' test variable
    Dim pos As Double ' test variable for position
    Dim LogMsg  As String
    Dim SuccessRecenter As Boolean
    DisplayProgress "Autofocus move initial position", RGB(0, &HC0, 0)
    
    StopScanCheck
    ' Recenter and move where it should be


    Z = posTempZ
    X = Lsm5.Hardware.CpStages.PositionX
    Y = Lsm5.Hardware.CpStages.PositionY

    'recenter only after activation of new track
    If ActiveAutofocus Then
    
        If AutofocusAlgorithm.Value = "external" Then
            SaveSetting "OnlineImageAnalysis", "macro", "code", "wait"     'this causes to wait as defaults
            If FilePath = "" Then
                MsgBox "Define an Outputfolder to save AF image for external image analysis!"
                Exit Function
            End If
            If Not SaveAFImage Then
                MsgBox "Tick Save AFImage to use external image analysis for focus!"
                Exit Function
            End If
        End If
        
        If AutofocusHRZ Then
            Lsm5.Hardware.CpHrz.Leveling
        End If
        
        ' Acquire image and calculate center of mass stored in XMass, YMass and ZMass
        DisplayProgress "Autofocus Activate Tracks", RGB(0, &HC0, 0)
        Time = Timer
        If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
            MsgBox "No track selected for Autofocus! Cannot Autofocus!"
            Exit Function
        End If
        
        Time = Round(Timer - Time, 2)
        LogMsg = "% AutofocusButton: time activate AF tracks " & Time
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        '''''center
        DisplayProgress "Autofocus: Recenter prior AF acquisition.... ", RGB(0, &HC0, 0)
        DoEvents
        Sleep (200)
        Time = Timer
        If Not Recenter_pre(posTempZ, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        
        Time = Round(Timer - Time, 2)
        LogMsg = "% AutofocusButton: center Z (pre AFImg) " & posTempZ
        pos = Lsm5.Hardware.CpFocus.Position
        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
            LogMsg = LogMsg & ", Obtained Z " & pos & "; actual position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        Else
            LogMsg = LogMsg & ", Obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
            & "; actual position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        End If
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        '''''''acquire
        DisplayProgress "Autofocus: Acquire AFimg.... ", RGB(0, &HC0, 0)
        Time = Timer
        'Just acquire image and save
        If Not MicroscopeIO.Autofocus_StackShift(AutofocusDoc) Then
            Exit Function
        End If
        
        DisplayProgress "Autofocus compute", RGB(0, &HC0, 0)
        
        If SaveAFImage Then
            If Not SaveDsRecordingDoc(AutofocusDoc, FilePath & "\" & "AFImg.lsm") Then
                Exit Function
            End If
        End If
        
        If Not ComputeNewCoordinatesAfterAF(AutofocusDoc, X, Y, Z, DeltaZ, "AutoFocus", AutofocusAlgorithm.Value) Then
            Exit Function
        End If
        LogMsg = "% AutofocusButton: computed position XYZ " & X & " " & Y & " " & Z
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        ''''''recenter
        DisplayProgress "Autofocus: Recenter after AF acquisition...", RGB(0, &HC0, 0)
        Time = Timer
        If Not Recenter_post(posTempZ, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        
        pos = Lsm5.Hardware.CpFocus.Position
        LogMsg = "% AutofocusButton: wait return to center Z (post AFImg) " & posTempZ
        Time = Round(Timer - Time, 2)
        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
            LogMsg = LogMsg & ", Obtained Z " & pos & "; actual position " & pos & ", Time required " & Time & ", success within rep. " & SuccessRecenter
        Else
            LogMsg = LogMsg & ", Obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
            & "; actual position " & pos & ", Time required " & Time & ", success within rep. " & SuccessRecenter
        End If
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        'move X and Y if tracking is on
        If ScanFrameToggle And AutofocusTrackXY Then
            If Not FailSafeMoveStageXY(X, Y) Then
                Exit Function
            End If
        End If
        
        If AutofocusHRZ Then
            Lsm5.Hardware.CpHrz.Position = 0
        End If
    End If

    ''''Acquisition
    If ActivateTrack(GlobalAcquisitionRecording, "Acquisition", DeltaZ) Then
        Dim Offset As Double
        If ActiveAutofocus Then
            Offset = AutofocusZOffset
        Else
            Offset = 0
        End If

        DisplayProgress "Autofocus: Center AQimg at ZOffset position...", RGB(0, &HC0, 0)
        '''''''center
        Time = Timer
        If Not Recenter_pre(Z + Offset, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        Time = Round(Timer - Time, 2)
        LogMsg = "% AutofocusButton: center Z + Offset (pre AQimg) " & Z + Offset & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        ''''''''''Acquire
        DisplayProgress "Autofocus: Center AQimg at ZOffset position...", RGB(0, &HC0, 0)
        Time = Timer
        If Not ScanToImage(AutofocusDoc) Then
            Exit Function
        End If
        
        Time = Timer - Time
        LogMsg = "% AutofocusButton: Time acquire AQImg " & Round(Time, 2)
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        ''''''''''recenter
        DisplayProgress "Autofocus: Recenter after AQimg ...", RGB(0, &HC0, 0)
        Time = Timer
        If Not Recenter_post(Z + Offset, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        Time = Round(Timer - Time, 2)
        pos = Lsm5.Hardware.CpFocus.Position
        LogMsg = "% AutofocusButton: wait return to center Z + Offset (post AQImg) " & Z + Offset
        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
            LogMsg = LogMsg & ", Obtained Z " & pos & "; actual position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        Else
            LogMsg = LogMsg & ", Obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
            & "; actual position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
        End If
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    End If

    If AutofocusHRZ Then
        Lsm5.Hardware.CpHrz.Position = 0
    End If

    '''Update position to the position without offset and move there
    If AutofocusTrackZ Then
        'wait that slice recentered after acquisition
        DisplayProgress "Autofocus: move to new Z...", RGB(0, &HC0, 0)
        posTempZ = Z
    Else
        DisplayProgress "Autofocus: return to initial Z...", RGB(0, &HC0, 0)
    End If
    
    Time = Timer
    Recenter_pre posTempZ, SuccessRecenter, ZEN
    pos = Lsm5.Hardware.CpFocus.Position
    ' move stage to posTempZ
    If ZEN = "2011" Or ZEN = "2010" Then
        If Round(pos, PrecZ) <> Round(posTempZ, PrecZ) Then
            If Not FailSafeMoveStageZ(posTempZ) Then
                Exit Function
            End If
        End If
        Recenter_post posTempZ, SuccessRecenter, ZEN
    End If
    Time = Round(Timer - Time, 2)
    pos = Lsm5.Hardware.CpFocus.Position
    LogMsg = "% AutofocusButton: wait return to center Z (end) " & posTempZ
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
        LogMsg = LogMsg & ", Obtained Z " & pos & "; actual position " & pos & ", time required " & Time & ", succes within rep. " & SuccessRecenter
    Else
        LogMsg = LogMsg & ", Obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
        & "; actual position " & pos & ", Time required " & Time & ", succes within rep. " & SuccessRecenter
    End If
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
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
    If MultipleLocationToggle.Value And Lsm5.Hardware.CpStages.MarkCount < 1 Then
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
            
            If ActiveMicropilot Then
                nt = MicropilotRepetitions
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
        
        If ActiveMicropilot Then
            nt = MicropilotRepetitions.Value
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
    Dim i As Integer
    Dim initPos As Boolean   'if False and gridsize correspond positions are taken from file positionsGrid.csv
    Dim initValid As Boolean 'if False and gridsize correspond positions are taken from file validGrid.csv
    Dim SuccessRecenter As Boolean
    
    initPos = True
    initValid = True
    StartSetting = False
    BleachingActivated = False
    AutomaticBleaching = False                                  'We do not do FRAps or FLIPS in this case. Bleaches can still be done with the "ExtraBleach" button.
    Set FileSystem = New FileSystemObject
    ' Do some checking
    If TrackingToggle And TrackingChannelString = "" Then
        MsgBox ("Select a channel for tracking, or uncheck the tracking button")
        Exit Function
    End If
    If MultipleLocationToggle.Value And Lsm5.Hardware.CpStages.MarkCount < 1 Then
        MsgBox ("Select at least one location in the stage control window, or uncheck the multiple location button")
        Exit Function
    End If
    'This loads value of Databasename
    SetDatabase
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start acquisition.")
        Exit Function
    Else
        If Not CheckDir(GlobalDataBaseName) Then
            Exit Function
        End If
        LogFileNameBase = GlobalDataBaseName & "\AutofocusScreen.log"
        If LogCode And LogFileNameBase <> "" Then
            On Error GoTo ErrorHandleLogFile
            LogFileName = LogFileNameBase
            Close
            SafeOpenTextFile LogFileName, LogFile, FileSystem
            LogFile.WriteLine "% ZEN software version " & ZEN & " " & Version
            LogFile.Close
            Log = True
        Else
            Log = False
        End If
    End If

    If Not AcquisitionTracksOn And Not ActiveAutofocus And Not ActiveAlterImage Then
        MsgBox ("Nothing to do! Check at least one imaging option!")
        Exit Function
    End If
    If Not AcquisitionTracksOn Then
        If MsgBox("Acquisition Track has not been clicked!! Do you want to continue", VbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    
    ' do not log if logfilename has not been defined
    If LogCode And LogFileName = "" Then
        Log = False
    End If
    'As default we do not overwrite files
    OverwriteFiles = False
    
    '''''''''''''''''''''''
    '***Set up GridScan***'
    '''''''''''''''''''''''
    If ActiveGridScan Then
        'Load starting position from stage
        Dim MarkCount As Long
        MarkCount = Lsm5.Hardware.CpStages.MarkCount
        If MarkCount = 0 Then  ' No marked position
            MsgBox "GridScan: Use stage to Mark the initial position "
            ScanStop = True
            StopAcquisition
            Exit Function
        End If
        If GridScan_nColumn.Value * GridScan_nRow.Value * GridScan_nColumnsub.Value * GridScan_nRowsub.Value > 10000 Then
            MsgBox "GridScan: Maximal number of locations is 10000. Please change Numbers  X and/or Y."
            ScanStop = True
            StopAcquisition
            Exit Function
        End If
        
        If CheckPosFile(GlobalDataBaseName & "\positionsGrid.csv", GridScan_nRow.Value, GridScan_nColumn.Value, _
            GridScan_nRowsub.Value, GridScan_nColumnsub.Value) Then
            If MsgBox("Position file " & "positionsGrid.csv exists. Do you want to reset positions?", VbYesNo) = vbNo Then
                 If LoadPosFile(GlobalDataBaseName & "\positionsGrid.csv", posGridX, posGridY, posGridZ) Then
                    initPos = False
                    FocusMapPresent = True
                 End If
            End If
        End If
        
        If initPos Then
            DisplayProgress "Initialize all grid positions. First Gridpoint is first Marked point on stage....", RGB(0, &HC0, 0)
            'MsgBox " GridScan: Uses as initial position the first Marked point on stage "
            'Store starting position for later restart. This is the first marked point
            Lsm5.Hardware.CpStages.MarkGetZ 0, XStart, YStart, ZStart
            Sleep (1000)
            ReDim posGridX(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value, 1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value)
            ReDim posGridY(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value, 1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value)
            ReDim posGridZ(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value, 1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value)
            MakeGrid posGridX, posGridY, posGridZ, XStart, YStart, ZStart, GridScan_dColumn.Value, GridScan_dRow.Value, _
            GridScan_dColumnsub.Value, GridScan_dRowsub.Value, GridScan_refColumn.Value, GridScan_refRow.Value
            DisplayProgress "Initialize all grid positions...DONE", RGB(0, &HC0, 0)
            WritePosFile GlobalDataBaseName & "\positionsGrid.csv", posGridX, posGridY, posGridZ
            FocusMapPresent = False
        End If
        
        If CheckPosFile(GlobalDataBaseName & "\validGrid.csv", GridScan_nRow.Value, GridScan_nColumn.Value, _
            GridScan_nRowsub.Value, GridScan_nColumnsub.Value) Then
            If MsgBox("Valid file " & "validGrid.csv exists. Do you want to reset valid positions?", VbYesNo) = vbNo Then
                 If LoadValidFile(GlobalDataBaseName & "\validGrid.csv", posGridXY_Valid) Then
                    initValid = False
                 End If
            End If
        End If
        
        If initValid Then
            ReDim posGridXY_Valid(1 To GridScan_nRow.Value, 1 To GridScan_nColumn.Value, 1 To GridScan_nRowsub.Value, 1 To GridScan_nColumnsub.Value) ' A position may be active or not
            WriteValidFile GlobalDataBaseName & "\validGrid.csv", posGridXY_Valid

            If GridScan_validGridDefault Then
                If isValidGridDefault(GlobalDataBaseName & "\validGridDefault.txt") Then
                    MakeValidGrid posGridXY_Valid, GlobalDataBaseName & "\validGridDefault.txt"
                Else
                    Exit Function
                End If
            Else
               MakeValidGrid posGridXY_Valid
            End If
            WriteValidFile GlobalDataBaseName & "\validGrid.csv", posGridXY_Valid
        End If
    End If
    '''''''''''''''''''''''''''
    '***End Set up GridScan***'
    '''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''
    '***Set up MultiLocationScan***'
    ''''''''''''''''''''''''''''''''
    If MultipleLocationToggle Then
    
        If FileExist(GlobalDataBaseName & "\" & "PositionsMultiLoc.txt") Then
            MsgBox ("File Exist")
        End If
        
        If Lsm5.Hardware.CpStages.MarkCount > 0 Then
            ReDim posGridX(1 To 1, 1 To Lsm5.Hardware.CpStages.MarkCount, 1 To 1, 1 To 1)
            ReDim posGridY(1 To 1, 1 To Lsm5.Hardware.CpStages.MarkCount, 1 To 1, 1 To 1)
            ReDim posGridZ(1 To 1, 1 To Lsm5.Hardware.CpStages.MarkCount, 1 To 1, 1 To 1)
            ReDim posGridXY_Valid(1 To 1, 1 To Lsm5.Hardware.CpStages.MarkCount, 1 To 1, 1 To 1) ' A well may be active or not
            For i = 1 To Lsm5.Hardware.CpStages.MarkCount
                Lsm5.Hardware.CpStages.MarkGetZ i - 1, posGridX(1, i, 1, 1), posGridY(1, i, 1, 1), _
                posGridZ(1, i, 1, 1)
                posGridXY_Valid(1, i, 1, 1) = True
            Next i
        End If
    End If
    
  
    If SingleLocationToggle And Not ActiveGridScan Then
            ReDim posGridX(1 To 1, 1 To 1, 1 To 1, 1 To 1)
            ReDim posGridY(1 To 1, 1 To 1, 1 To 1, 1 To 1)
            ReDim posGridZ(1 To 1, 1 To 1, 1 To 1, 1 To 1)
            ReDim posGridXY_Valid(1 To 1, 1 To 1, 1 To 1, 1 To 1) 'A well may be active or not
            Lsm5.Hardware.CpStages.GetXYPosition posGridX(1, 1, 1, 1), posGridY(1, 1, 1, 1)
            posGridZ(1, 1, 1, 1) = Lsm5.Hardware.CpFocus.Position
            posGridXY_Valid(1, 1, 1, 1) = 1
    End If
    
    ''''
    'load acquisition settings again
    '''''
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Recenter (posTempZ)
    
    Set GlobalAutoFocusRecording = Lsm5.CreateBackupRecording
    Set GlobalAcquisitionRecording = Lsm5.CreateBackupRecording
    Set GlobalMicropilotRecording = Lsm5.CreateBackupRecording
    Set GlobalBleachRecording = Lsm5.CreateBackupRecording
    Set GlobalAltRecording = Lsm5.CreateBackupRecording
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    GlobalAutoFocusRecording.Copy Lsm5.DsRecording
    GlobalAcquisitionRecording.Copy Lsm5.DsRecording
    GlobalMicropilotRecording.Copy Lsm5.DsRecording
    GlobalBleachRecording.Copy Lsm5.DsRecording
    GlobalAltRecording.Copy Lsm5.DsRecording
    GlobalBackupRecording.Copy Lsm5.DsRecording ' this will not be changed remains always the same
    GlobalBackupSampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime
    Dim NrTracks As Long
    ReDim GlobalBackupActiveTracks(Lsm5.DsRecording.TrackCount)
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       GlobalBackupActiveTracks(i) = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire
    Next i
    If Not Recenter_post(posTempZ, SuccessRecenter, ZEN) Then
        Exit Function
    End If
    
    'SaveSettings
    If GlobalDataBaseName <> "" Then
        SetDatabase
        SaveSettings GlobalDataBaseName & "\AutofocusScreen.ini"
    End If
    StartSetting = True
    Exit Function
ErrorHandleDataBase:
    MsgBox "Could not create directory " & GlobalDataBaseName
    Exit Function
ErrorHandleLogFile:
    MsgBox "Could not create LogFile " & LogFileName
    Exit Function
End Function



''''''
'   StartAcquisition(BleachingActivated)
'   Perform many things (TODO: write more). Pretty much the whole macro runs through here
''''''
Private Sub StartAcquisition(BleachingActivated As Boolean)
    
    'measure time required
    Dim rettime, DiffTime As Double
    Dim GlobalPrvTime As Double
    Dim StartTime As Double
    
    'Counters
    Dim GridPos As GridPosType
    Dim RowMax As Long
    Dim ColMax  As Long
    Dim RowSubMax As Long
    Dim ColSubMax As Long
    Dim StartCol As Long
    Dim StartColSub As Long
    Dim EndCol As Long
    Dim EndColSub As Long
    Dim StepCol As Integer    'forward or backward step
    Dim StepColSub As Integer 'forward or backward step
    Dim Cnt As Integer        'a local counter
    Dim TotPos As Long        'total number of positions
    
    Dim SuccessRecenter As Boolean
    'coordinates
    Dim previousZ As Double   'remember position of previous position in Z
    
    HighResExperimentCounter = 1
    HighResCounter = 0
    
    ' ActiveMicropilot  refers to the MicroPilot
    Erase HighResArrayX()
    Erase HighResArrayY()
    Erase HighResArrayZ()
    SaveSetting "OnlineImageAnalysis", "macro", "code", "2"     'this causes to do anything as defaults
    SaveSetting "OnlineImageAnalysis", "macro", "offsetx", 0
    SaveSetting "OnlineImageAnalysis", "macro", "offsety", 0
    SaveSetting "OnlineImageAnalysis", "macro", "offsetz", 0
    SaveSetting "OnlineImageAnalysis", "macro", "deltaZ", ""

    'Coordinates
    Dim X As Double              ' x value where to move the stage (this is used as reference)
    Dim Y As Double              ' y value where to move the stage
    Dim Z As Double              ' z value where to move the stage
    Dim Xold As Double
    Dim Yold As Double
    Dim Zold As Double
    
    'test variables
    Dim Success As Integer       ' Check if something was sucessfull
    Dim SuccessAF As Boolean     ' Check if AF was succesful
    Dim LocationSoFarBest As Integer
    Dim soFarBestGoodCellsPerImage As Integer
    
    'Recording stuff
    Dim FileNameID As String ' ID name of file (Well/Position, Subpositio, Timepoint)
    Dim FilePath As String   ' full path of file to save (changes through function)
    Dim RecordingDoc As DsRecordingDoc  ' contains the images
    Dim FcsData As AimFcsData           ' contains the data of the FCS measurement
    Dim Scancontroller As AimScanController ' the controller
  
        'do once leveling
    If AutofocusHRZ Then
        Lsm5.Hardware.CpHrz.Leveling ' not sure if this is needed
        Sleep (1000)
    End If
    
    
    ' Set the offset in z-stack to 0; otherwise there can be errors...
    Recenter (Lsm5.Hardware.CpFocus.Position)
    ' set up the imaging
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    NewRecord RecordingDoc, "W001_P001_T001", 0
'    Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
'    ' set up RecordingDoc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (20)
'            DoEvents
'        Wend
'    End If

    InitializeStageProperties
    SetStageSpeed 9, True

    Running = True  'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.
    ChangeButtonStatus False ' disable buttons
    TotPos = 1
    previousZ = posGridZ(1, 1, 1, 1)
    Do While Running   'As long as the macro is running we're in this loop. At everystop one will save actual location, and repetition
                
        RowMax = UBound(posGridX, 1)
        ColMax = UBound(posGridX, 2)
        
        RowSubMax = UBound(posGridX, 3)
        ColSubMax = UBound(posGridX, 4)
        
        GlobalPrvTime = CDbl(GetTickCount) * 0.001
        For GridPos.Row = 1 To RowMax
            'Meander
            If GridPos.Row Mod 2 = 0 Then
                StartCol = ColMax
                EndCol = 1
                StepCol = -1
            Else
                StartCol = 1
                EndCol = ColMax
                StepCol = 1
            End If
            GridPos.Col = StartCol
            For GridPos.Col = StartCol To EndCol Step StepCol
                ' Move in the subGrid
                For GridPos.RowSub = 1 To RowSubMax
                    'Meander in the subgrid
                    If GridPos.RowSub Mod 2 = 0 Then
                        StartColSub = ColSubMax
                        EndColSub = 1
                        StepColSub = -1
                    Else
                        StartColSub = 1
                        EndColSub = ColSubMax
                        StepColSub = 1
                    End If
                    GridPos.ColSub = StartColSub
                    For GridPos.ColSub = StartColSub To EndColSub Step StepColSub
                        ' Here comes the check for good or bad location ...
                        If posGridXY_Valid(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) Then
                            'define actual positions and move there
                            X = posGridX(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
                            Y = posGridY(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
                            Z = posGridZ(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
                            'In gridscan mode use initially Z of previous position to find new position
                            If RepetitionNumber = 1 And ActiveGridScan And Not FocusMapPresent Then
                                Z = previousZ
                            Else
                                Z = posGridZ(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
                            End If
                            'move in X and Y
                            Xold = Lsm5.Hardware.CpStages.PositionX
                            Yold = Lsm5.Hardware.CpStages.PositionY
                            If Round(Xold, PrecXY) <> Round(X, PrecXY) Or Round(Yold, PrecXY) <> Round(Y, PrecXY) Then
                                If Not FailSafeMoveStageXY(X, Y) Then
                                    StopAcquisition
                                    Exit Sub
                                End If
                            End If
                            
                            Recenter_pre Z, SuccessRecenter, ZEN
                            If Round(Lsm5.Hardware.CpFocus.Position, PrecZ) <> Round(Z, PrecZ) Then 'Need to move now! May cause problems!
                                If Not FailSafeMoveStageZ(Z) Then
                                    StopAcquisition
                                    Exit Sub
                                End If
                            End If
                            Recenter_post Z, SuccessRecenter, ZEN
                            DoEvents
                        Else ' jump to next location
                            GoTo NextLocation
                        End If
                        
                        ' Show position of stage
                        LocationTextLabel.Caption = SetLocationTextLabel(X, Y, Z, GridPos, TotPos, RepetitionNumber)
                        
                        
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
                        
                        If Not ImagingWorkFlow(RecordingDoc, FcsData, StartTime, GridPos, TotPos, RepetitionNumber) Then
                            StopAcquisition
                            Exit Sub
                        End If
        
NextLocation:
                        TotPos = TotPos + 1
                        If ScanPause Then
                            If Not Pause Then ' Pause is true is Resume
                                ScanStop = True
                                StopAcquisition
                                Exit Sub
                            End If
                        End If
                        If ScanStop Then
                            StopAcquisition
                            Exit Sub
                        End If
                        previousZ = posGridZ(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
                   Next GridPos.ColSub
            Next GridPos.RowSub
        Next GridPos.Col
    Next GridPos.Row
    TotPos = 1
    If StopAfterRepetition.Value Then
        ScanStop = True
        StopAcquisition
        Exit Sub
    End If
    ' DONE WITH THE IMAGING....NOW POSTPROCESSING...
    
    If AutomaticBleaching Then
        FillBleachTable     ' Updating the bleaching table before the next acquisitions, just in case there were changes n the bleaching window
    End If
    
        
    If (RepetitionNumber < LoopingRepetitions.Value) Then
        
        If (LoopingInterval) Then
            ' do nothing => leave GlobalPrvTime as the time that set at the beginng of the position loop
        Else ' delay
            GlobalPrvTime = CDbl(GetTickCount) * 0.001    'Reset the time to NOW
        End If
        
        rettime = CDbl(GetTickCount) * 0.001
        DiffTime = rettime - GlobalPrvTime
        'TODO: Check this
        'This loops define the waiting delay before going back to the first location
        Do While (DiffTime <= CDbl(LoopingRepetitionTime.Value * LoopingTimerUnit)) And Not (BleachTable(RepetitionNumber + 1) = True)
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
            DisplayProgress "Waiting " & CStr(CInt(LoopingRepetitionTime.Value * LoopingTimerUnit - DiffTime)) + " s before scanning repetition  " & (RepetitionNumber + 1), RGB(&HC0, &HC0, 0)
            rettime = CDbl(GetTickCount) * 0.001
            DiffTime = rettime - GlobalPrvTime
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
Private Function ImagingWorkFlow(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, StartTime As Double, GridPos As GridPosType, TotPos As Long, RepetitionNumber As Integer) As Boolean
    
    ImagingWorkFlow = False
    Dim Xnew As Double
    Dim Ynew As Double
    Dim Znew As Double
    Dim Xold As Double
    Dim Yold As Double
    Dim Zold As Double
    Dim DeltaZ As Double
    DeltaZ = -1
    Dim FileNameID As String
    Dim FilePath As String
    Dim FilePathAF  As String
    Dim Cnt As Integer
    Dim Time As Double
    Dim Offset As Double ' a localyy used Zoffset variable
    Dim pos As Double ' a tmp variable for position
    Dim Sample0Z As Double
    Dim LogMsg As String
    Dim SuccessRecenter As Boolean
    Dim WarningAcq As Boolean
    Xnew = posGridX(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
    Ynew = posGridY(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
    If RepetitionNumber = 1 And ActiveGridScan And Not FocusMapPresent Then
        Znew = Lsm5.Hardware.CpFocus.Position
    Else
        Znew = posGridZ(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub)
    End If
    Xold = Xnew
    Yold = Ynew
    Zold = Znew

    FileNameID = fileName(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub, RepetitionNumber)

    
    FilePath = DatabaseTextbox.Value & BackSlash & TextBoxFileName.Value & UnderScore & FileNameID
    FilePathAF = DatabaseTextbox.Value & BackSlash & TextBoxFileName.Value & UnderScore & "AFImg_" & FileNameID
'
'    If AutofocusHRZ Then
'        Lsm5.Hardware.CpHrz.Leveling
'    End If
    
    LogMsg = vbCrLf & vbCrLf & "% StartButton: Acquire image " & FilePath & vbCrLf _
    & "% StartButton: Imaging position Row " & GridPos.Row & ", Col " & GridPos.Col & ", Row (subpos) " & GridPos.RowSub & ", Col (subpos) " & GridPos.ColSub & vbCrLf _
    & "% StartButton: Current position  XYZ " & Round(Xold, PrecXY) & ", " & Round(Yold, PrecXY) & ", " & Round(Zold, PrecZ)
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    
    
    ' At every positon and repetition  check if Autofocus needs to be required. Update of positions in Z is only done at the end of acquisition
    If ActiveAutofocus And ((RepetitionNumber - 1) Mod AFeveryNth = 0) Then    ' Perform Autofocus if active
        If AutofocusAlgorithm.Value = "external" Then
            SaveSetting "OnlineImageAnalysis", "macro", "code", "wait"     'this causes to do anything as defaults
            If Not SaveAFImage Then
                MsgBox "Tick Save AFImage if you want to use an external image analysis for computing the focus!"
                Exit Function
            End If
        End If
        
        StopScanCheck 'stop any running jobs
        ' Acquire image and calculate center of mass stored in XMass, YMass and ZMass
        DisplayProgress "Autofocus Activate Tracks", RGB(0, &HC0, 0)
        If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
            MsgBox "No track selected for Autofocus! Cannot Autofocus!"
            Exit Function
        End If

        DisplayProgress "Autofocus center Z", RGB(0, &HC0, 0)
        Time = Timer
        pos = Lsm5.Hardware.CpFocus.Position
        If Not Recenter_pre(Zold, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        
        Time = Round(Timer - Time, 2)
        LogMsg = "% StartButton:  center Z (pre AFimg) " & Zold & ", time required " & Time & ", Success " & SuccessRecenter

        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        ' take a z-stack and finds the brightest plane:
        If Not MicroscopeIO.Autofocus_StackShift(RecordingDoc) Then
           Exit Function
        End If
        

        
        If SaveAFImage Then
            If Not SaveDsRecordingDoc(RecordingDoc, FilePathAF & ".lsm") Then
                Exit Function
            End If
        End If
        
        If Not ComputeNewCoordinatesAfterAF(RecordingDoc, Xnew, Ynew, Znew, DeltaZ, "Autofocus", AutofocusAlgorithm.Value) Then
            Exit Function
        End If
        
        LogMsg = "% StartButton:  center of mass XYZ  " & Xnew - Xold & ", " & Ynew - Yold & ", " & Znew - Zold & ". Computed position XYZ " & Xnew & ", " & Ynew & ", " & Znew
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        If AutofocusTrackXY.Value And ScanFrameToggle.Value Then
            DisplayProgress "Autofocus move XY stage", RGB(0, &HC0, 0)
            If Not FailSafeMoveStageXY(Xnew, Ynew) Then
                Exit Function
            End If
            posGridX(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) = Xnew
            posGridY(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) = Ynew
        Else
            Xnew = Xold
            Ynew = Yold
        End If
        
        'wait for recentering
        DisplayProgress "Autofocus: recentering Z after AF", RGB(0, &HC0, 0)
        Time = Timer
        If Not Recenter_post(Zold, SuccessRecenter, ZEN) Then
            Exit Function
        End If
            
        Time = Round(Timer - Time, 2)
        LogMsg = "% StartButton:  wait to return center Z (post AFimg) " & Zold
        pos = Lsm5.Hardware.CpFocus.Position
        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
            LogMsg = LogMsg & ", obtained Z " & pos & ", position " & pos & ", time required " & Time & ", success within rep. " & SuccessRecenter
        Else
             LogMsg = LogMsg & ", obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
             & ", position " & pos & ", time required " & Time & ", success within rep. " & SuccessRecenter
        End If
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        If ScanPause Then
            If Not Pause Then ' Pause is true if Resume
                Exit Function
            End If
        End If
    End If '(RepetitionNumber - 1) Mod AFeveryNth = 0

    FilePath = FilePath & ".lsm"

    LocationTextLabel.Caption = SetLocationTextLabel(Xnew, Ynew, Znew, GridPos, TotPos, RepetitionNumber)

       
    If ScanPause Then
        If Not Pause Then ' Pause is true if Resume
            Exit Function
        End If
    End If

    '''''''''''''''''''''''''''''''''''''
    '*Begin Normal acquisition imaging**'
    '''''''''''''''''''''''''''''''''''''
    DisplayProgress "Acquiring  Location   " & TotPos & "/" & UBound(posGridX, 1) * UBound(posGridX, 2) * UBound(posGridX, 3) * UBound(posGridX, 4) & vbCrLf & _
                    "                   Repetition  " & RepetitionNumber & "/" & LoopingRepetitions.Value, RGB(&HC0, &HC0, 0)
    Time = Timer
    If ActivateTrack(GlobalAcquisitionRecording, "Acquisition", DeltaZ) Then       ' If DeltaZ > 0 then new number of slices is set
        LogMsg = "% Startbutton: Time activate AQ track " & Round(Timer - Time, 2)
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        If ActiveAutofocus Then
            Offset = AutofocusZOffset
        Else
            Offset = 0
        End If
        
        DisplayProgress "Acquisition: recentering Z + Offset ", RGB(0, &HC0, 0)
        'center the slide
        Time = Timer
        'Sleep (200)
        If Not Recenter_pre(Znew + Offset, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        'Sleep (200)
    
        
        Time = Round(Timer - Time, 2)
        LogMsg = "% Startbutton: Center Z + Offset (pre AQimg) " & Znew + Offset & ", time required " & Time & ", repeats " & Round(Time / 0.4)
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        DoEvents
        DisplayProgress "Acquiring  Location   " & TotPos & "/" & UBound(posGridX, 1) * UBound(posGridX, 2) * UBound(posGridX, 3) * UBound(posGridX, 4) & vbCrLf & _
                    "                   Repetition  " & RepetitionNumber & "/" & LoopingRepetitions.Value, RGB(&HC0, &HC0, 0)
        Time = Timer
        If Not ScanToImage(RecordingDoc) Then
            Exit Function
        End If
        LogMsg = "% Startbutton: Time acquire AQ track " & Round(Timer - Time)
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
        
        ''''''''''''''''''''''''''''''''''''
        '*** Save Acquisition Image *******'
        ''''''''''''''''''''''''''''''''''''
        RecordingDoc.SetTitle TextBoxFileName.Value & FileNameID
        'this is the name of the file to be saved
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
        
        Time = Timer
        'wait that slice recentered after acquisition
        If Not Recenter_post(Znew + Offset, SuccessRecenter, ZEN) Then
            Exit Function
        End If
        
        LogMsg = "% StartButton:  recenter Z (post AQImg) " & Znew + Offset
        pos = Lsm5.Hardware.CpFocus.Position
        Time = Round(Timer - Time, 2)
        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
            LogMsg = LogMsg & ", obtained Z " & pos & ", position " & pos & ", time required " & Time & ", success within rep. " & SuccessRecenter
        Else
            LogMsg = LogMsg & ", obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
            & ", position " & pos & ", time required " & Time & ", success within rep. " & SuccessRecenter
        End If
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    End If
    
    If ScanPause Then
        If Not Pause Then ' Pause is true is Resume
            Exit Function
        End If
    End If

    ''''''''''''''''''''''''''''''
    '*Begin Alternative imaging**'
    ''''''''''''''''''''''''''''''
    If ActiveAlterImage.Value And ((RepetitionNumber - 1) Mod RoundAlterTrack = 0) Then
        Dim FilePathAlt As String ' name of path for alternative imaging
        Dim FileNameAlt As String ' name of file for alternative imaging
        Dim AcquireAltImage As Boolean
        AcquireAltImage = False
        'if we have subpositions
        If GridScan_nColumnsub.Value * GridScan_nRowsub.Value > 1 Then
            If ((GridPos.RowSub - 1) * UBound(posGridX, 4) + GridPos.ColSub) Mod RoundAlterLocation = 0 Then
                AcquireAltImage = True
            End If
        ElseIf ((GridPos.Row - 1) * UBound(posGridX, 2) + GridPos.Col) Mod RoundAlterLocation = 0 Then
            AcquireAltImage = True
        End If

        If AcquireAltImage Then
            Time = Timer
            If ActiveAutofocus Then
                Offset = AutofocusZOffset
            Else
                Offset = 0
            End If
            DisplayProgress "Addition acquisition: prepeare settings at ZOffset position...", RGB(0, &HC0, 0)
            ' setup acquisition paramneters
            If Not ActivateTrack(GlobalAltRecording, "AlterAcquisition", DeltaZ) Then            'An additional control....
                MsgBox "No track selected for Additional Acquisition! Cannot Acquire!"
                Exit Function
            End If
            'center the slide
            If Not Recenter_pre(Znew + Offset, SuccessRecenter, ZEN) Then
                Exit Function
            End If

            FilePathAlt = DatabaseTextbox.Value & BackSlash & TextBoxFileName.Value & UnderScore & "Alt_" & FileNameID & ".lsm" ' fullpath of alternative file
            FileNameAlt = TextBoxFileName.Value & UnderScore & "Alt_" & FileNameID & ".lsm"
            
            LogMsg = "% Start button: Additional acquisition " & FilePathAlt
            LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
            
            DisplayProgress "Additional acquisition: acquire...", RGB(0, &HC0, 0)
            If Not StartAlternativeImaging(RecordingDoc, FilePathAlt, FileNameAlt) Then
                    Exit Function
            End If
             
            
            ''' Recenter
            DisplayProgress "Additional acquisition:  wait recenter ...", RGB(0, &HC0, 0)
            
            If Not Recenter_post(Znew + Offset, SuccessRecenter, ZEN) Then
                Exit Function
            End If
            
            LogMsg = "% StartButton:  wait to return center Z (post AltImg) " & Znew + Offset
            pos = Lsm5.Hardware.CpFocus.Position
            If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
                LogMsg = LogMsg & ", obtained Z " & pos & ", position " & pos & ", success within rep." & SuccessRecenter
            Else
                LogMsg = LogMsg & ", obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos _
                & ", position " & pos & ", success within rep." & SuccessRecenter
            End If
            LogMsg = LogMsg & vbCrLf & "% Startbutton:  time for additional acquisition + centering " & Round(Timer - Time)
            LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
            
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

    If Not ActiveMicropilot Then ' without MicroPilot
        ' TODO: Revise all this code
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
            If GridPos.Row = UBound(posGridX, 1) And GridPos.Col = UBound(posGridX, 2) Then
                If GridPos.RowSub = UBound(posGridX, 3) And GridPos.Col = UBound(posGridX, 4) Then  'Allows again to do an extrableach at the end (Why???)
                    ExtraBleachButton.Caption = "Bleach"
                    ExtraBleachButton.BackColor = &H8000000F
                End If
            End If
        End If
        ' todo:
        ' but where is the bleaching image stored ?? Bleaching should be revised !!
    End If
                    
    ''''redefine new position
    Xold = Xnew
    Yold = Ynew
    Zold = Znew
    'compute potential new positions for later acquistion
    If TrackingToggle And Not ActiveGridScan Then 'This is if we're doing some postacquisition tracking (not possible with Grid) this is done before Micropilot analysis
        DisplayProgress "Tracking and computing new coordinates of " & vbCrLf & _
                        "Well/Position Row: " & GridPos.Row & ", Column: " & GridPos.Col & vbCrLf & _
                        "subposition   Row: " & GridPos.RowSub & ", Column: " & GridPos.ColSub & vbCrLf, RGB(&HC0, &HC0, 0)
        DoEvents
        Time = Timer
        If Not ComputeNewCoordinatesAfterAF(RecordingDoc, Xnew, Ynew, Znew, DeltaZ, "Tracking") Then
            Exit Function
        End If
        LogMsg = " StartButton: Time compute center of mass AQ img " & Round(Timer - Time, 2)
        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    End If



    ' COMMUNICATION WITH MICROPILOT: START *****************
      
    If ActiveMicropilot Or ActiveBleach Then
        SaveSetting "OnlineImageAnalysis", "macro", "filepath", FilePath
        'Wait for anything to sto
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
        SaveSetting "OnlineImageAnalysis", "macro", "code", "1"
        'recenter
        Recenter_pre Zold, SuccessRecenter, ZEN
        
        If Not MicroscopePilot(RecordingDoc, GridPos, Xold, Yold, Zold, FileNameID, HighResArrayX, HighResArrayY, HighResArrayZ, HighResArrayDeltaZ, FcsData) Then
            Exit Function
        End If
    End If
    ' one could monitor weather this position was any good at all here. Goodpositions
    ' COMMUNICATION WITH MICROPILOT: END *****************
    
    If AutofocusTrackZ And ActiveAutofocus Then
        posGridZ(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) = Znew
        Recenter_pre Znew, SuccessRecenter, ZEN
    Else
        Recenter_pre Zold, SuccessRecenter, ZEN
    End If
    
    If TrackingToggle Then
        'move to new position
        If PostTrackXY.Value Then
            If Not FailSafeMoveStageXY(Xnew, Ynew) Then
                Exit Function
            End If
        End If
        ' update positions for next acquistion
        posGridX(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) = Xnew
        posGridY(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) = Ynew
        If AutofocusTrackZ Then
            posGridZ(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub) = Znew
        End If
    Else ' no location tracking
        Lsm5.Hardware.CpHrz.Leveling   'This I think puts the HRZ to its resting position, and moves the focuswheel correspondingly. Do we need this?
    End If
    ''  End: Setting new (x,y)z positions *******************************
    If Log Then
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        LogFile.Close
    End If

    ImagingWorkFlow = True

End Function
    
Private Function SetLocationTextLabel(X As Double, Y As Double, Z As Double, GridPos As GridPosType, TotPos As Long, iRep As Integer) As String
    Dim Caption As String
    Caption = "Locations : " & TotPos & "/" & UBound(posGridX, 1) * UBound(posGridX, 2) * UBound(posGridX, 3) * UBound(posGridX, 4) & "; X= " & X & ", Y = " & Y & ", Z = " & Z & vbCrLf & _
                                "Repetition :" & iRep & "/" & LoopingRepetitions.Value & vbCrLf & _
                                "Well/Position Row: " & GridPos.Row & "/" & UBound(posGridX, 1) & "; Column: " & GridPos.Col & "/" & UBound(posGridX, 2) & vbCrLf

                                
     If MultipleLocationToggle Or ActiveGridScan Then
        If GridScan_nRowsub.Value > 1 Or GridScan_nColumnsub.Value > 1 Then
            Caption = Caption & "Subposition   Row: " & GridPos.RowSub & "/" & UBound(posGridX, 3) & "; Column: " & GridPos.ColSub & "/" & UBound(posGridX, 4) & vbCrLf
        End If
    End If
    SetLocationTextLabel = Caption
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
    
    GetCurrentPositionOffsetButton.Enabled = True
    AutofocusButton.Enabled = True
    GlobalPrvTime = CDbl(GetTickCount) * 0.001
    rettime = GlobalPrvTime
    DiffTime = rettime - GlobalPrvTime
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

        DisplayProgress "Pause " & CStr(CInt(DiffTime)) & " s", RGB(&HC0, &HC0, 0)
        rettime = CDbl(GetTickCount) * 0.001
        DiffTime = rettime - GlobalPrvTime
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
        ActiveGridScan.Value = False
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
        
        ActiveGridScan.Value = False ' currently not compatible
        
        
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

    
    AutofocusTrack1.Visible = False
    AutofocusTrack1.Enabled = False
    AutofocusTrack1.Value = False
    AcquisitionTrack1.Visible = False
    AcquisitionTrack1.Enabled = False
    AcquisitionTrack1.Value = False
    MicropilotTrack1.Visible = False
    MicropilotTrack1.Enabled = False
    MicropilotTrack1.Value = False
    AlterTrack1.Visible = False
    AlterTrack1.Enabled = False
    AlterTrack1.Value = False
    BleachTrack1.Visible = False
    BleachTrack1.Enabled = False
    BleachTrack1.Value = False
    
    AutofocusTrack2.Visible = False
    AutofocusTrack2.Enabled = False
    AutofocusTrack2.Value = False
    AcquisitionTrack2.Visible = False
    AcquisitionTrack2.Enabled = False
    AcquisitionTrack2.Value = False
    MicropilotTrack2.Visible = False
    MicropilotTrack2.Enabled = False
    MicropilotTrack2.Value = False
    AlterTrack2.Visible = False
    AlterTrack2.Enabled = False
    AlterTrack2.Value = False
    BleachTrack2.Visible = False
    BleachTrack2.Enabled = False
    BleachTrack2.Value = False
    
    AutofocusTrack3.Visible = False
    AutofocusTrack3.Enabled = False
    AutofocusTrack3.Value = False
    AcquisitionTrack3.Visible = False
    AcquisitionTrack3.Enabled = False
    AcquisitionTrack3.Value = False
    MicropilotTrack3.Visible = False
    MicropilotTrack3.Enabled = False
    MicropilotTrack3.Value = False
    AlterTrack3.Visible = False
    AlterTrack3.Enabled = False
    AlterTrack3.Value = False
    BleachTrack3.Visible = False
    BleachTrack3.Enabled = False
    BleachTrack3.Value = False
   
    AutofocusTrack4.Visible = False
    AutofocusTrack4.Enabled = False
    AutofocusTrack4.Value = False
    AcquisitionTrack4.Visible = False
    AcquisitionTrack4.Enabled = False
    AcquisitionTrack4.Value = False
    MicropilotTrack4.Visible = False
    MicropilotTrack4.Enabled = False
    MicropilotTrack4.Value = False
    AlterTrack4.Visible = False
    AlterTrack4.Enabled = False
    AlterTrack4.Value = False
    BleachTrack4.Visible = False
    BleachTrack4.Enabled = False
    BleachTrack4.Value = False

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
                        AutofocusTrack1.Visible = True
                        AutofocusTrack1.Caption = TrackName
                        AutofocusTrack1.Enabled = ActiveAutofocus.Value
                        AutofocusTrack1.BackColor = Color
                        
                        AcquisitionTrack1.Visible = True
                        AcquisitionTrack1.Caption = TrackName
                        AcquisitionTrack1.Enabled = True
                        AcquisitionTrack1.BackColor = Color
                        
                        MicropilotTrack1.Visible = True
                        MicropilotTrack1.Caption = TrackName
                        MicropilotTrack1.Enabled = ActiveMicropilot.Value
                        MicropilotTrack1.BackColor = Color
                        
                        AlterTrack1.Visible = True
                        AlterTrack1.Caption = TrackName
                        AlterTrack1.Enabled = ActiveAlterImage.Value
                        AlterTrack1.BackColor = Color
                        
                        BleachTrack1.Visible = True
                        BleachTrack1.Caption = TrackName
                        BleachTrack1.Enabled = ActiveBleach.Value
                        BleachTrack1.BackColor = Color
                        
                    End If
                    If GoodTracks = 2 Then
                        AutofocusTrack2.Visible = True
                        AutofocusTrack2.Caption = TrackName
                        AutofocusTrack2.Enabled = ActiveAutofocus
                        AutofocusTrack2.BackColor = Color
                        
                        AcquisitionTrack2.Visible = True
                        AcquisitionTrack2.Caption = TrackName
                        AcquisitionTrack2.Enabled = True
                        AcquisitionTrack2.BackColor = Color
                        
                        MicropilotTrack2.Visible = True
                        MicropilotTrack2.Caption = TrackName
                        MicropilotTrack2.Enabled = ActiveMicropilot.Value
                        MicropilotTrack2.BackColor = Color
                        
                        AlterTrack2.Visible = True
                        AlterTrack2.Caption = TrackName
                        AlterTrack2.Enabled = ActiveAlterImage.Value
                        AlterTrack2.BackColor = Color
                        
                        BleachTrack2.Visible = True
                        BleachTrack2.Caption = TrackName
                        BleachTrack2.Enabled = ActiveBleach.Value
                        BleachTrack2.BackColor = Color
                        
                    End If
                    If GoodTracks = 3 Then
                        AutofocusTrack3.Visible = True
                        AutofocusTrack3.Caption = TrackName
                        AutofocusTrack3.Enabled = ActiveAutofocus
                        AutofocusTrack3.BackColor = Color
                        
                        AcquisitionTrack3.Visible = True
                        AcquisitionTrack3.Caption = TrackName
                        AcquisitionTrack3.Enabled = True
                        AcquisitionTrack3.BackColor = Color
                        
                        MicropilotTrack3.Visible = True
                        MicropilotTrack3.Caption = TrackName
                        MicropilotTrack3.Enabled = ActiveMicropilot.Value
                        MicropilotTrack3.BackColor = Color
                        
                        AlterTrack3.Visible = True
                        AlterTrack3.Caption = TrackName
                        AlterTrack3.Enabled = ActiveAlterImage.Value
                        AlterTrack3.BackColor = Color
                        
                        BleachTrack3.Visible = True
                        BleachTrack3.Caption = TrackName
                        BleachTrack3.Enabled = ActiveBleach.Value
                        BleachTrack3.BackColor = Color
                        
                    End If
                    If GoodTracks = 4 Then
                        AutofocusTrack4.Visible = True
                        AutofocusTrack4.Caption = TrackName
                        AutofocusTrack4.Enabled = ActiveAutofocus
                        AutofocusTrack4.BackColor = Color
                        
                        AcquisitionTrack4.Visible = True
                        AcquisitionTrack4.Caption = TrackName
                        AcquisitionTrack4.Enabled = True
                        AcquisitionTrack4.BackColor = Color
                        
                        MicropilotTrack4.Visible = True
                        MicropilotTrack4.Caption = TrackName
                        MicropilotTrack4.Enabled = ActiveMicropilot.Value
                        MicropilotTrack4.BackColor = Color
                        
                        AlterTrack4.Visible = True
                        AlterTrack4.Caption = TrackName
                        AlterTrack4.Enabled = ActiveAlterImage.Value
                        AlterTrack4.BackColor = Color
                        
                        BleachTrack4.Visible = True
                        BleachTrack4.Caption = TrackName
                        BleachTrack4.Enabled = ActiveBleach.Value
                        BleachTrack4.BackColor = Color
                        
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
        If GlobalAcquisitionRecording.ScanMode = "Stack" Or GlobalAcquisitionRecording.ScanMode = "ZScanner" Then
            PostTrackZ.Enabled = True
        Else
            PostTrackZ.Enabled = False
            PostTrackZ.Value = False
        End If
    End If
    PostTrackXY.Visible = Enable
    PostTrackZ.Visible = Enable
    PostAcquisitionLabel.Visible = Enable

End Sub
    

'''''''
' AutofocusTrackZ_Click()
'   Activate post-acquisition Z-tracking.
'   Inactivates Autofocus
'''''''
Private Sub PostTrackZ_Click()
    If PostTrackZ.Value = True Then
        ActiveAutofocus.Value = False 'inactivate Autofocus
        SwitchEnableAutofocusPage (False)
    Else
        ActiveAutofocus.Value = True
        SwitchEnableAutofocusPage (True)
    End If
End Sub


'fills popup menu for chosing a track for post-acquisition tracking
' TODO: move in form
Private Sub FillTrackingChannelList()
    Dim T As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim channel As DsDetectionChannel
    Dim Track As DsTrack
    
    ReDim ActiveChannels(Lsm5.Constants.MaxActiveChannels)  'ActiveChannels is a dynamic array (variable size), ReDim defines array size required next
                                                            'Array size is (MaxActiveChannels gets) the total max number of active channels in all tracks
    ComboBoxTrackingChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    
    If ActivateTrack(GlobalAcquisitionRecording, "Acquisition") Then
        For T = 1 To TrackNumber 'This loop goes through all tracks and will collect all activated channels to display them in popup menu
            Set Track = GlobalAcquisitionRecording.TrackObjectByMultiplexOrder(T - 1, Success)
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
        Next T
        ComboBoxTrackingChannel.Value = ComboBoxTrackingChannel.List(0) 'initially displayed text in popup menu is a blank line (first channel is 1).
    End If
End Sub

Private Sub ComboBoxTrackingChannel_Change()        'Sets the name of the channel for PostAcquisition tracking.
If AcquisitionTracksOn Then
    TrackingChannelString = ActiveChannels(ComboBoxTrackingChannel.ListIndex + 1)
End If
End Sub



Private Sub CommandTimeMin_Click()
    LoopingTimerUnit = 60
    LoopingRepetitionTime.Max = 60                        'When workings with minutes the maximum delay that can be set with the slider is 1 hour
    LoopingRepetitionTime.Value = BlockTimeDelay / 60
    CommandTimeMin.BackColor = &HFF8080
    CommandTimeSec.BackColor = &H8000000F
End Sub

Private Sub CommandTimeSec_Click()
    LoopingTimerUnit = 1
    LoopingRepetitionTime.Max = 180                       'When workings with seconds the maximum delay that can be set with the slider is 3 minutes
    LoopingRepetitionTime.Value = BlockTimeDelay
    CommandTimeSec.BackColor = &HFF8080
    CommandTimeMin.BackColor = &H8000000F
End Sub

Private Sub LoopingRepetitionTime_Click()
    BlockTimeDelay = LoopingRepetitionTime.Value * LoopingTimerUnit                      'BlockTimedelay gets the value of the slider in seconds
End Sub

Private Sub LoopingRepetitions_Change()
    If Not Running Then
        BlockRepetitions = LoopingRepetitions.Value
    ElseIf Not (LoopingRepetitions.Value <= (RepetitionNumber + 1)) Then
        BlockRepetitions = LoopingRepetitions.Value
    Else
        LoopingRepetitions.Value = RepetitionNumber + 1
        BlockRepetitions = LoopingRepetitions.Value
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
    If AutofocusForm.TextBoxFileName.Value <> "" Then
        UnderScore = "_"
    Else
        UnderScore = ""
    End If
End Sub

'''''
'   ActivateTrack(ByRef Recording As DsRecording, ByVal Track As String) As Boolean
'   Wrapper function to call activate functions for track given a Track Name
'
'''''
Public Function ActivateTrack(ByRef Recording As DsRecording, ByVal Track As String, Optional DeltaZ As Double = -1) As Boolean
    Select Case Track
        Case "Autofocus":
            ActivateTrack = ActivateAutofocusTrack(Recording)
        Case "Acquisition":
            ActivateTrack = ActivateAcquisitionTrack(Recording, DeltaZ)
        Case "AlterAcquisition":
            ActivateTrack = ActivateAlterAcquisitionTrack(Recording, DeltaZ)
        Case "Micropilot":
            ActivateTrack = ActivateMicropilotTrack(Recording, DeltaZ)
        Case "Bleach":
            ActivateTrack = ActivateBleachTrack(Recording, DeltaZ)
        Case Else:
            MsgBox "ActivateTrack: Not able to find appropriate imging track to activate! Exit function now!"
    End Select
End Function
 
'''''''
'   ActivateAutofocusTrack(Recording As DsRecording, posZ As Double, pixelDwell As Double)
'   Check which Track should be used for Autofocus and update passed DsRercoding.
'   posZ is position with respect to which calculate central slice (often the CpFocus still moves after acquisition)
'   This sets also the Z acquisition parameters for Acquisition document. For this one uses parameters of the AutofocusForm
'       [Recording] In/Out - a DsRecording
'   TODO: test
''''''
Private Function ActivateAutofocusTrack(Recording As DsRecording) As Boolean
    Dim i As Integer
    Dim iZoom As Integer
    Dim TrackSuccess As Integer
    Dim FunSuccess As Boolean
    Dim pixelDwell As Double
    Dim ZoomPixelSlice(1 To 9, 1 To 3) As Double
    iZoom = -1
    
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
    
    FunSuccess = False
    ' Set all tracks to non-acquisition first
    For i = 1 To TrackNumber
       Recording.TrackObjectByMultiplexOrder(i - 1, TrackSuccess).Acquire = 0
    Next i
    
    For i = 1 To TrackNumber
        If AutofocusTrack1.Value = True And i = 1 Then
            FunSuccess = True
            Exit For
        ElseIf AutofocusTrack2.Value = True And i = 2 Then
            FunSuccess = True
            Exit For
        ElseIf AutofocusTrack3.Value = True And i = 3 Then
            FunSuccess = True
            Exit For
        ElseIf AutofocusTrack4.Value = True And i = 4 Then
            FunSuccess = True
            Exit For
        End If
    Next i
    
    If FunSuccess Then
        AutofocusTrack = i - 1
        Recording.TrackObjectByMultiplexOrder(AutofocusTrack, Success).Acquire = True
        If AutofocusMaxSpeed.Value Then
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
    
    If AutofocusLowZoom.Value Then                   ' Changes the zoom if necessary
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
        If AutofocusHRZ.Value Then                'Piezo
            Recording.SpecialScanMode = "ZScanner"
        Else
            Recording.SpecialScanMode = "FocusStep"
        End If
        'highspeed does set the pixeldwell time to its minimal possible value
        If AutofocusMaxSpeed.Value Then
            'compute maximal possible pixwelDwell for given zoom
            For i = 1 To UBound(ZoomPixelSlice, 1) - 1
                If Recording.ZoomX < ZoomPixelSlice(i, 1) And Recording.ZoomX >= ZoomPixelSlice(i + 1, 1) Then
                    iZoom = i
                    pixelDwell = ZoomPixelSlice(i, 2)
                End If
            Next i
            'do biderectional scanning (tocheck if it is fine with Zscan)
            If ScanFrameToggle Then
                Recording.ScanDirection = 1
            End If
        Else
            pixelDwell = GlobalBackupSampleObservationTime
            Recording.ScanDirection = GlobalBackupRecording.ScanDirection
        End If
        
        If iZoom < 0 Then
            iZoom = 1
        End If
        
        '''''''''''''''''''''''''''
        '**Setting for line scan**'
        '''''''''''''''''''''''''''
        If ScanLineToggle.Value Then
            Recording.ScanMode = "ZScan"             'This acquires  single X-Z image, like with "Range Select" button Z-stack Window.
            Recording.SamplesPerLine = AutofocusLineSize.Value
            Recording.LinesPerFrame = 1

            If AutofocusHRZ Then
                Recording.SpecialScanMode = "ZScanner"
            Else
                If AutofocusFastZline And AutofocusZStep.Value < ZoomPixelSlice(i, 3) Then
                    Recording.SpecialScanMode = "OnTheFly" 'aka: Fast Z-line in Z-Stack menu
                    For i = iZoom To UBound(ZoomPixelSlice, 1)
                        If AutofocusZStep.Value < ZoomPixelSlice(i, 3) Then
                            pixelDwell = ZoomPixelSlice(i, 2)
                            Exit For
                        End If
                    Next i
                 Else
                    If AutofocusZStep.Value > ZoomPixelSlice(i, 3) And (AutofocusFastZline And Not AutofocusHRZ) Then
                        DisplayProgress "Highest Z Step of 1.54 um with no piezo and Fast " & _
                        "Z line has been reached. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
                    End If
                    Recording.SpecialScanMode = "FocusStep"
                 End If
            End If
        End If
        
        ''''''''''''''''''''''''''''
        '**Setting for frame scan**'
        ''''''''''''''''''''''''''''
        If ScanFrameToggle.Value Then
            Recording.ScanMode = "Stack"                       'This is defining to acquire a Z stack of Z-Y images
            Recording.SamplesPerLine = AutofocusFrameSize.Value  'If doing frame autofocussing it uses the userdefined frame size
            Recording.LinesPerFrame = AutofocusFrameSize.Value
        End If
        pixelDwell = pixelDwell * 256 / Recording.SamplesPerLine
    End If  ' If SystemName = "LSM"
    
    Sleep (100)
    ' set the pixelDwellTime globally
    NoFrames = CLng(AutofocusZRange.Value / AutofocusZStep.Value) + 1   'Calculates the number of frames per stack. Clng converts it to a long and rounds up the fraction
    If NoFrames > 2048 Then                                         'overwrites the userdefined value if too many frames have been defined by the user
        NoFrames = 2048
    End If
    Recording.FrameSpacing = AutofocusZStep.Value
    Recording.FramesPerStack = NoFrames
    Recording.TimeSeries = True   ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
    Recording.StacksPerRecord = 1 ' why only one and not more

    Recording.TrackObjectByMultiplexOrder(AutofocusTrack, 1).SampleObservationTime = pixelDwell
    Lsm5.DsRecording.Copy Recording
    Lsm5.DsRecording.TrackObjectByMultiplexOrder(AutofocusTrack, 1).SampleObservationTime = pixelDwell
    Lsm5.DsRecording.FrameSpacing = AutofocusZStep.Value
    Lsm5.DsRecording.FramesPerStack = NoFrames
    
    ' need to do it twice:  set new pixelDwell and FrameSpacing This is asolutely required
    Lsm5.DsRecording.TrackObjectByMultiplexOrder(AutofocusTrack, 1).SampleObservationTime = pixelDwell
    Lsm5.DsRecording.FrameSpacing = AutofocusZStep.Value
    Lsm5.DsRecording.SpecialScanMode = Recording.SpecialScanMode

    ActivateAutofocusTrack = FunSuccess
End Function

'''''''''
' ActivateAcquisitionTrack()
' If any of the checkboxes in the AutoFocusForm Acquisition are checked activates themin DsRecording
'   [Recording] In/Out - a DsRecording
' TODO: Test
''''''''''
Private Function ActivateAcquisitionTrack(Recording As DsRecording, Optional DeltaZ As Double = -1) As Boolean
    Dim i As Integer
    Dim TrackSuccess As Integer
    Dim FunSuccess As Boolean
    Dim ToActivate() As Boolean
    ReDim ToActivate(1 To TrackNumber)
    FunSuccess = False

    For i = 1 To TrackNumber
        ToActivate(i) = False
        If AcquisitionTrack1.Value = True And i = 1 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf AcquisitionTrack2.Value = True And i = 2 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf AcquisitionTrack3.Value = True And i = 3 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf AcquisitionTrack4.Value = True And i = 4 Then
            ToActivate(i) = True
            FunSuccess = True
        End If
    Next i


    Recording.TimeSeries = True   ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
    Recording.StacksPerRecord = 1 ' why only one and not more
    'can't put Lsm5.DsRecording here. as it is not followed. Why?
    If FunSuccess Then
        Lsm5.DsRecording.Copy Recording
        'Activate the tracks
        For i = 1 To TrackNumber
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, 1).Acquire = ToActivate(i)
        Next i
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
        'change number of Slices
        If DeltaZ > 0 Then
            Lsm5.DsRecording.FramesPerStack = CLng(DeltaZ / Recording.FrameSpacing) + 1
        End If
    End If
    If Not ScanStop Then
        ActivateAcquisitionTrack = FunSuccess
    End If
End Function



'''''''
'   ActivateAlterAcquisitionTrack
'   Check which track has been activated and for AlternativeAcquisitionTrack set the track properties accordingly
'   TODO: Test
''''''
Private Function ActivateAlterAcquisitionTrack(Recording As DsRecording, Optional DeltaZ As Double = -1) As Boolean
    Dim i As Integer
    Dim FunSuccess As Boolean
    Dim ToActivate() As Boolean
    ReDim ToActivate(1 To TrackNumber)
    FunSuccess = False
    
    For i = 1 To TrackNumber
        ToActivate(i) = False
        If AlterTrack1.Value = True And i = 1 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf AlterTrack2.Value = True And i = 2 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf AlterTrack3.Value = True And i = 3 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf AlterTrack4.Value = True And i = 4 Then
            ToActivate(i) = True
            FunSuccess = True
        End If
    Next i
    
    If FunSuccess Then
        Recording.TimeSeries = True  ' This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure what is the reason for this
        Recording.StacksPerRecord = 1 ' This is time series stack!
            ' get and set the values from the Form
        Recording.ZoomX = AlterZoom.Value
        Recording.ZoomY = AlterZoom.Value
        Recording.ScanMode = "Stack"
        Recording.FrameSpacing = CDbl(AlterZStep.Value)
        Recording.FramesPerStack = CDbl(AlterZSlices.Value)
        Recording.SamplesPerLine = AlterFrameSize.Value
        Recording.LinesPerFrame = AlterFrameSize.Value
        Recording.SpecialScanMode = GlobalAcquisitionRecording.SpecialScanMode
        'Activate the tracks
        For i = 1 To TrackNumber
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, 1).Acquire = ToActivate(i)
        Next i
        Lsm5.DsRecording.Copy Recording
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
        'change number of Slices
        If DeltaZ > 0 Then
            Lsm5.DsRecording.FramesPerStack = CLng(DeltaZ / Recording.FrameSpacing) + 1
        End If
    End If
    If Not ScanStop Then
        ActivateAlterAcquisitionTrack = FunSuccess
    End If
End Function



'''''''''
' ActivateMicropilotTrack()
' Micropilotpage. This is extra track for micropilot
' TODO: Test and change name
''''''''''
Private Function ActivateBleachTrack(Recording As DsRecording, Optional DeltaZ As Double = -1) As Boolean
    Dim i As Integer
    Dim FunSuccess As Boolean
    Dim ToActivate() As Boolean
    ReDim ToActivate(1 To TrackNumber)
    
    FunSuccess = False
    ' Set all tracks to non-acquisition first

    For i = 1 To TrackNumber
        ToActivate(i) = False
        If BleachTrack1.Value = True And i = 1 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf BleachTrack2.Value = True And i = 2 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf BleachTrack3.Value = True And i = 3 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf BleachTrack4.Value = True And i = 4 Then
            ToActivate(i) = True
            FunSuccess = True
        End If
    Next i
    
    If FunSuccess Then
        Recording.ScanMode = "Stack"
        Recording.SamplesPerLine = BleachFrameSize.Value
        Recording.LinesPerFrame = BleachFrameSize.Value
        Recording.ZoomX = BleachZoom.Value
        Recording.ZoomY = BleachZoom.Value
        Recording.FramesPerStack = 1
        
        Set Track = Recording.TrackObjectBleach(Success)
        If Success Then
            Track.Acquire = True
            Track.UseBleachParameters = True            'Bleach parameters are lasers lines, bleach iterations... stored in the bleach control window
        End If
        
        Lsm5.DsRecording.Copy Recording
        Recording.StacksPerRecord = BleachRepetitions.Value
        Track.TimeBetweenStacks = BleachRepetitionTime.Value
        
        Lsm5.DsRecording.Copy Recording
        Lsm5.DsRecording.TimeSeries = True
        Lsm5.DsRecording.StacksPerRecord = BleachRepetitions.Value
        Lsm5.DsRecording.FramesPerStack = BleachZSlices.Value
        Lsm5.DsRecording.FrameSpacing = BleachZStep.Value
        Lsm5.DsRecording.ScanMode = "Stack"
        'set the correct dwelltime
        For i = 1 To TrackNumber
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, 1).Acquire = ToActivate(i)
        Next i
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
    Else
        Exit Function
    End If
    If Not ScanStop Then
        ActivateBleachTrack = FunSuccess
    End If
End Function



'''''''''
' ActivateBleachTrack()
' Micropilotpage. This is extra track for micropilot
' TODO: Test and change name
''''''''''
Private Function ActivateMicropilotTrack(Recording As DsRecording, Optional DeltaZ As Double = -1) As Boolean
    Dim i As Integer
    Dim FunSuccess As Boolean
    Dim ToActivate() As Boolean
    ReDim ToActivate(1 To TrackNumber)
    
    FunSuccess = False
    ' Set all tracks to non-acquisition first

    For i = 1 To TrackNumber
        ToActivate(i) = False
        If MicropilotTrack1.Value = True And i = 1 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf MicropilotTrack2.Value = True And i = 2 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf MicropilotTrack3.Value = True And i = 3 Then
            ToActivate(i) = True
            FunSuccess = True
        ElseIf MicropilotTrack4.Value = True And i = 4 Then
            ToActivate(i) = True
            FunSuccess = True
        End If
    Next i
    
    If FunSuccess Then
        Recording.ScanMode = "Stack"
        Recording.SamplesPerLine = MicropilotFrameSize.Value
        Recording.LinesPerFrame = MicropilotFrameSize.Value
        Recording.ZoomX = MicropilotZoom.Value
        Recording.ZoomY = MicropilotZoom.Value
        Recording.FrameSpacing = MicropilotZStep.Value
        Recording.FramesPerStack = MicropilotZSlices.Value
    
        Lsm5.DsRecording.Copy Recording
        Lsm5.DsRecording.TimeSeries = True
        Lsm5.DsRecording.StacksPerRecord = 1
        Lsm5.DsRecording.FrameSpacing = MicropilotZStep.Value
        Lsm5.DsRecording.FramesPerStack = MicropilotZSlices.Value
    
    
        Lsm5.DsRecording.ScanMode = "Stack"
        If AutofocusHRZ.Value Then
            Lsm5.DsRecording.SpecialScanMode = "ZScanner"
        Else
            Lsm5.DsRecording.SpecialScanMode = "FocusStep"
        End If
        'set the correct dwelltime
        For i = 1 To TrackNumber
            Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, 1).Acquire = ToActivate(i)
        Next i
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = GlobalBackupSampleObservationTime
        If DeltaZ > 0 Then
            Lsm5.DsRecording.FramesPerStack = CLng(DeltaZ / Recording.FrameSpacing) + 1
        End If
    End If
    If Not ScanStop Then
        ActivateMicropilotTrack = FunSuccess
    End If
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
 
    AutofocusMaxSpeed.Value = BlockHighSpeed
    AutofocusLowZoom.Value = BlockLowZoom
    AutofocusHRZ.Value = BlockHRZ
'    Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
'    If Position >= 0 Then
'        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
'    Else
'        Range = 0#
'    End If
'substituted29.06.2010 by Function Range
    If AutofocusZRange.Value > Range() * 0.9 Then
        AutofocusZRange.Value = Range() * 0.9
    End If
    If Abs(AutofocusZOffset.Value) > Range() * 0.9 Then
        AutofocusZOffset.Value = 0
    End If
    AutofocusZOffset.Value = AutofocusZOffset.Value
    AutofocusZRange.Value = AutofocusZRange.Value
    AutofocusZStep.Value = BlockZStep

End Sub


'''''
' TODO: All block values should use the checkboxes directly
'''''
Public Sub GetBlockValues()
   
    BlockHighSpeed = AutofocusMaxSpeed.Value
    BlockLowZoom = AutofocusLowZoom.Value
    HRZ = AutofocusHRZ.Value  ' this is for the piezo
    BlockZOffset = AutofocusZOffset.Value
    BlockZRange = AutofocusZRange.Value
    BlockZStep = AutofocusZStep.Value

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
    If AcquisitionTrack1.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, Success)
        Track1Speed = Track.SampleObservationTime
    End If
    If AcquisitionTrack2.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
        Track2Speed = Track.SampleObservationTime
    End If
    If AcquisitionTrack3.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(2, Success)
        Track3Speed = Track.SampleObservationTime
    End If
    If AcquisitionTrack4.Value = True Then
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



Private Sub AcquisitionTrack1_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Sub AcquisitionTrack2_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Sub AcquisitionTrack3_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

Private Sub AcquisitionTrack4_Change()
    TrackingToggle.Enabled = AcquisitionTracksOn
    If Not TrackingToggle.Enabled Then
        TrackingToggle.Value = False
    End If
    SwitchEnableTrackingToggle TrackingToggle.Value
End Sub

''''
'  AcquisitionTracksOn()
'  Checks if at least one track for acquisition is on
'''
Private Function AcquisitionTracksOn() As Boolean
    If AcquisitionTrack1 Then
        AcquisitionTracksOn = True
    End If
    If AcquisitionTrack2 Then
        AcquisitionTracksOn = True
    End If
    If AcquisitionTrack3 Then
        AcquisitionTracksOn = True
    End If
    If AcquisitionTrack4 Then
        AcquisitionTracksOn = True
    End If
End Function

'''
' Sets all acquisitions to off
'''
Private Function AcquisitionTracksSetOff() As Boolean
    AcquisitionTrack1.Value = 0
    AcquisitionTrack2.Value = 0
    AcquisitionTrack3.Value = 0
    AcquisitionTrack4.Value = 0
End Function

Public Function AutofocusTime() As Double
    Dim Speed As Double
    Dim Pixels As Long
    Dim FrameNumber As Integer
    Dim ScanDirection As Integer
    Dim i As Integer

    Speed = 0
    If AutofocusMaxSpeed.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, Success)
        Speed = 1.76 * 10 ^ -6
    Else
        If AutofocusTrack1.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
        If AutofocusTrack2.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
        If AutofocusTrack3.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
        If AutofocusTrack4.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, Success)
            Speed = Track.SampleObservationTime
        End If
    End If
    Pixels = 512
    AutofocusForm.GetBlockValues
    FrameNumber = CLng(AutofocusZRange.Value / AutofocusZStep.Value) + 1
    If Lsm5.DsRecording.ScanDirection = 0 Then
        ScanDirection = 1
    Else
        ScanDirection = 2
    End If
    If AutofocusHRZ.Value = True Then
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



''''
' Not anymore in use
''''
Private Sub CreateZoomDatabase(ZoomDatabaseName, HighResExperimentCounter, ZoomExpname)
            'Create ZoomDatabase
            Dim Start As Integer
            Dim bslash As String
            Dim pos As Long
            Dim NameLength As Long
            Dim MyPath As String
            
            Start = 1
            bslash = "\"
            pos = Start
            Do While pos > 0
                pos = InStr(Start, DatabaseTextbox.Value, bslash)
                If pos > 0 Then
                    Start = pos + 1
                End If
            Loop
            
            MyPath = DatabaseTextbox.Value + bslash
            NameLength = Len(DatabaseTextbox.Value)
            ZoomExpname = Strings.Right(DatabaseTextbox.Value, NameLength - Start + 1)
           ' NameLength = Len(Myname)
           ' Myname = Strings.Left(Myname, NameLength - 4)
            ZoomDatabaseName = MyPath & ZoomExpname & "_" & TextBoxFileName.Value & LocationName & "_R" & RepetitionNumber & "_Exp" & HighResExperimentCounter & "_zoom"
            ' Lsm5.NewDatabase (ZoomDatabaseName)
           ' ZoomDatabaseName = ZoomDatabaseName & "\" & Myname & "_zoom.mdb"
    
End Sub

Private Sub CreateAlterImageDatabase(AlterDatabaseName, MyPath)
        Dim Start As Integer
        Dim bslash As String
        Dim pos As Long
        Dim NameLength As Long
        Dim Myname As String

         Start = 1
         bslash = "\"
         pos = Start
         Do While pos > 0
             pos = InStr(Start, DatabaseTextbox.Value, bslash)
             If pos > 0 Then
                 Start = pos + 1
             End If
         Loop
         MyPath = Strings.Left(DatabaseTextbox.Value, Start - 1)
         NameLength = Len(DatabaseTextbox.Value)
         Myname = Strings.Right(DatabaseTextbox.Value, NameLength - Start + 1)
         NameLength = Len(Myname)
         ' Myname = Strings.Left(Myname, NameLength - 4)
         AlterDatabaseName = MyPath & Myname & "_additionalTracks"
        ' Lsm5.NewDatabase (AlterDatabaseName)
        '  AlterDatabaseName = AlterDatabaseName & "\" & Myname & "_additionalTracks"
         
End Sub



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
            If ScanStop Then
                Exit Function
            End If
        Wend
    End If
    DisplayProgress "Acquiring Additional Track...", RGB(0, &HC0, 0)
    ' take the image
    If Not ScanToImage(RecordingDoc) Then
        ScanStop = True
        Exit Function
    End If

    RecordingDoc.SetTitle name
    
    If Not SaveDsRecordingDoc(RecordingDoc, FilePath) Then
        ScanStop = True
        Exit Function
    End If
    StartAlternativeImaging = True
End Function








'''''
' AFTest1_Click()
' Perform repeatealy Autofocus with FastZline and acquisition with stage only.
' Uses No Z-track and Z-track
''''
Private Sub AFTest1_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest1Run
    StopAcquisition
End Sub

Private Function AFTest1Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    Dim FilePath As String
    Dim MaxTestRepeats As Integer
    Dim TestNr As Integer
    Dim pixelDwell As Double
    Dim i As Integer
    Log = True
    Dim Zold As Double
    Zold = posTempZ
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If
        
    AcquisitionTrack1.Value = AutofocusTrack1.Value
    AcquisitionTrack2.Value = AutofocusTrack2.Value
    AcquisitionTrack3.Value = AutofocusTrack3.Value
    AcquisitionTrack4.Value = AutofocusTrack4.Value
    AutofocusMaxSpeed.Value = True
    AutofocusFastZline = True
    AutofocusHRZ.Value = False
    AutofocusLowZoom.Value = False
        
    '''''''
    ' No Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = False
    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    GlobalBackupRecording.SpecialScanMode = "FocusStep"
    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest1_FastZLine_Stage_NoTrackZ") Then
        Exit Function
    End If
    
    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = True
    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    GlobalBackupRecording.SpecialScanMode = "FocusStep"
    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest1_FastZLine_Stage_TrackZ") Then
        Exit Function
    End If
    
    AFTest1Run = True
End Function


'''''
' AFTest2_Click()
' Perform repeatealy Autofocus with piezo and acquisition with piezo
' Uses No Z-track and Z-track
''''
Private Sub AFTest2_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest2Run
    StopAcquisition
End Sub

Private Function AFTest2Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    Log = True
    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.name) Then
        MsgBox ("No piezo availabe! Cannot start tests.")
        Exit Function
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If
        
    AcquisitionTrack1.Value = AutofocusTrack1.Value
    AcquisitionTrack2.Value = AutofocusTrack2.Value
    AcquisitionTrack3.Value = AutofocusTrack3.Value
    AcquisitionTrack4.Value = AutofocusTrack4.Value
    AutofocusMaxSpeed.Value = True
    AutofocusFastZline = False
    AutofocusHRZ.Value = True
    AutofocusLowZoom.Value = False
        
    '''''''
    ' No Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = False
    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
    GlobalAcquisitionRecording.SpecialScanMode = "ZScanner"
    GlobalBackupRecording.SpecialScanMode = "ZScanner"
    
    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest2_Piezo_Piezo_NoTrackZ") Then
        Exit Function
    End If
    
    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = True
    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
    GlobalAcquisitionRecording.SpecialScanMode = "ZScanner"
    GlobalBackupRecording.SpecialScanMode = "ZScanner"

    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest2_Piezo_Piezo_TrackZ") Then
        Exit Function
    End If
    AFTest2Run = True
End Function


'''''
' AFTest3_Click()
' Perform repeatealy Autofocus with stage and acquisition with stage
' Uses No Z-track and Z-track
''''
Private Sub AFTest3_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest3Run
    StopAcquisition
End Sub

Private Function AFTest3Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    Log = True
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If
        
    AcquisitionTrack1.Value = AutofocusTrack1.Value
    AcquisitionTrack2.Value = AutofocusTrack2.Value
    AcquisitionTrack3.Value = AutofocusTrack3.Value
    AcquisitionTrack4.Value = AutofocusTrack4.Value
    AutofocusMaxSpeed.Value = True
    AutofocusFastZline = False
    AutofocusHRZ.Value = False
    AutofocusLowZoom.Value = False
        
    '''''''
    ' No Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = False
    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
    GlobalBackupRecording.SpecialScanMode = "FocusStep"
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest3_Stage_Stage_NoTrackZ") Then
        Exit Function
    End If
    
    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = True
    GlobalBackupRecording.SpecialScanMode = "FocusStep"
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest3_Stage_Stage_TrackZ") Then
        Exit Function
    End If
    AFTest3Run = True
End Function

'''''
' AFTest4_Click()
' Perform repeatealy Autofocus with piezo and acquisition with stage
' Uses No Z-track and Z-track
''''
Private Sub AFTest4_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest4Run
    StopAcquisition
End Sub

Private Function AFTest4Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    Log = True
    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.name) Then
        MsgBox ("No piezo availabe! Cannot start tests.")
        Exit Function
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If
        
    AcquisitionTrack1.Value = AutofocusTrack1.Value
    AcquisitionTrack2.Value = AutofocusTrack2.Value
    AcquisitionTrack3.Value = AutofocusTrack3.Value
    AcquisitionTrack4.Value = AutofocusTrack4.Value
    AutofocusMaxSpeed.Value = True
    AutofocusFastZline = False
    AutofocusHRZ.Value = True
    AutofocusLowZoom.Value = False
        
    '''''''
    ' No Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = False
    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
    GlobalBackupRecording.SpecialScanMode = "FocusStep"
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    
    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest4_Piezo_Stage_NoTrackZ") Then
        Exit Function
    End If
    
    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = True
    GlobalBackupRecording.SpecialScanMode = "FocusStep"
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest4_Piezo_Stage_TrackZ") Then
        Exit Function
    End If
    AFTest4Run = True
End Function


'''''
' AFTest5_Click()
' Acquire reeatedly images with Fast-Z-Line
''''
Private Sub AFTest5_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest5Run
    StopAcquisition
End Sub

Private Function AFTest5Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If

    AutofocusTrackZ.Value = False
    AcquisitionTrack1.Value = False
    AcquisitionTrack2.Value = False
    AcquisitionTrack3.Value = False
    AcquisitionTrack4.Value = False
    AutofocusMaxSpeed.Value = True
    AutofocusHRZ.Value = False
    AutofocusLowZoom.Value = False
    AutofocusFastZline.Value = True
    AutofocusLineSize.Value = 256
    If Not RunTestFastZline(RecordingDoc, 1, AFTest_Repetitions.Value, 1, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
    AutofocusLineSize.Value = 128
    If Not RunTestFastZline(RecordingDoc, 2, AFTest_Repetitions.Value, 1, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
    AutofocusLineSize.Value = 64
    If Not RunTestFastZline(RecordingDoc, 3, AFTest_Repetitions.Value, 1, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
    AutofocusLineSize.Value = 256
    If Not RunTestFastZline(RecordingDoc, 4, AFTest_Repetitions.Value, 2, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If


    AutofocusLineSize.Value = 128
    If Not RunTestFastZline(RecordingDoc, 5, AFTest_Repetitions.Value, 2, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
    AutofocusLineSize.Value = 256
    If Not RunTestFastZline(RecordingDoc, 6, AFTest_Repetitions.Value, 3, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If


    AutofocusLineSize.Value = 128
    If Not RunTestFastZline(RecordingDoc, 7, AFTest_Repetitions.Value, 3, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
        
    AutofocusLineSize.Value = 256
    If Not RunTestFastZline(RecordingDoc, 8, AFTest_Repetitions.Value, 4, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
      
    AutofocusLineSize.Value = 128
    If Not RunTestFastZline(RecordingDoc, 9, AFTest_Repetitions.Value, 4, "AFTest5_FastZlineTest", 5000) Then
        Exit Function
    End If
    AFTest5Run = True
End Function



'''''
' AFTest6_Click()
' Perform repeatealy Autofocus with piezo and frame acquisition with piezo at multiposition
' Uses No Z-track and Z-track
''''
Private Sub AFTest6_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest6Run
    StopAcquisition
End Sub

Private Function AFTest6Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    Log = True
    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.name) Then
        MsgBox ("No piezo availabe! Cannot start tests.")
        Exit Function
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If
        
    AcquisitionTrack1.Value = AutofocusTrack1.Value
    AcquisitionTrack2.Value = AutofocusTrack2.Value
    AcquisitionTrack3.Value = AutofocusTrack3.Value
    AcquisitionTrack4.Value = AutofocusTrack4.Value
    AutofocusMaxSpeed.Value = True
    AutofocusFastZline = False
    AutofocusHRZ.Value = True
    AutofocusLowZoom.Value = False
        
        
    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = True

    MultipleLocationToggle.Value = True
    LoopingRepetitions = AFTest_Repetitions.Value
    LoopingRepetitionTime.Value = 0
    If Not StartSetting() Then
        Exit Function
    End If
    GlobalAcquisitionRecording.SpecialScanMode = "ZScanner"
    
    GlobalAcquisitionRecording.ScanMode = "Stack"                       'This is defining to acquire a Z stack of Z-Y images
    GlobalAcquisitionRecording.SamplesPerLine = 32  'If doing frame autofocussing it uses the userdefined frame size
    GlobalAcquisitionRecording.LinesPerFrame = 32
    If AutofocusZStep.Value > 0 Then
        GlobalAcquisitionRecording.FramesPerStack = Round(10 / AutofocusZStep.Value)
        GlobalAcquisitionRecording.FrameSpacing = AutofocusZStep.Value
    Else
        GlobalAcquisitionRecording.FramesPerStack = 10
        GlobalAcquisitionRecording.FrameSpacing = 10
    End If
    TextBoxFileName.Value = "Piezo"
    'Set counters back to 1
    RepetitionNumber = 1 ' first time point
    StartAcquisition BleachingActivated 'This is the main function of the macro
    AFTest6Run = True
End Function


'''''
' AFTest6_Click()
' Perform repeatealy Autofocus with piezo and frame acquisition with piezo at multiposition
' Uses No Z-track and Z-track
''''
Private Sub AFTest7_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    AFTest7Run
    StopAcquisition
End Sub

Private Function AFTest7Run() As Boolean
    Running = True
    Dim RecordingDoc As DsRecordingDoc
    Log = True
    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.name) Then
        MsgBox ("No piezo availabe! Cannot start tests.")
        Exit Function
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No outputfolder selected ! Cannot start tests.")
        Exit Function
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
        Exit Function
    End If
        
    AcquisitionTrack1.Value = AutofocusTrack1.Value
    AcquisitionTrack2.Value = AutofocusTrack2.Value
    AcquisitionTrack3.Value = AutofocusTrack3.Value
    AcquisitionTrack4.Value = AutofocusTrack4.Value
    AutofocusMaxSpeed.Value = True
    AutofocusFastZline = True
    AutofocusHRZ.Value = False
    AutofocusLowZoom.Value = False
        
        
    '''''''
    ' Z-Tracking, Acquistion after Autofocus
    '''''''
    AutofocusTrackZ.Value = True

    MultipleLocationToggle.Value = True
    LoopingRepetitions = AFTest_Repetitions.Value
    LoopingRepetitionTime.Value = 0
    If Not StartSetting() Then
        Exit Function
    End If
    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
    
    GlobalAcquisitionRecording.ScanMode = "Stack"                       'This is defining to acquire a Z stack of Z-Y images
    GlobalAcquisitionRecording.SamplesPerLine = 8  'If doing frame autofocussing it uses the userdefined frame size
    GlobalAcquisitionRecording.LinesPerFrame = 8
    If AutofocusZStep.Value > 0 Then
        GlobalAcquisitionRecording.FramesPerStack = Round(20 / AutofocusZStep.Value)
        GlobalAcquisitionRecording.FrameSpacing = AutofocusZStep.Value
    Else
        GlobalAcquisitionRecording.FramesPerStack = 10
        GlobalAcquisitionRecording.FrameSpacing = 10
    End If
    TextBoxFileName.Value = "FastZline"
    'Set counters back to 1
    RepetitionNumber = 1 ' first time point
    StartAcquisition BleachingActivated 'This is the main function of the macro
    AFTest7Run = True
End Function


Private Sub AFTestAll_Click()
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Running = True
    If Not AFTest1Run Then
        GoTo ScanStop
    End If
    If Not AFTest3Run Then
        GoTo ScanStop
    End If

    If Not AFTest5Run Then
        GoTo ScanStop
    End If
    
    If Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.name) Then
        If Not AFTest2Run Then
            GoTo ScanStop
        End If
        If Not AFTest4Run Then
            GoTo ScanStop
        End If
        If Not AFTest6Run Then
            GoTo ScanStop
        End If
        If Not AFTest7Run Then
            GoTo ScanStop
        End If
    End If
ScanStop:
    ScanStop = True
    StopAcquisition
End Sub


''''
'   RunTestAutofocusButton(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer) As Boolean
'   Using the actual setting for autofocusing function runs AutofocusButton. Save images and logfile on the GlobalDataBaseName directory
'       [RecordingDoc] - A recording where images are overwritten
'       [TestNr]       - Number of the test, this sets the name of the image files and logfiles.
'       [MaxTestRepeats] - Maximal number of tests for each repeat
''''
Private Function RunTestAutofocusButton(RecordingDoc As DsRecordingDoc, ResetPos As Boolean, MaxTestRepeats As Integer, Optional fileName As String = "AutofocusTest", Optional Pause As Integer = 1000) As Boolean

    Dim FilePath As String
    Dim TestRepeats As Integer
    Dim Zold As Double
    Dim pos As Double
    TestRepeats = 1
    LogFileName = GlobalDataBaseName & "\" & fileName & "_Log" & ".txt"
    
    If Log Then
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        LogFile.WriteLine "% Autofocus Test. Repeated AutofocusButton executions. "
        LogFile.WriteLine "% MaxSpeed " & AutofocusMaxSpeed.Value & ", Zoom1 " & AutofocusLowZoom.Value & ", Piezo " & AutofocusHRZ.Value & ", AFTrackZ " & AutofocusTrackZ.Value & _
        ", AFTrackXY " & AutofocusTrackXY.Value & ", FastZLine" & AutofocusFastZline.Value
    End If
    Zold = posTempZ
    While TestRepeats < MaxTestRepeats + 1
        DisplayProgress "Running Test " & fileName & ". Repeat " & TestRepeats & "/" & MaxTestRepeats & ".......", RGB(0, &HC0, 0)
                
        FilePath = GlobalDataBaseName & "\" & fileName & "_" & TestRepeats
        If Log Then
            SafeOpenTextFile LogFileName, LogFile, FileSystem
            LogFile.WriteLine " "

            LogFile.WriteLine "% Save image in file " & FilePath & ".lsm"
            LogFile.Close
        End If
        DoEvents
        Sleep (Pause)
        DoEvents

        If ResetPos Then
            posTempZ = Round(Zold + (1 - 2 * Rnd) * 10, PrecZ)
        End If
        Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller

        DisplayProgress "Autofocus SetupScanWindow", RGB(0, &HC0, 0)
        If RecordingDoc Is Nothing Then
            Set RecordingDoc = Lsm5.NewScanWindow
            While RecordingDoc.IsBusy
                Sleep (100)
                DoEvents
            Wend
        End If
        If Not AutofocusButtonRun(RecordingDoc, GlobalDataBaseName & "\AFimg_" & fileName & "_" & TestRepeats & ".lsm") Then
            Exit Function
        End If
        'save file
        If ActivateTrack(GlobalAcquisitionRecording, "Acquisition") Then
            SaveDsRecordingDoc RecordingDoc, FilePath & ".lsm"
        End If
        TestRepeats = TestRepeats + 1
        If ScanStop Then
            Exit Function
        End If
    Wend
    If Log Then
        LogFile.Close
    End If
    RunTestAutofocusButton = True
End Function

''''
'   RunTestFastZline(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer, pixelDwell As Double, FrameSize As Integer, pause As Integer) As Boolean
'   Using the actual setting for autofocusing function runs AutofocusButton. Save images and logfile on the GlobalDataBaseName directory
'       [RecordingDoc] - A recording where images are overwritten
'       [TestNr]       - Number of the test, this sets the name of the image files and logfiles.
'       [MaxTestRepeats] - Maximal number of tests for each repeat
''''
Private Function RunTestFastZline(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer, Optional pixelDwellfactor As Double = 1, Optional fileName As String = "AutofocusTest", Optional Pause As Integer = 5000) As Boolean

    Dim FilePath As String
    Dim TestRepeats As Integer
    Dim SuccessRecenter As Boolean
    Dim Time As Double
    Dim pos As Double ' position temp variable
    TestRepeats = 1
    LogFileName = GlobalDataBaseName & "\" & fileName & TestNr & ".txt"
    
    If Log Then
        SafeOpenTextFile LogFileName, LogFile, FileSystem
        LogFile.WriteLine "% FastZlineTest " & TestNr & ". Repeated fast Zline executions. PixelDwellfactor: " & pixelDwellfactor & ", LineSize: " & AutofocusLineSize.Value & ", pause: " & Pause
        LogFile.WriteLine "% MaxSpeed " & AutofocusMaxSpeed.Value & ", Zoom1 " & AutofocusLowZoom.Value & ", Piezo " & AutofocusHRZ.Value & ", AFTrackZ " & AutofocusTrackZ.Value & _
        ", AFTrackXY " & AutofocusTrackXY.Value
    End If
    
    While TestRepeats < MaxTestRepeats + 1
        DisplayProgress "Running Test " & TestNr & ". Repeat " & TestRepeats & "/" & MaxTestRepeats & ".......", RGB(0, &HC0, 0)
        FilePath = GlobalDataBaseName & "\" & fileName & TestNr & "_" & TestRepeats
        If Log Then
            SafeOpenTextFile LogFileName, LogFile, FileSystem
            LogFile.WriteLine " "
            LogFile.WriteLine "% Save image in file " & FilePath & ".lsm"
            LogFile.Close
        End If
        DoEvents
        Sleep (Pause)
        DoEvents
        If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
            MsgBox "No track selected for Autofocus! Cannot Autofocus!"
            Exit Function
        End If
        Time = Timer
        Recenter_pre posTempZ, SuccessRecenter, ZEN
        
        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime * pixelDwellfactor
        
        Sleep (Pause)
        DoEvents
        If Log Then
            SafeOpenTextFile LogFileName, LogFile, FileSystem
            Time = Timer - Time
            'pos = Lsm5.Hardware.CpFocus.Position
            LogFile.WriteLine ("% AutofocusButton: center and wait 1st  Z = " & posTempZ & ", Time required " & Time & ", success Recenter " & SuccessRecenter)
'            Sleep (100)
'            If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
'                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ  " & posTempZ & "; obtained Central slide " & pos & "; position " & pos)
'            Else
'                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ  " & posTempZ & "; obtained Central slide " & _
'                Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos & "; position " & pos)
'            End If
            LogFile.Close
        End If
        
        If Not ScanToImage(RecordingDoc) Then
            Exit Function
        End If
        Time = Timer
        Recenter_post posTempZ, SuccessRecenter, ZEN
        DoEvents
        If Log Then
            SafeOpenTextFile LogFileName, LogFile, FileSystem
            Time = Timer - Time
            pos = Lsm5.Hardware.CpFocus.Position
            LogFile.WriteLine ("% AutofocusButton: recenter 1st  Z = " & posTempZ & ", Time required " & Time & ", waiting repeats (max 9) " & Round(Time / 0.4))
            If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ (after img) " & posTempZ & "; obtained Central slide " & Lsm5.Hardware.CpFocus.Position & "; position " & Lsm5.Hardware.CpFocus.Position)
            Else
                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ (after img) " & posTempZ & "; obtained Central slide " & _
                Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position & "; position " & Lsm5.Hardware.CpFocus.Position)
            End If
            LogFile.Close
        End If
        SaveDsRecordingDoc RecordingDoc, FilePath & ".lsm"
        TestRepeats = TestRepeats + 1
        If ScanStop Then
            Exit Function
        End If
    Wend
    If Log Then
        LogFile.Close
    End If
    RunTestFastZline = True
End Function


