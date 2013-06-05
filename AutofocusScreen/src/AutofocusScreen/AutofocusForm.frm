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
Option Explicit 'force to declare all variables

Private shlShell As Shell32.Shell
Private shlFolder As Shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Version Description''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' AutofocusScreen_ZEN_v3.0.0
'''''''''''''''''''''End: Version Description'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Version As String
Public posTempZ  As Double                  'This is position at start after pushing AutofocusButton
Private Const DebugCode = False             'sets key to run tests visible or not
Private Const ReleaseName = True            'this adds the ZEN version
Private Const LogCode = True                'sets key to run tests visible or not





''''''
' UserForm_Initialize()
'   Function called from e.g. AutoFocusForm.Show
'   Load and initialize form
'''''
Public Sub UserForm_Initialize()
    Version = " v3.0.0"
    Dim i As Integer
    'find the version of the software
    Dim VersionNr As String
    VersionNr = Lsm5.Info.VersionIs
    VersionNr = Left(VersionNr, 1)
    Select Case VersionNr
        Case "6":
            ZENv = 2010
        Case "7":
            ZENv = 2011
        Case "8":
            ZENv = 2012
        Case Default:
            MsgBox "Don't understand the verssion of ZEN used. Set to ZEN2010"
    End Select
    If ZENv > 2010 Then
        'On Error GoTo ErrorMsg
        Set ZEN = Lsm5.CreateObject("Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface")
        GoTo NoError
ErrorMsg:
        MsgBox "Version is ZEN" & ZENv & " but can't find Zeiss.Micro.AIM.ApplicationInterface." & vbCrLf _
        & "Using ZEN2010 settings instead." & vbCrLf _
        & "Check if Zeiss.Micro.AIM.ApplicationInterface.dll is registered?" _
        & "See also the manual how to register a dll into windows."
        ZENv = 2010
NoError:
    End If
    'Setting of some global variables
    LogFileNameBase = ""
    Log = LogCode
        
    ' set name of jobs. These must correspond to the names of buttons etc in the form!!
    
    'This variable contains the keys for the OnlineImageanalysis
    ReDim OiaKeyNames(11)
    OiaKeyNames(0) = "code"
    OiaKeyNames(1) = "fileAnalyzed"
    OiaKeyNames(2) = "filePath"
    OiaKeyNames(3) = "X"
    OiaKeyNames(4) = "Y"
    OiaKeyNames(5) = "Z"
    OiaKeyNames(6) = "deltaZ"
    OiaKeyNames(7) = "roiType"
    OiaKeyNames(8) = "roiAim"
    OiaKeyNames(9) = "roiX"
    OiaKeyNames(10) = "roiY"
    OiaKeyNames(11) = "unit"
    
    
    ''''This variable contains all the imagingJobs
    Set Jobs = New ImagingJobs
    ReDim JobNames(4)
    JobNames(0) = "Autofocus"
    JobNames(1) = "Acquisition"
    JobNames(2) = "AlterAcquisition"
    JobNames(3) = "Trigger1"
    JobNames(4) = "Trigger2"
            
    Set JobShortNames = New Collection
    JobShortNames.Add "AF", JobNames(0)
    JobShortNames.Add "AQ", JobNames(1)
    JobShortNames.Add "AL", JobNames(2)
    JobShortNames.Add "TR1", JobNames(3)
    JobShortNames.Add "TR2", JobNames(4)
    
    
    Jobs.initialize JobNames, Lsm5.DsRecording
    Jobs.setZENv ZENv
    


    
    Me.Caption = Me.Caption + Version + " for ZEN "
    
    If ReleaseName Then
        Me.Caption = Me.Caption + CStr(ZENv)
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
    Dim i As Integer
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
    wait (delay)
    LoopingTimerUnit = 1
    GlobalRepetitionSec.BackColor = &HFF8080
    BlockRepetitions = 1
    ReDim Preserve GlobalImageIndex(BlockRepetitions)
     LocationTextLabel.Caption = ""
    GlobalProject = "AutofocusScreen2.1"
    GlobalProjectName = GlobalProject + ".lvb"
    HelpNamePDF = "AutofocusScreen_help.pdf"
    UsedDevices40 bLSM, bLIVE, bCamera
    SystemVersionOffset         ' extra offset depending on macroscope

    ' Set standard values for Autofocus
    ' blSM is a flag to decide weather systen is LSM (ZEN is LSM for instance). LIVE is 5Live not anymore in use?
      
    'TODO: Check if GUI is available (ZEN2011 onward). How do you do this!!
    '
    SetDefaultRecordings
    'Set default value
    SwitchEnablePage AutofocusActive, "Autofocus"
    
    AcquisitionActive.Value = False
    SwitchEnablePage AcquisitionActive, "Acquisition"
    
    
    Set Reps = New ImagingRepetitions
    ReDim RepNames(2)
    RepNames(0) = "Global"    'this is Autofocus Acquisition and AlterAcquisition job
    RepNames(1) = "Trigger1"
    RepNames(2) = "Trigger2"
    
    
    For i = 0 To 2
        Reps.AddRepetition RepNames(i), CDbl(Me.Controls(RepNames(i) + "RepetitionTime")), _
        CInt(Me.Controls(RepNames(i) + "RepetitionNumber")), CBool(Me.Controls(RepNames(i) + "RepetitionInterval"))
    Next i
    
    
    'Set standard values for Looping
    GlobalRepetitionNumber = 300
    GlobalRepetitionTime = 1
    
    SwitchEnablePage False, "Trigger1"
    Trigger1Autofocus.Value = False

    
    'Set standard values for Gridscan
    GridScanActive.Value = False
    SwitchEnableGridScanPage (False)
    
    Set Grids = New ImagingGrids
    Grids.AddGrid "Global", 1, 1, 1, 1
    
    'Set standard values for Additional Acquisition
    AlterAcquisitionActive.Value = False
    SwitchEnablePage AlterAcquisitionActive, "AlterAcquisition"
    
    'Set Database name
    DatabaseTextbox.Value = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="OutputFolder")
    
    'Set repetition and locations
    'RepetitionNumber = 1
    'locationNumber = 1
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

    'AutofocusTrackZ.Visible = DebugCode
    MultiPage1.Pages("TestsPage").Visible = DebugCode
    
    
'    AlterImageInitialize = True
'    ZoomImageInitialize = True
    
    If ZENv = 2010 Then
        ZBacklash = 0.5
    ElseIf ZENv > 2010 Then
        ZBacklash = 0.5
    End If
    
    '''Contains all settings for the repetitions of the jobs
        
    Re_Initialize
End Sub


'''''
'   Re_Initialize()
'   Initializations that need to be performed only when clicking the "Reinitialize" button
'''''
Public Sub Re_Initialize()
    Dim Name As Variant
    Dim delay As Single
    Dim standType As String
    Dim count As Long
    Dim SuccessRecenter As Boolean
    AutoFindTracks
    SwitchEnablePage AutofocusActive, "Autofocus"
    SwitchEnablePage AcquisitionActive, "Acquisition"
    SwitchEnablePage AlterAcquisitionActive, "AlterAcquisition"
    SwitchEnablePage (AcquisitionOiaActive Or AutofocusOiaActive), "Trigger1"
    SwitchEnablePage (AcquisitionOiaActive Or AutofocusOiaActive), "Trigger2"
    
    PubSearchScan = False
    NoReflectionSignal = False
    PubSentStageGrid = False
    
    '  AutofocusForm.Caption = GlobalProject + " for " + SystemName
    BleachingActivated = False
    FocusMapPresent = False
    'This sets standard values for all task we want to do. This will be changed by the macro
    
    If Lsm5.Hardware.CpHrz.Exist(0) Then
        Lsm5.Hardware.CpHrz.Leveling
        While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
    SetDefaultRecordings
    Jobs.initialize JobNames, Lsm5.DsRecording
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Recenter_pre posTempZ, SuccessRecenter, ZENv
    If Not Recenter_post(posTempZ, SuccessRecenter, ZENv) Then
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
    '''UpdateJobs from current form
    For Each Name In JobNames
        UpdateJobFromForm Jobs, CStr(Name)
    Next Name
End Sub



''''
'   ButtonSaveSettings_Click()
'   Open a dialog to save setting of the macro
''''
Private Sub ButtonSaveSettings_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
  
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST
    Filter$ = "Settings (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
            
    
    FileName = CommonDialogAPI.ShowSave(Filter, Flags, "", DatabaseTextbox.Value, "Save AutofocusScreen settings")
    
    If FileName <> "" Then
        If Right(FileName, 4) <> ".ini" Then
            FileName = FileName & ".ini"
        End If
        SaveFormSettings FileName
    End If
    
End Sub

''''
'   ButtonSaveSettings_Click()
'   Open a dialog to save setting of the macro
''''
Private Sub ButtonLoadSettings_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
  
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST
    Filter$ = "Settings (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
            
    'Filter = "ini file (*.ini) |*.ini"
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DatabaseTextbox.Value, "Load AutofocusScreen settings")
    
    If FileName <> "" Then
        LoadFormSettings FileName
    End If
End Sub


'''''''
'   MultipleLocationToggle_Change()
'   Activate MultipleLocation and deactivate SingleLocation
'''''''
Private Sub MultipleLocationToggle_Change()
                
    If MultipleLocationToggle.Value = True Then
 
        If Lsm5.Hardware.CpStages.MarkCount = 0 Then
            MsgBox "To use MultipleLocations you need to define at least one position with the Stage (Not the positions) dialog!"
            MultipleLocationToggle.Value = False
        End If
    End If
    SingleLocationToggle.Value = Not MultipleLocationToggle.Value
    
    
End Sub


'''''''
'   SingleLocationToggle_Change()
'   Activate Singlelocation and deactivate MultipleLocation
'''''''
Private Sub SingleLocationToggle_Change()
    MultipleLocationToggle.Value = Not SingleLocationToggle
End Sub
  

''''
'   FocusMap_Click()
'   create a focusMap using teh Autofocus Channel
''''
Private Sub FocusMap_Click()
    ' This will run just in the AutofocusMode all the AcquisitionTracks are set off
'    SetDatabase
'    SaveFormSettings GlobalDataBaseName & "\tmpSettings.ini"
'    AcquisitionTracksSetOff
'    'change values
'    GlobalRepetitionNumber.Value = 1
'    BlockTimeDelay = 0
'    GlobalRepetitionSec_Click
'    AlterAcquisitionActive.Value = False
'    StartButton_Click
'    WritePosFile GlobalDataBaseName & "\" & TextBoxFileName.Value & "positionsGrid.csv", posGridX, posGridY, posGridZ
'    'Return to original values for the
'    LoadFormSettings GlobalDataBaseName & "\tmpSettings.ini"
End Sub




'''''
' Enable/disable a general set of functions common to all pages
'''''
Private Sub SwitchEnablePage(Enable As Boolean, JobName As String)

    Dim i As Integer
    For i = 1 To 4
        Me.Controls(JobName + "Track" + CStr(i)).Enabled = Enable
    Next i
    If JobName <> "Autofocus" Then
        Me.Controls(JobName + "ZOffset").Enabled = Enable
        Me.Controls(JobName + "ZOffsetLabel").Enabled = Enable
    End If
    
    
    Me.Controls(JobName + "Period").Enabled = Enable
    Me.Controls(JobName + "PeriodLabel").Enabled = Enable
    
    Me.Controls(JobName + "SetJob").Enabled = Enable
    Me.Controls(JobName + "PutJob").Enabled = Enable
    
    'For online image analysis
    If JobName <> "AlterAcquisition" Then
        Me.Controls(JobName + "TrackZ").Enabled = Enable
        Me.Controls(JobName + "TrackXY").Enabled = Enable
        Me.Controls(JobName + "OfflineTrack").Enabled = Enable And (Me.Controls(JobName + "TrackZ") Or Me.Controls(JobName + "TrackXY"))
        Me.Controls(JobName + "OfflineTrackChannel").Enabled = Enable And (Me.Controls(JobName + "TrackZ") Or Me.Controls(JobName + "TrackXY"))
        Me.Controls(JobName + "OiaActive").Enabled = Enable
        If Me.Controls(JobName + "OiaActive") Then
            Me.Controls(JobName + "OiaParallel").Enabled = Enable
            Me.Controls(JobName + "OiaSequential").Enabled = Enable
            Me.Controls(JobName + "OiaOfflineTracking").Enabled = Enable
        Else
            Me.Controls(JobName + "OiaParallel").Enabled = False
            Me.Controls(JobName + "OiaSequential").Enabled = False
        End If
    End If
    
    If JobName = "Autofocus" Then
        AutofocusPeriod.Enabled = Enable
        AutofocusPeriodLabel.Enabled = Enable
        AutofocusTrackZ.Enabled = Enable
        AutofocusTrackXY.Enabled = Enable
        AutofocusSaveImage.Enabled = Enable
    End If

    If JobName = "Trigger1" Or JobName = "Trigger2" Then
        Me.Controls(JobName + "Autofocus").Enabled = Enable
        Me.Controls(JobName + "RepetitionTime").Enabled = Enable
        Me.Controls(JobName + "RepetitionTimeLabel").Enabled = Enable
        Me.Controls(JobName + "RepetitionSec").Enabled = Enable
        Me.Controls(JobName + "RepetitionMin").Enabled = Enable
        Me.Controls(JobName + "RepetitionInterval").Enabled = Enable
        Me.Controls(JobName + "RepetitionNumber").Enabled = Enable
        Me.Controls(JobName + "RepetitionNumberLabel").Enabled = Enable
    End If
    
    Me.Controls(JobName + "Label1").Enabled = Enable
    Me.Controls(JobName + "Label2").Enabled = Enable
    Me.Controls(JobName + "Label1").Caption = Jobs.JobDescriptor1(JobName)
    Me.Controls(JobName + "Label2").Caption = Jobs.JobDescriptor2(JobName)

End Sub



'fills popup menu for chosing a track for post-acquisition tracking
' TODO: move in form
Private Sub FillTrackingChannelList(JobName As String)

    Dim iTrack As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim channel As DsDetectionChannel
    Dim Track As DsTrack
    Dim TrackOn As Boolean
    
    ReDim ActiveChannels(Lsm5.Constants.MaxActiveChannels)  'ActiveChannels is a dynamic array (variable size), ReDim defines array size required next
                                                            'Array size is (MaxActiveChannels gets) the total max number of active channels in all tracks
    Me.Controls(JobName + "OfflineTrackChannel").Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    For iTrack = 0 To Jobs.TrackNumber(JobName) - 1
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(iTrack, Success)
        If Jobs.GetAcquireTrack(JobName, iTrack) Then
            For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
                If Track.DetectionChannelObjectByIndex(c - 1, Success).Acquire Then 'if channel is activated
                    ca = ca + 1 'counter for active channels will increase by one
                    Me.Controls(JobName + "OfflineTrackChannel").AddItem Track.Name & " " & Track.DetectionChannelObjectByIndex(c - 1, Success).Name      'entry is added to combo box to chose track for post-acquisition tracking
                    ActiveChannels(ca) = Track.DetectionChannelObjectByIndex(c - 1, Success).Name & "-T" & Track.MultiplexOrder + 1  'adds entry to ActiveChannel Array with name of channel + name of track
                    TrackOn = True
                End If
            Next c
        End If
    Next iTrack
    
    If TrackOn Then
        Me.Controls(JobName + "OfflineTrackChannel").Value = Me.Controls(JobName + "OfflineTrackChannel").List(0) 'initially displayed text in popup menu is a blank line (first channel is 1)
    End If
End Sub

'''
'   TrackClick(JobName As String, thisTrack As Integer, Exclusive As Boolean)
'       Activate iTrack-th track for a specific JobName
'       If Exclusive all other tracks are inactivated
'''
Private Sub TrackClick(JobName As String, iTrack As Integer, Optional Exclusive As Boolean = False)
    Dim i As Integer
    Dim AutofocusTrackOn As Boolean

    If Me.Controls(JobName + "Track" + CStr(iTrack)).Value Then
        For i = 1 To TrackNumber
            If i <> iTrack And Exclusive Then
                Me.Controls(JobName + "Track" + CStr(i)).Value = Not Me.Controls(JobName + "Track" + CStr(iTrack)).Value
            End If
        Next i
        Jobs.SetAcquireTrack JobName, iTrack - 1, Me.Controls(JobName + "Track" + CStr(iTrack)).Value
        'CheckAutofocusTrack (thisTrack)
    Else
        Jobs.SetAcquireTrack JobName, iTrack - 1, Me.Controls(JobName + "Track" + CStr(iTrack)).Value
    End If
    If JobName <> "AlterAcquisition" Then
        FillTrackingChannelList JobName
    End If
    'ActiveChannels (ComboBoxTrackingChannel.ListIndex + 1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''AutofocusJob'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''
'   AutofocusActive_Click()
'       Activates Autofocus. If not toggled only Acquisition and/or alteracquisition track is used
'''''

Private Sub ActivatePage(JobName As String, Active As Boolean)
    SwitchEnablePage Active, JobName
End Sub

Private Sub AutofocusActive_Click()
    ActivatePage "Autofocus", AutofocusActive
End Sub

'''''
'   AcquisitionActive_Click()
'''''
Private Sub AcquisitionActive_Click()
    ActivatePage "Acquisition", AcquisitionActive
End Sub

'''''
'   AlterAcquisitionActive_Click()
'''''
Private Sub AlterAcquisitionActive_Click()
    SwitchEnablePage AlterAcquisitionActive, "AlterAcquisition"
    'The Acquisition has no time series as default
    Jobs.SetTimeSeries "AlterAcquisition", False

End Sub

''''''
'   Activte Tracks for Jobs (For Autofocus need to be Click as the tracks are exclusive)
''''''
Private Sub AutofocusTrack1_Click()
   TrackClick "Autofocus", 1, True
End Sub

Private Sub AutofocusTrack2_Click()
    TrackClick "Autofocus", 2, True
End Sub

Private Sub AutofocusTrack3_Click()
    TrackClick "Autofocus", 3, True
End Sub

Private Sub AutofocusTrack4_Click()
    TrackClick "Autofocus", 4, True
End Sub

Private Sub AcquisitionTrack1_Change()
   TrackClick "Acquisition", 1
End Sub

Private Sub AcquisitionTrack2_Change()
   TrackClick "Acquisition", 2
End Sub

Private Sub AcquisitionTrack3_Change()
   TrackClick "Acquisition", 3
End Sub

Private Sub AcquisitionTrack4_Change()
   TrackClick "Acquisition", 4
End Sub

Private Sub AlterAcquisitionTrack1_Change()
   TrackClick "AlterAcquisition", 1
End Sub

Private Sub AlterAcquisitionTrack2_Change()
   TrackClick "AlterAcquisition", 2
End Sub

Private Sub AlterAcquisitionTrack3_Change()
   TrackClick "AlterAcquisition", 3
End Sub

Private Sub AlterAcquisitionTrack4_Change()
   TrackClick "AlterAcquisition", 4
End Sub

'''
' ZOffset: This is offset added to current central slice position. This position depends on previous history
''''
Private Sub JobZOffsetChange(JobName As String)
    If Abs(Me.Controls(JobName + "AcquisitionZOffset").Value) > Range() * 0.9 Then
            Me.Controls(JobName + "AcquisitionZOffset").Value = 0
            MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um"
    End If
End Sub

Private Sub AcquisitionZOffset_Change()
    JobZOffsetChange "Acquistion"
End Sub

Private Sub AlterAcquisitionZOffset_Change()
    JobZOffsetChange "AlterAcquistion"
End Sub

Private Sub Trigger1ZOffset_Change()
    JobZOffsetChange "Trigger1"
End Sub

Private Sub Trigger2ZOffset_Change()
    JobZOffsetChange "Trigger2"
End Sub

''''
' TrackZ: If on the Z position will be updated to the latest Z position
''''
Private Sub JobTrackXYZChange(JobName As String)
    Me.Controls(JobName + "OfflineTrackChannel").Enabled = (Me.Controls(JobName + "TrackZ") Or Me.Controls(JobName + "TrackXY")) _
    And Me.Controls(JobName + "OfflineTrack")
    Me.Controls(JobName + "OfflineTrack").Enabled = Me.Controls(JobName + "TrackZ") Or Me.Controls(JobName + "TrackXY")
End Sub

Private Sub AutofocusTrackZ_Change()
    JobTrackXYZChange "Autofocus"
End Sub

Private Sub AcquisitionTrackZ_Change()
    JobTrackXYZChange "Acquisition"
End Sub

Private Sub Trigger1TrackZ_Change()
    JobTrackXYZChange "Trigger1"
End Sub

Private Sub Trigger2TrackZ_Change()
    JobTrackXYZChange "Trigger2"
End Sub

''''
' TrackXY: If on the XY position will be updated to the latest XY position
''''
Private Sub AutofocusTrackXY_Change()
    JobTrackXYZChange "Autofocus"
End Sub

Private Sub AcquisitionTrackXY_Change()
    JobTrackXYZChange "Acquisition"
End Sub

Private Sub Trigger1TrackXY_Change()
    JobTrackXYZChange "Trigger1"
End Sub

Private Sub Trigger2TrackXY_Change()
    JobTrackXYZChange "Trigger2"
End Sub

'''
' If OfflineTrack = True an internal analysis of center of mass is done
'''
Private Sub AutofocusOfflineTrack_Change()
    AutofocusOfflineTrackChannel.Enabled = AutofocusOfflineTrack
End Sub

Private Sub AcquisitionOfflineTrack_Change()
    AcquisitionOfflineTrackChannel.Enabled = AcquisitionOfflineTrack
End Sub

Private Sub Trigger1OfflineTrack_Change()
    Trigger1OfflineTrackChannel.Enabled = Trigger1OfflineTrack
End Sub

Private Sub Trigger2OfflineTrack_Change()
    Trigger2OfflineTrackChannel.Enabled = Trigger2OfflineTrack
End Sub


''''
' Online image analysis. If True then VBAMacro listen to external program (Fiji, Macropilot, Cellprofiler)
''''
Private Sub JobOiaActiveClick(JobName As String)
    If JobName = "Autofocus" Or JobName = "Acquisition" Then
        If AutofocusOiaActive Or AcquisitionOiaActive Then
            SwitchEnablePage True, "Trigger1"
            SwitchEnablePage True, "Trigger2"
        Else
            SwitchEnablePage False, "Trigger1"
            SwitchEnablePage False, "Trigger2"
        End If
    End If
    Me.Controls(JobName + "OiaParallel").Enabled = Me.Controls(JobName + "OiaActive")
    Me.Controls(JobName + "OiaSequential").Enabled = Me.Controls(JobName + "OiaActive")
End Sub

Private Sub AutofocusOiaActive_Click()
    JobOiaActiveClick "Autofocus"
End Sub

Private Sub AcquisitionOiaActive_Click()
    JobOiaActiveClick "Acquisition"
End Sub

Private Sub Trigger1OiaActive_Click()
    JobOiaActiveClick "Trigger1"
End Sub

Private Sub Trigger2OiaActive_Click()
    JobOiaActiveClick "Trigger2"
End Sub

'''''''
'  Sequential online image analysis. VBA Macro waits after acquisition of image for a change in registry code
'''''''
Private Sub AutofocusOiaSequential_Change()
    AutofocusOiaParallel.Value = Not AutofocusOiaSequential.Value
End Sub

Private Sub AcquisitionOiaSequential_Change()
    AcquisitionOiaParallel.Value = Not AcquisitionOiaSequential.Value
End Sub

Private Sub Trigger1OiaSequential_Change()
    Trigger1OiaParallel.Value = Not Trigger1OiaSequential.Value
End Sub

Private Sub Trigger2OiaSequential_Change()
    Trigger2OiaParallel.Value = Not Trigger2OiaSequential.Value
End Sub

'''''''
' Parallel online image analysis. VBA Macro reads before starting job in a text file with name of image file chopped of "_Txxx.lsm"
'''''''
Private Sub AutofocusOiaParallel_Change()
    AutofocusOiaSequential.Value = Not AutofocusOiaParallel.Value
End Sub

Private Sub AcquisitionOiaParallel_Change()
    AcquisitionOiaSequential.Value = Not AcquisitionOiaParallel.Value
End Sub

Private Sub Trigger1OiaParallel_Change()
    Trigger1OiaSequential.Value = Not Trigger1OiaParallel.Value
End Sub

Private Sub Trigger2OiaParallel_Change()
    Trigger2OiaSequential.Value = Not Trigger2OiaParallel.Value
End Sub

'''
' Load settings from ZEN into Form/Joblist
'''
Private Sub JobSetJobClick(JobName As String)
    Jobs.SetJob JobName, Lsm5.DsRecording, ZEN
    UpdateFormFromJob Jobs, JobName
    'If JobName = "Autofocus" Then
        Me.Controls(JobName + "Label1").Caption = Jobs.JobDescriptor1(JobName)
        Me.Controls(JobName + "Label2").Caption = Jobs.JobDescriptor2(JobName)
    'End If
End Sub

Private Sub AutofocusSetJob_Click()
    JobSetJobClick "Autofocus"
End Sub

Private Sub AcquisitionSetJob_Click()
    JobSetJobClick "Acquisition"
End Sub

Private Sub AlterAcquisitionSetJob_Click()
    JobSetJobClick "AlterAcquisition"
End Sub

Private Sub Trigger1SetJob_Click()
    JobSetJobClick "Trigger1"
End Sub

Private Sub Trigger2SetJob_Click()
    JobSetJobClick "Trigger2"
End Sub

'''
' Put settings from Job into ZEN
'''
Private Sub AutofocusPutJob_Click()
    Jobs.PutJob "Autofocus", ZEN
End Sub

Private Sub AcquisitionPutJob_Click()
    Jobs.PutJob "Acquisition", ZEN
End Sub

Private Sub AlterAcquisitionPutJob_Click()
    Jobs.PutJob "AlterAcquisition", ZEN
End Sub

Private Sub Trigger1PutJob_Click()
    Jobs.PutJob "Trigger1", ZEN
End Sub

Private Sub Trigger2PutJob_Click()
    Jobs.PutJob "Trigger2", ZEN
End Sub


'''''
' Looping/RepetitionSettings
'''''
Private Sub RepetitionTime(Name As String)
    If Me.Controls(Name + "RepetitionSec").Value Then
        Reps.setRepetitionTime Name, CDbl(Me.Controls(Name + "RepetitionTime").Value)
    ElseIf Me.Controls(Name + "RepetitionMin").Value Then
        Reps.setRepetitionTime Name, CDbl(Me.Controls(Name + "RepetitionTime").Value * 60)
    End If
End Sub

Private Sub RepetitionMin(Name As String)
    
    'if previously it was in sec divide by 60
    Me.Controls(Name + "RepetitionTime").Value = CDbl(Me.Controls(Name + "RepetitionTime").Value / 60)
    Me.Controls(Name + "RepetitionMin").BackColor = &HFF8080
    Me.Controls(Name + "RepetitionSec").BackColor = &H8000000F
    Me.Controls(Name + "RepetitionSec").Value = Not Me.Controls(Name + "RepetitionMin").Value
    Me.Controls(Name + "RepetitionTime").Max = 60
    RepetitionTime (Name)
End Sub


Private Sub RepetitionSec(Name As String)
    Me.Controls(Name + "RepetitionTime").Max = 360
    Me.Controls(Name + "RepetitionTime").Value = CDbl(Me.Controls(Name + "RepetitionTime").Value * 60)
    Me.Controls(Name + "RepetitionSec").BackColor = &HFF8080
    Me.Controls(Name + "RepetitionMin").BackColor = &H8000000F
    Me.Controls(Name + "RepetitionMin").Value = Not Me.Controls(Name + "RepetitionSec").Value
    RepetitionTime (Name)
End Sub

Public Sub GlobalRepetitionMin_Click()
    If GlobalRepetitionMin Then
        RepetitionMin ("Global")
    End If
End Sub


Private Sub Trigger1RepetitionMin_Click()
    If Trigger1RepetitionMin Then
        RepetitionMin ("Trigger1")
    End If
End Sub

Private Sub Trigger2RepetitionMin_Click()
    If Trigger2RepetitionMin Then
        RepetitionMin ("Trigger2")
    End If
End Sub

Public Sub GlobalRepetitionSec_Click()
    If GlobalRepetitionSec Then
        RepetitionSec ("Global")
    End If
End Sub

Private Sub Trigger1RepetitionSec_Click()
    If Trigger1RepetitionSec Then
        RepetitionSec ("Trigger1")
    End If
End Sub


Private Sub Trigger2RepetitionSec_Click()
    If Trigger2RepetitionSec Then
        RepetitionSec ("Trigger2")
    End If
End Sub

Private Sub GlobalRepetitionTime_Click()
    RepetitionTime ("Global")
End Sub

Private Sub Trigger1RepetitionTime_Click()
    RepetitionTime ("Trigger1")
End Sub

Private Sub Trigger2RepetitionTime_Click()
    RepetitionTime ("Trigger2")
End Sub

Private Sub RepetitionNumber(Name As String)
    Reps.setRepetitionNumber Name, CInt(Me.Controls(Name + "RepetitionNumber"))
End Sub

Private Sub GlobalRepetitionNumber_Change()
    RepetitionNumber ("Global")
End Sub

Private Sub Trigger1RepetitionNumber_Change()
    RepetitionNumber ("Trigger1")
End Sub

Private Sub Trigger2RepetitionNumber_Change()
    RepetitionNumber ("Trigger2")
End Sub

''''
' Set weather Interval or delay
'''
Private Sub RepetitionInterval(Name As String)
    Reps.setInterval "Global", Me.Controls(Name + "RepetitionInterval").Value
End Sub

Private Sub GlobalRepetitionInterval_Click()
    RepetitionInterval "Global"
End Sub

Private Sub Trigger1RepetitionInterval_Click()
    RepetitionInterval "Trigger1"
End Sub


Private Sub Trigger2RepetitionInterval_Click()
    RepetitionInterval "Trigger2"
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


''''
' GridScanActive_Click()
'   Set the grid scan on or off. Changes also
''
Private Sub GridScanActive_Click()
    SwitchEnableGridScanPage (GridScanActive.Value)
    If GridScanActive Then
        If MultipleLocationToggle.Value Then
            GridScan_nRow = 1
            GridScan_nColumn = Lsm5.Hardware.CpStages.MarkCount
            Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
        End If
    End If
End Sub


Private Sub GridScan_nRow_Click()
     Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
End Sub

Private Sub GridScan_nColumn_Click()
     Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
End Sub

Private Sub GridScan_nColumnSub_Click()
     Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
End Sub

Private Sub GridScan_nRowSub_Click()
     Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
End Sub


''''
'   SwitchEnableGridScanPage(Enable As Boolean)
'   Disable or enable all buttons and slider
'       [Enable] In - Sets the mini page enable status
''''
Public Sub SwitchEnableGridScanPage(Enable As Boolean)

    GridScan_validGridDefault.Enabled = Enable
    GridScan_posLabel.Enabled = Enable
    GridScan_nColumnLabel.Enabled = Enable And Not MultipleLocationToggle
    GridScan_nRowLabel.Enabled = Enable And Not MultipleLocationToggle
    GridScan_nColumn.Enabled = Enable And Not MultipleLocationToggle
    GridScan_nRow.Enabled = Enable And Not MultipleLocationToggle
    GridScan_dColumnLabel.Enabled = Enable And Not MultipleLocationToggle
    GridScan_dRowLabel.Enabled = Enable And Not MultipleLocationToggle
    GridScan_dColumn.Enabled = Enable And Not MultipleLocationToggle
    GridScan_dRow.Enabled = Enable And Not MultipleLocationToggle
    GridScan_refColumn.Enabled = Enable And Not MultipleLocationToggle
    GridScan_refRow.Enabled = Enable And Not MultipleLocationToggle
    GridScan_refColumnLabel.Enabled = Enable And Not MultipleLocationToggle
    GridScan_refRowLabel.Enabled = Enable And Not MultipleLocationToggle
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


'''''
'   Open a Dialog to set file name for storage of valid positions
'''''
Private Sub GridScanValidFileButton_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
  
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
            
    Filter = "Alle Dateien (*.*)" & Chr$(0) & "*.*"
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", "", "Select file containing valid grid positions")
    
    If Right(FileName, 3) <> "*.*" Then
        GridScanValidFile.Value = FileName
    Else
        GridScanValidFile.Value = ""
    End If
    
End Sub

'''''
'   Open a dialog to set filename where positions of grid are stored
'''''
Private Sub GridScanPositionFileButton_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
  
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
            
    Filter = "Alle Dateien (*.*)" & Chr$(0) & "*.*"
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", "", "Select file containing positions of grid")
       
    If Right(FileName, 3) <> "*.*" Then
        GridScanPositionFile.Value = FileName
    Else
        GridScanPositionFile.Value = ""
    End If
End Sub




''''
' Stop all jobs after current repetition of current job
''''
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
    StopAfterRepetition.Value = False
    StopAfterRepetition.BackColor = &H8000000F
    StopButton.BackColor = &H8000000F
    StopButton.Value = False
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


'''
' Pause a job
''''
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






'''''
'   CommandButtonNewDataBase_Click()
'   Open a Dialog to set output folder where to save the results. then cal SetDatabase to set global variables
'''''
Private Sub CommandButtonNewDataBase_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
  
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
            
    Filter = "Alle Dateien (*.*)" & Chr$(0) & "*.*"
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", "", "Select output folder")
    
    If Len(FileName) > 3 Then
        FileName = Left(FileName, Len(FileName) - 3)
        DatabaseTextbox.Value = FileName
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
        DatabaseLabel.Caption = "No output folder"
    End If

    If Not GlobalDataBaseName = "" Then
        If Right(GlobalDataBaseName, 1) <> "\" Then
            DatabaseTextbox.Value = DatabaseTextbox.Value + "\"
            GlobalDataBaseName = DatabaseTextbox.Value
        End If
        On Error GoTo ErrorHandleDataBase
        If Not CheckDir(GlobalDataBaseName) Then
            Exit Sub
        End If
        DatabaseLabel.Caption = GlobalDataBaseName
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
'   RestoreAcquisitionParameters()
'   Restores the image acquisition recording parameters from GlobalBackupRecording
'   recenter acquisition
'   Lsm5.DsRecording Out - Recording settings
''''''
Public Sub RestoreAcquisitionParameters()
    Dim i As Integer
    Dim pos As Double
    Dim time As Double
    Dim LogMsg As String
    Dim SuccessRecenter As Boolean
    
    time = Timer
    Lsm5.DsRecording.Copy GlobalBackupRecording
    Lsm5.DsRecording.FrameSpacing = GlobalBackupRecording.FrameSpacing
    Lsm5.DsRecording.FramesPerStack = GlobalBackupRecording.FramesPerStack
    For i = 0 To Lsm5.DsRecording.TrackCount - 1
       Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, 1).Acquire = GlobalBackupActiveTracks(i)
    Next i
    time = Round(Timer - time, 2)
    LogMsg = "% Restore settings: time to return to backuprecording " & time
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    
    Sleep (1000)
 
    time = Timer
    Recenter_pre posTempZ, SuccessRecenter, ZENv
    pos = Lsm5.Hardware.CpFocus.Position
    'move to posTempZ
    If ZENv = "2011" Or ZENv = "2010" Then
        If Round(pos, PrecZ) <> Round(posTempZ, PrecZ) Then
            If Not FailSafeMoveStageZ(posTempZ) Then
                Exit Sub
            End If
        End If
        Recenter_post posTempZ, SuccessRecenter, ZENv
    End If
    time = Round(Timer - time, 2)
    LogMsg = "% Restore settings: recenter Z " & posTempZ & ", Time required " & time & ", success within rep. " & SuccessRecenter & vbCrLf
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
'    Dim Y As Double
'    Dim Z As Double
'    Dim deltaZ As Double
'    Dim GridPos As GridPosType
'    Dim time As Double
'    Dim pos As Double
'    Dim LogMsg As String
'    Dim SuccessRecenter As Boolean
'    Running = True
'    Dim NewPicture As DsRecordingDoc
'    DisplayProgress "Get Current Position Offset - Autofocus", RGB(0, &HC0, 0)             'Gives information to the user
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    Z = posTempZ
'    X = Lsm5.Hardware.CpStages.PositionX
'    Y = Lsm5.Hardware.CpStages.PositionY
'
'    'recenter only after activation of new track
'    If AutofocusActive Then
'        StopScanCheck
'        If AutofocusHRZ Then
'            Lsm5.Hardware.CpHrz.Leveling
'        End If
'       'FailSafeMoveStageZ (posTempZ) 'move at position
'        ' Acquire image and calculate center of mass stored in XMass, YMass and ZMass
'        DisplayProgress "Autofocus Activate Tracks", RGB(0, &HC0, 0)
'        time = Timer
'        If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
'            MsgBox "No track selected for Autofocus! Cannot Autofocus!"
'            Exit Function
'        End If
'
'        LogMessage "% Get current position: time activate AF track " & Round(Timer - time), Log, LogFileName, LogFile, FileSystem
'
'        'DoEvents
'        'Sample0Z = Lsm5.DsRecording.Sample0Z
'        DisplayProgress "Autofocus: Recenter prior AF acquisition.... ", RGB(0, &HC0, 0)
'        DoEvents
'        time = Timer
'        If Not Recenter_pre(posTempZ, SuccessRecenter, ZENv) Then
'            Exit Function
'        End If
'        pos = Lsm5.Hardware.CpFocus.Position
'        time = Round(Timer - time, 2)
'        LogMsg = "% Get current position: center Z (pre AFimg) " & posTempZ & ", time required" & time & ", succes within rep. " & SuccessRecenter
'        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
'        'Use internal agorithm to compute Xmass etc.
'        If Not MicroscopeIO.Autofocus_StackShift(NewPicture) Then
'                Exit Function
'        End If
'
'        DisplayProgress "Autofocus compute", RGB(0, &HC0, 0)
'
'        If Not ComputeNewCoordinatesAfterAF(NewPicture, X, Y, Z, deltaZ, "Autofocus") Then
'            Exit Function
'        End If
'        AcquisitionZOffset.Value = posTempZ - Z
'
'        DisplayProgress "Autofocus: Recenter after AF acquisition...", RGB(0, &HC0, 0)
'
'        time = Timer
'        If Not Recenter_post(posTempZ, SuccessRecenter, ZENv) Then
'            Exit Function
'        End If
'        time = Round(Timer - time, 2)
'        LogMsg = "% Get current position: recenter Z (post AFImg) " & posTempZ
'        If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
'                LogMsg = LogMsg & "; obtained central slide " & pos & "; position " & pos & ", time required " & time & ", succes within rep. " & SuccessRecenter
'        Else
'            LogMsg = LogMsg & "; obtained central slide " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 _
'            - Lsm5.DsRecording.Sample0Z + pos & "; position " & pos & ", time required " & time & ", succes within rep. " & SuccessRecenter
'        End If
'        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
'
'        posTempZ = Z
'        time = Timer
'        If Not Recenter_pre(posTempZ, SuccessRecenter, ZENv) Then
'            Exit Function
'        End If
'        time = Round(Timer - time, 2)
'        LogMsg = "% Get current position: center Z (end) " & posTempZ & ", time required" & time & ", success" & SuccessRecenter
'        LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
'    End If
'    GetCurrentPositionOffsetButtonRun = True
End Function

'''''''
'   AutofocusButton_Click()
'   calls AutofocusButtonRun
''''''''
Public Sub AutofocusButton_Click()
    Dim node As AimExperimentTreeNode
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Dim RecordingDoc As DsRecordingDoc
    Dim SuccessRecenter As Boolean
    Running = True
    posTempZ = Lsm5.Hardware.CpFocus.Position
    Recenter_pre posTempZ, SuccessRecenter, ZENv
    Set GlobalAutoFocusRecording = Lsm5.CreateBackupRecording
    Set GlobalAcquisitionRecording = Lsm5.CreateBackupRecording
    Set GlobalTrigger1Recording = Lsm5.CreateBackupRecording
    Set GlobalBleachRecording = Lsm5.CreateBackupRecording
    Set GlobalAltRecording = Lsm5.CreateBackupRecording
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    GlobalAutoFocusRecording.Copy Lsm5.DsRecording
    GlobalAcquisitionRecording.Copy Lsm5.DsRecording
    GlobalTrigger1Recording.Copy Lsm5.DsRecording
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
    
    'Check if there is an existing document then start acquisition
    Set node = viewerGuiServer.ExperimentTreeNodeSelected
    If Not node Is Nothing Then
        If node.Type <> eExperimentTeeeNodeTypeLsm Then
            Lsm5.NewScanWindow
        End If
        Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
    End If
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
    Dim OiaSettings As Dictionary
    Dim FileName As String
    Dim time As Double
    Dim X As Double, Y As Double, Z As Double
    Dim NewCoord() As Double
    Dim deltaZ As Double
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
    

    'initialize offsets. Default is no offset
    ResetRegistry
    ReadOiaSettingsFromRegistry OiaSettings, OiaKeyNames
    FileName = "AF_T000.lsm"


    'recenter only after activation of new track
    If AutofocusActive Then
        OiaJobInitialize "Autofocus", OiaSettings, FilePath, FileName
        ExecuteJob "Autofocus", AutofocusDoc, FilePath, FileName, X, Y, Z, CInt(deltaZ)
    End If
    If AcquisitionActive Then
        FileName = "AQ_T000.lsm"
        OiaJobInitialize "Acquisition", OiaSettings, FilePath, FileName
        ExecuteJob "Acquisition", AutofocusDoc, FilePath, FileName, X, Y, Z + AcquisitionZOffset.Value, CInt(deltaZ)
    End If
    
    If AlterAcquisitionActive Then
        FileName = "AL_T000.lsm"
        OiaJobInitialize "AlterAcquisition", OiaSettings, FilePath, FileName
        ExecuteJob "AlterAcquisition", AutofocusDoc, FilePath, FileName, X, Y, Z + AlterAcquisitionZOffset.Value, CInt(deltaZ)
    End If
    
    posTempZ = Z
    AutofocusButtonRun = True

End Function




'''''
'   StartButton_Click()
'''''
Private Sub StartButton_Click()

    If Not StartSetting() Then
        ScanStop = True
        StopAcquisition
        Exit Sub
    End If
    
    StartJobOnGrid "Global", "Global", GlobalDataBaseName 'This is the main function of the macro
End Sub


''''''
'   StartSetting()
'   Setups and controls before start of experiment
'       Create list of positions for Grid or Multiposition
''''''
Private Function StartSetting() As Boolean
    Dim i As Integer
    Dim initPos As Boolean   'if False and gridsize correspond positions are taken from file positionsGrid.csv
    Dim SuccessRecenter As Boolean
    Dim X() As Double
    Dim Y() As Double
    Dim Z() As Double
    
    
    initPos = True
    StartSetting = False
    BleachingActivated = False
    AutomaticBleaching = False                                  'We do not do FRAps or FLIPS in this case. Bleaches can still be done with the "ExtraBleach" button.
    Set FileSystem = New FileSystemObject
    
    Dim MarkCount As Long
    MarkCount = Lsm5.Hardware.CpStages.MarkCount
    
    If MultipleLocationToggle.Value And MarkCount < 1 Then
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
            'On Error GoTo ErrorHandleLogFile
            LogFileName = LogFileNameBase
            Close
            SafeOpenTextFile LogFileName, LogFile, FileSystem
            LogFile.WriteLine "% ZEN software version " & ZENv & " " & Version
            LogFile.Close
            Log = True
        Else
            Log = False
        End If
    End If

    If Not AcquisitionTracksOn And Not AutofocusActive And Not AlterAcquisitionActive Then
        MsgBox ("Nothing to do! Check at least one imaging option!")
        Exit Function
    End If
    
    ' do not log if logfilename has not been defined
    If LogCode And LogFileName = "" Then
        Log = False
    End If
    'As default we do not overwrite files
    OverwriteFiles = False
    
    If MultipleLocationToggle Then
        If GridScanActive Then
            If MarkCount = 0 Then  ' No marked position
                MsgBox "GridScan: Use stage to Mark at least the initial position "
                ScanStop = True
                StopAcquisition
                Exit Function
            End If
            GridScan_nColumn.Value = MarkCount
            GridScan_nRow.Value = 1
            Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
        Else
            Grids.updateGrid "Global", 1, MarkCount, 1, 1
        End If
    End If
    
    If SingleLocationToggle Then
        If GridScanActive Then
            If MarkCount = 0 Then  ' No marked position
                MsgBox "GridScan: Use stage to Mark at least the initial position "
                ScanStop = True
                StopAcquisition
                Exit Function
            End If
            Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
        Else
            Grids.updateGrid "Global", 1, 1, 1, 1
        End If
    End If
    
    If SingleLocationToggle And Not GridScanActive Then
        Grids.setX "Global", Lsm5.Hardware.CpStages.PositionX, 1, 1, 1, 1
        Grids.setY "Global", Lsm5.Hardware.CpStages.PositionY, 1, 1, 1, 1
        Grids.setZ "Global", Lsm5.Hardware.CpFocus.Position, 1, 1, 1, 1
        Grids.setValid "Global", True, 1, 1, 1, 1
        GoTo GridReady
    End If
        
    
    If GridScan_nColumn.Value * GridScan_nRow.Value * GridScan_nColumnsub.Value * GridScan_nRowsub.Value > 10000 Then
        MsgBox "GridScan: Maximal number of locations is 10000. Please change Numbers  X and/or Y."
        ScanStop = True
        StopAcquisition
        Exit Function
    End If
    
    
    If GridScanPositionFile <> "" Then
        If Grids.isPositionGridFile("Global", GridScanPositionFile, Grids.numRow("Global"), Grids.numCol("Global"), _
        Grids.numRowSub("Global"), Grids.numCol("Global")) Then
            If Grids.loadPositionGridFile("Global", GridScanPositionFile) Then
                DisplayProgress "Loading grid positions from file. " & GridScanPositionFile & "....", RGB(0, &HC0, 0)
                initPos = False
            Else
                MsgBox "Not able to use " & GridScanPositionFile & ". Resetting the positions."
            End If
        End If
    End If
        
    If initPos Then
            DisplayProgress "Initialize all positions....", RGB(0, &HC0, 0)
            If MultipleLocationToggle.Value Then
                ReDim X(MarkCount - 1)
                ReDim Y(MarkCount - 1)
                ReDim Z(MarkCount - 1)
                For i = 0 To MarkCount - 1
                    Lsm5.Hardware.CpStages.MarkGetZ i, X(i), Y(i), Z(i)
                Next i
                Grids.makeGridFromManyPts "Global", X, Y, Z, GridScan_dColumnsub.Value, GridScan_dRowsub.Value
            Else
                Lsm5.Hardware.CpStages.MarkGetZ 0, XStart, YStart, ZStart
                Grids.updateGrid "Global", GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub
                Grids.makeGridFromOnePt "Global", XStart, YStart, ZStart, GridScan_dColumnsub.Value, GridScan_dRowsub.Value, _
                GridScan_refColumn.Value, GridScan_refRow.Value
                DisplayProgress "Initialize all grid positions...DONE", RGB(0, &HC0, 0)
            End If
    End If
        
    
    If GridScanValidFile <> "" Then
        Dim FormatValidFile As String
        FormatValidFile = Grids.isValidGridFile("Global", GridScanValidFile, GridScan_nRow, GridScan_nColumn, GridScan_nRowsub, GridScan_nColumnsub)
        If Grids.loadValidGridFile(Name, GridScanValidFile, FormatValidFile) Then
            
        Else
            MsgBox "Not able to use " & GridScanValidFile & " for loading valid positions."
        End If
    End If

GridReady:
    Grids.writePositionGridFile "Global", GlobalDataBaseName & "positionsGrid.csv"
    Grids.writeValidGridFile "Global", GlobalDataBaseName & "validGrid.csv"

    'SaveSettings
    If GlobalDataBaseName <> "" Then
        SetDatabase
        SaveFormSettings GlobalDataBaseName & "\AutofocusScreen.ini"
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



    
Private Function SetLocationTextLabel(X As Double, Y As Double, Z As Double, GridPos As GridPosType, TotPos As Long, iRep As Integer) As String
    Dim Caption As String
    Caption = "Locations : " & TotPos & "/" & UBound(posGridX, 1) * UBound(posGridX, 2) * UBound(posGridX, 3) * UBound(posGridX, 4) & "; X= " & X & ", Y = " & Y & ", Z = " & Z & vbCrLf & _
                                "Repetition :" & iRep & "/" & GlobalRepetitionNumber.Value & vbCrLf & _
                                "Well/Position Row: " & GridPos.Row & "/" & UBound(posGridX, 1) & "; Column: " & GridPos.Col & "/" & UBound(posGridX, 2) & vbCrLf

                                
     If MultipleLocationToggle Or GridScanActive Then
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





  

'''''
'   AutoFindTracks()
'   Set the names of the tracks and find possible tracks
'''''
Public Sub AutoFindTracks()

    Dim i, j As Integer
    Dim ChannelOK As Boolean
    Dim MaxTracks As Integer
    Dim iTrack As Integer
    Dim Name As Variant
    Dim ActiveJobs As Collection
    Dim Active() As Boolean
    Set ActiveJobs = New Collection

    
    For Each Name In JobNames
        ReDim Active(3)
        For i = 1 To 4
            Active(i - 1) = Me.Controls(Name + "Track" + CStr(i)).Value
            Me.Controls(Name + "Track" + CStr(i)).Visible = False
            Me.Controls(Name + "Track" + CStr(i)).Value = False
        Next i
        ActiveJobs.Add Active, Name
    Next Name

    
    'The next line and the following "if" should be removed when working with the LSM 2.8 software (where the lambda mode is not defined)
    Set Track = Lsm5.DsRecording.TrackObjectLambda(Success)
    If Success Then
        If Track.Acquire Then
            MsgBox ("This macro does not work in the Lambda Mode. Please switch to the Channel Mode and reinitialize the Macro.")
            Exit Sub
        End If
    End If
    
    'ConfiguredTracks = Lsm5.DsRecording.TrackCount
    MaxTracks = Lsm5.DsRecording.GetNormalTrackCount
    If MaxTracks > 4 Then
        MsgBox ("This Macro only accepts 4 different tracks")
    End If

    iTrack = 1
    For i = 0 To MaxTracks - 1
        If iTrack < 4 Then
            ChannelOK = False
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i, Success)
            For j = 0 To Track.DataChannelCount - 1
                If Track.DataChannelObjectByIndex(j, Success).Acquire = True Then
                    ChannelOK = True
                End If
            Next j
            If ChannelOK And (Not Track.IsLambdaTrack) And (Not Track.IsBleachTrack) Then
                For Each Name In JobNames
                    Me.Controls(Name + "Track" + CStr(iTrack)).Visible = True
                    Me.Controls(Name + "Track" + CStr(iTrack)).Value = ActiveJobs(Name)(i)
                    Me.Controls(Name + "Track" + CStr(iTrack)).Caption = Track.Name
                Next Name
                iTrack = iTrack + 1
            End If
        End If
    Next i
        
    If iTrack < 5 Then
        TrackNumber = iTrack - 1
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






''fills popup menu for chosing a track for post-acquisition tracking
'' TODO: move in form
'Private Sub FillTrackingChannelList()
'    Dim T As Integer
'    Dim c As Integer
'    Dim ca As Integer
'    Dim channel As DsDetectionChannel
'    Dim Track As DsTrack
'
'    ReDim ActiveChannels(Lsm5.Constants.MaxActiveChannels)  'ActiveChannels is a dynamic array (variable size), ReDim defines array size required next
'                                                            'Array size is (MaxActiveChannels gets) the total max number of active channels in all tracks
'    ComboBoxTrackingChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
'    ca = 0
'
'    If ActivateTrack(GlobalAcquisitionRecording, "Acquisition") Then
'        For T = 1 To TrackNumber 'This loop goes through all tracks and will collect all activated channels to display them in popup menu
'            Set Track = GlobalAcquisitionRecording.TrackObjectByMultiplexOrder(T - 1, Success)
'            If Track.Acquire Then 'if track is activated for acquisition
'                For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
'                    Set channel = Track.DetectionChannelObjectByIndex(c - 1, Success)
'                    If channel.Acquire Then 'if channel is activated
'                        ca = ca + 1 'counter for active channels will increase by one
'                        ComboBoxTrackingChannel.AddItem Track.Name & " " & channel.Name 'entry is added to combo box to chose track for post-acquisition tracking
'                        ActiveChannels(ca) = channel.Name & "-T" & Track.MultiplexOrder + 1  'adds entry to ActiveChannel Array with name of channel + name of track
'                    End If
'                Next c
'            End If
'        Next T
'        ComboBoxTrackingChannel.Value = ComboBoxTrackingChannel.List(0) 'initially displayed text in popup menu is a blank line (first channel is 1).
'    End If
'End Sub





Private Sub TextBoxFileName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'this is the enter key
        SetFileName
    End If
End Sub

Private Sub SetFileName()
    If TextBoxFileName <> "" Then
        If Right(TextBoxFileName, 1) <> "_" Then
            TextBoxFileName = TextBoxFileName & "_"
        End If
    End If
End Sub


 
'''''''

' TODO a long does it wait
'Wait time in sec?
Sub wait(PauseTime As Single)
    Dim Start As Single
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
       DoEvents    ' Yield to other processes.
       'Lsm5.DsRecording.StartScanTriggerIn
    Loop
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
            If ActiveChannelNumber = 1 Then AutofocusChannel = DataChannel.Name 'Gets the name of the first activated channel
        End If
    Next
    
    If ActiveChannelNumber > 1 Then 'if more than one channel is activated...
        MsgBox ("The track you selected has more than one active Channel. " & AutofocusChannel & " will be used to calculate autofocus parameters.")
    End If
End Sub







'''''
'   ChangeButtonStatus(Enable As Boolean)
'   Reset status of buttons on rightside of form
'''''
Private Sub ChangeButtonStatus(Enable As Boolean)
    StartButton.Enabled = Enable
    CloseButton.Enabled = Enable
    ReinitializeButton.Enabled = Enable
End Sub


'''''
' Sub StopScanCheck()
' This stop all running scans. Check is the wrong name
'''''
Public Sub StopScanCheck()
    Lsm5.StopScan
    DoEvents
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
FilePath As String, Name As String) As Boolean
    
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

    RecordingDoc.SetTitle Name
    
    If Not SaveDsRecordingDoc(RecordingDoc, FilePath) Then
        ScanStop = True
        Exit Function
    End If
    StartAlternativeImaging = True
End Function









''''''
''   CheckZRanges()
''   Check if Z movements are in agreement with range of microscope
''''''
'Public Function CheckZRanges() As Boolean
'    If ScanStop Then
'        Exit Function
'    End If
'
'    If Range() = 0 Then
'        MsgBox "Objective's working distance not defined! Cannot Autofocus!"
'        CheckZRanges = False
'        Exit Function
'    Else
'        CheckZRanges = True
'    End If
'
'    If AutofocusZRange.Value > Range() * 0.9 Then 'this is already tested in the slider could be removed
'        AutofocusForm.AutofocusZRange.Value = Range() * 0.9
'        MsgBox "Autofocus range is too large! Has been reduced to " + Str(AutofocusForm.AutofocusZRange.Value)
'    End If
'
''    If Abs(AcquisitionZOffset.Value) > Range() * 0.9 Then 'this is already tested in the slider could be removed
''        AutofocusForm.AcquisitionZOffset = 0
''        MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um. Has been put back to " + Str(AutofocusForm.AutofocusZOffset)
''    End If
'
'End Function
  




'''''''''
''   CommandButtonHelp_Click()
''   Look for Help file
''   TODO: Test
'''''''''
'Private Sub CommandButtonHelp_Click()
'
'    Dim dblTask As Double
'    Dim MacroPath As String
'    Dim MyPath As String
'    Dim MyPathPDF As String
'
'    Dim bslash As String
'    Dim Success As Integer
'    Dim pos As Integer
'    Dim Start As Integer
'    Dim count As Long
'    Dim ProjName As String
'    Dim indx As Integer
'    Dim AcrobatObject As Object
'    Dim AcrobatViewer As Object
'    Dim OK As Boolean
'    Dim StrPath As String
'    Dim ExecName As String
'
'    count = ProjectCount()
'    For indx = 0 To count - 1
'        MacroPath = ProjectPath(indx, Success)
'        ProjName = ProjectTitle(indx, Success)
'        If StrComp(ProjName, GlobalProjectName, vbTextCompare) = 0 Then
'            Start = 1
'            bslash = "\"
'            pos = Start
'            Do While pos > 0
'                pos = InStr(Start, MacroPath, bslash)
'                If pos > 0 Then
'                    Start = pos + 1
'                End If
'            Loop
'            MyPath = Strings.Left(MacroPath, Start - 1)
'            MyPathPDF = MyPath + HelpNamePDF
'
'            OK = False
'            On Error GoTo RTFhelp
'            OK = FServerFromDescription("AcroExch.Document", StrPath, ExecName)
'            dblTask = Shell(ExecName + " " + MyPathPDF, vbNormalFocus)
'
'RTFhelp:
'            If Not OK Then
'                MsgBox "Install Acrobat Viewer!"
'            End If
'            Exit For
'        End If
'    Next indx
'End Sub


'
'''''''
''   BleachRegion(XShift As Double, YShift As Double)
''       [XShift] In - Shifts origin of x
''       [YShift] In - Shifts origin of y
''   Todo: function is never been used and does not belong to form or being called. Check it
'''''''
'Private Sub BleachRegion(XShift As Double, YShift As Double)
'    Dim RecordingDoc As DsRecordingDoc
'    Dim Recording As DsRecording
'    Dim Track As DsTrack
'    Dim Laser As DsLaser
'    Dim DetectionChannel As DsDetectionChannel
'    Dim IlluminationChannel As DsIlluminationChannel
'    Dim DataChannel As DsDataChannel
'    Dim BeamSplitter As DsBeamSplitter
'    Dim Timers As DsTimers
'    Dim Markers As DsMarkers
'    Dim Success As Integer
'    Set Recording = Lsm5.DsRecording
'    Dim SampleObservationTime As Double
'    Dim SampleOX As Double
'    Dim SampleOY As Double
'
'
'    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
'
'    SampleOX = Recording.Sample0X
'    SampleOY = Recording.Sample0Y
'    Recording.Sample0X = XShift
'    Recording.Sample0Y = YShift
'    'x = Lsm5.Hardware.CpStages.PositionX - XShift
'    'y = Lsm5.Hardware.CpStages.PositionY - YShift
'    'Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
'    ' maybe wait here till it is finished moving
'    Recording.SpecialScanMode = "NoSpecialMode"
'    Recording.ScanMode = "Point"
'    Recording.TimeSeries = True
'    Recording.FramesPerStack = 1
'    Recording.StacksPerRecord = 50  ' timepoints x 1000
'    SampleObservationTime = Track.SampleObservationTime
'    MsgBox "SampleObservationTime = " + CStr(SampleObservationTime)
'    Track.SampleObservationTime = 0.0001 ' pixel-dwell time in seconds
'    Track.TimeBetweenStacks = 0#
'    'Timers.TimeInterval = 0#
'
'    TakeImage
'
'    Recording.Sample0X = SampleOX
'    Recording.Sample0Y = SampleOY
'    'x = Lsm5.Hardware.CpStages.PositionX + XShift
'    'y = Lsm5.Hardware.CpStages.PositionY + YShift
'    'Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
'    ' maybe wait here till it is finished moving
'    Recording.SpecialScanMode = "NoSpecialMode"
'    Recording.ScanMode = "Frame"
'    Recording.TimeSeries = False
'    Recording.FramesPerStack = 1
'    Recording.StacksPerRecord = 1  ' timepoints x 1000
'    Track.SampleObservationTime = SampleObservationTime ' pixel-dwell time in seconds
'    MsgBox "SampleObservationTime = " + CStr(SampleObservationTime)
'
'
'    'Recording.ScanMode = "Plane"
'    'Recording.FrameSpacing = 0.636243
'
'
'End Sub


'''''''
''   TakeImage()
''   Acquire an image. Allow to stop acquisition and displaqy progress. Nt used anymore
''''''''
'Private Sub TakeImage()
'
'    Dim ScanImage As DsRecordingDoc
'
'    Set ScanImage = Lsm5.StartScan
'
'    DisplayProgress "Taking Image.......", RGB(0, &HC0, 0)
'    Do While ScanImage.IsBusy ' Waiting until the image acquisition is done
'        Sleep (100)
'        If GetInputState() <> 0 Then
'            DoEvents
'            If ScanStop Then
'                StopAcquisition
'                Exit Sub
'            End If
'        End If
'    Loop
'    DisplayProgress "Taking Image...DONE.", RGB(0, &HC0, 0)
'End Sub
'

'Private Sub StartBleachButton_Click()
'
'    Dim Success As Integer
'    Dim nt As Integer
'
'    BleachingActivated = True
'    AutomaticBleaching = False
'
'    If TrackingToggle And TrackingChannelString = "" Then
'        MsgBox ("Select a channel for tracking, or uncheck the tracking button")
'        Exit Sub
'    End If
'    If MultipleLocationToggle.Value And Lsm5.Hardware.CpStages.MarkCount < 1 Then
'        MsgBox ("Select at least one location in the stage control window, or uncheck the multiple location button")
'        Exit Sub
'    End If
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No Output Folder selected ! Cannot start acquisition.")
'        Exit Sub
'    End If
'
'
'    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
'
'    If Success Then
'        If Track.BleachPositionZ <> 0 Then
'            MsgBox ("This macro does not enable to bleach at a different Z position. Please uncheck the corresponding box in the Bleach Control Window")
'            Exit Sub
'        End If
'
'        If Lsm5.IsValidBleachRoi Then
'
'            If ActiveMicropilot Then
'                nt = MicropilotRepetitions
'            Else
'                nt = BlockRepetitions
'            End If
'
'            If (Track.BleachScanNumber + 1) > nt Then
'                MsgBox ("Not enough repetitions to bleach; either increase the Number of Acquisitions, or, when using MicroPilot, the Cycles")
'                Exit Sub
'            End If
'
'            FillBleachTable
'            AutomaticBleaching = True
'           'Track.UseBleachParameters = True ' deleted 20100818 , can probably not work with Micropilot
'        Else
'            MsgBox ("A bleaching ROI needs to be defined to start the macro in the bleaching mode")
'            Exit Sub
'        End If
'    Else
'        MsgBox ("A bleach track needs to be defined to start the macro in the bleaching mode")
'        Exit Sub
'    End If
'
'    StartAcquisition BleachingActivated
'
'End Sub

'Private Sub FillBleachTable()  'Fills a table for the macro to know when the bleaches have to occur. This works for FRAPs (and FLIPS if working with LSM 3.2)
'
'    Dim i As Integer
'    Dim nt As Integer
'    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
'    If Success Then
'
'        If ActiveMicropilot Then
'            nt = MicropilotRepetitions.Value
'        Else
'            nt = BlockRepetitions
'        End If
'
'        ReDim BleachTable(nt)               'The bleach table contains as many timepoints as blockrepetitions
'
'        'When working with the Lsm 2.8, remove all this test, except the one indicated line
'        If Track.EnableBleachRepeat Then
'            For i = Track.BleachScanNumber + 1 To nt Step Track.BleachRepeat
'                BleachTable(i) = True
'            Next
'        Else
'        '    BleachTable(Track.BleachScanNumber + 1) = True  'This is the only line to be kept when working with the Lsm 2.8
'        End If
'
'    End If
'End Sub

''''
'' Not used at the moment
''''
'Public Function SetGetLaserPower(power As Double)
'
'    Dim Recording As DsRecording
'    Dim Track As DsTrack
'    Dim IlluminationChannel As DsIlluminationChannel
'
'    Set Recording = Lsm5.DsRecording
'    Set Track = Recording.TrackObjectByIndex(0, Success)
'    Set IlluminationChannel = Track.IlluminationObjectByIndex(0, Success)
'
'    If (power > 0) Then
'        IlluminationChannel.power = power
'    End If
'
'    power = IlluminationChannel.power
'
'End Function
'
'
'Public Function MeasureExposure(fractionMax As Double, fractionSat As Double)
'
''    Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
' '   DoEvents
'
'    'Dim ColMax As Integer
'    Dim iRow As Integer
'    Dim nRow As Integer
'    Dim iFrame As Integer
'    Dim gvRow As Variant  ' gv = gray value
'    Dim iCol As Long
'    Dim nCol As Long
'    Dim bitDepth As Long
'    Dim channel As Integer
'    Dim gvMax As Double
'    Dim gvMaxBitRange As Double
'    Dim nSaturatedPixels As Long
'    Dim maxGV_nSat(2) As Double
'
'
'    DisplayProgress "Measuring Exposure...", RGB(0, &HC0, 0)
'
'    'ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
'
'    nRow = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
'    'MsgBox "nRow = " + CStr(nRow)
'
''        ElseIf SystemName = "LSM" Then
''            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
''            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
''        Else
''            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
'''            Exit Sub
' '       End If
' '   End If
'
'    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
'    'ReDim Intline(nLines) As Long
'
'    iFrame = 0
'    gvMax = -1
'
'    iRow = 0
'    channel = 0
'    bitDepth = 0 ' leaves the internal bit depth
'    gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
'    'MsgBox "nCol = " + CStr(nCol)
'    'MsgBox "bytes per pixel = " + CStr(bitDepth)
'
'    ' todo: is there another function to find this out??
'    If (bitDepth = 1) Then
'        gvMaxBitRange = 255
'    ElseIf (bitDepth = 2) Then
'        gvMaxBitRange = 65536
'    End If
'
'    nSaturatedPixels = 0
'
'    For iRow = 0 To nRow - 1
'        gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
'        For iCol = 0 To nCol - 1            'Now I'm scanning all the pixels in the line
'
'            If (gvRow(iCol) > gvMax) Then
'                gvMax = gvRow(iCol)
'            End If
'
'            If (gvRow(iCol) = gvMaxBitRange) Then
'                nSaturatedPixels = nSaturatedPixels + 1
'                ' TODO: measure neighbouring saturated pixels
'            End If
'
'        Next iCol
'    Next iRow
'
'    fractionMax = gvMax / gvMaxBitRange
'    fractionSat = nSaturatedPixels / (nRow * nCol)
'
'    'MsgBox "maximal gray value in image = " + CStr(gvMax)
'    'MsgBox "fractional brightness of maximal gray value in image = " + CStr(fractionMax)
'    'MsgBox "number of saturated pixles = " + CStr(nSaturatedPixels)
'    'MsgBox "fraction of saturated pixles = " + CStr(fractionSat)
'
'    DisplayProgress "Measuring Exposure...DONE", RGB(0, &HC0, 0)
'
'End Function

