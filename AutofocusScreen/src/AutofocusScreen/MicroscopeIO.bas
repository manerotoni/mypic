Attribute VB_Name = "MicroscopeIO"
Option Explicit
Option Base 0
Public SystemVersion As String

Public Declare Function GetInputState Lib "user32" () As Long ' Check if mouse or keyboard has been pushed

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function RegOpenKeyEx _
    Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long

Public Declare Function RegCloseKey _
    Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegQueryValueEx _
    Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'These are the name of the jobs
Public JobNames() As String
Public JobShortNames As Collection
'These are the name of the repetitions
Public RepNames() As String

'contains a list of keys that will be used for image analysis
Public OiaKeyNames() As String


''''''''''''''''''''
'''''CONSTANTS''''''
''''''''''''''''''''
Public Const VK_SPACE = &H20
Public Const VK_RETURN = &HD
Public Const VK_CANCEL = &H3
Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_ESCAPE = &H1B
Public Const VK_PAUSE = &H13
Public Const VK_ADD = &H6B
Public Const VK_SUBTRACT = &H6D
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const ERROR_SUCCESS = 0&

Public Const vbOKOnly = 0   '  Display OK button only.
Public Const VbOKCancel = 1 '  Display OK and Cancel buttons.
Public Const VbAbortRetryIgnore = 2  ' Display Abort, Retry, and Ignore buttons.
Public Const VbYesNoCancel = 3  '  Display Yes, No, and Cancel buttons.
Public Const VbYesNo = 4 '  Display Yes and No buttons.
Public Const VbRetryCancel = 5   ' Display Retry and Cancel buttons.
Public Const VbCritical = 16 ' Display Critical Message icon.
Public Const VbQuestion = 32 ' Display Warning Query icon.
Public Const VbExclamation = 48  ' Display Warning Message icon.
Public Const VbInformation = 64  ' Display Information Message icon.
Public Const VbDefaultButton1 = 0    ' First button is default.
Public Const VbDefaultButton2 = 256  ' Second button is default.
Public Const VbDefaultButton3 = 512  ' Third button is default.
Public Const VbDefaultButton4 = 768   'Fourth button is default.
Public Const VbApplicationModal = 0  ' Application modal; the user must respond to the message box before continuing work in the current application.
Public Const VbSystemModal = 4096   '  System modal; all applications are suspended until the user responds to the message box.
'The first group of values (0–5) describes the number and type of buttons displayed in the dialog box; the second group (16, 32, 48, 64) describes the icon style; the third group (0, 256, 512) determines which button is the default; and the fourth group (0, 4096) determines the modality of the message box. When adding numbers to create a final value for the buttons argument, use only one number from each group.

'Note   These constants are specified by Visual Basic for Applications. As a result, the names can be used anywhere in your code in place of the actual values.
'Return Values
Public Const vbOK = 1   '  OK
Public Const vbCancel = 2    ' Cancel
Public Const vbAbort = 3 ' Abort
Public Const vbRetry = 4 '  Retry
Public Const vbIgnore = 5   '  Ignore
Public Const vbYes = 6  '  Yes
Public Const vbNo = 7    ' No

Public Const PrecZ = 2                     'precision of Z passed for stage movements i.e. Z = Round(Z, PrecZ)
Public Const PrecXY = 2                    'precision of X and Y passed for stage movements

Public ZBacklash  As Double           'ToDo: is it still recquired?.
                                           'Has to do with the movements of the focus wheel that are "better"
                                           'if they are long enough. For amoment a test did not gave significant differences This is required for ZEN2010
Public ZENv As Integer            'String variable indicating the version of ZEN used 2010 ir 2011 (2012)
Public ZEN As Object             'Object containing Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface (for ZEN > 2011)

'''''''''''''''''''''
'''GLOBAL VARIABLE'''
'''''''''''''''''''''
Public RowG As Integer
Public ColG As Integer
Public RowSubG As Integer
Public ColSubG As Integer
Public X11 As Double
Public X12 As Double
Public X21 As Double
Public X22 As Double

Public ScanStop As Boolean
Public ScanPause As Boolean
Public Running As Boolean
Public ExtraBleach As Boolean
Public AutomaticBleaching As Boolean
Public BleachTable() As Boolean
Public BleachStartTable() As Double
Public BleachStopTable() As Double

Public ZOffset As Double
'Public PositionData As Workbook
'position variables
Public ZShift As Double
Public XShift As Double
Public YShift As Double
Public XStart As Double ' Stores starting X position of Acquisition
Public YStart As Double ' Stores starting Y position of Acquisition
Public ZStart As Double


'Filehandling variables
Public OverwriteFiles As Boolean
Public NoReflectionSignal As Boolean
Public PubSentStageGrid As Boolean
Public BleachingActivated As Boolean
Public FocusMapPresent As Boolean

Public flgEvent As Integer
Public flg As Integer
Public toContinue As Integer


Public GlobalProjectName As String
Public GlobalProject As String
Public GlobalHelpName As String

Public GlobalPrvTime As Double
Public GlobalMacroKey As String
Public GlobalCorrectionOffset As Double

'newPublic29.06.2010
Public NoFrames As Long

' Public BlockAutoConfiguration As String
Public BlockTimeIndex As Long
' Public BlockAutoConfigurationUse As Boolean

Public TimerName As String
Public BlockTimeDelay As Double
Public SelectedTimeButton As Integer
Public TimerButton1 As Double
Public TimerButton2 As Double
Public TimerButton3 As Double
Public TimerButton4 As Double
Public TimerButton5 As Double
Public TimerButton6 As Double
Public LoopingTimerUnit As Integer
Public BlockRepetitions As Long

Public TimerKey As String

Public GlobalHighRes As Boolean
Public GlobalDataBaseName As String
Public GlobalFileName As String
Public GlobalImageIndex() As Long
Public GlobalStripeIndex() As Long
Public PubSearchScan As Boolean

Public BlockIsSingle As Boolean
Public BlockSingleTrack As String
Public BlockSingleTrackIndex As Long
Public BlockMultiTrack As String
Public BlockMultiTrackIndex As Long


     
Public Track As DsTrack
Public TrackNumber As Integer
Public TrackName As String
Public Success As Integer
Public IsAutofocusTrackSelected As Boolean
Public AutofocusTrack As Integer ' number of AutofocusTrack
Public IsAcquisitionTrackSelected As Boolean
Public ActiveChannels() As String

Public LocationName As String

Public DoNotGoOn As Boolean
Public ChangeFocus As Boolean
Public FocusChanged As Boolean
Public Try As Long
Public SystemName As String
          
Public BackupRecording As DsRecording             ' To remove
          
Public GlobalBackupRecording As DsRecording       ' A backupRecording from initial setup (this will not be changed after Re_initialize)
Public GlobalAutoFocusRecording As DsRecording    ' A global variable for AutofocusRecording
Public GlobalAcquisitionRecording As DsRecording  ' A global variable for AcquisitionRecording
Public GlobalTrigger1Recording As DsRecording   ' A global variable for Trigger1
Public GlobalBleachRecording As DsRecording       ' A global variable for Bleaching

Public GlobalAltRecording As DsRecording          ' A global variable for AlternativeTrack
Public GlobalBackupActiveTracks() As Boolean


Public GlobalBackupSampleObservationTime As Double  ' Stores pixelDwell time

Public ImageNumber As Long
Public Const OFS_MAXPATHNAME = 128
Public Const OF_EXIST = &H4000
Public flgBreak As Boolean
Public Const WM_COMMAND = &H111

Public tools As Lsm5Tools
Public Stage As CpStages

Public TileX As Integer
Public TileY As Integer
Public Overlap As Double

Public AcquisitionController As AimAcquisitionController40.AimScanController  'Debugging 20110131

'''
'RecordingDoc used globally
'''
Public GlobalRecordingDoc As DsRecordingDoc
'''
'FcsData used globally
'''
Public GlobalFcsData As AimFcsData

'Grid positions
Public posGridX() As Double ' they are initiated during acquisition
Public posGridY() As Double ' they are initiated during acquisition
Public posGridZ() As Double ' initiated during acquistion
Public posGridXY_Valid() As Boolean ' they are initiated during acquisition

Public posGridXsub() As Double ' they are initiated during acquisition
Public posGridYsub() As Double ' they are initiated during acquisition
Public posGridZsub() As Double ' initiated during acquistion
Public posGridXYsub_valid() As Boolean ' they are initiated during acquisition

' Counters for HighresImaging 'TODO remove global variables
Public HighResExperimentCounter As Integer
Public HighResCounter As Integer
Public HighResArrayX() As Double ' this is an array of values why do you need to store values?
Public HighResArrayY() As Double
Public HighResArrayZ() As Double
Public HighResArrayDeltaX() As Double ' width of acquisition in um. Default is as set from the main window 'not implemented yet (better to define ROIs)
Public HighResArrayDeltaY() As Double ' width of acquisition in um. Default is as set from the main window 'not implemented yet (better to define ROIs)
Public HighResArrayDeltaZ() As Double ' size of acquisition in um. Default is a set from the main window

Public HelpNamePDF As String

Public GlobalStageControlZValues As Boolean

Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        Offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, _
ByVal wStyle As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
(ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long





Public Sub DisplayProgress(State As String, Color As Long)       'Used to display in the progress bar what the macro is doing
    If (Color & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Then
        AutofocusForm.ProgressLabel.ForeColor = 0
    Else
        AutofocusForm.ProgressLabel.ForeColor = &HFFFFFF
    End If
    AutofocusForm.ProgressLabel.BackColor = Color
    AutofocusForm.ProgressLabel.Caption = State
    DoEvents
End Sub


'''''
'   Range() As Double
'   Returs maximal range of Objective movement in um
'''''
Public Function Range() As Double
    Dim RevolverPosition As Long
    RevolverPosition = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
    If RevolverPosition >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(RevolverPosition) * 1000# ' the # is a double declaration
    Else
        Range = 0#
    End If
End Function



''''
' TODO: Why not use Lsm5.StartScan?
''''
Public Sub ScanToImageOld(RecordingDoc As DsRecordingDoc) ' new routine to scan overwrite the same image, even with several z-slices
   ' Dim AcquisitionController As AimAcquisitionController40.AimScanController 'now public
    Dim image As AimImage
    
    If Not RecordingDoc Is Nothing Then
        Set image = RecordingDoc.RecordingDocument.image(0, True)

        If Not image Is Nothing Then
            Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
            AcquisitionController.DestinationImage(0) = image
            AcquisitionController.DestinationImage(1) = Nothing
            AcquisitionController.StartGrab eGrabModeSingle
        End If
    End If
    
End Sub

'''''
'   ScanToImage ( RecordingDoc As DsRecordingDoc) As Boolean
'   scan overwrite the same image, even with several z-slices
'''''
Public Function ScanToImage(RecordingDoc As DsRecordingDoc) As Boolean

    Dim ProgressFifo As IAimProgressFifo ' what is this?
    Dim gui As Object, treenode As Object
    'Set gui = Lsm5.ViewerGuiServer
    If Not RecordingDoc Is Nothing Then
        Set treenode = RecordingDoc.RecordingDocument.image(0, True)
        'Set treenode = Lsm5.NewDocument why not this?
        Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller ' public variable
        AcquisitionController.DestinationImage(0) = treenode 'EngelImageToHechtImage(GlobalSingleImage).Image(0, True)
        AcquisitionController.DestinationImage(1) = Nothing
        Set ProgressFifo = AcquisitionController.DestinationImage(0)
        Lsm5.tools.CheckLockControllers True
        AcquisitionController.StartGrab eGrabModeSingle
        'Set RecordingDoc = Lsm5.StartScan this does not overwrite
        If Not ProgressFifo Is Nothing Then ProgressFifo.Append AcquisitionController
    End If
    Sleep (200)
    While AcquisitionController.IsGrabbing
        Sleep (200) ' this sometimes hangs if we use GetInputState. Try now without it and test if it does not hang
        DoEvents
        If ScanStop Then
            Exit Function
        End If
    Wend
    ScanToImage = True
End Function



'''''
'   SystemVersionOffset()
'   Calculate an offset added to z-stack changes
'       [GlobalCorrectionOffset] Global Out - Offset added to shift in zStack
'   TODO: Do we still need it. Only for Axioskop does the Offset change
'''''
Public Sub SystemVersionOffset(Optional Tmp As Boolean) ' tmp is a hack to hide function from menu
    SystemVersion = Lsm5.Info.VersionIs
    If StrComp(SystemVersion, "2.8", vbBinaryCompare) >= 0 Then
        If Lsm5.Info.IsAxioskop Then

        ElseIf Lsm5.Info.IsAxioplan Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioplan2 Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioplan2i Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioVert Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxiovert100M Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxiovert200M Then
            GlobalCorrectionOffset = 0
        Else
            GlobalCorrectionOffset = 0
        End If
    Else
        If Lsm5.Info.IsAxioskop Then

        ElseIf Lsm5.Info.IsAxioplan Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioVert Then
            GlobalCorrectionOffset = 0
        Else
            GlobalCorrectionOffset = 0
        End If
    End If
End Sub



''''' ' this should move to function
'   FailSafeMoveStage(Optional Mark As Integer = 0)
'   Moves stage and wait till it is finished
'       [x] In - x-position
'       [y] In - y-position
'''''
Public Function FailSafeMoveStageXY(X As Double, Y As Double) As Boolean
    
    FailSafeMoveStageXY = False


    Lsm5.Hardware.CpStages.SetXYPosition X, Y
    'TODO Check this
    Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
        Sleep (200)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                ScanStop = True
                Exit Function
            End If
        End If
    Loop
    
    FailSafeMoveStageXY = True
    
End Function


'''''
'   FailSafeMoveStageZ(z As Double)
'   Moves focus and wait till it is finished
'       [z] In - z-position )
'''''
Public Function FailSafeMoveStageZ(Z As Double) As Boolean
    FailSafeMoveStageZ = False
    If ZBacklash <> 0 Then
        Lsm5.Hardware.CpFocus.position = Z - ZBacklash ' move at correct position
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
            Sleep (20)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    FailSafeMoveStageZ = False
                    Exit Function
                End If
            End If
        Loop
    End If
    Lsm5.Hardware.CpFocus.position = Z  ' move at correct position
    Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
        Sleep (20)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                FailSafeMoveStageZ = False
                Exit Function
            End If
        End If
    Loop

    FailSafeMoveStageZ = True
End Function


'''''
'   MoveToNextLocation(Optional Mark As Integer = 0)
'   Moves to next location as set in the stage (mark)
'   Default will cycle through all positions sequentially starting from actual position
'       [Mark] In - Number of position where to move.
'''''
Public Sub MoveToNextLocation(Optional Mark As Integer = 0)
        Dim MarkCount As Long
        Dim count As Long
        Dim idx As Long
        Dim dX As Double
        Dim dY As Double
        Dim dZ As Double
        Dim i As Integer
        Lsm5.Hardware.CpStages.MarkMoveToZ (Mark)
        'Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ (0)  'old code Moves to the first location marked in the stage control. How to move to next point?
        ' the points were deleted and readded at the end of list in the Acquisition function
        'TODO: Check code
        Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
            Sleep (100)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    AutofocusForm.StopAcquisition
                    Exit Sub
                End If
            End If
        Loop
End Sub


Private Sub MovetoCorrectZPosition(ZOffset As Double)
Const ZBacklash = -50
Dim ZFocus As Double
Dim Zbefore As Double
Dim X As Double
Dim Y As Double
     ZFocus = Lsm5.Hardware.CpFocus.position + ZOffset + ZShift
       Lsm5.Hardware.CpFocus.position = ZFocus + ZBacklash    'Moves down -50uM (ZBacklash) with the focus wheel
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            Sleep (20)
            DoEvents
        Loop
        Lsm5.Hardware.CpFocus.position = ZFocus                     'Moves up to the focus position with the focus wheel
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            Sleep (20)
            DoEvents
        Loop
''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the AutofocusForm.AutofocusHRZ.Value can acces the position, and also to wait it is done...
        Sleep (100)
        DoEvents
End Sub


''''
'   WaitForRecentering(Z As Double, Success As Boolean) As Boolean
'   calls the microscope specific WaitForRecentering
'''
Public Function WaitForRecentering(Z As Double, Optional Success As Boolean = False, Optional ZENv As Integer = 2011) As Boolean
    If ZENv = 2010 Then
        If Not WaitForRecentering2010(Z, Success) Then
            Exit Function
        End If
    End If
    If ZENv > 2010 Then
        If Not WaitForRecentering2011(Z, Success) Then
            Exit Function
        End If
    End If
    WaitForRecentering = True
End Function



''''
'   WaitForRecentering2010(Z As Double) As Boolean
'   Helping function to check if after acquisition focus returns to its correct position
'       [Z] - is value where the central slice should be.
'   Additional remarks: Lsm5.Hardware.CpFocus.Position is not updated correctly after acquisition (CpFocus needs to return to working position) on the other hand
'   Lsm5.DsRecording.Sample0Z keeps track correctly of the position
'''
Public Function WaitForRecentering2010(Z As Double, Optional Success As Boolean = False) As Boolean
    Dim Cnt As Integer
    Dim MaxCnt As Integer
    MaxCnt = 6
    Cnt = 0
    ' Wait up to 4 sec for centering
    ' position central slice is Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position (or the real actual position)
    ' this waits for central slice at Z
    Dim pos As Double
    Dim Sample0Z As Double
    pos = Lsm5.Hardware.CpFocus.position
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        While Round(pos, 1) <> Round(Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
            GoTo FailedWaiting
        End If
'        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
'            pos - Z, 1) And Cnt < 10 'this is slow and not required and makes it slow
'            Sleep (400)
'            DoEvents
'            Cnt = Cnt + 1
'            If ScanStop Then
'                Exit Function
'            End If
'        Wend
    Else
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
            pos - Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
            GoTo FailedWaiting
        End If
    End If
    Success = True
    WaitForRecentering2010 = True
    Exit Function
FailedWaiting:
    DoEvents
    Success = False
    WaitForRecentering2010 = True
End Function

''''
'   WaitForRecentering2011(Z As Double, Success As Boolean) As Boolean
'   Helping function to check if after acquisition focus returns to its correct position
'       [Z] - is value where the central slice should be.
'       [Success] - Tells if central slide has been found before maximal number of iterations
'   Additional remarks: Lsm5.Hardware.CpFocus.Position is not updated correctly after acquisition (CpFocus needs to return to working position) on the other hand
'   Lsm5.DsRecording.Sample0Z keeps track correctly of the position
'''
Public Function WaitForRecentering2011(Z As Double, Optional Success As Boolean = False) As Boolean
    Dim Cnt As Integer
    Dim MaxCnt As Integer
    MaxCnt = 6
    Cnt = 0
    ' Wait up to 4 sec for centering
    ' Note pculiarity of centering
    ' position central slice is Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position (or the real actual position)
    ' this waits for central slice at Z
    Dim pos As Double
    Dim Sample0Z As Double
    pos = Lsm5.Hardware.CpFocus.position
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        While Round(pos, 1) <> Round(Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
            GoTo FailedWaiting
        End If
        
    Else
    
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
            pos - Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        
        If Cnt = MaxCnt Then
            Success = False
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
            GoTo FailedWaiting
        End If
    End If
    DoEvents
    Success = True
    WaitForRecentering2011 = True
    Exit Function
FailedWaiting:
    DoEvents
    Success = False
    WaitForRecentering2011 = True
End Function

''''
'   Recenter(Z As Double)
'   Sets the central slice. This slice is then maintained even when framespacing is changing.
'       [Z]     - Absolute position of central slice
'   position central slice is Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position
''''
Public Function Recenter_pre(Z As Double, Optional Success As Boolean = False, Optional ZENv As Integer = 2011) As Boolean
    If Not Recenter(Z, ZENv) Then
        Exit Function
    End If
    If Not WaitForRecentering(Z, Success, ZENv) Then
        Exit Function
    End If
    If Not ScanStop Then
        Recenter_pre = True
    End If
End Function

Public Function Recenter_post(Z As Double, Optional Success As Boolean = False, Optional ZENv As Integer = 2011) As Boolean
    If Not WaitForRecentering(Z, Success, ZENv) Then
        Exit Function
    End If
    
    If Not ScanStop Then
        Recenter_post = True
    End If
End Function

Public Function Recenter(Z As Double, Optional ZENv As Integer = 2011) As Boolean
    Dim i As Integer
    If ZENv = 2010 Then
        For i = 1 To 1
            If Not Recenter2010(Z) Then
                Exit Function
            End If
            Sleep (200)
        Next i
    End If
    If ZENv > 2010 Then
        If Not Recenter2011(Z) Then
            Exit Function
        End If
    End If
    Recenter = True
End Function

Public Function Recenter2010(Z As Double) As Boolean
    Dim MoveStage As Boolean
    Dim pos As Double
    Dim Sample0Z As Double
    pos = Lsm5.Hardware.CpFocus.position
    MoveStage = True ' this is the only difference to 2011 version
    
    If Lsm5.DsRecording.SpecialScanMode = "ZScanner" Or (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Then
        MoveStage = True
    End If
    Dim Tmp As Integer
    
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
    Sleep (100)
    DoEvents
    If MoveStage Then
        If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
        End If
    End If
    DoEvents
    Recenter2010 = True
End Function

Public Function Recenter2011(Z As Double) As Boolean
    Dim MoveStage As Boolean
    Dim FramesPerStack As Integer
    Dim pos As Double
    pos = Lsm5.Hardware.CpFocus.position
    MoveStage = False 'only move stage when required
    
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        MoveStage = True
    End If
        
    'Center slide
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
    Sleep (100)
    DoEvents
    If MoveStage Then
        If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
        End If
    End If
    DoEvents
    
    'to be tested

    'ZEN.gui.Acquisition.ZStack.MarkCenterPosition.Execute

    Recenter2011 = True
End Function








''''
' Compute the centerofmass of image stored in RecordingDoc
'   Use channel with name TrackingChannel
''''
Public Function MassCenter(RecordingDoc As DsRecordingDoc, TrackingChannel As String) As Vector

    On Error GoTo ErrorHandle:
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp")
    Dim Match As MatchCollection
    
    Dim scrline As Variant
    Dim spl As Long
    Dim bpp As Long
    Dim ColMax As Long
    Dim LineMax As Long
    Dim FrameNumber As Integer
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim Intline() As Long
    Dim IntCol() As Long
    Dim IntFrame() As Long
    
    Dim channel As Integer
    Dim frame As Long
    Dim Line As Long
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
    
   
    DoEvents
    
    
    'Find the channel to track
    Dim Rec As DsRecordingDoc
    Dim FoundChannel As Boolean
    FoundChannel = False
    RegEx.Pattern = "(\w+\d+) (\w+)"
    Dim name_channel As String
    If RegEx.test(TrackingChannel) Then
        Set Match = RegEx.Execute(TrackingChannel)
        name_channel = Match.Item(0).SubMatches.Item(1)
    End If
        
    For channel = 0 To RecordingDoc.GetDimensionChannels - 1
        If RecordingDoc.ChannelName(channel) = name_channel Then ' old Code: Left(TrackingChannel,4)
            FoundChannel = True
            Exit For
        End If
    Next channel
        
    If Not FoundChannel Then
        MsgBox (" Was not able to find channel for tracking!!")
        Exit Function
    End If


    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
    ColMax = RecordingDoc.Recording.SamplesPerLine
    LineMax = RecordingDoc.Recording.LinesPerFrame
    
    If RecordingDoc.Recording.ScanMode = "ZScan" Then
        LineMax = 1
    End If
    If RecordingDoc.Recording.ScanMode = "ZScan" Or RecordingDoc.Recording.ScanMode = "Stack" Then
        FrameNumber = RecordingDoc.Recording.FramesPerStack
    Else
        FrameNumber = 1
    End If
    
    'Gets the pixel size (it is in meter)
    pixelSize = RecordingDoc.Recording.SampleSpacing * 1000000
    'Gets the distance between frames in Z (in um)
    frameSpacing = RecordingDoc.Recording.frameSpacing
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    ReDim Intline(LineMax) As Long
    ReDim IntCol(ColMax) As Long
    ReDim IntFrame(FrameNumber) As Long
        

    
    'Compute center of mass
    'Reads the pixel values and fills the tables with the projected (integrated) pixels values in the three directions
    For frame = 1 To FrameNumber
        For Line = 1 To LineMax
            bpp = 0
            'channel = 0: This will allow to do the tracking on a different channel
            scrline = RecordingDoc.ScanLine(channel, 0, frame - 1, Line - 1, spl, bpp) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 2 To ColMax               'Now I'm scanning all the pixels in the line
                Intline(Line - 1) = Intline(Line - 1) + scrline(Col - 1)
                IntCol(Col - 1) = IntCol(Col - 1) + scrline(Col - 1)
                IntFrame(frame - 1) = IntFrame(frame - 1) + scrline(Col - 1)
            Next Col
        Next Line
    Next frame
    'First it finds the minimum and maximum porjected (integrated) pixel values in the 3 dimensions
    MinColValue = 4095 * LineMax * FrameNumber          'The maximum values are initially set to the maximum possible value
    minLineValue = 4095 * ColMax * FrameNumber
    minFrameValue = 4095 * LineMax * ColMax
    MaxColValue = 0                                     'The maximun values are initialliy set to 0
    MaxLineValue = 0
    MaxframeValue = 0
    For Line = 1 To LineMax
        If Intline(Line - 1) < minLineValue Then
            minLineValue = Intline(Line - 1)
        End If
        If Intline(Line - 1) > MaxLineValue Then
            MaxLineValue = Intline(Line - 1)
        End If
    Next Line
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
    For Line = 1 To LineMax
        LineValue = Intline(Line - 1) - Threshold                           'Subtracs the threshold
        PosValue = LineValue + Abs(LineValue)                               'Makes sure that the value is positive or zero. If LineValue is negative, the Posvalue = 0; if Line value is positive, then Posvalue = 2*LineValue
        LineWeight = LineWeight + (pixelSize * (Line - MidLine)) * PosValue 'Calculates the weight of the Thresholded projected pixel values according to their position relative to the center of the image and sums them up
        LineSum = LineSum + PosValue                                        'Calculates the sum of the thresholded pixel values
    Next Line
    If LineSum = 0 Then
         MassCenter.Y = 0
    Else
         MassCenter.Y = Round(LineWeight / LineSum, PrecXY)                                       'Normalizes the weights to get the center of mass
    End If

    ColSum = 0
    ColWeight = 0
    MidCol = (ColMax + 1) / 2
    Threshold = MinColValue + (MaxColValue - MinColValue) * 0.8
    For Col = 1 To ColMax
        ColValue = IntCol(Col - 1) - Threshold
        PosValue = ColValue + Abs(ColValue)
        ColWeight = ColWeight + (pixelSize * (Col - MidCol)) * PosValue
        ColSum = ColSum + PosValue
    Next Col
    If ColSum = 0 Then
         MassCenter.X = 0
    Else
         MassCenter.X = Round(ColWeight / ColSum, PrecXY)
    End If

    FrameSum = 0
    FrameWeight = 0
    MidFrame = (FrameNumber + 1) / 2
    Threshold = minFrameValue + (MaxframeValue - minFrameValue) * 0.8
    For frame = 1 To FrameNumber
        FrameValue = IntFrame(frame - 1) - Threshold
        PosValue = FrameValue + Abs(FrameValue)
        FrameWeight = FrameWeight + (frameSpacing * (frame - MidFrame)) * PosValue
        FrameSum = FrameSum + PosValue
    Next frame
    
    If FrameSum = 0 Then
        MassCenter.Z = 0
    Else
        MassCenter.Z = Round(FrameWeight / FrameSum, PrecZ)
    End If
    Exit Function
ErrorHandle:
    MsgBox ("Error in MicroscopeIO.MassCenter " + TrackingChannel + " " + Err.Description)
End Function



''''''
' SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String) As Boolean
' Copied and adapted from MultiTimeSeries macro
''''''
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
    'Set Export = New AimImageExport
    Export.FileName = FileName
    Export.format = eAimExportFormatLsm5
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
    AutofocusForm.StopAcquisition
End Function



'''''
'   UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
'   Ask which system is the macro runnning on
'       [bLSM]  In/Out - True if LSM system
'       [bLive] In/Out - True for LIVE system
'       [bCamera] In/Out - True if Camera is used
''''
Public Sub UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
    Dim Scancontroller As AimScanController
    Dim TrackParameters As AimTrackParameters
    Dim Size As Long
    Dim lTrack As Long
    Dim eDeviceMode As Long

    bLSM = False
    bLIVE = False
    bCamera = False
    Set Scancontroller = Lsm5.ExternalDsObject.Scancontroller
    Set TrackParameters = Scancontroller.TrackParameters
    If TrackParameters Is Nothing Then Exit Sub
    Size = TrackParameters.GetTrackArraySize
    For lTrack = 0 To Size - 1
            eDeviceMode = TrackParameters.TrackDeviceMode(lTrack)
            Select Case eDeviceMode
                Case eAimDeviceModeLSM
                    bLSM = True

                Case eAimDeviceModeLSM_ChannelMode
                    bLSM = True

                Case eAimDeviceModeLSM_NDD
                    bLSM = True

                Case eAimDeviceModeLSM_DD
                    bLSM = True

                Case eAimDeviceModeSpectralImager
                    bLSM = True
                    Exit Sub

                Case eAimDeviceModeRtScanner
                    bLIVE = True
                    Exit Sub

                Case eAimDeviceModeCamera1
                    bCamera = True
                    Exit Sub

            End Select
    Next lTrack
End Sub






'''''
'
'''''
Public Function SubImagingWorkFlowFcs(FcsData As AimFcsData, HighResArrayX() As Double, _
 HighResArrayY() As Double, HighResArrayZ() As Double, fileDir As String, FileName As String, Optional pixelSizeXY As Double = 0, Optional pixelSizeZ As Double = 0) As Boolean


    Dim i As Long
    ClearFcsPositionList
    Dim Recording As DsRecordingDoc
    For i = 1 To UBound(HighResArrayX)
        SetFcsPosition HighResArrayX(i), HighResArrayY(i), HighResArrayZ(i), i - 1
    Next i
    
    If Not CheckDir(fileDir) Then
        Exit Function
    End If
    
    NewFcsRecord FcsData, "fcs_" & FileName, 1
    
    If Not FcsMeasurement(FcsData) Then
        Exit Function
    End If
        
    Set Recording = Lsm5.DsRecordingActiveDocObject
    Recording.SetTitle "fcs_" & FileName
    
    SaveFcsPositionList fileDir & FileName & "_fcsPos.txt", pixelSizeXY, pixelSizeZ
    'save measurement
    SaveFcsMeasurement FcsData, fileDir & FileName & ".fcs"
    
End Function





