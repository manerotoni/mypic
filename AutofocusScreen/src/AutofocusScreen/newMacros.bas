Attribute VB_Name = "newMacros"
Option Explicit
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

''''''''''''''''''''''''
'Debug and LogVariables'
''''''''''''''''''''''''
Public LogFile As TextStream 'This is the file where a log of the procedure is saved
Public LogFileName As String
Public LogFileNameBase As String
Public FileSystem As FileSystemObject
Public Log     As Boolean          'If true we log data during the macro



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


'''''''''''''''''''''
'''GLOBAL VARIABLE'''
'''''''''''''''''''''



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
Public RepetitionNumber As Integer ' number of repetition
Public locationNumber As Long      ' number of location global

Public ZOffset As Double
Public TrackingChannelString As String
'Public PositionData As Workbook
'position variables
Public XMass As Double
Public YMass As Double
Public ZMass As Double
Public ZShift As Double
Public XShift As Double
Public YShift As Double
Public XStart As Double ' Stores starting X position of Acquisition
Public YStart As Double ' Stores starting Y position of Acquisition
Public ZStart As Double
Public HRZBefore As Double
Public HRZ As Boolean

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
Public TimerUnit As Integer
Public BlockRepetitions As Long

Public TimerKey As String

Public GlobalHighRes As Boolean
Public GlobalDataBaseName As String
Public GlobalFileName As String
Public GlobalImageIndex() As Long
Public GlobalStripeIndex() As Long
Public BlockZOffset As Double
Public BlockZRange As Double
Public BlockZStep As Double
Public BlockHighSpeed As Boolean
Public BlockLowZoom As Boolean
Public BlockHRZ As Boolean
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
Public GlobalZoomRecording As DsRecording         ' A global variable for Micropilot
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
Public RecordingDocpub As DsRecordingDoc


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


Public Sub Autofocus_Setup()
        AutofocusForm.Show
End Sub


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


''''
' TODO: Why not use Lsm5.StartScan?
''''
Public Sub ScanToImage(RecordingDoc As DsRecordingDoc) ' new routine to scan overwrite the same image, even with several z-slices
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
'   ScanToImageNew ( RecordingDoc As DsRecordingDoc) As Boolean
'   scan overwrite the same image, even with several z-slices
'''''
Public Function ScanToImageNew(RecordingDoc As DsRecordingDoc) As Boolean

    Dim ProgressFifo As IAimProgressFifo ' what is this?
    Dim gui As Object, treenode As Object
    'Set gui = Lsm5.ViewerGuiServer
    ScanToImageNew = False
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
    ScanToImageNew = True
End Function



'''''
'   SystemVersionOffset()
'   Calculate an offset added to z-stack changes
'       [GlobalCorrectionOffset] Global Out - Offset added to shift in zStack
'   TODO: Do we still need it. Only for Axioskop does the Offset change
'''''
Public Sub SystemVersionOffset(Optional tmp As Boolean) ' tmp is a hack to hide function from menu
    SystemVersion = Lsm5.Info.VersionIs
    If StrComp(SystemVersion, "2.8", vbBinaryCompare) >= 0 Then
        If Lsm5.Info.IsAxioskop Then
            If AutofocusForm.CheckBoxHighSpeed Then
                GlobalCorrectionOffset = 15
            Else
                GlobalCorrectionOffset = 1.2
            End If
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
            If AutofocusForm.CheckBoxHighSpeed Then
                GlobalCorrectionOffset = 15
            Else
                GlobalCorrectionOffset = 1.2
            End If
        ElseIf Lsm5.Info.IsAxioplan Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioVert Then
            GlobalCorrectionOffset = 0
        Else
            GlobalCorrectionOffset = 0
        End If
    End If
End Sub

'''''''
' Autofocus_StackShift ( NewPicture As DsRecordingDoc )
' Performs image scan as in GlobalAutofocusRecording, calculation of signal centroid (mass)
' global variables [ZMass] and [XMass], [YMasss] (FrameScan).
'                  GlobalAutofocusRecording is set in function
' This function does not change the focus just computes it
'       [NewPicture] In/Out - Contains the image
'''''''
Public Function Autofocus_StackShift(NewPicture As DsRecordingDoc) As Boolean
    Dim pixelDwell As Double
    Dim BigZStep As Double
    Dim LogMsg As String
    Dim Time As Double
    Dim Cnt As Integer

    
    
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    DisplayProgress "Autofocus SetupScanWindow", RGB(0, &HC0, 0)
    If NewPicture Is Nothing Then
        Set NewPicture = Lsm5.NewScanWindow
        While NewPicture.IsBusy
            Sleep (100)
            DoEvents
        Wend
    End If
    
    'Dim FramesPerStack As Double
    'FramesPerStack = Lsm5.DsRecording.FramesPerStack
    'Lsm5.DsRecording.FramesPerStack = 1
    
    'If Not ScanToImageNew(NewPicture) Then
    '    Exit Function
    'End If
    
    'Lsm5.DsRecording.FramesPerStack = FramesPerStack
    
    DisplayProgress "Autofocus: CheckZRange", RGB(0, &HC0, 0)
    'checks again if Zranges are good
    If Not AutofocusForm.CheckZRanges() Then
        Autofocus_StackShift = False
        Exit Function
    End If
    
    SystemVersionOffset         ' extra offset depending on macroscope

    ''''''''''''''''''
    '** Autofocus ***'
    ''''''''''''''''''
    
    DisplayProgress "Autofocus reset Z-position", RGB(0, &HC0, 0)
    If AutofocusForm.CheckBoxHRZ Then
        Lsm5.Hardware.CpHrz.Position = 0                ' center the piezo focus (or bring it down again ?)
    End If
    
    Time = Timer
    DisplayProgress "Autofocus acquire", RGB(0, &HC0, 0)
    '''Check a last time that AF stack number and step is correct when in Fast Z-line mode
    If (Not AutofocusForm.CheckBoxHRZ.Value) And AutofocusForm.ScanLineToggle.Value And AutofocusForm.CheckBoxFastZline.Value Then
        If Lsm5.DsRecording.SpecialScanMode = "FocusStep" Then
             DisplayProgress "Highest Z Step of 1.54 um with no piezo and Fast Z line has been reached. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
        End If
        If AutofocusForm.BSliderZStep.Value > Round(Lsm5.DsRecording.FrameSpacing, 3) Then
            DisplayProgress "Autofocus acquire. Highest Z Step with no piezo and Fast Z line " + CStr(Round(Lsm5.DsRecording.FrameSpacing, 3)) + " um. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
            Lsm5.DsRecording.SpecialScanMode = "FocusStep"
            Lsm5.DsRecording.FrameSpacing = AutofocusForm.BSliderZStep.Value
        End If
    End If

    If Not ScanToImageNew(NewPicture) Then
        Exit Function
    End If
    
    LogMsg = "% Autofocus_stackshift: acquire time " & Round(Timer - Time, 2)
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    
    Time = Timer
    DisplayProgress "Autofocus compute", RGB(0, &HC0, 0)
    
    ' Computes XMass, YMass and ZMass
    AutofocusForm.MassCenter ("Autofocus")
    
    LogMsg = "% Autofocus_stackshift: compute time " & Round(Timer - Time, 2)
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem

    If Not ScanStop Then
        Autofocus_StackShift = True
    End If
End Function


'''''''
'   ComputeShiftedCoordinates(XMass, ....)
'   Calculates new coordinates after translation
'       [XMass], [YMass], [ZMass]    In - Translation vector
'       [x], [y], [z] Out - Shifted coordinates. Depends on stage build up and actual position. Positions are rounded up to PrecXY and PrecZ
''''''
Public Function ComputeShiftedCoordinates(ByVal XMass As Double, ByVal YMass As Double, ByVal ZMass As Double, ByRef X As Double, ByRef Y As Double, ByRef Z As Double)

    If AreStageCoordinateExchanged Then
        X = X - YMass
        Y = Y - XMass
    Else
        X = X + XMass
        Y = Y - YMass
    End If
        
    Z = Z + ZMass
    X = Round(X, PrecXY)
    Y = Round(Y, PrecXY)
    Z = Round(Z, PrecZ)
End Function

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
        Lsm5.Hardware.CpFocus.Position = Z - ZBacklash ' move at correct position
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
    Lsm5.Hardware.CpFocus.Position = Z  ' move at correct position
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

''''''
'   Autofocus_MoveAcquisition
'   Move stage and Z. To Z a ZOffset from the Autofocusform is added
'       [XShift] In
'       [YShift] In
'       [ZShift] In
''''''
Public Function Autofocus_MoveAcquisition(XShift As Double, YShift As Double, ZShift As Double, ZOffset As Double) As Boolean
    
    Dim ZFocus As Double
    Dim Zbefore As Double
    Dim X As Double
    Dim Y As Double
        
    '''''''''''''''''''''''''''''''''''''''
    ' Moving to the correct position in Z
    ' Defines the new focus position as the actual position plus the shift and goes back to the object position (that's why you need the offset)
    ZFocus = Lsm5.Hardware.CpFocus.Position + ZOffset + ZShift
    ' Why do you need to move downward first ? Todo check if recquired step
    Lsm5.Hardware.CpFocus.Position = ZFocus + ZBacklash     'Moves down -50uM (ZBacklash) with the focus wheel
    Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
        Sleep (20)
        DoEvents
    Loop
    Lsm5.Hardware.CpFocus.Position = ZFocus                     'Moves up to the focus position with the focus wheel
    Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
        Sleep (20)
        DoEvents
    Loop
    
    ' Todo: one might add a lot of controls here, to wait to be sure the focus wheel can acces the position, and also to wait it is done...
    Sleep (100)
    DoEvents
   
    'This is moving the x and y position
    'This we want only to do when xy-focus is set
    'Moving to the correct position in X and Y
    If AutofocusForm.ScanFrameToggle Then
        ' Todo: check whether it moves in the correct direction
        If AutofocusForm.CheckBoxAutofocusTrackXY Then
            X = Lsm5.Hardware.CpStages.PositionX - XShift
            Y = Lsm5.Hardware.CpStages.PositionY - YShift
            Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(X, Y)
        End If
         
        Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            If ScanStop Then
                Lsm5.StopScan
                AutofocusForm.StopAcquisition
                DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                Autofocus_MoveAcquisition = False
                Exit Function
            End If
            DoEvents
            Sleep (5)
        Loop
    
    End If
    
    ' center all z-stacks again!
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * Lsm5.DsRecording.FramesPerStack / 2
    
    Autofocus_MoveAcquisition = True
    
End Function

'''''
'   MoveToNextLocation(Optional Mark As Integer = 0)
'   Moves to next location as set in the stage (mark)
'   Default will cycle through all positions sequentially starting from actual position
'       [Mark] In - Number of position where to move.
'''''
Public Sub MoveToNextLocation(Optional Mark As Integer = 0)
        Dim Markcount As Long
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
     ZFocus = Lsm5.Hardware.CpFocus.Position + ZOffset + ZShift
       Lsm5.Hardware.CpFocus.Position = ZFocus + ZBacklash    'Moves down -50uM (ZBacklash) with the focus wheel
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            Sleep (20)
            DoEvents
        Loop
        Lsm5.Hardware.CpFocus.Position = ZFocus                     'Moves up to the focus position with the focus wheel
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            Sleep (20)
            DoEvents
        Loop
''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the AutofocusForm.CheckBoxHRZ.Value can acces the position, and also to wait it is done...
        Sleep (100)
        DoEvents
End Sub

''''
'   Autofocus_MoveAcquisition_HRZ(ZOffset As Double)
'   Allow to use HRZ for Move Z-stage (not used at the moment)
''''
Public Sub Autofocus_MoveAcquisition_HRZ(ZOffset As Double)
    Dim NoZStack As Boolean
    Const ZBacklash = -50
    Dim ZFocus As Double
    Dim Zbefore As Double
    Dim X As Double
    Dim Y As Double

    AutofocusForm.RestoreAcquisitionParameters
    
    Set GlobalBackupRecording = Nothing
    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
    
    NoZStack = True
    If GlobalAcquisitionRecording.ScanMode = "ZScan" Or GlobalAcquisitionRecording.ScanMode = "Stack" Then  'Looks if a Z-Stack is going to be acquired
        NoZStack = False
    End If

    'Moving to the correct position in Z
    If AutofocusForm.CheckBoxHRZ.Value And NoZStack Then                                            'If using HRZ for autofocusing and there is no Zstack for image acquisition
        Lsm5.Hardware.CpHrz.Stepsize = 0.2
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
        DoEvents
     '   ZFocus = Lsm5.Hardware.CpHrz.Position + ZShift - ZOffset
     
     'Defines the new focus position as the actual position plus the shift and goes back to the object position (that's why you need the offset)
  
        ZFocus = Lsm5.Hardware.CpHrz.Position + ZOffset + ZShift
       
        Lsm5.Hardware.CpHrz.Position = ZFocus                     'Moves up to the focus position with the focus wheel
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            Sleep (20)
            DoEvents
        Loop
''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the HRZ can acces the position, and also to wait it is done...
        
        DoEvents

    Else                                        'either there is a Z stack for image acquisition or we're using the focuswheel for autofocussing
        If AutofocusForm.CheckBoxHRZ.Value Then                             ' Now I'm not sure with the signs and... I some point I just tried random combinations...
            ZFocus = Lsm5.Hardware.CpHrz.Position - ZOffset - ZShift '         'ZBefore corresponds to the position where the focuswheel was before doing anything. Zshift is the calculated shift
        Else                                    'If the HRZ is not calibrated the Z shift might be wrong
            ZFocus = Zbefore + ZShift
        End If
       
        Lsm5.Hardware.CpHrz.Position = ZFocus                     'Moves up to the focus position with the focus wheel
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            Sleep (20)
            DoEvents
        Loop
    End If

    'Moving to the correct position in X and Y
 
    If AutofocusForm.ScanFrameToggle Then
        If AutofocusForm.CheckBoxAutofocusTrackXY Then
            X = Lsm5.Hardware.CpStages.PositionX - XShift  'the fact that it is "-" in this line and "+" in the next line  probably has to do with where the XY of the origin is set (top right corner and not botom left, I think)
            Y = Lsm5.Hardware.CpStages.PositionY - YShift
            Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(X, Y)
        End If
         
        Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
            If ScanStop Then
                Lsm5.StopScan
                AutofocusForm.StopAcquisition
                DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                Exit Sub
            End If
            DoEvents
            Sleep (5)
        Loop
    End If
    

    DisplayProgress "Autofocus 14", RGB(0, &HC0, 0)
    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
    DisplayProgress "Autofocus 15", RGB(0, &HC0, 0)
End Sub



