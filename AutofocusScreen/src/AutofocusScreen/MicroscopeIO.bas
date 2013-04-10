Attribute VB_Name = "MicroscopeIO"
Option Explicit
Public LegacyCode As Boolean

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

''''''''''
'''TYPES''
''''''''''
Public Type RepetitionType     'Contains Number and Time interval for repetition of acquisition protocol
    Number As Integer   'Number of repetitions
    Time As Double      'Interval between repetitions
    Interval As Boolean ' If Interval is True than one computes interval between first and second image = Time other wise
End Type

''''Position on grid
Public Type GridPosType
    Row As Long
    Col As Long
    RowSub As Long
    ColSub As Long
End Type

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

Public ZEN As String
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
Public LoopingTimerUnit As Integer
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
Public GlobalMicropilotRecording As DsRecording         ' A global variable for Micropilot
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
Public Sub SystemVersionOffset(Optional tmp As Boolean) ' tmp is a hack to hide function from menu
    SystemVersion = Lsm5.Info.VersionIs
    If StrComp(SystemVersion, "2.8", vbBinaryCompare) >= 0 Then
        If Lsm5.Info.IsAxioskop Then
            If AutofocusForm.AutofocusMaxSpeed Then
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
            If AutofocusForm.AutofocusMaxSpeed Then
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
    
    'If Not ScanToImage(NewPicture) Then
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
    If AutofocusForm.AutofocusHRZ Then
        Lsm5.Hardware.CpHrz.Position = 0                ' center the piezo focus (or bring it down again ?)
    End If
    
    Time = Timer
    DisplayProgress "Autofocus acquire", RGB(0, &HC0, 0)
    '''Check a last time that AF stack number and step is correct when in Fast Z-line mode
    If (Not AutofocusForm.AutofocusHRZ.Value) And AutofocusForm.ScanLineToggle.Value And AutofocusForm.AutofocusFastZline.Value Then
        If Lsm5.DsRecording.SpecialScanMode = "FocusStep" Then
             DisplayProgress "Highest Z Step of 1.54 um with no piezo and Fast Z line has been reached. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
        End If
        If AutofocusForm.AutofocusZStep.Value > Round(Lsm5.DsRecording.FrameSpacing, 3) Then
            DisplayProgress "Autofocus acquire. Highest Z Step with no piezo and Fast Z line " + CStr(Round(Lsm5.DsRecording.FrameSpacing, 3)) + " um. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
            Lsm5.DsRecording.SpecialScanMode = "FocusStep"
            Lsm5.DsRecording.FrameSpacing = AutofocusForm.AutofocusZStep.Value
        End If
    End If

    If Not ScanToImage(NewPicture) Then
        Exit Function
    End If
    
    LogMsg = "% Autofocus_stackshift: acquire time " & Round(Timer - Time, 2)
    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
    
    Time = Timer
    DisplayProgress "Autofocus compute", RGB(0, &HC0, 0)
    
    ' Computes XMass, YMass and ZMass
    MassCenter ("Autofocus")
    
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
    Dim Xpre As Integer
    Dim Ypre As Integer
    
    If MirrorX Then
        Xpre = -1
    Else
        Xpre = 1
    End If
    
    If MirrorY Then
        Ypre = -1
    Else
        Ypre = 1
    End If
    
    If ExchangeXY Then ' not sure about this
        X = X + Xpre * YMass
        Y = Y + Ypre * XMass
    Else
        X = X + Xpre * XMass
        Y = Y + Ypre * YMass
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
        If AutofocusForm.AutofocusTrackXY Then
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
''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the AutofocusForm.AutofocusHRZ.Value can acces the position, and also to wait it is done...
        Sleep (100)
        DoEvents
End Sub


''''
'   WaitForRecentering(Z As Double, Success As Boolean) As Boolean
'   calls the microscope specific WaitForRecentering
'''
Public Function WaitForRecentering(Z As Double, Optional Success As Boolean = False, Optional ZEN As String = "2011") As Boolean
    If ZEN = "2010" Then
        If Not WaitForRecentering2010(Z, Success) Then
            Exit Function
        End If
    End If
    If ZEN = "2011" Then
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
    pos = Lsm5.Hardware.CpFocus.Position
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
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
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
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
            pos - Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
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
    pos = Lsm5.Hardware.CpFocus.Position
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
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
            GoTo FailedWaiting
        End If
        
    Else
    
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
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
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
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
Public Function Recenter_pre(Z As Double, Optional Success As Boolean = False, Optional ZEN As String = "2011") As Boolean
    If Not Recenter(Z, ZEN) Then
        Exit Function
    End If
    If Not WaitForRecentering(Z, Success, ZEN) Then
        Exit Function
    End If
    If Not ScanStop Then
        Recenter_pre = True
    End If
End Function

Public Function Recenter_post(Z As Double, Optional Success As Boolean = False, Optional ZEN As String = "2011") As Boolean
    If Not WaitForRecentering(Z, Success, ZEN) Then
        Exit Function
    End If
    
    If Not ScanStop Then
        Recenter_post = True
    End If
End Function

Public Function Recenter(Z As Double, Optional ZEN As String = "2011") As Boolean
    Dim i As Integer
    If ZEN = "2010" Then
        For i = 1 To 1
            If Not Recenter2010(Z) Then
                Exit Function
            End If
            Sleep (200)
        Next i
    End If
    If ZEN = "2011" Then
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
    pos = Lsm5.Hardware.CpFocus.Position
    MoveStage = True ' this is the only difference to 2011 version
    
    If Lsm5.DsRecording.SpecialScanMode = "ZScanner" Or (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Then
        MoveStage = True
    End If
    Dim tmp As Integer
    
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
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
    pos = Lsm5.Hardware.CpFocus.Position
    MoveStage = False 'only move stage when required
    
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        MoveStage = True
    End If
        
    'Center slide
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + pos - Z
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
    Recenter2011 = True
End Function

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
    If AutofocusForm.AutofocusHRZ.Value And NoZStack Then                                            'If using HRZ for autofocusing and there is no Zstack for image acquisition
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
        If AutofocusForm.AutofocusHRZ.Value Then                             ' Now I'm not sure with the signs and... I some point I just tried random combinations...
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
        If AutofocusForm.AutofocusTrackXY Then
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



'''''
'   MakeGrid( posGridX() As Double, posGridY() As Double, posGridXY_valid() )
'   Create a grid
'       [posGridX] In/Out - Array where X grid positions are stored
'       [posGridY] In/Out - Array where Y grid positions are stored
'       [posGridXY_valid] In/Out - Array that says if position is valid
'       [locationNumbersMainGrid] In/Out - location number on main grid
'''''
Public Sub MakeGrid(posGridX() As Double, posGridY() As Double, posGridZ() As Double, XStart As Double, YStart As Double, ZStart As Double, dX As Double, dY As Double, dXsub As Double, dYsub As Double, refCol As Integer, refRow As Integer)
        ' A row correspond to Y movement and Column to X shift
        ' entries are posGridX(row, column)!! this what is
        'counters
        'Make main grid
        Dim iRow As Long
        Dim iCol As Long
        Dim iRowSub As Long
        Dim iColSub As Long
        'Make grid and subgrid
        For iRow = 1 To UBound(posGridX, 1)
            For iCol = 1 To UBound(posGridX, 2)
                For iRowSub = 1 To UBound(posGridX, 3)
                    For iColSub = 1 To UBound(posGridX, 4)
                        posGridX(iRow, iCol, iRowSub, iColSub) = Round(XStart + (1 - refCol) * dX + (iCol - 1) * dX + (iColSub - 1 - (UBound(posGridX, 4) - 1) / 2) * dXsub, PrecXY)
                        posGridY(iRow, iCol, iRowSub, iColSub) = Round(YStart + (1 - refRow) * dY + (iRow - 1) * dY + (iRowSub - 1 - (UBound(posGridX, 3) - 1) / 2) * dYsub, PrecXY)
                        posGridZ(iRow, iCol, iRowSub, iColSub) = Round(ZStart, PrecZ)
                    Next iColSub
                Next iRowSub
            Next iCol
        Next iRow
End Sub

Public Function isValidGridDefault(ByVal sFile As String) As Boolean
    Dim CellBase As String
    Dim Default As String
    Dim last_entry  As String
    Dim Active As Boolean
    Dim GoodMatch As Boolean
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp")
    Dim Match As MatchCollection
    'Well--Row--Col--(Row,Col)--TypeofWell
    CellBase = "(\d+)--(\d+)--(\d+)--(\S+)--(\S+)"
    'valid Row--Col
    Default = "(\d+) (\d+)--(\d+)"
    If FileExist(sFile) Then
        Close
        Dim iFileNum As Integer
        Dim Fields As String
        Dim FieldEntries() As String
        iFileNum = FreeFile()
        Open sFile For Input As iFileNum
        Do While Not EOF(iFileNum)
            On Error GoTo ErrorHandle:
            Line Input #iFileNum, Fields
            GoodMatch = False
            RegEx.Pattern = Default
            If RegEx.Test(Fields) Then
                Set Match = RegEx.Execute(Fields)
                Active = (Match.Item(0).SubMatches.Item(0) = "1")
                GoodMatch = True
            Else
                RegEx.Pattern = CellBase
                If RegEx.Test(Fields) Then
                    Set Match = RegEx.Execute(Fields)
                    Active = (Match.Item(0).SubMatches.Item(4) <> "empty")
                    GoodMatch = True
                Else
                    MsgBox "validGridDefault.txt has not the correct format! " & vbCrLf & "The format should be " & vbCrLf & "(In)active(0 or 1) Row(>=1)--Col(>=1) e.g." & vbCrLf & "0 1--1" & vbCrLf & "1 1--2" & vbCrLf _
                    & "or CellBase format Well--Row--Col--(Row,Col)--Identifier Identifier = empty=> InactiveWell e.g." & vbCrLf & "1--2--1--(1,1)--empty" & vbCrLf _
                    & "1--1--2--(1,2)--MyCoolGene"
                    Exit Function
                End If
            End If
        Loop
        isValidGridDefault = True
    Else
        MsgBox " File validGridDefault.txt needs to be in the output folder"
    End If
    Exit Function
ErrorHandle:
    MsgBox " Could not load validGridDefault.txt"
End Function


'''''
'   MakeValidGrid( posGridX() As Double, posGridY() As Double, posGridXY_valid() )
'   Create a grid
'       [posGridX] In/Out - Array where X grid positions are stored
'       [posGridY] In/Out - Array where Y grid positions are stored
'       [posGridXY_valid] In/Out - Array that says if position is valid
'       [locationNumbersMainGrid] In/Out - location number on main grid
'''''
Public Sub MakeValidGrid(posGridXY_Valid() As Boolean, Optional ByVal sFile As String)
    ' A row correspond to Y movement and Column to X shift
    ' entries are posGridX(row, column)!! this what is
    ' counters
    ' Make main grid
    Dim CellBase As String
    Dim Default As String
    Dim last_entry  As String
    Dim Active As Boolean
    Dim GoodMatch As Boolean
    'Well--Row--Col--(Row,Col)--TypeofWell
    CellBase = "(\d+)\-\-(\d+)\-\-(\d+)\-\-(\S+)\-\-(\S+)"
    'valid Row--Col
    Default = "(\d+) (\d+)\-\-(\d+)"
    Dim iRow As Long
    Dim iCol As Long
    Dim iRowSub As Long
    Dim iColSub As Long
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp")
    Dim Match As MatchCollection
    Dim Rec As DsRecordingDoc
    Dim FoundChannel As Boolean
    FoundChannel = False
    
    ' All points are true as default
    'Make grid and subgrid
    For iRow = 1 To UBound(posGridXY_Valid, 1)
        For iCol = 1 To UBound(posGridXY_Valid, 2)
            For iRowSub = 1 To UBound(posGridXY_Valid, 3)
                For iColSub = 1 To UBound(posGridXY_Valid, 4)
                    posGridXY_Valid(iRow, iCol, iRowSub, iColSub) = True
                Next iColSub
            Next iRowSub
        Next iCol
    Next iRow
    
    'File is either a Cellbase file or a series of 1 and zeros vertically ordered
    If sFile <> "" And FileExist(sFile) Then
        Close
        On Error GoTo ErrorHandle
        Dim iFileNum As Integer
        Dim Fields As String
        Dim FieldEntries() As String
        iFileNum = FreeFile()
        Open sFile For Input As iFileNum
        ' Go till end of file and fill the grid validity
        Do While Not EOF(iFileNum)
            On Error GoTo ErrorHandle:
            Line Input #iFileNum, Fields
            GoodMatch = False
            RegEx.Pattern = Default
            If RegEx.Test(Fields) Then
                Set Match = RegEx.Execute(Fields)
                Active = (Match.Item(0).SubMatches.Item(0) = "1")
                GoodMatch = True
            Else
                RegEx.Pattern = CellBase
                If RegEx.Test(Fields) Then
                    Set Match = RegEx.Execute(Fields)
                    Active = (Match.Item(0).SubMatches.Item(4) <> "empty")
                    GoodMatch = True
                End If
            End If
            'if this gridposition exists then update activity
            If GoodMatch And CInt(Match.Item(0).SubMatches.Item(1)) <= UBound(posGridXY_Valid, 1) And CInt(Match.Item(0).SubMatches.Item(1)) >= LBound(posGridXY_Valid, 1) _
            And CInt(Match.Item(0).SubMatches.Item(2)) <= UBound(posGridXY_Valid, 2) And CInt(Match.Item(0).SubMatches.Item(2)) >= LBound(posGridXY_Valid, 2) Then
                For iRowSub = 1 To UBound(posGridXY_Valid, 3)
                    For iColSub = 1 To UBound(posGridXY_Valid, 4)
                        posGridXY_Valid(CInt(Match.Item(0).SubMatches.Item(1)), CInt(Match.Item(0).SubMatches.Item(2)), iRowSub, iColSub) = Active
                    Next iColSub
                Next iRowSub
            End If
        Loop
    End If
    Exit Sub
ErrorHandle:
    MsgBox "MakeValidGrid error!"
End Sub


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
    
   
    'Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
    DoEvents
    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
    If AutofocusForm.ScanFrameToggle And SystemName = "LIVE" Then ' binning only with LIVE device
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
        Dim RegEx As VBScript_RegExp_55.RegExp
        Set RegEx = CreateObject("vbscript.regexp")
        Dim Match As MatchCollection
        Dim Rec As DsRecordingDoc
        Dim FoundChannel As Boolean
        FoundChannel = False
        RegEx.Pattern = "(\w+\d+)-(\w+)"
        Dim name_channel As String
        If RegEx.Test(TrackingChannelString) Then
            Set Match = RegEx.Execute(TrackingChannelString)
            name_channel = Match.Item(0).SubMatches.Item(0)
        End If
        
        For channel = 0 To Lsm5.DsRecordingActiveDocObject.GetDimensionChannels - 1
            If Lsm5.DsRecordingActiveDocObject.ChannelName(channel) = name_channel Then ' old Code: Left(TrackingChannel,4)
                FoundChannel = True
                Exit For
            End If
        Next channel
        
        If Not FoundChannel Then
            MsgBox (" Was not able to find channel for tracking!!")
            Exit Sub
        End If
    End If
    
    'Tracking is not the correct word. It just does center of mass on an additional channel
    'Reads the pixel values and fills the tables with the projected (integrated) pixels values in the three directions
    For frame = 1 To FrameNumber
        For Line = 1 To LineMax
            bpp = 0
            'channel = 0: This will allow to do the tracking on a different channel
            scrline = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, frame - 1, Line - 1, spl, bpp) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
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
        LineWeight = LineWeight + (PixelSize * (Line - MidLine)) * PosValue 'Calculates the weight of the Thresholded projected pixel values according to their position relative to the center of the image and sums them up
        LineSum = LineSum + PosValue                                        'Calculates the sum of the thresholded pixel values
    Next Line
    If LineSum = 0 Then
        YMass = 0
    Else
        YMass = Round(LineWeight / LineSum, PrecXY)                                       'Normalizes the weights to get the center of mass
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
        XMass = Round(ColWeight / ColSum, PrecXY)
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
        ZMass = Round(FrameWeight / FrameSum, PrecZ)
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
    
   
    'Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
    DoEvents
    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
    If AutofocusForm.ScanFrameToggle And SystemName = "LIVE" Then ' binning only with LIVE device
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
        For Line = 1 To LineMax
            bpp = 0 ' bytesperpixel
            'channel = 0: This will allow to do the tracking on a different channel
            scrline = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, frame - 1, Line - 1, spl, bpp) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 2 To ColMax               'Now I'm scanning all the pixels in the line
                Intline(Line - 1) = Intline(Line - 1) + scrline(Col - 1)
                IntCol(Col - 1) = IntCol(Col - 1) + scrline(Col - 1)
                IntFrame(frame - 1) = IntFrame(frame - 1) + scrline(Col - 1)
            Next Col
        Next Line
    Next frame
    
    'First it finds the minimum and maximum projected (integrated) pixel values in the 3 dimensions
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
        LineWeight = LineWeight + (PixelSize * (Line - MidLine)) * PosValue 'Calculates the weight of the Thresholded projected pixel values according to their position relative to the center of the image and sums them up
        LineSum = LineSum + PosValue                                        'Calculates the sum of the thresholded pixel values
    Next Line
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

'''
'   StorePositioninHighResArray(HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, HighResCounter As Integer)
'   TODO: Test stricter way of passing arguments
''''
Public Function StorePositioninHighResArray(ByVal Xref As Double, ByVal Yref As Double, ByVal Zref As Double, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double) As Boolean
    ' store postion from windows registry in array
    Dim Xoffset()  As String 'the string containinig the X-positions
    Dim Yoffset()  As String 'the string containinig the Y-positions
    Dim ZOffset() As String  'the string containinig the Z-positions
    Dim Xnew As Double
    Dim Ynew As Double
    Dim Znew As Double
    Dim PixelSize As Double
    Dim unit As String
    Dim LowBound As Integer
    Dim i As Integer
    unit = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="unit")

    If unit = "um" Or unit = Chr(181) & "m" Then   'has correct pixelSize of um
        PixelSize = 1
    Else                                    'Default is pixel even when Key does not exist
        PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
    End If
    
    Xoffset() = Split(GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="offsetx"), ",")
    If isArrayEmpty(Xoffset) Then
        ReDim Xoffset(0)
        Xoffset(0) = 0
    End If
    Yoffset() = Split(GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="offsety"), ",")
    If isArrayEmpty(Yoffset) Then
        ReDim Yoffset(0)
        Yoffset(0) = 0
    End If
    ZOffset() = Split(GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="offsetz"), ",")
    If isArrayEmpty(ZOffset) Then
        ReDim ZOffset(UBound(Xoffset))
        For i = 0 To UBound(Xoffset)
            ZOffset(i) = 0
        Next i
    End If
        
    If UBound(ZOffset) <> UBound(Xoffset) Then 'Z position as not been set
        
        ReDim Preserve ZOffset(UBound(Xoffset))
        For i = 0 To UBound(Xoffset)
            ZOffset(i) = ZOffset(0)
        Next i
    End If
    
    If UBound(Xoffset) <> UBound(Yoffset) Then
        MsgBox ("StorePositioninHighResArray: nr of values in registry offsetx and offsety is not the same!")
        Exit Function
    End If
    'ZOffset() = Split(GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="offsetz"), ",") for a later time point we can also set Z
    If isArrayEmpty(HighResArrayX) Then
        LowBound = 1
        ReDim HighResArrayX(1 To UBound(Xoffset) + 1)
        ReDim HighResArrayY(1 To UBound(Xoffset) + 1)
        ReDim HighResArrayZ(1 To UBound(Xoffset) + 1)
    Else
        LowBound = UBound(HighResArrayX) + 1
        ReDim Preserve HighResArrayX(1 To UBound(HighResArrayX) + UBound(Xoffset) + 1)
        ReDim Preserve HighResArrayY(1 To UBound(HighResArrayY) + UBound(Yoffset) + 1)
        ReDim Preserve HighResArrayZ(1 To UBound(HighResArrayZ) + UBound(ZOffset) + 1)
    End If
    
    For i = 0 To UBound(Xoffset)
        Xnew = Xref
        Ynew = Yref
        Znew = Zref
        ComputeShiftedCoordinates CDbl(Xoffset(i)) * PixelSize, CDbl(Yoffset(i)) * PixelSize, CDbl(ZOffset(i)), Xnew, Ynew, Znew
        HighResArrayX(LowBound + i) = Xnew  ' this needs to be unified with computing internal AF
        HighResArrayY(LowBound + i) = Ynew ' needs to be unified with computing internal AF
        HighResArrayZ(LowBound + i) = Znew
    Next i
    Dim Xoff As Double
    Xoff = Xref - CDbl(Xoffset(0)) * PixelSize
    DisplayProgress "StorePositioninHighResArray - Position stored", RGB(0, &HC0, 0)
End Function


''''''
'   MicroscopePilot(RecordingDoc As DsRecordingDoc, BleachingActivated As Boolean, HighResExperimentCounter As Integer, HighResCounter As Integer _
'   HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double)
'   TODO: test stricter way of passing arguments
'   [code] passed by MacroPilot can be
'       0, wait, Wait, DoNothing or doNothing: Macro waits, but no image is ready to analyse
'       1, newImage: Macro waits and there is a new image available
'       2, "nothing", "Nothing", "DoNothing", "doNothing", "donothing": Do nothing
'       4, "storePosition", "StorePosition", "storePosition", "StorePosition", "store", "Store": Store positions and exit function
'       5, "imagePositions", "ImagePositions", "imagePosition", "ImagePosition", "imageposition", "image", "Image": Image all stored positions
''''''
Public Function MicroscopePilot(RecordingDoc As DsRecordingDoc, GridPos As GridPosType, Xref As Double, Yref As Double, Zref As Double, FileNameID As String, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double) As Boolean
    Dim code As String
    Dim Repetitions As RepetitionType
    If Not LegacyCode Then
        code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="code")
        Select Case code
            Case "1", "newImage", "0", "wait", "Wait":
                'Wait for image analysis to finish
                DisplayProgress "Waiting for image analysis...", RGB(0, &HC0, 0)
                Do While (code = "1" Or code = "wait" Or code = "Wait" Or code = "0")
                    Sleep (50)
                    code = GetSetting(appname:="OnlineImageAnalysis", section:="macro", _
                              Key:="Code")
                    DoEvents
                    If ScanStop Then
                        Exit Function
                    End If
'                    If GetInputState() <> 0 Then
'                        DoEvents
'                        If ScanStop Then
'                            Exit Function
'                        End If
'                    End If
                Loop
        End Select
        
        Select Case code
            Case "2", "nothing", "Nothing", "DoNothing", "doNothing", "donothing":  'Nothing to do
            Case "4", "storePosition", "StorePosition", "storePosition", "StorePosition", "store", "Store": 'store positions for later processing
                DisplayProgress "Registry Code 4 (storePosition): store positions and do nothings ...", RGB(0, &HC0, 0)
                StorePositioninHighResArray Xref, Yref, Zref, HighResArrayX, HighResArrayY, HighResArrayZ
                If AutofocusForm.MicropilotMaxPositions.Value <> "" Then
                    If UBound(HighResArrayX) = CInt(AutofocusForm.MicropilotMaxPositions.Value) Then
                        Repetitions.Number = CInt(AutofocusForm.MicropilotRepetitions.Value)
                        Repetitions.Time = CDbl(AutofocusForm.MicropilotRepetitionTime.Value)
                        Repetitions.Interval = True
                        SubImagingWorkFlow RecordingDoc, GlobalMicropilotRecording, "Micropilot", AutofocusForm.MicropilotAutofocus, AutofocusForm.MicropilotZOffset, Repetitions, _
                        HighResArrayX, HighResArrayY, HighResArrayZ, GridPos, FileNameID
                        Erase HighResArrayX
                        Erase HighResArrayY
                        Erase HighResArrayZ
                    End If
                End If
            Case "5", "imagePositions", "ImagePositions", "imagePosition", "ImagePosition", "imageposition", "image", "Image":
                DisplayProgress "Registry Code 5 (imagePositions): store positions and do imaging ...", RGB(0, &HC0, 0)
                StorePositioninHighResArray Xref, Yref, Zref, HighResArrayX, HighResArrayY, HighResArrayZ
                ' BatchHighresImagingRoutine
                ' HERE THE IMAGES ARE ACQUIRED
                Repetitions.Number = CInt(AutofocusForm.MicropilotRepetitions.Value)
                Repetitions.Time = CDbl(AutofocusForm.MicropilotRepetitionTime.Value)
                Repetitions.Interval = True
                SubImagingWorkFlow RecordingDoc, GlobalMicropilotRecording, "Micropilot", AutofocusForm.MicropilotAutofocus, AutofocusForm.MicropilotZOffset, Repetitions, _
                HighResArrayX, HighResArrayY, HighResArrayZ, GridPos, FileNameID
                Erase HighResArrayX
                Erase HighResArrayY
                Erase HighResArrayZ
            Case Else
                MsgBox ("Invalid OnlineImageAnalysis Code = " & code)
                Exit Function
        End Select
        MicroscopePilot = True
    End If
    'Reset Code to 0 in Windows Registry
    'SaveSetting "OnlineImageAnalysis", "Cinput", "Code", 0 ' this should be done by the image analysis software
      

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
    'Reset Code to 0 in Windows Registry
    'SaveSetting "OnlineImageAnalysis", "Cinput", "Code",
            
End Function

'''''
'   BatchHighresImagingRoutine(RecordingDoc As DsRecordingDoc, HighResArrayX() As Double, HighResArrayY() As Double, HighResArrayZ() As Double, _
'   HighResCounter As Integer, HighResExperimentCounter As Integer)
'   TODO: Test stricter way of passing arguments
'''''
Public Function SubImagingWorkFlow(RecordingDoc As DsRecordingDoc, Recording As DsRecording, RecordingName As String, Autofocus As Boolean, ByVal ZOffset As Double, Repetitions As RepetitionType, HighResArrayX() As Double, _
 HighResArrayY() As Double, HighResArrayZ() As Double, GridPos As GridPosType, inFileNameID As String) As Boolean
    

    Dim Succes As Integer
    Dim SuccessRecenter As Boolean
    Dim LocationLabel As String
    'Timer and Looping Variables
    Dim iPos As Integer
    Dim iRep As Integer
    Dim StartTime As Double
    Dim NewTime As Double
    Dim DiffTime As Double
    Dim TimeLog As Double  'used in the the log
    
    'file name variables
    Dim FileNameID  As String
    Dim fullpathname As String
    Dim BackSlash As String
    Dim UnderScore As String
    Dim LogMsg As String
    'position variables
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim posZ As Double 'actual position Z

    ' set up the imaging
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    
    If RecordingDoc Is Nothing Then
        Set RecordingDoc = Lsm5.NewScanWindow
        While RecordingDoc.IsBusy
            Sleep (20)
            DoEvents
        Wend
    End If
    
    If Right(AutofocusForm.DatabaseTextbox.Value, 1) = "\" Then
        BackSlash = ""
    Else
        BackSlash = "\"
    End If
    
    If AutofocusForm.TextBoxFileName.Value <> "" Then
        UnderScore = "_"
    Else
        UnderScore = ""
    End If
    
    
    LocationLabel = AutofocusForm.LocationTextLabel.Caption
    For iRep = 1 To Repetitions.Number
        StartTime = CDbl(GetTickCount) * 0.001
        For iPos = 1 To UBound(HighResArrayX)  ' Postition loop
                ' Move to Positon in x, y, z for Highresscan
                DisplayProgress RecordingName & " - Move to Position " & iPos, RGB(0, &HC0, 0)
                X = HighResArrayX(iPos)
                Y = HighResArrayY(iPos)
                Z = HighResArrayZ(iPos)
                If Not FailSafeMoveStageXY(HighResArrayX(iPos), HighResArrayY(iPos)) Then
                    Exit Function
                End If
                'center acquisition (this should be already at correct position after AF)
                Recenter_pre Z, SuccessRecenter, ZEN
                AutofocusForm.LocationTextLabel.Caption = LocationLabel & _
                RecordingName & " Locations: " & iPos & "/" & UBound(HighResArrayX) & _
                "; X = " & X & ", Y = " & Y & ", Z = " & Z & vbCrLf & _
                "Repetition :" & iRep & "/" & Repetitions.Number
                DoEvents
                'Autofocus. This does an extra Autofocus also for the HighresImaging with the same parameters as Autofocus
                If Autofocus Then
                    DisplayProgress RecordingName & " - Autofocus activate track and recenter...", RGB(0, &HC0, 0)
                    If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
                        MsgBox "No track selected for Autofocus! Cannot Autofocus!"
                        Exit Function
                    End If
                    
                    '''center
                    TimeLog = Timer
                    If Not Recenter_pre(Z, SuccessRecenter, ZEN) Then
                        Exit Function
                    End If
                    LogMsg = "% " & RecordingName & " : recenter Z (pre AFImg) " & Z & ", time required " & Round(Timer - TimeLog)
                    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
                    
                    ' take a z-stack and finds the brightest plane:
                    DisplayProgress RecordingName & " - Autofocus acquire ...", RGB(0, &HC0, 0)
                    If Not MicroscopeIO.Autofocus_StackShift(RecordingDoc) Then
                       Exit Function
                    End If
                    
                    'wait for recentering
                    DisplayProgress RecordingName & " - Wait for recentering after acquisition AF...", RGB(0, &HC0, 0)
                    TimeLog = Timer
                    If Not Recenter_post(Z, SuccessRecenter, ZEN) Then
                        Exit Function
                    End If
                    
                    LogMsg = "% " & RecordingName & ": recenter Z (post AFImg) " & Z
                    posZ = Round(Lsm5.Hardware.CpFocus.Position, PrecZ)
                    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusForm.AutofocusHRZ Then
                        LogMsg = LogMsg & ", Obtained Z " & posZ & "; actual position " & posZ & ", Time required " & Round(Timer - TimeLog) & ", success within rep. " & SuccessRecenter
                    Else
                        LogMsg = LogMsg & ", Obtained Z " & Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + posZ _
                        & "; actual position " & posZ & ", Time required " & Round(Timer - TimeLog) & ", success within rep. " & SuccessRecenter
                    End If
                    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
                        
                        
                    ' move the xyz to the right position
                    ComputeShiftedCoordinates XMass, YMass, ZMass, X, Y, Z
                    If AutofocusForm.AutofocusTrackXY.Value And AutofocusForm.ScanFrameToggle.Value Then
                        DisplayProgress RecordingName & " - Autofocus move XY stage", RGB(0, &HC0, 0)
                        If Not FailSafeMoveStageXY(X, Y) Then
                            Exit Function
                        End If
                    End If
                    LogMsg = "% " & RecordingName & ":  center of mass  " & XMass & ", " & YMass & ", " & ZMass & ", computed position " & X & ", " & Y & ", " & Z
                    LogMessage LogMsg, Log, LogFileName, LogFile, FileSystem
                    HighResArrayX(iPos) = X
                    HighResArrayY(iPos) = Y
                    HighResArrayZ(iPos) = Z
                End If
                
                DisplayProgress RecordingName & " - activate micropilot acquisition track and recenter ...", RGB(0, &HC0, 0)
                If Not AutofocusForm.ActivateTrack(Recording, RecordingName) Then
                    MsgBox " No Track selected for  " & RecordingName & "! Macro stops here"
                    Exit Function
                End If
                
                ' set central slice before moving
                If Not Recenter_pre(Z + ZOffset, SuccessRecenter, ZEN) Then
                    Exit Function
                End If
                
                DisplayProgress RecordingName & " - Acquisition position " & iPos & "/" & UBound(HighResArrayX) & " Repetition " & iRep _
                & "/" & Repetitions.Number, RGB(&HC0, &HC0, 0)

                If Not ScanToImage(RecordingDoc) Then
                    Exit Function
                End If
                
                fullpathname = AutofocusForm.DatabaseTextbox.Value & BackSlash & "HRExp_" & AutofocusForm.TextBoxFileName.Value & UnderScore & inFileNameID & "\"
                If Not CheckDir(fullpathname) Then
                    Exit Function
                End If
                FileNameID = FileName(GridPos.Row, GridPos.Col, GridPos.RowSub, GridPos.ColSub, iRep)
                fullpathname = fullpathname & AutofocusForm.TextBoxFileName.Value & UnderScore & "HRPos" & ZeroString(3 - Len(CStr(iPos))) & iPos & _
                 "_" & FileNameID & ".lsm"
                
                DisplayProgress RecordingName & " - SaveImage", RGB(0, &HC0, 0)
                If Not SaveDsRecordingDoc(RecordingDoc, fullpathname) Then
                    Exit Function
                End If

                'Here the Location-tracking in High-resmode code needs be added! This only if one wants to track extra in the highresmode
          
        Next iPos ' End of postions loop
        'TODO Check
        If Not Repetitions.Interval Then
            StartTime = CDbl(GetTickCount) * 0.001
        End If
        DiffTime = CDbl(GetTickCount) * 0.001 - StartTime
        While DiffTime <= Repetitions.Time And iRep < Repetitions.Number
            Sleep (10)
            If GetInputState() <> 0 Then
            DoEvents
                If ScanStop Then
                    SubImagingWorkFlow = False
                    Exit Function
                End If
            End If
            DiffTime = CDbl(GetTickCount) * 0.001 - StartTime
            DisplayProgress "Waiting " & CStr(CInt(Repetitions.Time - DiffTime)) + " s before scanning repetition  " & (iRep + 1), RGB(&HC0, &HC0, 0)
        Wend
    Next iRep  ' End of time repetition loop
    
'                If BleachingActivated Then
'
'                    DisplayProgress "Bleaching...", &HFF00FF
'
'                    Set Track = Lsm5.DsRecording.TrackObjectBleach(Success)
'                    If Success Then ' This Needs to be checked again
'                        'move stage
'                        If Not FailSafeMoveStageZ(Z + MicropilotZOffset.Value) Then
'                            Exit Function
'                        End If
'                        Track.Acquire = True
'                        Lsm5.DsRecording.TimeSeries = True
'                        Lsm5.DsRecording.StacksPerRecord = MicropilotRepetitions.Value
'                        Lsm5.DsRecording.FramesPerStack = 1
'                        Track.TimeBetweenStacks = MicropilotRepetionTime.Value
'                        If Not WaitForRecentering(Z + MicropilotZOffset.Value, SuccessRecenter, ZEN) Then
'                            Exit Function
'                        End If
'
'                        DoEvents
'                        Track.UseBleachParameters = True            'Bleach parameters are lasers lines, bleach iterations... stored in the bleach control window
'                        'BleachStartTable(RepetitionNumber) = GetTickCount      'Get the time right before bleach to store this in the image metadata
'                        Set RecordingDoc = Lsm5.StartScan
'                        'TODO Check
'                        Do While RecordingDoc.IsBusy
'                            Sleep (100)
'                            If GetInputState() <> 0 Then
'                                DoEvents
'                                If ScanStop Then
'                                    Exit Function
'                                End If
'                            End If
'                        Loop
'
'                        Track.UseBleachParameters = False  'switch off the bleaching
'                        Lsm5.DsRecording.TimeSeries = False
'                    Else
'                        MsgBox ("Could not set bleach track. Did not bleach.")
'                    End If
'
'
'                    'Save Image
'                    FileNameID = FileName(Row, Col, RowSub, ColSub, -1)
'                    ' e.g. name_Bleach_HRExp001_HRPos001_W001_P001.lsm
'                    fullpathname = DatabaseTextbox.Value & BackSlash & TextBoxFileName.Value & "_Bleach" & "_HRExp" & ZeroString(3 - Len(CStr(HighResExperimentCounter))) & _
'                    HighResExperimentCounter & "_HRPos" & ZeroString(3 - Len(CStr(highrespos))) _
'                    & highrespos & "_" & FileNameID & ".lsm"
'                    SaveDsRecordingDoc RecordingDoc, fullpathname
'                    DisplayProgress "Micropilot Code Bleach - SaveImage", RGB(0, &HC0, 0)
'

    
    SubImagingWorkFlow = True
End Function

