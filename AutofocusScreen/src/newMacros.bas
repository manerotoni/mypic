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
Public RepetitionNumber As Long

Public ZOffset As Double
Public MultipleLocation As Boolean
Public LocationTracking As Boolean
Public TrackingChannelString As String
'Public PositionData As Workbook
Public FrameAutofocussing As Boolean
Public XMass As Double
Public YMass As Double
Public ZMass As Double
Public ZShift As Double
Public XShift As Double
Public YShift As Double
Public Zbefore As Double
Public HRZBefore As Double
Public HRZ As Boolean
Public NoReflectionSignal As Boolean
Public PubSentStageGrid As Boolean
Public BleachingActivated As Boolean

Public flgUserChange As Boolean
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
Public AutofocusTrack As Integer
Public IsAcquisitionTrackSelected As Boolean
Public ActiveChannels() As String

Public LocationName As String

Public DoNotGoOn As Boolean
Public ChangeFocus As Boolean
Public FocusChanged As Boolean
Public Try As Long
Public SystemName As String
          
Public GlobalBackupRecording As DsRecording ' TODO: why two variables
Public BackupRecording As DsRecording       ' TODO: why two variables
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



Public posGridX(10000) As Double
Public posGridY(10000) As Double
Public locationNumbersMainGrid(10000) As Integer
Public posGridXY_valid(10000) As Integer
Public nGoodCells As Integer
Public minGoodCellsPerImage As Integer
Public minGoodCellsPerWell As Integer
Public nGoodCellsPerWell As Integer

        
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
        offset As Long
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



Sub A_Setup()
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
End Sub

''''
'   AutoStore()
'   This was used to store values of the macro parameteres in the registry.
'   This could be used for saving of the macro and resuisng them but is not working in its present state.
''''
Public Sub AutoStore()
    Dim myKey As String
    Dim Success As Boolean
    Dim storeOK As Boolean
    Dim idx As Long
    Dim lockNo As Long
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    AutofocusForm.GetBlockValues
    storeOK = True
    myKey = "UI\" + GlobalMacroKey + "\AutoStore"
    Success = tools.RegExistKey(myKey)
    If Success Then
        Success = tools.RegDeleteKey(myKey)
    End If
    Success = tools.RegCreateKey(myKey)
'    tools.RegStringValue(myKey, "BlockAutoConfiguration") = BlockAutoConfiguration
    tools.RegLongValue(myKey, "BlockTimeIndex") = BlockTimeIndex
'    tools.RegLongValue(myKey, "BlockAutoConfigurationUse") = BlockAutoConfigurationUse
    tools.RegDoubleValue(myKey, "BlockHighSpeed") = BlockHighSpeed
    tools.RegDoubleValue(myKey, "BlockLowZoom") = BlockLowZoom
    tools.RegDoubleValue(myKey, "BlockHRZ") = BlockHRZ
   
    tools.RegDoubleValue(myKey, "BlockZOffset") = BlockZOffset
    tools.RegDoubleValue(myKey, "BlockZRange") = BlockZRange
    tools.RegDoubleValue(myKey, "BlockZStep") = BlockZStep
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

Public Sub ScanToImageNew(RecordingDoc As DsRecordingDoc) ' new routine to scan overwrite the same image, even with several z-slices
    'new changed on 30.05.2011 should then also scan and keep all the tracks...
    'Dim AcquisitionController As AimAcquisitionController40.AimScanController 'now public
    Dim ProgressFifo As IAimProgressFifo

    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Dim WasLocked As Double
    
    Dim gui As Object, treenode As Object
    Set gui = Lsm5.ViewerGuiServer
    
    If Not RecordingDoc Is Nothing Then
        Set treenode = RecordingDoc.RecordingDocument.image(0, True)
        'Set treenode = Lsm5.NewDocument
    
        Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
        AcquisitionController.DestinationImage(0) = treenode 'EngelImageToHechtImage(GlobalSingleImage).Image(0, True)
        AcquisitionController.DestinationImage(1) = Nothing
        Set ProgressFifo = AcquisitionController.DestinationImage(0)
        Lsm5.tools.CheckLockControllers True
        AcquisitionController.StartGrab eGrabModeSingle 'TODO why not use Lsm5.Start
        'Set RecordingDoc = Lsm5.StartScan
        If Not ProgressFifo Is Nothing Then ProgressFifo.Append AcquisitionController
    End If
    
End Sub


'''''
'  AutoRecall()
'  This was used to read values of the macro parameteres in the registry. This could be used for saving of the macro and resuisng them but
'  is not working in its present state.
'''''
Public Sub AutoRecall()
    Dim myKey As String
    Dim Success As Boolean
    Dim idx As Long
    Dim lockNo As Long
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString As String
'    Dim Position As Long
   ' Dim Range As Double
    
    myKey = "UI\" + GlobalMacroKey + "\AutoStore"
    Success = tools.RegExistKey(myKey)
    If Success Then
'        Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
'        If Position >= 0 Then
'            Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
'        Else
'            Range = 0#
'        End If
'substituted29.06.2010 by Function Range
    
'        BlockAutoConfiguration = tools.RegStringValue(myKey, "BlockAutoConfiguration")
        BlockTimeIndex = tools.RegLongValue(myKey, "BlockTimeIndex")
'        BlockAutoConfigurationUse = tools.RegLongValue(myKey, "BlockAutoConfigurationUse")
        BlockHighSpeed = tools.RegDoubleValue(myKey, "BlockHighSpeed")
        BlockLowZoom = tools.RegDoubleValue(myKey, "BlockLowZoom")
        HRZ = tools.RegDoubleValue(myKey, "BlockHRZ")
        
        BlockZOffset = tools.RegDoubleValue(myKey, "BlockZOffset")
        BlockZRange = tools.RegDoubleValue(myKey, "BlockZRange")
        BlockZStep = tools.RegDoubleValue(myKey, "BlockZStep")
        If BlockZRange > Range * 0.9 Then
            BlockZRange = Range * 0.9
        End If
        If Abs(BlockZOffset) > Range * 0.9 Then
            BlockZOffset = 0
        End If
        
        AutofocusForm.SetBlockValues
      
'        AutofocusForm.Re_Initialize
    Else
    End If
End Sub

''''''
'   CopyRecording(Destination As DsRecording, Source As DsRecording)
'   TODO: Do we need this function?
'''''
Public Sub CopyRecording(Destination As DsRecording, Source As DsRecording)
    Destination.Copy Source
    Destination.FramesPerStack = Source.FramesPerStack ' why only this
End Sub

'''''
' StoreAcquisitionParameters()
' stores the whole set of scan parameters.
' it uses GlobalBackupRecording and BackupRecording
' TODO: Why 2 backuprecording are needed
'''''''
Public Sub StoreAcquisitionParameters()
    Set GlobalBackupRecording = Lsm5.CreateBackupRecording
    Set BackupRecording = Lsm5.CreateBackupRecording
    CopyRecording GlobalBackupRecording, Lsm5.DsRecording ' TODO why do we need CopyRecording. This is done with CreateBackupRecording!
    CopyRecording BackupRecording, Lsm5.DsRecording
End Sub


''''''
'   RestoreAcquisitionParameters()
'   Restores the image acquisition recording parameters from GlobalBackupRecording
'   Lsm5.DsRecording Out - Recording setting
''''''
Public Sub RestoreAcquisitionParameters()
     CopyRecording Lsm5.DsRecording, GlobalBackupRecording
End Sub

'''''
'   SystemVersionOffset()
'   Calculate an offset added to z-stack changes
'       [GlobalCorrectionOffset] Global Out - Offset added to shift in zStack
'   TODO: Do we still need it. Only for Axioskop does the Offset change
'''''
Public Sub SystemVersionOffset()
    SystemVersion = Lsm5.Info.VersionIs
    If StrComp(SystemVersion, "2.8", vbBinaryCompare) >= 0 Then
        If Lsm5.Info.IsAxioskop Then
            If BlockHighSpeed Then
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
            If BlockHighSpeed Then
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
' Autofocus_StackShift ( ZRange As Double, ZStep As Double, HighSpeed As Boolean, ZOffset As Double, NewPicture As DsRecordingDoc )
' Performs image scan, calculation of signal centroid (mass) and assign the
' global variables [ZShift] (LineScan) + [XShift] and [YShift] (FrameScan). This function does not change the focus just computes it
'       [ZRange]    In/Out  - The range in um over which the scan is made. Changed if to big
'       [ZStep]     In      - ZStep size in um
'       [HighSpeed] In      - Corresponds to BlockHighSpeed. Corresponds Range to MaxSpeed toggle
'       [ZOffset]   In/Out  - zOffset is checked weather it fits the
'       [NewPicture] In/Out - Contains the image
'''''''
Public Function Autofocus_StackShift(ZRange As Double, ZStep As Double, HighSpeed As Boolean, ZOffset As Double, NewPicture As DsRecordingDoc) As Boolean
    Dim BigZStep As Double
    Dim locStep As Double
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    If NewPicture Is Nothing Then
        Set NewPicture = Lsm5.NewScanWindow
        While NewPicture.IsBusy
            Sleep (100)
            DoEvents
        Wend
    End If

    AutofocusForm.ActivateAutofocusTrack HighSpeed
    
    If Not IsAutofocusTrackSelected Then
        MsgBox "No track selected for Autofocus! Cannot Autofocus!"
        ScanStop = True
        Autofocus_StackShift = False
        Exit Function
    End If
    DoEvents
    
    If Range() = 0 Then
        MsgBox "Objective's working distance not defined! Cannot Autofocus!"
        Exit Function
    End If
    If ZRange > Range() * 0.9 Then 'this is already tested in the slider could be removed
        ZRange = Range * 0.9
        MsgBox "Autofocus range is too large! Has been reduced to " + Str(ZRange)
    End If
    If Abs(ZOffset) > Range() * 0.9 Then 'this is already tested in the slider could be removed
        ZOffset = 0
        MsgBox "ZOffset has to be less than the working distance of the objective: " + CStr(Range) + " um"
    End If
    SystemVersionOffset

    'Now this is a code specific for the DoAutofocus (is not in the SetAutofocus).
    'This is to move to the offset position, with the focuswheel
    
    If Not GettingZmap Then DisplayProgress "Autofocus 1", RGB(0, &HC0, 0)       'I added this at some points for troubleshooting.
    
    Lsm5.Hardware.CpHrz.Position = 0  ' center the piezo focus
  
    
ZStackagain: 'this refers to goto lines (not used anymore)

    Zbefore = Lsm5.Hardware.CpFocus.Position        'To remember the position of the focuswheel
    'Lsm5.DsRecording.SpecialScanMode = "ZScanner" ' is taken care of in AutofocusForm.AutofocusSetting
   
    If ZOffset <= Range * 0.9 Then
        
        Lsm5.Hardware.CpFocus.Position = Zbefore - ZOffset + GlobalCorrectionOffset + ZBacklash 'Move down 50um (=ZBacklash) below the position of the offset
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
           Sleep (20)  '20ms
           DoEvents
        Loop
        Lsm5.Hardware.CpFocus.Position = Zbefore - ZOffset + GlobalCorrectionOffset            'Moves up to the position of the offset
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
           Sleep (20)
           DoEvents
        Loop
    
    End If

    'Lsm5.DsRecording.FrameSpacing = ZStep
    'NoFrames = CLng(ZRange / ZStep) + 1                     'Calculates the number of frames per stack. Clng converts it to a long and rounds up the fraction
    'Lsm5.DsRecording.FramesPerStack = NoFrames
    'If NoFrames > 2048 Then                                 'overwrites the userdefined value if too many frames have been defined by the user
    '    NoFrames = 2048
    'End If
    'Lsm5.DsRecording.Sample0Z = ZStep * NoFrames / 2        'Distance of the actual focus to the first Z position of the image (or line) to acquire in the stack.
                                                            'I think this is only valid for the focus wheel and not the HRZ
    
    AutofocusForm.AutofocusSetting HRZ, BlockHighSpeed, BlockZStep

    Lsm5.DsRecording.FrameSpacing = ZStep
    NoFrames = CLng(ZRange / ZStep) + 1                     'Calculates the number of frames per stack. Clng converts it to a long and rounds up the fraction
    Lsm5.DsRecording.FramesPerStack = NoFrames
    If NoFrames > 2048 Then                                 'overwrites the userdefined value if too many frames have been defined by the user
        NoFrames = 2048
    End If
    Lsm5.DsRecording.Sample0Z = ZStep * NoFrames / 2        'Distance of the actual focus to the first Z position of the image (or line) to acquire in the stack.
    locStep = Lsm5.DsRecording.FrameSpacing
    
    ' check that the ZStep has been set correctly otherwise remove on the fly = Fast Z line. Microscope with Fast Zline can nonly make small number of steps
    If HRZ = False Then
        If ZStep > Round(Lsm5.DsRecording.FrameSpacing, 3) Then
            DisplayProgress "Highest Z Step with no piezo and Fast Z line " + CStr(Round(Lsm5.DsRecording.FrameSpacing, 3)) + " um. Autofocus uses slower Focus Step", RGB(&HC0, &HC0, 0)
            Lsm5.DsRecording.SpecialScanMode = "FocusStep"
            Lsm5.DsRecording.FrameSpacing = ZStep
        End If
    End If
    '!!!!!!!!!!!!!!!!!!!!!! potential error source!!!!!!!!!!!!!!!!!!
    If PubSearchScan Then ' todo: what is this?
        
        '  BigZStep = Range * 0.7 / 200
        If HRZ And SystemName = "LIVE" Then

        '   If Range > 1000 Then Range = 600  deleted 30.06.2010

            Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
            Lsm5.DsRecording.FramesPerStack = 1201
            Lsm5.DsRecording.Sample0Z = Range / 2
            Lsm5.DsRecording.FrameSpacing = Range / 1200
            Sleep (100)
        
        Else
        
            BigZStep = Range * 0.7 / 200
            Lsm5.DsRecording.SpecialScanMode = "FocusStep"
            NoFrames = CLng(Range * 0.7 / BigZStep) + 1
            Lsm5.DsRecording.FramesPerStack = NoFrames
            Lsm5.DsRecording.FrameSpacing = BigZStep
            Lsm5.DsRecording.Sample0Z = BigZStep * NoFrames / 2
            Sleep (20)
            
        End If
    
    End If
        
    ' Here the Stack is acquired ***
    'DisplayProgress "Acquiring AF stack...", RGB(&HC0, 0, 0)
    'Set NewPicture = Lsm5.StartScan
    ScanToImageNew NewPicture
    
    While AcquisitionController.IsGrabbing
        Sleep (100)
        DoEvents
        If ScanStop Then
            AutofocusForm.StopAcquisition
            Autofocus_StackShift = False
            Exit Function
        End If
    Wend
    
    'Lsm5.tools.WaitForScanEnd False, 20
    ' ******************************
    

    If Not GettingZmap Then DisplayProgress "Autofocus 6", RGB(0, &HC0, 0)
    
    AutofocusForm.MassCenter ("Autofocus")
    
    If AreStageCoordinateExchanged Then
           XShift = YMass
           YShift = XMass
    Else
           XShift = -XMass
           YShift = YMass
    End If
    
    ZShift = ZMass
    
    
    'check if Z shift makes sense
    If PubSearchScan = True Then Exit Function ' TODO What is this?
    Autofocus_StackShift = True

End Function


''''''
'   Autofocus_MoveAcquisition(ZOffset As Double)
'   Add offset to z determined by the Autofocus. This uses stage stepping
'       [ZOffset] In - Value of ZOffset in um
''''''
Public Sub Autofocus_MoveAcquisition(ZOffset As Double)
    
    Dim NoZStack As Boolean
    Const ZBacklash = -50   'why do we need this. TODO?
    Dim ZFocus As Double
    Dim Zbefore As Double
    Dim x As Double
    Dim y As Double
    
    RestoreAcquisitionParameters
    DoEvents ' this releases window to push some butttons
    
    AutofocusForm.ActivateAcquisitionTrack ' Check if Acquisition
    If Lsm5.DsRecording.ScanMode = "ZScan" Or Lsm5.DsRecording.ScanMode = "Stack" Then  'Looks if a Z-Stack is going to be acquired
        NoZStack = False
    Else
        NoZStack = True
    End If
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
    If FrameAutofocussing Then
        ' Todo: check whether it moves in the correct direction
        x = Lsm5.Hardware.CpStages.PositionX - XShift
        y = Lsm5.Hardware.CpStages.PositionY - YShift
        
        Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
         
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
    
    ' center all z-stacks again!
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.FrameSpacing * Int(Lsm5.DsRecording.FramesPerStack / 2)
    
    DisplayProgress "Autofocus 14", RGB(0, &HC0, 0)
    'Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
    DisplayProgress "Autofocus 15", RGB(0, &HC0, 0)

End Sub

Private Sub MovetoCorrectZPosition(ZOffset As Double)
Const ZBacklash = -50
Dim ZFocus As Double
Dim Zbefore As Double
Dim x As Double
Dim y As Double
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
''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the HRZ can acces the position, and also to wait it is done...
        Sleep (100)
        DoEvents
End Sub


Public Sub Autofocus_MoveAcquisition_HRZ(ZOffset As Double)
    Dim NoZStack As Boolean
    Const ZBacklash = -50
    Dim ZFocus As Double
    Dim Zbefore As Double
    Dim x As Double
    Dim y As Double

    RestoreAcquisitionParameters
    
    Set GlobalBackupRecording = Nothing
    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
    AutofocusForm.ActivateAcquisitionTrack
    If Lsm5.DsRecording.ScanMode = "ZScan" Or Lsm5.DsRecording.ScanMode = "Stack" Then  'Looks if a Z-Stack is going to be acquired
        NoZStack = False
    Else
        NoZStack = True
    End If

    'Moving to the correct position in Z
    If HRZ And NoZStack Then                                            'If using HRZ for autofocusing and there is no Zstack for image acquisition
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
        If HRZ Then                             ' Now I'm not sure with the signs and... I some point I just tried random combinations...
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
 
    If FrameAutofocussing Then
        x = Lsm5.Hardware.CpStages.PositionX - XShift  'the fact that it is "-" in this line and "+" in the next line  probably has to do with where the XY of the origin is set (top right corner and not botom left, I think)
        y = Lsm5.Hardware.CpStages.PositionY - YShift
        Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
         
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



Public Sub PutStagePositionsInArray()
    ReDim GlobalXpos(GlobalPositionsStage)
    ReDim GlobalYpos(GlobalPositionsStage)
    ReDim GlobalZpos(GlobalPositionsStage)
    For idpos = 0 To GlobalPositionsStage - 1
    Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).GetMarkZ idpos, GlobalXpos(idpos + 1), GlobalYpos(idpos + 1), GlobalZpos(idpos + 1)
    '           GlobalXpos(idpos) = Lsm5.Hardware.CpStages.PositionX
    '           GlobalYpos(idpos) = Lsm5.Hardware.CpStages.PositionY
    '           GlobalZpos(idpos) = Lsm5.Hardware.CpFocus.Position
                
    Next idpos
End Sub
