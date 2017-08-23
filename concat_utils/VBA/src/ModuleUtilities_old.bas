Attribute VB_Name = "ModuleUtilities"
Option Explicit

'''''''''
'Minimize button for Macro window
''''''
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
 
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'''''''''


Public Const WM_COMMAND = &H111

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


Public Const OFS_MAXPATHNAME = 128
Public Const OF_EXIST = &H4000

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

Public Const eEventFocus = 5
Public Const eEventStage = 6

Public Const eEventUpdate = 25


Public GlobalSystemVersion As Long
Public GlobalMacroVersion As String

Public GlobalPath As String
Public GlobalMacrosPath As String

Public GlobalProjectName As String
Public GlobalHelpName As String
Public GlobalHelpNamePDF As String
Public GlobalHelpName1 As String
Public GlobalHelpNamePDF1 As String
Public GlobalHelpName2 As String
Public GlobalHelpNamePDF2 As String
Public GlobalHelpName3 As String
Public GlobalHelpNamePDF3 As String
Public GlobalHelpName4 As String
Public GlobalHelpNamePDF4 As String
Public GlobalHelpName5 As String
Public GlobalHelpNamePDF5 As String
Public GlobalHelpName6 As String
Public GlobalHelpNamePDF6 As String
Public GlobalHelpName7 As String
Public GlobalHelpNamePDF7 As String
Public GlobalHelpName8 As String
Public GlobalHelpNamePDF8 As String
Public GlobalHelpName9 As String
Public GlobalHelpNamePDF9 As String
Public GlobalHelpName10 As String
Public GlobalHelpNamePDF10 As String
Public GlobalHelpName11 As String
Public GlobalHelpNamePDF11 As String
Public GlobalHelpName12 As String
Public GlobalHelpNamePDF12 As String
Public GlobalHelpName14 As String
Public GlobalHelpNamePDF14 As String
Public GlobalHelpNamePDF15 As String
Public GlobalHelpName15 As String

Public GlobalHelpNamePDF16 As String
Public GlobalHelpNamePDF17 As String
Public GlobalHelpNamePDF18 As String
Public GlobalHelpNamePDF19 As String
Public GlobalHelpNamePDF20 As String
Public GlobalHelpNamePDF21 As String
Public GlobalHelpNamePDF22 As String
Public GlobalHelpNamePDF23 As String
Public GlobalHelpNamePDF24 As String
Public GlobalHelpNamePDF25 As String
Public GlobalHelpNamePDF26 As String
Public GlobalHelpNamePDF27 As String
Public GlobalHelpNamePDF28 As String
Public GlobalHelpNamePDF29 As String
Public GlobalHelpNamePDF30 As String

Public GlobalErrorFile As String
Public GlobalTimelineFile As String

Public GlobalHelpNameScale As String
Public GlobalMacroKey As String

Public GlobalAutoStoreKey As String

Public GlobalIsStage As Boolean
Public tools As Lsm5Tools
Public Stage As CpStages
Public GlobalOptions As Lsm5Options

Public ScanInterrupt As Boolean

Public flgUserChange As Boolean
Public User_flg As Boolean

Public flgEvent As Integer

Public GlobalIsFRET As Boolean
Public GlobalPi As Double

Public GlobalProgressString As String
Public GlobalColor As Long

Public GlobalRecallLocations As Boolean

Public GlobalSampleObservationTime(13) As Double
Public GlobalIsDSP As Boolean

Public GlobalStageCounter As Long
Public GlobalStageText As String

Public GlobalSystemGroup As String
Public GlobalIsDuo As Boolean

Public X11 As Double
Public X12 As Double
Public X21 As Double
Public X22 As Double

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


Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
(ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long



Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function RegOpenKeyEx _
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
    
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Any) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Any) As Long
Public Declare Function GetModuleHandle Lib "kernel32" (ByVal lpModuleName As String) As Long
Public Declare Function SetWindowLong Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOPMOST = &H8&

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Sub FormatUserForm(UserFormCaption As String)

    Dim hWnd            As Long
    Dim exLong          As Long

    hWnd = FindWindowA(vbNullString, UserFormCaption)
    exLong = GetWindowLongA(hWnd, -16)
    If (exLong And &H20000) = 0 Then
        SetWindowLongA hWnd, -16, exLong Or &H20000
    Else
    End If

End Sub

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function



Public Sub BleachROI(Optional opt As Boolean)

'The first example generates a rectangular region and bleaches this region.

    Dim Left As Long
    Dim Right As Long
    Dim Top As Long
    Dim Bottom As Long
    Dim Width As Long
    Dim Height As Long
    Dim Row As Long
    Dim Column As Long
    Dim BytesPerRow As Long

'The bleach rectangle

    Left = 10
    Top = 10
    Width = 100
    Height = 100

' 32-bit alignment
    BytesPerRow = ((Width + 31) / 32) * 4

'Generate the bitmask memory
    Dim Mask() As Byte
    ReDim Mask(BytesPerRow * Height)
    
'Fill the bitmask
    For Row = 0 To Height - 1
    For Column = 0 To BytesPerRow - 1
    Mask(Row * BytesPerRow + Column) = 255
    Next Column
    Next Row

'Transfer the bitmask to the scan-controller
    Lsm5.ExternalCpObject.pHardwareObjects.pScanController.SetBleachRoi _
Left, Top, Width, Height, CVar(Mask)

'Start the bleach - Use the scan-controller and not "Lsm5.Bleach" cause
'the latter would overwrite the region with the
'region currently stored in the "DS" - The first argument has no meaning.
'The second argument must be "0".

    Lsm5.ExternalCpObject.pHardwareObjects.pScanController.Bleach 0, 0

End Sub


Public Sub ReadMask(Optional opt As Boolean)

'The second example uses the region from the vector overlay. The vector
'overlay returns a mask with one byte per
'pixel. One has to convert to the format "one bit per pixel" which the
'scan-controller accepts.

    Dim IndexSource As Long
    Dim IndexDestination As Long
    Dim MaskByte As Byte
    Dim Factor As Long
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    Dim Width As Long
    Dim Height As Long
    Dim MaskSource() As Byte
    Dim BytesPerRow As Long
    Dim Row As Long
    Dim Column As Long
'Get current overlay mask
    MaskSource(0) = _
Lsm5.DsRecordingActiveDocObject.VectorOverlay.MakeRoiMask(Left, Top, _
Right, Bottom, 0, 0, Lsm5.DsRecordingActiveDocObject.GetDimensionX, _
Lsm5.DsRecordingActiveDocObject.GetDimensionY, 1)

    Width = Right - Left
    Height = Bottom - Top
'32 bit alignment (scan controller wants it - but vector overlay mask
' requires no special aligment)
    BytesPerRow = ((Width + 31) / 32) * 4

'Create memory for the bitmask
    Dim MaskDestination() As Byte
    ReDim MaskDestination(BytesPerRow * Height)
       
'Convert One-byte-per-pixel to One-bit-per-pixel

    IndexSource = 0
    IndexDestination = 0
    For Row = 0 To Height - 1
        Factor = 1
        MaskByte = 0
        For Column = 0 To Width - 1
            MaskByte = MaskByte + Factor * MaskSource(IndexSource)
            IndexSource = IndexSource + 1
            Factor = Factor * 2
            If (Factor > 255) Then
                MaskDestination(IndexDestination) = MaskByte
                IndexDestination = IndexDestination + 1
                Factor = 1
                MaskByte = 0
            End If
        Next Column
        If (Factor > 1) Then
            MaskDestination(IndexDestination) = MaskByte
            IndexDestination = IndexDestination + 1
        End If
    Next Row

'Transfer the bitmask to the scan-controller
    Lsm5.ExternalCpObject.pHardwareObjects.pScanController.SetBleachRoi _
Left, Top, Width, Height, CVar(MaskDestination)

'Start the bleach
    Lsm5.ExternalCpObject.pHardwareObjects.pScanController.Bleach 0, 0
    
End Sub
    
Public Sub DisplayHelp(HelpNamePDF As String, HelpName As String)
    Dim dblTask As Double
    Dim MacroPath As String
    Dim MyPath As String
    Dim MyPathPDF As String
    
    Dim bslash As String
    Dim Success As Integer
    Dim pos As Integer
    Dim Start As Integer
    Dim Count As Long
    Dim ProjName As String
    Dim indx As Integer
    Dim AcrobatObject As Object
    Dim AcrobatViewer As Object
    Dim OK As Boolean
    Dim StrPath As String
    Dim ExecName As String
    
    Count = ProjectCount()
    For indx = 0 To Count - 1
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
            MyPath = MyPath + HelpName
            OK = False
            On Error GoTo RTFhelp
            OK = FServerFromDescription("AcroExch.Document", StrPath, ExecName)
            dblTask = Shell(ExecName + " " + MyPathPDF, vbNormalFocus)
            
'            Set AcrobatViewer = CreateObject("AcroExch.app")
'            If Not AcrobatViewer Is Nothing Then
'                Set AcrobatObject = CreateObject("AcroExch.AVDoc")
'                If Not AcrobatObject Is Nothing Then
'                    OK = AcrobatViewer.Show
'                    If OK Then
'                        OK = AcrobatObject.Open(MyPathPDF, MyPathPDF)
'                    End If
'                Else
'                    OK = False
'                End If
'            Else
'                OK = False
'            End If
RTFhelp:
            If Not OK Then
                MsgBox "Install Acrobat Viewer!"
'                dblTask = Shell("C:\Program Files\Windows NT\Accessories\wordpad.exe " + MyPath, vbNormalFocus)
            End If
            Exit For
        End If
    Next indx
End Sub

    
    
Function FServerFromDescription(strName As String, StrPath As String, ExecName As String) As Boolean
    Dim lngResult As Long
    Dim strTmp As String
    Dim hKeyServer As Long
    Dim strBuffer As String
    Dim cb As Long
    Dim i As Integer
    
    FServerFromDescription = False
    
    strTmp = VBA.Space(255)
    strTmp = strName + "\CLSID"
    lngResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTmp, 0&, KEY_READ, hKeyServer)
    
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
    strBuffer = VBA.Space(255)
    cb = Len(strBuffer)
    
    lngResult = RegQueryValueEx(hKeyServer, "", 0&, REG_SZ, ByVal strBuffer, cb)
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
    
    lngResult = RegCloseKey(hKeyServer)
    strTmp = VBA.Space(255)
    strTmp = "CLSID\" + Strings.Left(strBuffer, cb - 1) + "\LocalServer32"
    strBuffer = VBA.Space(255)
    cb = Len(strBuffer)
    lngResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTmp, 0&, KEY_READ, hKeyServer)
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
        
    lngResult = RegQueryValueEx(hKeyServer, "", 0&, REG_SZ, ByVal strBuffer, cb)
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
    StrPath = Strings.Left(strBuffer, cb - 1)
    ExecName = StrPath
    lngResult = RegCloseKey(hKeyServer)
    
    i = Len(StrPath)
    
    Do Until (i = 0)
        If (VBA.Mid(StrPath, i, 1) = "\") Then
            StrPath = Strings.Left(StrPath, i - 1)
            FServerFromDescription = True
            Exit Do
        End If
        i = i - 1
    Loop

error_exit:
    If (Not hKeyServer = 0) Then lngResult = RegCloseKey(hKeyServer)

End Function


Public Sub CopyRecording(destination As DsRecording, Source As DsRecording)
    Dim Ts As DsTrack
    Dim Td As DsTrack
    Dim DataS As DsDataChannel
    Dim DataD As DsDataChannel
    Dim DetS As DsDetectionChannel
    Dim DetD As DsDetectionChannel
    Dim IlS As DsIlluminationChannel
    Dim IlD As DsIlluminationChannel
    Dim BS As DsBeamSplitter
    Dim BD As DsBeamSplitter
    Dim lT As Long
    Dim lI As Long
    Dim Success As Integer
   
    '''''''''''''''''''''''''''start inserted lines
    destination.Copy Source
    
    destination.SpecialScanMode = Source.SpecialScanMode
    destination.ScanMode = Source.ScanMode
    
    For lI = 1 To destination.TrackCount - Source.TrackCount
        destination.TrackRemove destination.TrackCount - lI
    Next lI
    
    For lI = 1 To destination.LaserCount - Source.LaserCount
        destination.LaserRemove destination.LaserCount - lI
    Next lI
    
    For lI = 1 To destination.TimersCount - Source.TimersCount
        destination.TimersRemove destination.TimersCount - lI
    Next lI
    
    For lI = 1 To destination.MarkersCount - Source.MarkersCount
        destination.MarkersRemove destination.MarkersCount - lI
    Next lI
    
    For lT = 0 To destination.TrackCount - 1
        If lT < Source.TrackCount Then
            Set Ts = Source.TrackObjectByIndex(lT, Success)
            Set Td = destination.TrackObjectByIndex(lT, Success)
            
            For lI = 1 To Td.BeamSplitterCount - Ts.BeamSplitterCount
                Td.BeamSplitterRemove Td.BeamSplitterCount - lI
            Next lI
                        
            For lI = 1 To Td.DataChannelCount - Ts.DataChannelCount
                Td.DataChannelRemove Td.DataChannelCount - lI
            Next lI
            
            For lI = 1 To Td.DetectionChannelCount - Ts.DetectionChannelCount
                Td.DetectionChannelRemove Td.DetectionChannelCount - lI
            Next lI
            
            For lI = 1 To Td.IlluminationChannelCount - Ts.IlluminationChannelCount
                Td.IlluminationChannelRemove Td.IlluminationChannelCount - lI
            Next lI
            
            
        End If
    Next lT

   '''''''''''''''''''''''''''end inserted lines
    If GlobalSystemVersion >= 30 Then
        NewCopyRecording destination, Source
    Else
        OldCopyRecording destination, Source
    End If
End Sub

Public Sub NewCopyRecording(destination As DsRecording, Source As DsRecording)
    Dim Ts As DsTrack
    Dim Td As DsTrack
    Dim DataS As DsDataChannel
    Dim DataD As DsDataChannel
    Dim DetS As DsDetectionChannel
    Dim DetD As DsDetectionChannel
    Dim IlS As DsIlluminationChannel
    Dim IlD As DsIlluminationChannel
    Dim BS As DsBeamSplitter
    Dim BD As DsBeamSplitter
    Dim lT As Long
    Dim lI As Long
    Dim Success As Integer
    
        destination.Copy Source
        For lT = 0 To destination.TrackCount - 1
        
            Set Ts = Source.TrackObjectByIndex(lT, Success)
            Set Td = destination.TrackObjectByIndex(lT, Success)
            Td.DataChannelCount
        Next lT

        destination.Objective = Source.Objective
        For lT = 0 To destination.TrackCount - 1
        
            Set Ts = Source.TrackObjectByIndex(lT, Success)
            Set Td = destination.TrackObjectByIndex(lT, Success)
            
            Td.Collimator1Value = Ts.Collimator1Value
            Td.Collimator2Value = Ts.Collimator2Value
            Td.SpiCenterWavelength = Ts.SpiCenterWavelength
            
            For lI = 0 To Td.DataChannelCount - 1
                Set DataS = Ts.DataChannelObjectByIndex(lI, Success)
                Set DataD = Td.DataChannelObjectByIndex(lI, Success)
'                DataD.ColorRef = DataS.ColorRef
            Next lI
            
            For lI = 0 To Td.DetectionChannelCount - 1
                Set DetS = Ts.DetectionChannelObjectByIndex(lI, Success)
                Set DetD = Td.DetectionChannelObjectByIndex(lI, Success)
                DetD.Filter1 = DetS.Filter1
                DetD.Filter2 = DetS.Filter2
                DetD.DetectorGain = DetS.DetectorGain
                DetD.AmplifierGain = DetS.AmplifierGain
                DetD.AmplifierOffset = DetS.AmplifierOffset
                DetD.PinholeDiameter = DetS.PinholeDiameter
                DetD.DetectorGainABC1 = DetS.DetectorGainABC1
                DetD.DetectorGainABC2 = DetS.DetectorGainABC2
                DetD.AmplifierGainABC1 = DetS.AmplifierGainABC1
                DetD.AmplifierGainABC2 = DetS.AmplifierGainABC2
                DetD.AmplifierOffsetABC1 = DetS.AmplifierOffsetABC1
                DetD.AmplifierOffsetABC2 = DetS.AmplifierOffsetABC2
                DetD.SpiWavelengthStart1 = DetS.SpiWavelengthStart1
                DetD.SpiWavelengthEnd1 = DetS.SpiWavelengthEnd1
                DetD.SpiWavelengthStart2 = DetS.SpiWavelengthStart2
                DetD.SpiWavelengthEnd2 = DetS.SpiWavelengthEnd2
                DetD.SpiSpectralScanChannels = DetS.SpiSpectralScanChannels
                
            Next lI
            
            For lI = 0 To Td.IlluminationChannelCount - 1
                Set IlS = Ts.IlluminationObjectByIndex(lI, Success)
                Set IlD = Td.IlluminationObjectByIndex(lI, Success)
                IlD.Acquire = IlS.Acquire
                IlD.Power = IlS.Power
                IlD.DetectionChannelName = IlS.DetectionChannelName
                IlD.PowerABC1 = IlS.PowerABC1
                IlD.PowerABC2 = IlS.PowerABC2
            Next lI
            
            For lI = 0 To Td.BeamSplitterCount - 1
                Set BS = Ts.BeamSplitterObjectByIndex(lI, Success)
                Set BD = Td.BeamSplitterObjectByIndex(lI, Success)
                If Success Then
                    BD.Filter = BS.Filter
                End If
            Next lI
            
        Next lT

End Sub

Public Sub OldCopyRecording(destination As DsRecording, Source As DsRecording)
    Dim Ts As DsTrack
    Dim Td As DsTrack
    Dim DataS As DsDataChannel
    Dim DataD As DsDataChannel
    Dim DetS As DsDetectionChannel
    Dim DetD As DsDetectionChannel
    Dim IlS As DsIlluminationChannel
    Dim IlD As DsIlluminationChannel
    Dim BS As DsBeamSplitter
    Dim BD As DsBeamSplitter
    Dim lT As Long
    Dim lI As Long
    Dim Success As Integer

        destination.Copy Source
        destination.Objective = Source.Objective
        For lT = 0 To destination.TrackCount - 1
        
            Set Ts = Source.TrackObjectByIndex(lT, Success)
            Set Td = destination.TrackObjectByIndex(lT, Success)
            
'            TD.Collimator1Position = TS.Collimator1Position
'            TD.Collimator2Position = TS.Collimator2Position
            
            For lI = 0 To Td.DataChannelCount - 1
                Set DataS = Ts.DataChannelObjectByIndex(lI, Success)
                Set DataD = Td.DataChannelObjectByIndex(lI, Success)
                DataD.ColorRef = DataS.ColorRef
            Next lI
            
            For lI = 0 To Td.DetectionChannelCount - 1
                Set DetS = Ts.DetectionChannelObjectByIndex(lI, Success)
                Set DetD = Td.DetectionChannelObjectByIndex(lI, Success)
                DetD.Filter1 = DetS.Filter1
                DetD.Filter2 = DetS.Filter2
                DetD.DetectorGain = DetS.DetectorGain
                DetD.AmplifierGain = DetS.AmplifierGain
                DetD.AmplifierOffset = DetS.AmplifierOffset
                DetD.PinholeDiameter = DetS.PinholeDiameter
            Next lI
            
            For lI = 0 To Td.IlluminationChannelCount - 1
                Set IlS = Ts.IlluminationObjectByIndex(lI, Success)
                Set IlD = Td.IlluminationObjectByIndex(lI, Success)
                IlD.Acquire = IlS.Acquire
                IlD.Power = IlS.Power
                IlD.DetectionChannelName = IlS.DetectionChannelName
            Next lI
            
            For lI = 0 To Td.BeamSplitterCount - 1
                Set BS = Ts.BeamSplitterObjectByIndex(lI, Success)
                Set BD = Td.BeamSplitterObjectByIndex(lI, Success)
                BD.Filter = BS.Filter
            Next lI
            
        Next lT


End Sub


Public Sub CheckDiskSpace(lpRootPathName As String, lFreeSpace As Double, lSpace As Long)
    Dim lpSectorsPerCluster As Long
    Dim lpBytesPerSector As Long
    Dim lpNumberOfFreeClusters As Long
    Dim lpTotalNumberOfClusters As Long
    
    lSpace = GetDiskFreeSpace(lpRootPathName, lpSectorsPerCluster, lpBytesPerSector, _
                            lpNumberOfFreeClusters, lpTotalNumberOfClusters)
    lFreeSpace = CDbl(lpSectorsPerCluster) * CDbl(lpBytesPerSector) * CDbl(lpNumberOfFreeClusters)

End Sub

Public Sub AutofocusRange(ZRange As Double, OffRange As Double, MsgZRange As String, MsgOffRange As String)

    Dim Position As Long
    Dim Range As Double

    Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
    If Position >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
    Else
        Range = 0#
    End If
    OffRange = Range
    ZRange = Range
    MsgOffRange = "Zoffset has to be less or equal to the working distance of the objective: " + CStr(OffRange) + " um"
    MsgZRange = "ZRange has to be less or equal to the working distance of the objective: " + CStr(ZRange) + " um"

End Sub

Public Sub MoveToLoc(x As Double, y As Double)
    Dim X1 As Double
    Dim Y1 As Double
    Dim cnt As Long
    Dim WaitTime As Long
    Dim PreviousX As Double
    Dim PreviousY As Double
    Dim CurrentX As Double
    Dim CurrentY As Double
    Dim Msg, Style, Title, Response
    Dim StageTime As Long
    
    cnt = 0
    On Error GoTo retry
retry:
    If cnt > 1000 Then GoTo Finish
    cnt = cnt + 1
    
    X1 = Stage.PositionX
    Y1 = Stage.PositionY
    Stage.PositionX = x
    Stage.PositionY = y
    
        PreviousX = X1
        PreviousY = Y1
        WaitTime = 0
        
        While Lsm5.Hardware.CpStages.IsBusy
            DoEvents
            Sleep 100
            If ScanInterrupt Then Exit Sub
            WaitTime = WaitTime + 1
            If WaitTime > 100 Then
                CurrentX = Stage.PositionX
                CurrentY = Stage.PositionY
            
                If (Abs(CurrentX - PreviousX) < 1) And (Abs(CurrentY - PreviousY) < 1) Then
'                    StageTime = GetTickCount
                
                    If GlobalStageCounter = 1 Then
                        Msg = "The stage did not move for more than 10 seconds:" _
                        + Strings.Chr(13) + Strings.Chr(10) + "Instance " + CStr(GlobalStageCounter) _
                        + ", Time = " + Strings.Format(Now, "hh:mm:ss mm/dd/yy") + Strings.Chr(13) + Strings.Chr(10) + _
                        "X = " + CStr(CurrentX) + Strings.Chr(13) + Strings.Chr(10) + "Y = " + CStr(CurrentY)
                        GlobalStageText = Msg
                    Else
                        Msg = Strings.Chr(13) + Strings.Chr(10) + "Instance " + CStr(GlobalStageCounter) _
                        + ", Time = " + Strings.Format(Now, "hh:mm:ss mm/dd/yy") + Strings.Chr(13) + Strings.Chr$(10) + _
                        "X = " + CStr(CurrentX) + Strings.Chr(13) + Strings.Chr(10) + "Y = " + CStr(CurrentY)
                        GlobalStageText = GlobalStageText + Msg

                    End If
'                    Print #GlobaliFile, GlobalStageText

                    GlobalStageCounter = GlobalStageCounter + 1
                    Style = vbInfoText + VbExclamation ' Define buttons.
                    Title = "Stage Did't Move"  ' Define title.
'                    Response = MsgBox(Msg, Style, Title)
''                    StageErrorForm.TextBox1.Text = GlobalStageText
''                    StageErrorForm.Show 0
                
'                    ScanInterrupt = True
'                    Exit Sub
                End If
                PreviousX = CurrentX
                PreviousY = CurrentY
                WaitTime = 0
            End If
        Wend
Finish:
'        UpdateCurrPos StartStop.LabelFocus, StartStop.LabelCurrentPos

End Sub


Public Sub ReadLoc(x As Double, y As Double)
    Dim cnt As Long
    
    cnt = 0
    On Error GoTo retry
retry:
    If cnt > 1000 Then GoTo Finish
    cnt = cnt + 1
    x = Stage.PositionX
    y = Stage.PositionY
Finish:
End Sub

Public Sub MoveToLocUsingScreenCoordinates(x As Double, y As Double)
    Dim X1 As Double
    Dim Y1 As Double
    Dim cnt As Long
    Dim WaitTime As Long
    Dim PreviousX As Double
    Dim PreviousY As Double
    Dim CurrentX As Double
    Dim CurrentY As Double
        
    cnt = 0
    On Error GoTo retry
retry:
    If cnt > 1000 Then GoTo Finish
    cnt = cnt + 1
        
    X1 = GetStagePositionX
    Y1 = GetStagePositionY
    
    SetStagePositionXY x, y
    
       
    PreviousX = X1
    PreviousY = Y1
    WaitTime = 0
    
    While Lsm5.Hardware.CpStages.IsBusy
        DoEvents
        Sleep 100
        If ScanInterrupt Then Exit Sub
        WaitTime = WaitTime + 1
        If WaitTime > 50 Then
            CurrentX = GetStagePositionX
            CurrentY = GetStagePositionY
        
            If (Abs(CurrentX - PreviousX) < 1) And (Abs(CurrentY - PreviousY) < 1) Then
                MsgBox "The stage did not move for more than 5 seconds." + Chr(13) + Chr(10) + _
                       "Assuming that end of the stage positionning range has been reached.", VbExclamation
                ScanInterrupt = True
                Exit Sub
            End If
            PreviousX = CurrentX
            PreviousY = CurrentY
            WaitTime = 0
        End If
    Wend
Finish:
End Sub

Public Sub WaitWhileImageBusy(Image As DsRecordingDoc)
If Not Image Is Nothing Then
    Do While Image.IsBusy
        DoEvents
        Sleep 20
        If ScanInterrupt Then Exit Sub
    Loop
End If
End Sub
Public Function GetStagePositionX() As Double

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim x As Double
    Dim y As Double
    DoEvents
    GetStagePositionX = 0#
    Dim cnt As Long
    
    cnt = 0
    On Error GoTo retry
retry:
    If cnt > 1000 Then GoTo nostage
    cnt = cnt + 1
        
    x = Lsm5.Hardware.CpStages.PositionX
    y = Lsm5.Hardware.CpStages.PositionY
    GetStagePositionX = x
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    If bExchangeXY Then
        GetStagePositionX = y
    End If
    
    If bMirrorX Then
        GetStagePositionX = -GetStagePositionX
    End If
    Exit Function
oldversion:
    GetStagePositionX = X11 * x + X12 * y
nostage:
End Function

Public Function GetStagePositionY() As Double


    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim x As Double
    Dim y As Double

    GetStagePositionY = 0#
    Dim cnt As Long
    
    cnt = 0
    On Error GoTo retry
retry:
    If cnt > 1000 Then GoTo nostage
    cnt = cnt + 1
        
    x = Lsm5.Hardware.CpStages.PositionX
    y = Lsm5.Hardware.CpStages.PositionY
    GetStagePositionY = y
     
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    
    If bExchangeXY Then
        GetStagePositionY = x
    End If
    
    If bMirrorY Then
'    If Not bMirrorY Then
        GetStagePositionY = -GetStagePositionY
    End If
    Exit Function
oldversion:
   GetStagePositionY = X21 * x + X22 * y
nostage:
End Function

Public Sub SetStagePositionXY(XP As Double, YP As Double)

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim dExchange As Double
    Dim x As Double
    Dim y As Double
    
    x = XP
    y = YP
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    
    If bMirrorX Then
        x = -x
    End If
    
    If bMirrorY Then
'    If Not bMirrorY Then
        y = -y
    End If
    
    On Error GoTo nostage
    
    If bExchangeXY Then
        Lsm5.Hardware.CpStages.PositionY = x
        Lsm5.Hardware.CpStages.PositionX = y
    Else
        Lsm5.Hardware.CpStages.PositionX = x
        Lsm5.Hardware.CpStages.PositionY = y
    End If
            
    Exit Sub
oldversion:

    Lsm5.Hardware.CpStages.PositionX = X11 * x + X21 * y
    Lsm5.Hardware.CpStages.PositionY = X12 * x + X22 * y
    
nostage:

End Sub

Public Sub ConvertToStagePositionXY(XP As Double, YP As Double, xnew As Double, ynew As Double)

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim dExchange As Double
    Dim x As Double
    Dim y As Double
    
    x = XP
    y = YP
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    
    If bMirrorX Then
        x = -x
    End If
'    If Not bMirrorY Then
    If bMirrorY Then
        y = -y
    End If
        
    If bExchangeXY Then
        ynew = x
        xnew = y
    Else
        xnew = x
        ynew = y
    End If
            
    Exit Sub
oldversion:

    xnew = X11 * x + X21 * y
    ynew = X12 * x + X22 * y
    
nostage:

End Sub

Public Sub ConvertStageToScreenXY(XP As Double, YP As Double, xnew As Double, ynew As Double)

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim dExchange As Double
    Dim x As Double
    Dim y As Double
    
    x = XP
    y = YP
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    
    If bExchangeXY Then
        ynew = x
        xnew = y
    Else
        xnew = x
        ynew = y
    End If
    
    If bMirrorX Then
        xnew = -xnew
    End If
'    If Not bMirrorY Then
    If bMirrorY Then
        ynew = -ynew
    End If
        
            
    Exit Sub
oldversion:

    xnew = X11 * x + X12 * y
    ynew = X21 * x + X22 * y
    
nostage:

End Sub

Public Sub SetStagePositionX(XP As Double)

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim x As Double

    x = XP
    
    On Error GoTo oldversion
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    If bMirrorX Then
        x = -x
    End If
    On Error GoTo nostage
    
    If bExchangeXY Then
        Lsm5.Hardware.CpStages.PositionY = x
    Else
        Lsm5.Hardware.CpStages.PositionX = x
    End If
        
    Exit Sub
oldversion:

    Lsm5.Hardware.CpStages.PositionX = x
nostage:

End Sub


Public Sub SetStagePositionY(YP As Double)

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim dExchange As Double
    Dim y As Double

    y = YP
    
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    
    If Not bMirrorY Then
        y = -y
    End If
    On Error GoTo nostage
    
    If bExchangeXY Then
        Lsm5.Hardware.CpStages.PositionX = y
    Else
        Lsm5.Hardware.CpStages.PositionY = y
    End If
        
    Exit Sub
oldversion:

    Lsm5.Hardware.CpStages.PositionY = y
nostage:

End Sub

Public Sub GetPathAndVersion(Path As String, ThisSystemVersion As Long, pathUp As String)

    Dim OK As Boolean
    Dim SystemVersion As String
    Dim Count As Long
    Dim MacroPath As String
    Dim ProjName As String
    Dim Success As Integer
    Dim pos As Integer
    Dim Start As Integer
    Dim indx As Integer
    Dim bslash As String
    Dim path1 As String
    Dim lngth As Long

    Count = ProjectCount()
    For indx = 0 To Count - 1
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
            Path = Strings.Left(MacroPath, Start - 1)
            lngth = Strings.Len(Path)
            path1 = Strings.Left(Path, lngth - 1)
            Start = 1
            bslash = "\"
            pos = Start
            Do While pos > 0
                pos = InStr(Start, path1, bslash)
                If pos > 0 Then
                    Start = pos + 1
                End If
            Loop
            pathUp = Strings.Left(path1, Start - 1)
            
            Exit For
        End If
    Next indx
            
    SystemVersion = Lsm5.Info.VersionIs
    SystemVersion = Replace(SystemVersion, ",", ".")
    If StrComp(SystemVersion, "5.5", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 55
    ElseIf StrComp(SystemVersion, "5.0", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 50
    ElseIf StrComp(SystemVersion, "4.5", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 45
    ElseIf StrComp(SystemVersion, "4.0", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 40
    ElseIf StrComp(SystemVersion, "3.5", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 35
    ElseIf StrComp(SystemVersion, "3.2", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 32
        
    ElseIf StrComp(SystemVersion, "3.0", vbBinaryCompare) >= 0 Then
        ThisSystemVersion = 30
    Else
        If StrComp(SystemVersion, "2.8", vbBinaryCompare) >= 0 Then
            ThisSystemVersion = 28
        Else
            ThisSystemVersion = 25
        End If
    End If
    
    End Sub

Public Sub Wait(PauseTime As Single)
    Dim Start As Single
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
       DoEvents    ' Yield to other processes.
       'Lsm5.DsRecording.StartScanTriggerIn
    Loop
End Sub


Public Sub FillComboList(TheComboBox As MSForms.ComboBox, key As String)
    Dim Count As Integer
    Dim idx As Integer
    Dim flgUserChangeSaved As Boolean
    Dim flgUserChangeSaved1 As Boolean
    
    Dim theString As String
    flgUserChangeSaved1 = flgUserChange
    flgUserChangeSaved = User_flg
    User_flg = False
    flgUserChange = False
    
    TheComboBox.Clear
    Count = tools.RegCountSubKeys(key)
    If Count > 0 Then
        For idx = 0 To Count - 1
            theString = tools.RegSubkeyName(idx, key)
            TheComboBox.AddItem tools.RegSubkeyName(idx, key)
        Next idx
    End If
    User_flg = flgUserChangeSaved
    flgUserChange = flgUserChangeSaved1
 
End Sub


Public Sub SetComboValues(AComboBox As MSForms.ComboBox, Conf As String, ConfIndex As Long, Keep As Boolean)
' Add 01/30/03 start
    Dim Count As Long
'    Keep = True
    Count = AComboBox.ListCount
    If ConfIndex < Count Then
        AComboBox.Text = Conf
        ConfIndex = AComboBox.ListIndex
        If ConfIndex < 0 Then
            ConfIndex = -1
            Conf = ""
            AComboBox.Text = Conf
            Keep = False

        End If
    Else
        ConfIndex = -1
        Conf = ""
        Keep = False

    End If
    
End Sub


Public Sub DisplayProgress(ProgressLabel As MSForms.Label, state As String, Color As Long)
    If (Color & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Then
        ProgressLabel.ForeColor = 0
    Else
        ProgressLabel.ForeColor = &HFFFFFF
    End If
    ProgressLabel.BackColor = Color
    ProgressLabel.Caption = state
End Sub

'Sub Macro1()
'
'    Dim DS As Document
'    Dim HechtImage As RecordingDocument
'    Dim EngelImage As DsRecordingDoc
'
'    Set DS = Lsm5.ExternalDsObject
'
'    Set HechtImage = DS.MakeNewImageDocument(100, 100, 1, 1, 2, 1, 1)
'
'    Set EngelImage = HechtImageToEngelImage(HechtImage)
'
'End Sub



'Function HechtImageToEngelImage(HechtImage As RecordingDocument) As DsRecordingDoc
'
'    Dim Ds As Document
'    Dim EngelImage As DsRecordingDoc
'    Dim index As Long
'    Dim ImageIndex As Long
'    Dim title As String
'    Dim OriginalTitle As String
'    Dim Success As Integer
'
'    Set Ds = Lsm5.ExternalDsObject
'    If Not HechtImage Is Nothing Then
'        OriginalTitle = HechtImage.title
'        For index = 1 To 1000000
'            title = "XXXXXX" + CStr(index)
'            If (Ds.RecordingDocuments.Item(title) Is Nothing) Then
'                HechtImage.SetTitle title
'                title = HechtImage.title
'                For ImageIndex = 0 To 1000000
'                    Set EngelImage = Lsm5.DsRecordingDocObject(ImageIndex, Success)
'                    If EngelImage Is Nothing Then Exit For
'                    If EngelImage.title = title Then
'                        EngelImage.SetTitle OriginalTitle
'                        Set HechtImageToEngelImage = EngelImage
'                        HechtImage.SetTitle OriginalTitle
'                        Exit Function
'                    End If
'                Next ImageIndex
'            End If
'        Next index
'    End If
'    Set HechtImageToEngelImage = Nothing
'
'End Function
Function HechtImageToEngelImage(HechtImage As RecordingDocument) As DsRecordingDoc

    Dim DS As Object
    Dim EngelImage As DsRecordingDoc
    Dim OtherHechtImage As RecordingDocument
    Dim index As Long
    Dim ImageIndex As Long
    Dim Title As String
    Dim OriginalTitle As String
    Dim Success As Integer

    Set DS = Lsm5.ExternalDsObject
    If Not HechtImage Is Nothing Then
        OriginalTitle = HechtImage.Title
        For index = 1 To DS.RecordingDocuments.Count + 1
            Title = "XXXXXX" + CStr(index)
            If (DS.RecordingDocuments.Item(Title) Is Nothing) Then
                HechtImage.SetTitle Title
                Title = HechtImage.Title
                For ImageIndex = 0 To DS.RecordingDocuments.Count - 1
                    Set EngelImage = Lsm5.DsRecordingDocObject(ImageIndex, Success)
                    If EngelImage Is Nothing Then Exit For
                    Set OtherHechtImage = EngelImage.RecordingDocument
                    If OtherHechtImage Is Nothing Then Exit For
                    If OtherHechtImage.Title = Title Then
                        EngelImage.SetTitle OriginalTitle
                        Set HechtImageToEngelImage = EngelImage
                        HechtImage.SetTitle OriginalTitle
                        Exit Function
                    End If
                Next ImageIndex
                Set HechtImageToEngelImage = Nothing
                Exit Function
            End If
        Next index
    End If
    Set HechtImageToEngelImage = Nothing

End Function


Function EngelImageToHechtImage(EngelImage As DsRecordingDoc) As RecordingDocument

    Dim DS As Document
    Dim HechtImage As RecordingDocument
    Dim Found As Boolean
    Dim index As Long
    Dim ImageIndex As Long
    Dim Title As String
    Dim OriginalTitle As String
    Dim Success As Integer
    If Not EngelImage Is Nothing Then
        Set EngelImageToHechtImage = EngelImage.RecordingDocument
    End If
'    Set Ds = Lsm5.ExternalDsObject
'    If Not EngelImage Is Nothing Then
'        OriginalTitle = EngelImage.title
'
'        For index = 1 To 1000000
'            title = "XXXXXX" + CStr(index)
'            Found = False
'
'            For ImageIndex = 0 To Ds.RecordingDocuments.Count - 1
'                If Not (Lsm5.DsRecordingDocObject(ImageIndex, Success) Is Nothing) Then
'                    If Lsm5.DsRecordingDocObject(ImageIndex, Success).title = title Then
'                        Found = True
'                        Exit For
'                    End If
'                End If
'            Next ImageIndex
'            If Not Found Then
'                EngelImage.SetTitle title
'                title = EngelImage.title
'                Set HechtImage = Ds.RecordingDocuments.Item(title)
'                If HechtImage Is Nothing Then Exit For
'                HechtImage.SetTitle OriginalTitle
'                Set EngelImageToHechtImage = HechtImage
'                EngelImage.SetTitle OriginalTitle
'                Exit Function
'            End If
'        Next index
'    End If
'    Set EngelImageToHechtImage = Nothing

End Function


Sub Heapsort(arr() As Double, hcount As Long, art() As Long)
Dim i As Long
Dim L As Long
Dim Ir As Long
Dim Rra As Double
Dim Tra As Double
Dim J As Long
ReDim art(hcount + 1)
For i = 1 To hcount
    art(i) = i
Next i

If hcount > 1 Then
  L = CInt(hcount / 2) + 1
  Ir = hcount
Cont:
  If L > 1 Then
    L = L - 1
    Rra = arr(art(L))
    Tra = art(L)
  Else
    Rra = arr(art(Ir))
    Tra = art(Ir)
    art(Ir) = art(1)
    Ir = Ir - 1
    If Ir = 1 Then
      art(1) = Tra
      GoTo Done
    End If
  End If
  i = L
  J = L + L
back:
  If J <= Ir Then
    If J < Ir Then
      If arr(art(J)) < arr(art(J + 1)) Then
        J = J + 1
      End If
    End If
    If Rra < arr(art(J)) Then
      art(i) = art(J)
      i = J
      J = J + J
    Else
      J = Ir + 1
    End If
    GoTo back
  End If
  art(i) = Tra
  GoTo Cont
Done:
End If
End Sub


Sub WaitSeconds(seconds As Double)
    Dim Start As Double
    Start = RunTime
    While RunTime < Start + seconds
    Wend
End Sub

Function RunTime() As Double
    Dim secTime As Currency
    Dim secFreq As Currency
    Dim Time As Double
    Dim frequency As Double
    
    QueryPerformanceFrequency secFreq
    QueryPerformanceCounter secTime
    
    Time = secTime
    frequency = secFreq

    If frequency = 0 Then
        RunTime = 0
    Else
        RunTime = Time / frequency
    End If
End Function

Public Function IsLaserOn(Wavelength As Long) As Boolean
    Dim indx As Long
    Dim Count As Long
    
    Dim countLines As Integer
    Dim jndx As Integer
    
    Dim Wave As Long
    Dim LaserName As String
    
    Dim OK As Boolean
    Dim Power As Double
    Dim Attenuation As Double
    Dim OnOff As Integer
    
    If GlobalSystemVersion < 35 Then
'        For indx = 0 To count - 1
'            If Wavelength = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).lWavelength Then
'                IsLaserOn = (Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateOn) _
'                         Or (Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateStandby)
'                Exit Function
'            End If
'        Next indx
        Count = Lsm5.Hardware.CpLaserLines.Count
        For indx = 0 To Count - 1
            OK = Lsm5.Hardware.CpLaserLines.LineInfo(indx, Wave, Attenuation, OnOff, LaserName)
            If Wavelength = Wave Then
                OK = Lsm5.Hardware.CpLasers.Select(LaserName)
                If OK Then
                    IsLaserOn = Lsm5.Hardware.CpLaserLines.IsLaserOn(Wave)
'                    OK = Lsm5.Hardware.CpLasers.IsIgnited
'                    OK = Lsm5.Hardware.CpLasers.Status
'                    If OnOff = 1 Then
'                        IsLaserOn = True
'                    Else
'                        IsLaserOn = False
'                    End If
                    Exit Function
                End If
            End If
        Next indx
    Else
        Count = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.lCount
        For indx = 0 To Count - 1
            If Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).Group = GlobalSystemGroup Then
                If Wavelength = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).lWavelength Then
                    IsLaserOn = (Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateOn) _
                             Or (Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateStandby)
                    Exit Function
                End If
            End If
        Next indx
    End If
    IsLaserOn = False
End Function

Public Sub FindLaserNameAndPower(LaserWave As Long, TheLaserName As String, LaserPower As Double, IsLaserOn As Boolean)
    Dim Count As Integer
    Dim indx As Integer
    Dim countLines As Integer
    Dim jndx As Integer
    
    Dim Wave As Long
    Dim IndexArray() As Long
    
    Dim Attenuation As Double
    Dim OnOff As Integer
    Dim LaserName As String
    
    Dim OK As Boolean
    Dim Power As Double
    Dim Done As Boolean
'    Count = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.lCount
'    For indx = 0 To Count - 1
'        If Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).Group = GlobalSystemGroup Then
'
'            Wave = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).lWavelength
'    'to check if the laser is on :
'    'if Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateOn  then ...
'    'possible states
'                '   eLaserStateOff = 0,
'                '   eLaserStateOn = 1
'                '   eLaserStateStandby = 2
'                '   eLaserStateBusy = 3
'                '   eLaserStateNone = 4
'        End If
'    Next indx
IsLaserOn = False
TheLaserName = ""
LaserPower = 0

If GlobalSystemVersion < 35 Then
    Count = Lsm5.Hardware.CpLaserLines.Count
    For indx = 0 To Count - 1
        OK = Lsm5.Hardware.CpLaserLines.LineInfo(indx, Wave, Attenuation, OnOff, LaserName)
        If Wave = LaserWave Then
            OK = Lsm5.Hardware.CpLasers.Select(LaserName)
            If OK Then
                LaserPower = Lsm5.Hardware.CpLasers.PowerWatt
                If OnOff = 1 Then
                    IsLaserOn = True
                Else
                    IsLaserOn = False
                End If
                TheLaserName = LaserName
                Exit Sub
            End If
        End If
    Next indx
        For indx = 0 To Count - 1
            OK = Lsm5.Hardware.CpLaserLines.LineInfo(indx, Wave, Attenuation, OnOff, LaserName)
            If LaserWave = Wave Then
                OK = Lsm5.Hardware.CpLasers.Select(LaserName)
                If OK Then
                    LaserPower = Lsm5.Hardware.CpLasers.PowerWatt
                    IsLaserOn = Lsm5.Hardware.CpLaserLines.IsLaserOn(Wave)
                    TheLaserName = LaserName

'                    OK = Lsm5.Hardware.CpLasers.IsIgnited
'                    OK = Lsm5.Hardware.CpLasers.Status
'                    If OnOff = 1 Then
'                        IsLaserOn = True
'                    Else
'                        IsLaserOn = False
'                    End If
                    Exit Sub
                End If
            End If
        Next indx
    
Else
    Count = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.lCount
    For indx = 0 To Count - 1
        If Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).Group = GlobalSystemGroup Then
            Wave = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).lWavelength
            If Wave = LaserWave Then
                IsLaserOn = (Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateOn) _
                         Or (Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.lState = eLaserStateStandby)
            
                TheLaserName = Lsm5.ExternalCpObject.pHardwareObjects.pAttenuators.pItem(CVar(indx)).pChLaser.strName
                OK = Lsm5.Hardware.CpLasers.Select(TheLaserName)
                OK = Lsm5.Hardware.CpLasers.IsPowerChangeable
                If OK Then
                    LaserPower = Lsm5.Hardware.CpLasers.Power
                Else
                    LaserPower = 0
                End If
                Exit For
            End If
        End If
    Next indx
End If
End Sub

Public Sub InitializeSampleObservationTime(Mo)
    Dim IsLIVE As Boolean
    If GlobalSystemVersion >= 30 Then
        If GlobalIsDuo Then
            IsLIVE = IsLIVEorLSM
        End If
        If (GlobalSystemGroup = "LIVE") Or IsLIVE Then
            GlobalSampleObservationTime(1) = 493000 * 10 ^ -8
            GlobalSampleObservationTime(2) = 493000 * 10 ^ -8
            GlobalSampleObservationTime(3) = 244000 * 10 ^ -8
            GlobalSampleObservationTime(4) = 120000 * 10 ^ -8
            GlobalSampleObservationTime(5) = 58962 * 10 ^ -8
            GlobalSampleObservationTime(6) = 39310 * 10 ^ -8
            GlobalSampleObservationTime(7) = 23585 * 10 ^ -8
            GlobalSampleObservationTime(8) = 15725 * 10 ^ -8
            GlobalSampleObservationTime(9) = 11782 * 10 ^ -8
            GlobalSampleObservationTime(10) = 7860 * 10 ^ -8
            GlobalSampleObservationTime(11) = 5896 * 10 ^ -8
            GlobalSampleObservationTime(12) = 3928 * 10 ^ -8
            GlobalSampleObservationTime(13) = 2968 * 10 ^ -8


        Else
            GlobalSampleObservationTime(1) = 20484 * 10 ^ -8
            GlobalSampleObservationTime(2) = 10240 * 10 ^ -8
            GlobalSampleObservationTime(3) = 5120 * 10 ^ -8
            GlobalSampleObservationTime(4) = 2560 * 10 ^ -8
            GlobalSampleObservationTime(5) = 1280 * 10 ^ -8
            GlobalSampleObservationTime(6) = 640 * 10 ^ -8
            GlobalSampleObservationTime(7) = 320 * 10 ^ -8
            GlobalSampleObservationTime(8) = 256 * 10 ^ -8
            GlobalSampleObservationTime(9) = 160 * 10 ^ -8
            GlobalSampleObservationTime(10) = 128 * 10 ^ -8
            GlobalSampleObservationTime(11) = 96 * 10 ^ -8
            GlobalSampleObservationTime(12) = 80 * 10 ^ -8
            GlobalSampleObservationTime(13) = 64 * 10 ^ -8
        End If
     Else
        GlobalSampleObservationTime(1) = 7168 * 10 ^ -8
        GlobalSampleObservationTime(2) = 7168 * 10 ^ -8
        GlobalSampleObservationTime(3) = 3584 * 10 ^ -8
        GlobalSampleObservationTime(4) = 1792 * 10 ^ -8
        GlobalSampleObservationTime(5) = 896 * 10 ^ -8
        GlobalSampleObservationTime(6) = 448 * 10 ^ -8
        GlobalSampleObservationTime(7) = 224 * 10 ^ -8
        GlobalSampleObservationTime(8) = 176 * 10 ^ -8
        GlobalSampleObservationTime(9) = 112 * 10 ^ -8
        GlobalSampleObservationTime(10) = 88 * 10 ^ -8
        GlobalSampleObservationTime(11) = 64 * 10 ^ -8
        GlobalSampleObservationTime(12) = 56 * 10 ^ -8
        GlobalSampleObservationTime(13) = 40 * 10 ^ -8
    End If
End Sub


Public Sub UpdateFocus(LabelFocus As MSForms.Label)
    LabelFocus.Caption = " Z=" + Strings.Format(Lsm5.Hardware.CpFocus.Position, "0.00")
       
End Sub

Public Sub UpdateCurrPos(LabelFocus As MSForms.Label, LabelCurrentPos As MSForms.Label)
    Dim x As Double
    Dim y As Double
    
    x = GetStagePositionX
    y = GetStagePositionY
'    x = Stage.PositionX
'    y = Stage.PositionY
    LabelCurrentPos.Caption = "Current Pos: X=" + Strings.Format(x, "0.00") + _
    " Y=" + Strings.Format(y, "0.00")
    LabelFocus.Caption = " Z=" + Strings.Format(Lsm5.Hardware.CpFocus.Position, "0.00")
        
End Sub

Public Sub UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
Dim ScanController As AimScanController
Dim TrackParameters As AimTrackParameters
Dim Size As Long
Dim lTrack As Long
Dim eDeviceMode As Long
'Dim MyRecord As DsRecording

    bLSM = False
    bLIVE = False
    bCamera = False
    Set ScanController = Lsm5.ExternalDsObject.ScanController
    Set TrackParameters = ScanController.TrackParameters
    If TrackParameters Is Nothing Then Exit Sub
'    Set MyRecord = Lsm5.DsRecording
'    Size = MyRecord.TrackCount
    
    Size = TrackParameters.GetTrackArraySize
    For lTrack = 0 To Size - 1
        If TrackParameters.IsTrackUsed(lTrack) Then
            eDeviceMode = TrackParameters.TrackDeviceMode(lTrack)
            Select Case eDeviceMode
                Case eAimDeviceModeLSM
                Case eAimDeviceModeLSM_ChannelMode
                Case eAimDeviceModeLSM_NDD
                Case eAimDeviceModeLSM_DD
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
        End If
    Next lTrack
End Sub


Public Sub CoordinateConversion(bExchangeXY As Boolean, bMirrorX As Boolean, bMirrorY As Boolean)
    Dim bLSM As Boolean
    Dim bLIVE As Boolean
    Dim bCamera As Boolean
    Dim lSystem As Long
    If GlobalSystemVersion = 32 Then
        Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisState bExchangeXY, bMirrorX, bMirrorY
    ElseIf GlobalSystemVersion > 32 Then
        UsedDevices40 bLSM, bLIVE, bCamera
        If bLSM Then
            lSystem = 0
        ElseIf bLIVE Then
            lSystem = 1
        ElseIf bCamera Then
            lSystem = 3
        End If
        Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisStateS lSystem, bExchangeXY, bMirrorX, bMirrorY
    End If

End Sub


Public Sub RemoveInvalidCharacters(InputString As String, OutputString As String)
    OutputString = Replace(InputString, " ", "_")
    OutputString = Replace(OutputString, "\", "_")
    OutputString = Replace(OutputString, "/", "_")
    OutputString = Replace(OutputString, "*", "_")
    OutputString = Replace(OutputString, "|", "_")
    OutputString = Replace(OutputString, "?", "_")
    OutputString = Replace(OutputString, ":", "_")
    OutputString = Replace(OutputString, "<", "_")
    OutputString = Replace(OutputString, ">", "_")
    OutputString = Replace(OutputString, Chr(34), "_")

End Sub

Public Function IsLIVEorLSM()
    Dim Td As DsTrack
    Dim DetS As DsDetectionChannel
    Dim lT As Long
    Dim lI As Long
    Dim Count As Long
    Dim OK As Boolean
    Dim Success As Integer
    OK = False
    IsLIVEorLSM = False
    For lT = 0 To Lsm5.DsRecording.TrackCount - 1
        Set Td = Lsm5.DsRecording.TrackObjectByIndex(lT, Success)
        If Td.Acquire Then
            Count = Td.DetectionChannelCount
            For lI = 0 To Count - 1
                Set DetS = Td.DetectionChannelObjectByIndex(lI, Success)
                If DetS.Acquire Then
                    If DetS.Name = "ChL1" Or DetS.Name = "ChL2" Then
                        IsLIVEorLSM = True
                    Else
                        IsLIVEorLSM = False
                    End If
                    OK = True
                    Exit For
                End If
                If OK Then Exit For
            Next lI
            If OK Then Exit For
        End If
    Next lT
End Function

Public Function TransferPicture(Source As AimImageBitmap) As AimImageBitmap
    Dim x As Long
    Dim y As Long
    Dim Picture As New AimImageBitmap
    Set TransferPicture = New AimImageBitmap
    If Source Is Nothing Then
        TransferPicture.Cleanup
    Else
        Picture.Data = Source.Data
        
        x = Picture.GetLogicalWidth
        y = Picture.GetLogicalHeight
        
        If (x < 1) Or (y < 1) Then
            TransferPicture.Clenaup
        Else
            If x > y Then
                y = 200 * y / x
                x = 200
            Else
                x = 200 * x / y
                y = 200
            End If
        End If
        
        TransferPicture.Create x, y, eAimImageBitmapFormatBGRA
        TransferPicture.Copy Picture, True, False, False
    End If
End Function

