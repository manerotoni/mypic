Attribute VB_Name = "FileHandling"
Option Explicit
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_OVERWRITEPROMPT = &H2

Public Type ImageName
     BaseName As String
     ListOfNames() As String
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type BROWSEINFO
    hwndOwner As Long
    lpRoot As Long
    lpstrDirectory As String
    lpstrTitle As String
    ilFlags As Long
    lpCallback As Long
    lParam As Long
    iImage As Long
End Type
    
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal List As Long, ByVal lpPath As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
(ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

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



Public Sub CheckDiskSpace(lpRootPathName As String, lFreeSpace As Double, lSpace As Long)
    Dim lpSectorsPerCluster As Long
    Dim lpBytesPerSector As Long
    Dim lpNumberOfFreeClusters As Long
    Dim lpTotalNumberOfClusters As Long
    
    lSpace = GetDiskFreeSpace(lpRootPathName, lpSectorsPerCluster, lpBytesPerSector, _
                            lpNumberOfFreeClusters, lpTotalNumberOfClusters)
    lFreeSpace = CDbl(lpSectorsPerCluster) * CDbl(lpBytesPerSector) * CDbl(lpNumberOfFreeClusters)

End Sub


Public Function OpenFileNameBox(Name As String, FileType As String, FileExtension As String) As Boolean
    Dim filebox As OPENFILENAME
    Dim result As Long
    Dim Directory As String
    Dim index As String
    
    OpenFileNameBox = False
    On Error GoTo ErrorExit
    
    Directory = CurDir
    If Len(Directory) > 0 Then
        If Not Right(Directory, 1) = "\" Then
            Directory = Directory + "\"
        End If
    End If
    Name = Directory + Name
    
    With filebox
        .lStructSize = Len(filebox)
        .hwndOwner = GetFocus
        .hInstance = 0
        If FileType = "" Then
            .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        Else
            .lpstrFilter = FileType + " (" + FileExtension + ")" + vbNullChar + FileExtension + vbNullChar _
                           + "All Files (*.*)" + vbNullChar + "*.*" + vbNullChar & vbNullChar
        End If
        .nMaxCustomFilter = 0
        .nFilterIndex = 1
        .lpstrFile = Name + Space(256) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .lpstrInitialDir = vbNullChar
        .lpstrTitle = "Open" & vbNullChar
        .FLAGS = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
    End With
    
    result = GetOpenFileName(filebox)
    OpenFileNameBox = result <> 0
    If result <> 0 Then
        Name = VBA.Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
        Directory = Name
        index = Len(Directory)
        Do While index > 0
            If Mid(Directory, index, 1) = "\" Then Exit Do
            If Mid(Directory, index, 1) = ":" Then Exit Do
            index = index - 1
        Loop
        If index > 0 Then
            Directory = Left(Directory, index)
            ChDir Directory
        End If
    End If
ErrorExit:
End Function

Public Function SaveFileNameBox(Name As String, FileType As String, FileExtension As String) As Boolean
    Dim filebox As OPENFILENAME  ' open file dialog structure
    Dim result As Long           ' result of opening the dialog
    
    With filebox
        .lStructSize = Len(filebox)
            .hwndOwner = 0 'Me.hWnd
        .hInstance = 0
        If FileType = "" Then
            .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
        Else
            .lpstrFilter = FileType + " (" + FileExtension + ")" + vbNullChar + FileExtension + vbNullChar _
                           + "All Files (*.*)" + vbNullChar + "*.*" + vbNullChar & vbNullChar
        End If
                
        .nMaxCustomFilter = 0
        .nFilterIndex = 1
        .lpstrFile = Name + Space(256) & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = Space(256) & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .lpstrInitialDir = vbNullChar
        .lpstrTitle = "Save" & vbNullChar
        .FLAGS = OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
    End With
    
    result = GetSaveFileName(filebox)
    SaveFileNameBox = result <> 0
    If result <> 0 Then
        Name = VBA.Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
    End If
End Function

Public Function BrowseForFolder(Name As String, Title As String) As Boolean
    Dim bi As BROWSEINFO
    Dim result As Long
    Dim Directory As String
    Dim index As String
    Dim NewName As String
    
    BrowseForFolder = False
    On Error GoTo ErrorExit
    
    If Len(Name) > 0 Then
        NewName = Name
        ChDir NewName
    Else
        Directory = CurDir
        If Len(Directory) > 0 Then
            If Not Right(Directory, 1) = "\" Then
                Directory = Directory + "\"
            End If
        End If
        NewName = Directory + Name
    End If
    
    bi.hwndOwner = GetFocus
    bi.lpRoot = 0
    bi.lpstrDirectory = NewName + Space(4096) & vbNullChar
    bi.lpstrTitle = Title
    bi.ilFlags = 0
    bi.lpCallback = 0
    bi.lParam = 0
    bi.iImage = 0
    
    result = SHBrowseForFolder(bi)
    If result <> 0 Then
        NewName = Space(4096) & vbNull
        If SHGetPathFromIDList(result, NewName) Then
            
            Directory = VBA.Left(NewName, InStr(NewName, vbNullChar) - 1)
            Name = Directory
            index = Len(Directory)
            Do While index > 0
                If Mid(Directory, index, 1) = "\" Then Exit Do
                If Mid(Directory, index, 1) = ":" Then Exit Do
                index = index - 1
            Loop
            If index > 0 Then
                ChDir Directory
            End If
            BrowseForFolder = True
        End If
    End If
ErrorExit:
End Function

Public Function ListFiles(Files() As String, PathNames() As String, NumberFiles As Long, Directory As String, Mask As String) As Boolean

    Dim FindData As WIN32_FIND_DATA
    Dim Path As String
    Dim PathWithMask As String
    Dim Name As String
    Dim Handle As Long
    Dim index As Long
    
    ListFiles = False
    
    ReDim Files(0)
    ReDim PathNames(0)
    NumberFiles = 0
    
    Path = Directory
    If Len(Path) > 0 Then
        If Not Right(Path, 1) = "\" Then
            Path = Path + "\"
        End If
    End If
    PathWithMask = Path + Mask
    
    index = 0
    Handle = FindFirstFile(PathWithMask, FindData)
    If Handle = INVALID_HANDLE_VALUE Then
        ListFiles = True
        Exit Function
    End If

    Do
        NumberFiles = NumberFiles + 1
        ReDim Preserve Files(NumberFiles + 1)
        ReDim Preserve PathNames(NumberFiles + 1)
        Name = FindData.cFileName
        Name = VBA.Left(Name, InStr(Name, vbNullChar) - 1)
        Files(NumberFiles - 1) = Name
        PathNames(NumberFiles - 1) = Path + Name
    Loop While FindNextFile(Handle, FindData)
    FindClose Handle
    
    ListFiles = True
End Function

