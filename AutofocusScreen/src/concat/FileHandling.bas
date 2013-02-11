Attribute VB_Name = "FileHandling"
Option Explicit
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_OVERWRITEPROMPT = &H2

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
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal list As Long, ByVal lpPath As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long
Public Declare Function GetFocus Lib "user32" () As Long


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

