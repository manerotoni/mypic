Attribute VB_Name = "Functions"
Option Explicit


'''''''''
'Minimize button for Macro window
''''''
Private Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
 
Private Declare Function GetWindowLongA Lib "user32" _
(ByVal hWnd As Long, _
ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLongA Lib "user32" _
(ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Sub FormatUserForm(UserFormCaption As String)
     
    Dim hWnd            As Long
    Dim exLong          As Long
     
    hWnd = FindWindowA(vbNullString, UserFormCaption)
    exLong = GetWindowLongA(hWnd, -16)
    If (exLong And &H20000) = 0 Then
        SetWindowLongA hWnd, -16, exLong Or &H20000
    Else
    End If
     
End Sub
''''''''


'''''
'   FileExist(ByVal Pathname)
'   Check if file is present or not
'''''
Public Function FileExist(ByVal PathName As String) As Boolean
    If (Dir(PathName) = "") Then
        FileExist = False
     Else
        FileExist = True
     End If
End Function

Public Function CheckDir(ByVal PathName As String) As Boolean
    On Error GoTo ErrorDir
    If Dir(GlobalDataBaseName, vbDirectory) = "" Then
        MkDir GlobalDataBaseName
    End If
    CheckDir = True
    Exit Function
ErrorDir:
    MsgBox "Was not able to create Directory " & PathName & "  please check disc/pathname!"
End Function

''''
' Tries to open a file. If already open resume to next command
''''
Public Function SafeOpenTextFile(ByVal PathName As String, ByRef File As TextStream, ByVal FileSystem As FileSystemObject) As Boolean
    On Error Resume Next
    Set File = FileSystem.OpenTextFile(PathName, 8, True)
    On Error GoTo ErrorHandle
    SafeOpenTextFile = True
    Exit Function
ErrorHandle:
    SafeOpenTextFile = False 'file is already open
End Function

'''''
'   FileName(iPosition As Integer, iSubposition As Integer, iRepetition As Integer ) As String
'   Returns string by concatanating well, and sublocation and timepoint. A negative point will omit the string
'       [Row] In - Row
'       [Col] In - Col
'       [RowSub]  In - subrow
'       [ColSub]  In - subcol
'       [iRepetition] In - time point
'''''
Public Function FileName(Row As Long, Col As Long, RowSub As Long, ColSub As Long, iRepetition As Integer) As String
    'convert numbers into a string
    Dim iWell As Long
    Dim iPosition As Long

    Dim name As String
    Dim nrZero As Integer
    Dim maxZeros As Integer
    maxZeros = 3
    name = ""
    iWell = (Row - 1) * UBound(posGridX, 2) + Col
    iPosition = (RowSub - 1) * UBound(posGridX, 4) + ColSub
    If iWell >= 0 Then
        nrZero = maxZeros - Len(CStr(iWell))
        name = name + "W" + ZeroString(nrZero) + CStr(iWell)
    End If
    If iPosition >= 0 Then
        nrZero = maxZeros - Len(CStr(iPosition))
        name = name + "_P" + ZeroString(nrZero) + CStr(iPosition)
    End If
    If iRepetition >= 0 Then
        nrZero = maxZeros - Len(CStr(iRepetition))
        name = name + "_T" + ZeroString(nrZero) + CStr(iRepetition)
    End If
    FileName = name
End Function

'''''
'   ZeroString(NrofZeros As Integer) As String
'   Returns a string of zeros
'       [NrofZeros] In - Length of string
'''''
Public Function ZeroString(NrofZeros As Integer) As String
    'convert numbers into a string
    Dim i As Integer
    Dim name As String
    name = ""
    If NrofZeros > 0 Then
        For i = 1 To NrofZeros
            name = name + "0"
        Next i
    End If
        
    ZeroString = name
End Function

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


'''''
'  isArrayEmpty(parArray As Variant) As Boolean
'  Returns false if not an array or dynamic array that has not been initialised (ReDim) or has been erased (Erase)
'''''
Public Function isArrayEmpty(parArray As Variant) As Boolean

    If IsArray(parArray) = False Then isArrayEmpty = True
    On Error Resume Next
    If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False

End Function

'''''
'   FServerFromDescription(strName As String, StrPath As String, ExecName As String) As Boolean
'   TODO: What is this
'''''
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

