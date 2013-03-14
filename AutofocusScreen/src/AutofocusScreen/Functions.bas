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


''''
' CheckDir
' Check that directory exists
''''
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
    If FileExist(PathName) Then
        ' file exist we try to open it
        On Error Resume Next
        Set File = FileSystem.OpenTextFile(PathName, 8, True)
        On Error GoTo FileIsOpen
        SafeOpenTextFile = True
        Exit Function
    Else
        On Error Resume Next
        Set File = FileSystem.OpenTextFile(PathName, 8, True)
        On Error GoTo FileIsNotAccessible
        SafeOpenTextFile = True
        Exit Function
    End If
FileIsOpen:
    SafeOpenTextFile = True 'file is already open
    Exit Function
FileIsNotAccessible:
    SafeOpenTextFile = False
End Function


''''
'   LogMessage(ByVal Msg As String, ByVal Log As Boolean, ByVal PathName As String, ByRef File As TextStream, ByVal FileSystem As FileSystemObject)
'   Write Msg to a File if Log is on otherwise it does nothing
''''
Public Function LogMessage(ByVal Msg As String, ByVal Log As Boolean, ByVal PathName As String, ByRef File As TextStream, ByVal FileSystem As FileSystemObject)
    If Log Then
        If SafeOpenTextFile(PathName, File, FileSystem) Then
            File.WriteLine (Msg)
        End If
    End If
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

''''
' CheckPosFile
' Check that entries of first line correspond to 4 specific values
''''
Public Function CheckPosFile(ByVal sFile As String, ByVal Row As Integer, ByVal Col As Integer, ByVal RowSub As Integer, ByVal ColSub As Integer) As Boolean
    If FileExist(sFile) Then
        Close
        On Error GoTo ErrorPosFile:
        Dim iFileNum As Integer
        Dim Fields As String
        Dim FieldEntries() As String
        iFileNum = FreeFile()
        Open sFile For Input As iFileNum
        Line Input #iFileNum, Fields
        While Left(Fields, 1) = "%"
            Line Input #iFileNum, Fields
        Wend
        FieldEntries = Split(Fields, " ")
        If FieldEntries(0) = Row And FieldEntries(1) = Col And FieldEntries(2) = RowSub And FieldEntries(3) = ColSub Then
            CheckPosFile = True
        End If
        Close #iFileNum
    End If
    Exit Function
ErrorPosFile:
    If Err.Number = 70 Then
        MsgBox ("WritePosFile: Was not able to load position file " & sFile & ". File is open by another program.")
    End If
End Function

'''''''
' LoadPosFile
' Function loads a file and write entries in Double arrays
''''''''
Public Function LoadPosFile(ByVal sFile As String, _
  posGridX() As Double, posGridY() As Double, posGridZ() As Double) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Input As iFileNum
    Line Input #iFileNum, Fields
    While Left(Fields, 1) = "%"
        Line Input #iFileNum, Fields
    Wend
    FieldEntries = Split(Fields, " ")
    ReDim posGridX(1 To CDbl(FieldEntries(0)), 1 To CDbl(FieldEntries(1)), 1 To CDbl(FieldEntries(2)), 1 To CDbl(FieldEntries(3)))
    ReDim posGridY(1 To CDbl(FieldEntries(0)), 1 To CDbl(FieldEntries(1)), 1 To CDbl(FieldEntries(2)), 1 To CDbl(FieldEntries(3)))
    ReDim posGridZ(1 To CDbl(FieldEntries(0)), 1 To CDbl(FieldEntries(1)), 1 To CDbl(FieldEntries(2)), 1 To CDbl(FieldEntries(3)))
    For iRow = 1 To UBound(posGridX, 1)
      For iCol = 1 To UBound(posGridX, 2)
          Line Input #iFileNum, Fields
          While Left(Fields, 1) = "%"
            Line Input #iFileNum, Fields
          Wend
          FieldEntries = Split(Fields, " ")
          For iRowSub = 1 To UBound(posGridX, 3)
              For iColSub = 1 To UBound(posGridX, 4)
                  posGridX(iRow, iCol, iRowSub, iColSub) = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * UBound(posGridX, 3) * 3))
                  posGridY(iRow, iCol, iRowSub, iColSub) = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * UBound(posGridX, 3) * 3 + 1))
                  posGridZ(iRow, iCol, iRowSub, iColSub) = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * UBound(posGridX, 3) * 3 + 2))
              Next iColSub
          Next iRowSub
      Next iCol
    Next iRow
    Close #iFileNum
    LoadPosFile = True
    Exit Function
ErrorPosFile:
    If Err.Number = 70 Then
        MsgBox ("WritePosFile: Was not able to load position file " & sFile & ". File is open by another program.")
    Else
        MsgBox ("WritePosFile: Was not able to load position file " & sFile)
    End If
    MsgBox ("Was not able to load position file " & sFile)
End Function
    
'''''''
' LoadValidFile
' Function loads a file and write entries in Double arrays
''''''''
Public Function LoadValidFile(ByVal sFile As String, posGridXY_Valid() As Boolean) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Input As iFileNum
    Line Input #iFileNum, Fields
    While Left(Fields, 1) = "%"
        Line Input #iFileNum, Fields
    Wend
    FieldEntries = Split(Fields, " ")
    ReDim posGridXY_Valid(1 To CDbl(FieldEntries(0)), 1 To CDbl(FieldEntries(1)), 1 To CDbl(FieldEntries(2)), 1 To CDbl(FieldEntries(3)))
    For iRow = 1 To UBound(posGridX, 1)
      For iCol = 1 To UBound(posGridX, 2)
          Line Input #iFileNum, Fields
          While Left(Fields, 1) = "%"
            Line Input #iFileNum, Fields
          Wend
          FieldEntries = Split(Fields, " ")
          For iRowSub = 1 To UBound(posGridX, 3)
              For iColSub = 1 To UBound(posGridX, 4)
                  posGridXY_Valid(iRow, iCol, iRowSub, iColSub) = CBool(FieldEntries((iColSub - 1) + (iRowSub - 1) * UBound(posGridX, 3)))
              Next iColSub
          Next iRowSub
      Next iCol
    Next iRow
    Close #iFileNum
    LoadValidFile = True
    Exit Function
ErrorPosFile:
    If Err.Number = 70 Then
        MsgBox ("WritePosFile: Was not able to load position file " & sFile & ". File is open by another program.")
    Else
        MsgBox ("WritePosFile: Was not able to load position file " & sFile)
    End If
    MsgBox ("Was not able to load position file " & sFile)
End Function
    
'''''''
' WritePosFile
' Function loads a file and write entries in Double arrays
''''''''
Public Function WritePosFile(ByVal sFile As String, _
  posGridX() As Double, posGridY() As Double, posGridZ() As Double) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Dim Line As String
    Dim LineComm As String
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Output As iFileNum
    Print #iFileNum, "%nrRows nrColumns nrsubRows nrsubColumns"
    Print #iFileNum, UBound(posGridX, 1) & " " & UBound(posGridX, 2) & " " & UBound(posGridX, 3) & " " & UBound(posGridX, 4)
    For iRow = 1 To UBound(posGridX, 1)
      For iCol = 1 To UBound(posGridX, 2)
          LineComm = "%Row: " & iRow & ", Col: " & iCol & " "
          Print #iFileNum, LineComm
          Line = ""
          LineComm = "%Rowsub Colsub: "
          For iRowSub = 1 To UBound(posGridX, 3)
              For iColSub = 1 To UBound(posGridX, 4)
                 LineComm = LineComm & iRowSub & " " & iColSub & ", "
                 Line = Line & posGridX(iRow, iCol, iRowSub, iColSub) & " " & posGridY(iRow, iCol, iRowSub, iColSub) & " " _
                  & posGridZ(iRow, iCol, iRowSub, iColSub) & " "
              Next iColSub
          Next iRowSub
        Print #iFileNum, LineComm
        Print #iFileNum, Line
      Next iCol
    Next iRow
    Close #iFileNum
    WritePosFile = True
    Exit Function
ErrorPosFile:
    If Err.Number = 70 Then
        MsgBox ("WritePosFile: Was not able to load position file " & sFile & ". File is open by another program")
    Else
        MsgBox ("WritePosFile: Was not able to load position file " & sFile)
    End If
    Close #iFileNum
End Function
    
'''''''
' WritePosFile
' Function loads a file and write entries in Double arrays
''''''''
Public Function WriteValidFile(ByVal sFile As String, posGridXY_Valid() As Boolean) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Dim Line As String
    Dim LineComm As String
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Output As iFileNum
    Print #iFileNum, "%nrRows nrColumns nrsubRows nrsubColumns"
    Print #iFileNum, UBound(posGridXY_Valid, 1) & " " & UBound(posGridXY_Valid, 2) & " " & UBound(posGridXY_Valid, 3) & " " & UBound(posGridXY_Valid, 4)
    For iRow = 1 To UBound(posGridXY_Valid, 1)
      For iCol = 1 To UBound(posGridXY_Valid, 2)
          LineComm = "%Row: " & iRow & ", Col: " & iCol & " "
          Print #iFileNum, LineComm
          Line = ""
          LineComm = "%Rowsub Colsub: "
          For iRowSub = 1 To UBound(posGridXY_Valid, 3)
              For iColSub = 1 To UBound(posGridXY_Valid, 4)
                 LineComm = LineComm & iRowSub & " " & iColSub & ", "
                 Line = Line & -posGridXY_Valid(iRow, iCol, iRowSub, iColSub) * 1 & " "
              Next iColSub
          Next iRowSub
        Print #iFileNum, LineComm
        Print #iFileNum, Line
      Next iCol
    Next iRow
    Close #iFileNum
    WriteValidFile = True
    Exit Function
ErrorPosFile:
    If Err.Number = 70 Then
        MsgBox ("WriteValidFile: Was not able to load position file " & sFile & ". File is open by another program")
    Else
        MsgBox ("WriteValidFile: Was not able to load position file " & sFile)
    End If
    Close #iFileNum
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

