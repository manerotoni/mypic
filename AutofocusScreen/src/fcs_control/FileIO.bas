Attribute VB_Name = "FileIO"
''''
' A list of functions to open and write text files, check their existance etc
''''

''''''''''''''''''''''''
'Debug and LogVariables'
''''''''''''''''''''''''
Public LogFile As TextStream 'This is the file where a log of the procedure is saved
Public LogFileName As String
Public LogFileNameBase As String
Public FileSystem As FileSystemObject
Public Log     As Boolean          'If true we log data during the macro



'''''
'   ZeroString(NrofZeros As Integer) As String
'   Returns a string of zeros
'       [NrofZeros] In - Length of string
'''''
Public Function ZeroString(NrofZeros As Integer) As String
    'convert numbers into a string
    Dim i As Integer
    Dim Name As String
    Name = ""
    If NrofZeros > 0 Then
        For i = 1 To NrofZeros
            Name = Name + "0"
        Next i
    End If
        
    ZeroString = Name
End Function

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

    Dim Name As String
    Dim nrZero As Integer
    Dim maxZeros As Integer
    maxZeros = 3
    Name = ""
    iWell = (Row - 1) * UBound(posGridX, 2) + Col
    iPosition = (RowSub - 1) * UBound(posGridX, 4) + ColSub
    If iWell >= 0 Then
        nrZero = maxZeros - Len(CStr(iWell))
        Name = Name + "W" + ZeroString(nrZero) + CStr(iWell)
    End If
    If iPosition >= 0 Then
        nrZero = maxZeros - Len(CStr(iPosition))
        Name = Name + "_P" + ZeroString(nrZero) + CStr(iPosition)
    End If
    If iRepetition >= 0 Then
        nrZero = maxZeros - Len(CStr(iRepetition))
        Name = Name + "_T" + ZeroString(nrZero) + CStr(iRepetition)
    End If
    FileName = Name
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
' LoadPosFile(ByVal sFile As String, posGridX() As Double, posGridY() As Double, posGridZ() As Double)
'       [sFile] In  - Output file name
'       [posGridX], [posGridY], [posGridZ] In  - Array where to write coordinates of positions
'   Function loads a file and write entries in Double arrays. First entry of file is structure of grid
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
'   LoadValidFile(ByVal sFile As String, posGridXY_Valid() As Boolean) As Boolean
'       [sFile] In - name of file
'       [posGridXY_Valid] In/Out - the valid positions to image in the grid
'   Function loads a file and write entries in Double arrays
'   First non-commented line is structure of array
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
'   WritePosFile (ByVal sFile As String, posGridX() As Double, posGridY() As Double, posGridZ() As Double)
'       [sFile] In  - Output file name
'       [posGridX], [posGridY], [posGridZ] In  - Array with coordinates of positions
'   Write out position of grid. The first uncommented line is the structure of the grid
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
'   WriteValidFile (ByVal sFile As String, posGridXY_Valid() As Boolean) As Boolean
'       [sFile] In - The filename
'        [posGridXY_Valid] In - The valid positions
'   Write which position is active/valid, i.e. imaged
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


