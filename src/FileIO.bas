Attribute VB_Name = "FileIO"
''''
' A list of functions to open and write text files, check their existance etc
''''
Option Explicit
''''''''''''''''''''''''
'Debug and LogVariables'
''''''''''''''''''''''''
Public LogFile As TextStream 'This is the file where a log of the procedure is saved
Public ErrFile As TextStream 'This is the file where a log of the procedure is saved
Public LogFileName As String
Public ErrFileName As String
Public LogFileNameBase As String
Public ErrFileNameBase As String
Public FileSystem As FileSystemObject
Public Log     As Boolean          'If true we log data during the macro


'''
' Separates different part of file name
'''
Public Const FNSep = "_"
'''''
' Variables that are set to build the name of the files
'''''
Public BackSlash As String
Public UnderScore As String


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


'''''
''   LogMessage(ByVal Msg As String, ByVal Log As Boolean, ByVal PathName As String, ByRef File As TextStream, ByVal FileSystem As FileSystemObject)
''   Write Msg to a File if Log is on otherwise it does nothing
'''''
'Public Function LogMessage(ByVal msg As String, ByVal Log As Boolean, ByVal PathName As String, ByRef File As TextStream, ByVal FileSystem As FileSystemObject)
'    If Log Then
'        If SafeOpenTextFile(PathName, File, FileSystem) Then
'            File.WriteLine (msg)
'        End If
'    End If
'End Function


''''''
''   FileName(iPosition As Integer, iSubposition As Integer, iRepetition As Integer ) As String
''   Returns string by concatanating well, and sublocation and timepoint. A negative point will omit the string
''       [Row] In - Row
''       [Col] In - Col
''       [RowSub]  In - subrow
''       [ColSub]  In - subcol
''       [iRepetition] In - time point
''''''
'Public Function FileName(Row As Long, Col As Long, RowSub As Long, ColSub As Long, iRepetition As Integer) As String
'    'convert numbers into a string
'    Dim iWell As Long
'    Dim iPosition As Long
'
'    Dim Name As String
'    Dim nrZero As Integer
'    Dim maxZeros As Integer
'    maxZeros = 3
'    Name = ""
'    iWell = (Row - 1) * UBound(posGridX, 2) + Col
'    iPosition = (RowSub - 1) * UBound(posGridX, 4) + ColSub
'    If iWell >= 0 Then
'        nrZero = maxZeros - Len(CStr(iWell))
'        Name = Name + "W" + ZeroString(nrZero) + CStr(iWell)
'    End If
'    If iPosition >= 0 Then
'        nrZero = maxZeros - Len(CStr(iPosition))
'        Name = Name + "_P" + ZeroString(nrZero) + CStr(iPosition)
'    End If
'    If iRepetition >= 0 Then
'        nrZero = maxZeros - Len(CStr(iRepetition))
'        Name = Name + "_T" + ZeroString(nrZero) + CStr(iRepetition)
'    End If
'    FileName = Name
'End Function



