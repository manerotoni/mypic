VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogManager 
   Caption         =   "ErrorLog"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7485
   OleObjectBlob   =   "LogManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MaxSizeLog = 10000

Public Sub UserForm_Initialize()
    FormatUserForm (Me.Caption)
End Sub



Public Function UpdateErrorLog(Text As String)
    Dim iFileNum
    Dim ErrText As String
    ErrText = VBA.Left(ErrorLogLabel.Caption, MaxSizeLog)
    ErrorLogLabel.Caption = Now & " " & Text & vbCrLf & ErrText
    LogManager.Show
    'write to ErrorFile
    If Log Then
        If SafeOpenTextFile(ErrFileName, ErrFile, FileSystem) Then
            ErrFile.WriteLine Now & " " & Text
            ErrFile.Close
        Else
            Log = False
        End If
    End If
End Function

Public Function UpdateWarningLog(Text As String)
    Dim iFileNum
    Dim ErrText As String
'    ErrText = VBA.Left(ErrorLogLabel.Caption, MaxSizeLog)
'    ErrorLogLabel.Caption = Now & " " & Text & vbCrLf & ErrText
'    LogManager.Show

    'write to ErrorFile
    If Log Then
        If SafeOpenTextFile(ErrFileName, ErrFile, FileSystem) Then
            ErrFile.WriteLine Now & " " & " Warning: " & CurrentFileName & " " & Text
            ErrFile.Close
        Else
            Log = False
        End If
    End If
End Function


Public Function UpdateLog(Text As String, Optional Level As Integer = 0)
    Dim iFileNum
    Dim ErrText As String
    'write to Logfile
    If Log Then
        If Level <= LogLevel Then
            LogFileBuffer = LogFileBuffer & Now & " " & Text
        
            If Len(LogFileBuffer) > 10000 Then
                If SafeOpenTextFile(LogFileName, LogFile, FileSystem) Then
                    LogFile.WriteLine LogFileBuffer
                    LogFile.Close
                Else
                    Log = False
                End If
                LogFileBuffer = ""
            Else
                LogFileBuffer = LogFileBuffer & vbCrLf
            End If
        End If
        'force output
        If Level = -1 Then
            If SafeOpenTextFile(LogFileName, LogFile, FileSystem) Then
                LogFile.WriteLine LogFileBuffer
                LogFile.Close
            Else
                Log = False
            End If
        End If

    End If
End Function

Public Function ResetLog()
    LogManager.Caption = ""
    LogManager.Hide
End Function


