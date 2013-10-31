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

Public Function UpdateErrorLog(Text As String)
    Dim iFileNum
    Dim ErrText As String
    ErrText = Left(ErrorLogLabel.Caption, MaxSizeLog)
    ErrorLogLabel.Caption = Now & " Error: " & Text & vbCrLf & ErrText
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

Public Function UpdateLog(Text As String)
    Dim iFileNum
    Dim ErrText As String
    'write to Logfile
    If Log Then
        If SafeOpenTextFile(LogFileName, LogFile, FileSystem) Then
            LogFile.WriteLine Now & " " & Text
            LogFile.Close
        Else
            Log = False
        End If
    End If
End Function

Public Function ResetLog()
    LogManager.Caption = ""
    LogManager.Hide
End Function


