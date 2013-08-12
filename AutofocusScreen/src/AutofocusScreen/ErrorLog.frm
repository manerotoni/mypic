VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorLog 
   Caption         =   "ErrorLog"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8010
   OleObjectBlob   =   "ErrorLog.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MaxSizeLog = 10000

Public Function UpdateLog(Text As String)
    Dim iFileNum
    Dim ErrText As String
    ErrText = Left(ErrorLogLabel.Caption, MaxSizeLog)
    ErrorLogLabel.Caption = Text & vbCrLf & ErrText
    ErrorLog.Show
    'write to ErrorFile
    If Log Then
        If SafeOpenTextFile(ErrFileName, ErrFile, FileSystem) Then
            ErrFile.WriteLine Text
            ErrFile.Close
        Else
            Log = False
        End If
    End If
End Function

Public Function ResetLog()
    ErrorLogLabel.Caption = ""
    ErrorLog.Hide
End Function


