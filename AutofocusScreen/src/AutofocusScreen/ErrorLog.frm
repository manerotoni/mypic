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
    Dim Log As String
    Log = Left(ErrorLogLabel.Caption, MaxSizeLog)
    ErrorLogLabel.Caption = Text & vbCrLf & Log
End Function
