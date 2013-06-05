VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TimeStampForm 
   Caption         =   "Modify Time Stamp"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "TimeStampForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TimeStampForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim TimeStampLoaded As Boolean
Private Sub CanceButton1_Click()
    TimeStampChange = False
    SaveWindowPosition
    Unload TimeStampForm
End Sub

Private Sub OkButton_Click()
    TimeStampChange = True
    GlobalTimeStampDate = DTPicker1.Value
    GlobalTimeStampTime = DTPicker2.Value
    Unload TimeStampForm
End Sub

Private Sub UserForm_Activate()
    If Not TimeStampLoaded Then
        LoadWindowPosition
    End If
    TimeStampLoaded = True

End Sub

Private Sub UserForm_Initialize()
    DTPicker1.Value = GlobalTimeStampDate
    DTPicker2.Value = GlobalTimeStampTime
    TimeStampChange = False
End Sub

Public Function LoadWindowPosition()
    Dim PosKey As String
    
    PosKey = Lsm5.tools.GetWindowPositionKey() + "\" + Caption
    Left = Lsm5.tools.RegLongValue(PosKey, "Left")
    Top = Lsm5.tools.RegLongValue(PosKey, "Top")
    If Left < 1 Then Left = 0
    If Top < 1 Then Top = 0
    
    If Left = 0 And Top = 0 Then
                'Center frm
                Left = 300
                Top = 300
'    SaveWindowPosition
                Exit Function
    End If
End Function



Public Sub SaveWindowPosition()
    Dim PosKey As String
    
    PosKey = Lsm5.tools.GetWindowPositionKey() + "\" + Caption
    Lsm5.tools.RegLongValue(PosKey, "Left") = CInt(Left)
    Lsm5.tools.RegLongValue(PosKey, "Top") = CInt(Top)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    TimeStampChange = False
    SaveWindowPosition

End Sub
