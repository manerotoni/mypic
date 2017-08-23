VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChannelNameForm 
   Caption         =   "Change Channel Names"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   OleObjectBlob   =   "ChannelNameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChannelNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChannelNameLoaded As Boolean

Private Sub ChannelBox_Change()
    If flgUserChange Then
        flgUserChange = False
        TextBox1.Value = GlobalImage.DisplayParameters.ChannelInformation.ChannelName(ChannelBox.ListIndex)
        flgUserChange = True
    End If
End Sub

Sub FillChannelList()
    flgUserChange = False

    Dim indx As Long
    ChannelBox.Clear
    For indx = 1 To GlobalNumberOfChannels
        ChannelBox.AddItem CStr(indx)
    Next indx
    ChannelBox.ListIndex = 0
    flgUserChange = True
    
End Sub



Private Sub OkButton_Click()
    SaveWindowPosition
    Unload Me
End Sub

Private Sub TextBox1_Change()
    If flgUserChange Then
        GlobalImage.DisplayParameters.ChannelInformation.ChannelName(ChannelBox.ListIndex) = TextBox1.Value
    End If
End Sub

Private Sub UserForm_Activate()
    If Not ChannelNameLoaded Then
        LoadWindowPosition
    End If
    ChannelNameLoaded = True

End Sub

Private Sub UserForm_Initialize()
    FillChannelList
    flgUserChange = False
    TextBox1.Value = GlobalImage.DisplayParameters.ChannelInformation.ChannelName(ChannelBox.ListIndex)
    flgUserChange = True

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
    SaveWindowPosition
End Sub
