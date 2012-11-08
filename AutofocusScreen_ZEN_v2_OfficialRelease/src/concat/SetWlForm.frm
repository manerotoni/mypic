VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetWlForm 
   Caption         =   "Set Wavelength"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   OleObjectBlob   =   "SetWlForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetWlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim SetWlLoaded As Boolean

Private Sub CanceButton1_Click()
    GlobalSetWlChange = False
    SaveWindowPosition
    Unload SetWlForm
End Sub

Private Sub DefaultButton_Click()
    SetDefaultWl 4, GlobalStartWlTmp, GlobalStepWlTmp
    BSlider1.Value = GlobalStartWlTmp
    BSlider2.Value = GlobalStepWlTmp

End Sub

Private Sub OkButton_Click()
    GlobalSetWlChange = True
    GlobalStartWlTmp = BSlider1.Value
    GlobalStepWlTmp = BSlider2.Value
    Unload SetWlForm
End Sub


Private Sub UserForm_Activate()
    If Not SetWlLoaded Then
        LoadWindowPosition
    End If
    SetWlLoaded = True

End Sub

Private Sub UserForm_Initialize()
    BSlider1.Value = GlobalStartWlTmp
    BSlider2.Value = GlobalStepWlTmp
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
    GlobalSetWlChange = False
    SaveWindowPosition

End Sub
