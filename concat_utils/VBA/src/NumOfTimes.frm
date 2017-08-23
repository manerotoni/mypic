VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NumOfTimes 
   Caption         =   "Enter Number Of Time Images"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   OleObjectBlob   =   "NumOfTimes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NumOfTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NumOfTimesLoaded As Boolean

Private Sub CanceButton1_Click()
    TimeNumberChange = False
    SaveWindowPosition
    Unload NumOfTimes
End Sub

Private Sub OkButton_Click()
    TimeNumberChange = True
    GlobalNumberOfStacks = BSlider1.Value
    GlobalTimeIntv = BSlider2.Value
    Unload NumOfTimes
End Sub

Private Sub UserForm_Activate()
    If Not NumOfTimesLoaded Then
        LoadWindowPosition
    End If
    NumOfTimesLoaded = True

End Sub

Private Sub UserForm_Initialize()
    BSlider1.Value = GlobalNumberOfStacks
    BSlider2.Value = GlobalTimeIntv
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    TimeNumberChange = False
    SaveWindowPosition
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

