VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CorrectFocusForm 
   Caption         =   "Correct Focus"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   OleObjectBlob   =   "CorrectFocusForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CorrectFocusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub ChangeFocusButton_Click()
ChangeFocus = True
FocusChanged = True
RestoreAquisitionParameters
Lsm5Vba.Application.ThrowEvent eRootReuse, 0
DoEvents
AutofocusForm.ActivateAcquisitionTrack
 While ChangeFocus = True
                       DoEvents
                        Sleep (100)
Wend

End Sub

Private Sub GoOnButton_Click()
ChangeFocus = False
Unload CorrectFocusForm
DoNotGoOn = False
End Sub



Private Sub UserForm_Activate()
Dim i As Long
Dim j As Long
i = 5
ChangeFocus = False
FocusChanged = False
While i >= 1
    
    Label1.Caption = "The focus couldnot be found. Do You want to change the Position of the FocusWheel manually?" _
                     + " If you donot press the Change Button, then prgramme will go on automatically in " _
                     + CStr(i) + " sec."
    Sleep (1000)
    DoEvents
    If ChangeFocus Then Exit Sub
    i = i - 1
Wend

If i = 0 Then Unload CorrectFocusForm
DoNotGoOn = False
End Sub
