VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PumpForm 
   Caption         =   "Start imaging and pump"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   OleObjectBlob   =   "PumpForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PumpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Change_Settings_Click()
    Pump = True
    PauseEndAcquisition = PumpForm.Pump_interval_Jobs
    PumpForm.Hide
End Sub

Private Sub Pump_interval_Jobs_Change()

End Sub

Private Sub Start_Imaging_Click()
    Pump = True
    PauseEndAcquisition = PumpForm.Pump_interval_Jobs
    PumpForm.Hide
    DoEvents
    AutofocusForm.Execute_StartButton
End Sub
