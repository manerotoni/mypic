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

Private Sub UserForm_Initialize()
    Dim strIconPath As String
    Dim lngIcon As Long
    Dim lnghWnd As Long
    ' Change to the path and filename of an icon file
    Debug.Print Application.ProjectFilePath
    strIconPath = Application.ProjectFilePath & "\resources\micronaut_mc.ico"
    ' Get the icon from the source
    lngIcon = ExtractIcon(0, strIconPath, 0)
    ' Get the window handle of the userform
    lnghWnd = FindWindow("ThunderDFrame", Me.Caption)
    'Set the big (32x32) and small (16x16) icons
    SendMessage lnghWnd, WM_SETICON, True, lngIcon
    SendMessage lnghWnd, WM_SETICON, False, lngIcon
    FormatUserForm (Me.Caption)
End Sub



Private Sub Change_Settings_Click()
    Pump = True
    PauseEndAcquisition = PumpForm.Pump_interval_Jobs
    PumpTime = PumpForm.Pump_time
    PumpWait = PumpForm.Pump_wait
    PumpIntervalTime = PumpForm.Pump_interval_time
    PumpIntervalDistance = PumpForm.Pump_interval_distance
    PumpForm.Hide
End Sub

Private Sub Start_Imaging_Click()
    Pump = True
    PauseEndAcquisition = PumpForm.Pump_interval_Jobs
    PumpTime = PumpForm.Pump_time
    PumpWait = PumpForm.Pump_wait
    PumpIntervalTime = PumpForm.Pump_interval_time
    PumpIntervalDistance = PumpForm.Pump_interval_distance
    PumpForm.Hide
    DoEvents
    PipelineConstructor.StartSetting
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        PumpForm.Hide
        Cancel = True
    End If
End Sub
