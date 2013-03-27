VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCSRunner 
   Caption         =   "FCSRunner"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   OleObjectBlob   =   "FCSRunner.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FCSRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force to declare all variables
Private Const Version = "v0.3"
Private IsAcquiring As Boolean
Private Counter As Integer


Private Sub TestCode1_Click()
    Dim Record As DsRecordingDoc
    Dim FcsData As AimFcsData
    Counter = Counter + 1
    NewRecord Record, "Lsm" & Counter, 0
    ScanToImage Record
    NewFcsRecord FcsData, "FCS" & Counter, 1
    FcsMeasurement FcsData
End Sub

''''
'   FCSButton_Click()
'   Save actual image, save FCSpoints and perform FCS
''''
Private Sub FCSButton_Click()
    FcsRecord False, False
End Sub

''''
'   PreImgFCSButton_Click()
'   Make new Image, save FCSpoints, and perform FCS
''''
Private Sub PreImgFCSButton_Click()
    FcsRecord True, False
End Sub

''''
'   PreImgFCSPostButton_Click()
'   Make new Image, save FCSpoints, perform FCS, make PostFcsImg
''''
Private Sub PreImgFCSPostImgButton_Click()
    FcsRecord True, True
End Sub

''''
'   FcsRecord(PreImg As Boolean, PostImg As Boolean)
'       [PreImg]    In - If True acquire an image prior FCS
'       [PosImg]    In - If True acquire and image ater FCS
'   Default save actual recording, save FCSpositions and perform FCS measurement
''''
Private Sub FcsRecord(PreImg As Boolean, PostImg As Boolean)
    Dim node As AimExperimentTreeNode
    Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
    Dim iExp As Integer
    Dim FileName As String
    Dim FileID As String
    Dim Record As DsRecordingDoc
    Dim FcsData As AimFcsData
    Dim pixelSize As Double
    If OutputFolderTextBox.Value = "" Then
        MsgBox "No Outputolder!"
        Exit Sub
    End If
    If GetFcsPositionListLength = 0 Then
        MsgBox "No points selected for FCS!"
        Exit Sub
    End If
    IsAcquiring = True
    iExp = 0
    Dim FileCol As String
    FileID = BaseNameTextBox.Value & "_" & iExp & "_preFCS.lsm"
    While FileExist(OutputFolderTextBox.Value & "\" & FileID)
        iExp = iExp + 1
        FileID = BaseNameTextBox.Value & "_" & iExp & "_preFCS.lsm"
    Wend
    FileCol = FileID
    'Scan and save record
    If PreImg Then
        NewRecord Record, FileID, 0
        FocusOnLastDocument
        If Not ScanToImage(Record) Then
            GoTo Abort
        End If
    Else
        Set Record = Lsm5.DsRecordingActiveDocObject
        Record.SetTitle FileID
    End If
    

    'Fcs Measurment
    FileID = BaseNameTextBox.Value & "_" & iExp & ".fcs"
    NewFcsRecord FcsData, FileID, 1
    If Not FcsMeasurement(FcsData) Then
        GoTo Abort
    End If
    
    'Save Files
    SaveDsRecordingDoc Record, OutputFolderTextBox.Value & "\" & BaseNameTextBox.Value & "_" & iExp & "_preFCS.lsm"
    pixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing
    'save positions in a file
    SaveFcsPositionList OutputFolderTextBox.Value & "\" & BaseNameTextBox.Value & "_" & iExp & ".txt", pixelSize

    SaveFcsMeasurement FcsData, OutputFolderTextBox.Value & "\" & FileID
    'node.NumberImages = 2 ??
    
    'Scan and save record after FCS
    If PostImg Then
        FileID = BaseNameTextBox.Value & "_" & iExp & "_postFCS.lsm"
        NewRecord Record, FileID, 0
        If Not ScanToImage(Record) Then
            GoTo Abort
        End If
        Record.SetTitle FileID
        SaveDsRecordingDoc Record, OutputFolderTextBox.Value & "\" & FileID
    End If
    Sleep (1000)
    FocusOnLastDocument
Abort:
    StopButtonState False
    ScanStop = False
    IsAcquiring = False
End Sub




Private Sub StopButton_Click()
    ScanStop = True
    If IsAcquiring Then
        StopButtonState ScanStop
    Else
        StopButtonState False
    End If
End Sub

Private Sub StopButtonState(State As Boolean)
    If State Then
        StopButton.BackColor = &HFF
    Else
        StopButton.BackColor = &H8000000F
    End If
End Sub



''''''
' UserForm_Initialize()
'   Function called from e.g. AutoFocusForm.Show
'   Load and initialize form
''
Private Sub UserForm_Initialize()
    'Setting of some global variables
    Me.Caption = Me.Caption + " " + Version
    FormatUserForm (Me.Caption) ' make minimizing button available
    Me.Show
    LoadPointButton_Click
End Sub




Private Sub OutputFolderButton_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
  
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
            
    Filter = "Alle Dateien (*.*)" & Chr$(0) & "*.*"
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", "", "Select output folder")
    
    If Len(FileName) > 3 Then
        FileName = Left(FileName, Len(FileName) - 3)
        OutputFolderTextBox.Value = FileName
    End If
End Sub



Private Sub ClearAllPointsButton_Click()
    ListPoints.Clear
    ClearFcsPositionList
End Sub



Private Sub LoadPointButton_Click()
    Dim i As Long
    Dim PosX As Double
    Dim PosY As Double
    Dim PosZ  As Double
    ListPoints.ColumnCount = 3
    ListPoints.ColumnWidths = "80;80;80"
    ListPoints.Clear
    For i = 0 To GetFcsPositionListLength - 1
        ListPoints.AddItem
        GetFcsPosition PosX, PosY, PosZ, i
        ListPoints.List(i, 0) = CStr(Round(PosX * 10 ^ 6, PrecXY))
        ListPoints.List(i, 1) = CStr(Round(PosY * 10 ^ 6, PrecXY))
        ListPoints.List(i, 2) = CStr(Round(PosZ * 10 ^ 6, PrecXY))
    Next i
End Sub

Private Sub SetPointButton_Click()
    Dim MaxLength As Long
    MaxLength = GetFcsPositionListLength
    If XTextBox.Value <> "" And YTextBox.Value <> "" And ZTextBox.Value <> "" Then
        SetFcsPosition CDbl(XTextBox.Value) * 10 ^ -6, CDbl(YTextBox.Value) * 10 ^ -6, CDbl(ZTextBox.Value) * 10 ^ -6, MaxLength
    End If
    LoadPointButton_Click
End Sub

Private Sub AddPointButton_Click()
    Dim PosX As Double
    Dim PosY As Double
    Dim PosZ  As Double
    GetFcsPosition PosX, PosY, PosZ
    SetFcsPosition PosX, PosY, PosZ, GetFcsPositionListLength
    LoadPointButton_Click
End Sub
