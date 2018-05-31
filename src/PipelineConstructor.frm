VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PipelineConstructor 
   Caption         =   "MyPiC"
   ClientHeight    =   17775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16005
   OleObjectBlob   =   "PipelineConstructor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PipelineConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents EventMng As EventAdmin
Attribute EventMng.VB_VarHelpID = -1

'current pipeline of form and total number of pipelines
Private currPipeline As Integer
Const NrPipelines = 3
Private PipelineCaption(0 To NrPipelines - 1) As String
'letter code for microscopy plate. Specifies the row
Private Lett() As Variant
''version of pipelineConstructor
Public Version As String
Private TestedPipelines
Private positionOption As Integer







Private Sub ShowOiaKeysButton_Click()
'''''
' Display Keys used for AFM image analysis
'''''
    Dim AFMSettings As clsAFMSettings
On Error GoTo ShowOiaKeysButton_Click_Error

    Set AFMSettings = New clsAFMSettings
    AFMSettings.initializeDefault
    KeyReport.Show
    KeyReport.KeyReportLabel2.MultiLine = True
    KeyReport.KeyReportLabel2.Text = AFMSettings.createKeyReport

   Exit Sub

ShowOiaKeysButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ShowOiaKeysButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' USER FORM INITIALIZATION AND DEATH
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UserForm_Initialize()
    Dim i As Integer
    Dim strIconPath As String
    Dim lngIcon As Long
    Dim lnghWnd As Long
    
On Error GoTo UserForm_Initialize_Error

    Version = "v0.9.1"
    
    Me.Caption = Me.Caption + " " + Version

    'Contains name of the Grids two letter code
    Dim GridNames(2, 1) As String
    GridNames(0, 0) = "DE"
    GridNames(0, 1) = "Default"
    GridNames(1, 0) = "TR1"
    GridNames(1, 1) = "Trigger1"
    GridNames(2, 0) = "TR2"
    GridNames(2, 1) = "Trigger2"
    
    
    'find the version of the software and load ZEN object
    If ZenV > 2010 Then
        On Error GoTo errorMsg
        'in some cases this does not register properly
        'Set ZEN = Lsm5.CreateObject("Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface")
        'this should always work
        Set ZEN = Application.ApplicationInterface
        Dim TestBool As Boolean
        'Check if it works
        TestBool = ZEN.GUI.Acquisition.EnableTimeSeries.value
        ZEN.GUI.Acquisition.EnableTimeSeries.value = Not TestBool
        ZEN.GUI.Acquisition.EnableTimeSeries.value = TestBool
        GoTo NoError
errorMsg:
        MsgBox "Version is ZEN" & ZenV & " but can't find Zeiss.Micro.AIM.ApplicationInterface." & vbCrLf & "Using ZEN2010 settings instead." & vbCrLf & "Check if Zeiss.Micro.AIM.ApplicationInterface.dll is registered?" & "See also the manual how to register a dll into windows."
        ZenV = 2010
NoError:
    End If
    StageSettings MirrorX, MirrorY, ExchangeXY
    
    'a custom event manager
    Set EventMng = New EventAdmin
    EventMng.initialize
        
    ''Pipeline settings
    ReDim Pipelines(0 To NrPipelines - 1)
    For i = 0 To NrPipelines - 1
        Set Pipelines(i) = New APipeline
        Set Pipelines(i).Repetition = New ARepetition
        Set Pipelines(i).Grid = New AGrid
        Pipelines(i).Repetition.interval = True
        Pipelines(i).Grid.NameGrid = GridNames(i, 0)
        Pipelines(i).keepParent = True
        PipelineCaption(i) = GridNames(i, 1)
    Next i
    
    Erase ImgJobs
    Erase FcsJobs
    'initialize registry reader and registry values
    Set AFMSettings = New clsAFMSettings
    AFMSettings.initializeDefault
    AFMSettings.resetRegistry
    
    'default extension
    imgFileFormat = eAimExportFormatLsm5
    imgFileExtension = ".lsm"
    
    ''Form layout
    CurrentPipelineList.ColumnCount = 3
    CurrentPipelineList.ColumnWidths = "20;30;200"
    JobChoiceList.ColumnCount = 2
    JobChoiceList.ColumnWidths = "30;200"
    JobChoiceFrame.Visible = False
    PositionsList.ColumnCount = 5
    PositionsList.ColumnWidths = "20;25;35;35;50"

    Set FocusMethods = New Dictionary
    FocusMethods.Add AnalyseImage.No, "None"
    FocusMethods.Add AnalyseImage.CenterOfMassThr, "Center of Mass (thr)"
    FocusMethods.Add AnalyseImage.Peak, "Peak"
    FocusMethods.Add AnalyseImage.CenterOfMass, "Center of Mass"
    FocusMethods.Add AnalyseImage.AFM, "AFM"
    FocusMethods.Add AnalyseImage.AFMfast, "AFM fast"
    
    If DebugCode Then
        FocusMethods.Add AnalyseImage.FcsLoop, "Debug FcsLoop"
    End If
    For i = 0 To FocusMethods.count - 1
        FocusMethod.AddItem FocusMethods.item(i), i
    Next i
    
    FocusMethod.ListIndex = 0
    PlateType.AddItem "None"
    PlateType.AddItem "Single Well"
    PlateType.AddItem "2 Wells"
    PlateType.AddItem "4 Wells (1x4)"
    PlateType.AddItem "8 Wells (2x4)"
    PlateType.AddItem "96 Wells (8x12)"
    PlateType.ListIndex = 0
    Lett = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")

    PositionButton2.value = True
    PositionButton1.value = True
    currentImgJob = -1
    currentFcsJob = -1
    
    ToggleFrameButton (1)
    
    Me.Height = 465
    Me.Width = 430
    
    Load JobSetter
    Load PumpForm
    'read in the icon
    strIconPath = Application.ProjectFilePath & "\resources\mypic_logo_small.ico"
    ' Get the icon from the source
    lngIcon = ExtractIcon(0, strIconPath, 0)
    ' Get the window handle of the userform
    lnghWnd = FindWindow("ThunderDFrame", Me.Caption)
    'Set the big (32x32) and small (16x16) icons
    SendMessage lnghWnd, WM_SETICON, True, lngIcon
    SendMessage lnghWnd, WM_SETICON, False, lngIcon
    FormatUserForm Me.Caption
   On Error GoTo 0
   Exit Sub

UserForm_Initialize_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UserForm_Initialize of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim exitPipCon As Integer
On Error GoTo UserForm_QueryClose_Error

    exitPipCon = MsgBox("Exit PipelineConstructor?", VbOKCancel + VbQuestion, "PipCon exit")
    If exitPipCon = vbOK Then
        Unload JobSetter
        Unload PumpForm
        Erase ImgJobs
        Erase FcsJobs
        Erase Pipelines
        End
    Else
        Cancel = True
    End If
   On Error GoTo 0
   Exit Sub

UserForm_QueryClose_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UserForm_QueryClose of Form PipelineConstructor at line " & Erl & " "
End Sub

Public Function HideShowForms(OpenForms() As Boolean) As Boolean()
    Dim UForm As Object
    Dim i As Integer
On Error GoTo HideShowForms_Error

    If isArrayEmpty(OpenForms) Then
    
        For Each UForm In VBA.UserForms
            If isArrayEmpty(OpenForms) Then
                 ReDim OpenForms(0)
                 OpenForms(0) = UForm.Visible
            Else
                ReDim Preserve OpenForms(UBound(OpenForms) + 1)
                OpenForms(UBound(OpenForms)) = UForm.Visible
            End If
            If UForm.Visible = True Then
                UForm.Hide
            End If
        Next
        HideShowForms = OpenForms
    Else
        i = 0
        For Each UForm In VBA.UserForms
            If OpenForms(i) Then
                UForm.Show
            End If
            i = i + 1
        Next
    End If
   On Error GoTo 0
   Exit Function

HideShowForms_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure HideShowForms of Form PipelineConstructor at line " & Erl & " "
End Function


Private Sub CreditButton_Click()
    CreditForm.Show
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' USER FORM CHANGE OF FOCUS (DEFAULT, TRIGGER1, etc) SAVING, LOADING OF SETTINGS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub JobSetterButton_Click()

On Error GoTo JobSetterButton_Click_Error

    JobSetter.Show
    JobSetter.Repaint
    DoEvents
On Error GoTo 0
   Exit Sub

JobSetterButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure JobSetterButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub FrameButton1_Click()
    ToggleFrameButton 1
End Sub

Private Sub FrameButton2_Click()
    ToggleFrameButton 2
End Sub

Private Sub FrameButton3_Click()
    ToggleFrameButton 3
End Sub

Private Sub FrameButton4_Click()
    ToggleFrameButton 4
End Sub

Private Sub FrameButton5_Click()
    ToggleFrameButton 5
End Sub

''''
' update form according to button that has been clicked
''''
Public Sub ToggleFrameButton(ButtonNumber As Integer)
    Dim i As Integer

On Error GoTo ToggleFrameButton_Error

    For i = 1 To NrPipelines + 2
        Me.Controls("FrameButton" & i).value = False
        Me.Controls("FrameButton" & i).BackColor = &H8000000A
    Next i
    Me.Controls("FrameButton" & ButtonNumber).value = True
    Me.Controls("FrameButton" & ButtonNumber).BackColor = &HC000&
    
    Select Case ButtonNumber
        Case Is <= NrPipelines
            currPipeline = ButtonNumber - 1
            FramePipeline.Visible = True
            FramePositions.Visible = False
            FrameSaving.Visible = False
            FramePipelineTask.Caption = "Pipeline " & PipelineCaption(currPipeline) & " tasks"
            FramePipelineRepetitions.Caption = "Pipeline " & PipelineCaption(currPipeline) & " repetitions"
            FramePipelineTrigger.Caption = "Pipeline " & PipelineCaption(currPipeline) & " start/end conditions"
            UpdatePipelineList CurrentPipelineList, currPipeline
            UpdateRepetitionSettings currPipeline
            'sanity checks
            If CurrentPipelineList.ListCount > 0 Then
                If CurrentPipelineList.ListIndex = -1 Then
                    CurrentPipelineList.ListIndex = 0
                End If
                enableFrame FrameTaskOptions, True
                enableFrame FramePipelineRepetitions, True
                enableFrame FramePipelineTrigger, currPipeline > 0
                getPeriod
            Else
                enableFrame FrameTaskOptions, False
                enableFrame FramePipelineRepetitions, False
                enableFrame FramePipelineTrigger, False
            End If
    
            UpdateFocusEnabled
            
            keepParentButton.value = Pipelines(currPipeline).keepParent
            maxWait.value = Pipelines(currPipeline).maxWait
            optPtNumber.value = Pipelines(currPipeline).optPtNumber
            
        Case NrPipelines + 1 'position of
            FramePipeline.Visible = False
            FrameSaving.Visible = False
            FramePositions.Visible = True
            FramePositions.Left = 65
            FramePositions.Top = 25
            
        Case NrPipelines + 2 'saving
            FrameSaving.Visible = True
            FramePipeline.Visible = False
            FramePositions.Visible = False
            FrameSaving.Left = 73
            FrameSaving.Top = 25
        End Select

   On Error GoTo 0
   Exit Sub

ToggleFrameButton_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ToggleFrameButton of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub SaveSettings_Click()
    Dim FSO As FileSystemObject
    Dim Filter As String, FileName As String
    Dim Flags As Long
    Dim DefDir As String
   
On Error GoTo SaveSettings_Click_Error

    Flags = OFN_OVERWRITEPROMPT Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Configuration (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    
    FileName = CommonDialogAPI.ShowSave(Filter, Flags, "PipelineConstructor.ini", DefDir, "Save PipelineConstructor settings")
    If FileName = "" Then
        Exit Sub
    End If
    Set FSO = New FileSystemObject
    WorkingDir = FSO.GetParentFolderName(FileName) & "\"
    If Len(FileName) > 3 And VBA.Right(FileName, 4) <> ".ini" Then
        FileName = FileName & ".ini"
    End If
    SaveFormSettings FileName

   On Error GoTo 0
   Exit Sub

SaveSettings_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveSettings_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

''''
'   SaveSettings of PipelineConstructor in file name FileName.
''''
Public Sub SaveFormSettings(FileName As String)
    Dim iTsk As Integer, ipip As Integer, iSet As Integer
    Dim tskSettings As String
    Dim iFileNum As Long
    Dim arrTsk() As Variant
    Dim tskFieldNames() As String
    Dim tsk As Task
On Error GoTo SaveFormSettings_Error

    Close
    iFileNum = FreeFile()
    Open FileName For Output As iFileNum
    tskFieldNames = TaskFieldNames
    For ipip = 0 To UBound(Pipelines)
        With Pipelines(ipip)
            Print #iFileNum, "Pip " & ipip & " Reptime " & .Repetition.Time & " RepNr " & .Repetition.number & " RepInt " & .Repetition.interval
            Print #iFileNum, "Pip " & ipip & " keepParent " & .keepParent & " optPtNumber " & .optPtNumber & " maxWait " & .maxWait

            For iTsk = 0 To Pipelines(ipip).count - 1
                arrTsk = TaskToArray(.getTask(iTsk))
                Debug.Print "Variable type " & VarType(arrTsk(0))
                tskSettings = ""
                For iSet = 0 To UBound(arrTsk)
                    tskSettings = tskSettings & " " & tskFieldNames(iSet) & " " & arrTsk(iSet)
                Next iSet
                Print #iFileNum, "Pip " & ipip & " Tsk " & iTsk & tskSettings
            Next iTsk
        End With
    Next ipip
    Print #iFileNum, "PosSet " & positionOption & " nRow " & GridScan_nRow & " nColumn " & GridScan_nColumn & " dRow " & GridScan_dRow & " dColumn " & GridScan_dColumn & _
    " nRowSub " & GridScan_nRowsub & " nColumnSub " & GridScan_nColumnsub & " dRowSub " & GridScan_dRowsub & " dColumnSub " & _
    GridScan_dColumnsub & " FirstWell " & CInt(GridScan_FirstWell.value) & " WellUpperLeft " & CInt(GridScan_WellUpperLeft.value)
    Close #iFileNum

   On Error GoTo 0
   Exit Sub

SaveFormSettings_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveFormSettings of Form PipelineConstructor at line " & Erl & " "
 
End Sub



Private Sub LoadSettings_Click()
    Dim FSO As FileSystemObject
    Dim Filter As String, FileName As String
    Dim Flags As Long
    Dim DefDir As String

On Error GoTo LoadSettings_Click_Error

    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Images (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DefDir, "Load PipelineConstructor settings")
    If FileName = "" Then
        Exit Sub
    End If
    Set FSO = New FileSystemObject
    WorkingDir = FSO.GetParentFolderName(FileName) & "\"
    LoadFormSettings FileName
    ToggleFrameButton currPipeline + 1

   On Error GoTo 0
   Exit Sub

LoadSettings_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure LoadSettings_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Public Sub LoadFormSettings(FileName As String)
'TODO use regExp to remove several white spaces
    Dim iFileNum As Integer, ipip As Integer, iSet As Integer
    Dim tsk As Task
    Dim arr() As Variant
    Dim Fields As String
    Dim JobName As String
    Dim objRegExp As Object
On Error GoTo LoadFormSettings_Error

    Set objRegExp = CreateObject("vbscript.regexp")
    Dim FieldEntries() As String
    Close
    'On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open FileName For Input As iFileNum
    arr = TaskToArray(tsk)
    Pipelines(0).delAllTasks
    Pipelines(1).delAllTasks
    Pipelines(2).delAllTasks
    Do While Not EOF(iFileNum)
            Line Input #iFileNum, Fields
            While VBA.Left(Fields, 1) = "%"
                Line Input #iFileNum, Fields
            Wend
            If Fields <> "" Then
                With objRegExp
                    .Global = True
                    .Pattern = "\s+"
            
                    Fields = .Replace(Fields, " ")
                End With
                FieldEntries = Split(Fields, " ")
                If FieldEntries(0) = "Pip" Then
                    ipip = CInt(FieldEntries(1))
                    If FieldEntries(2) = "Reptime" Then
                        Pipelines(ipip).Repetition.Time = CDbl(FieldEntries(3))
                        Pipelines(ipip).Repetition.number = CInt(FieldEntries(5))
                        Pipelines(ipip).Repetition.interval = CBool(FieldEntries(7))
                    End If
                    If FieldEntries(2) = "keepParent" Then
                        Pipelines(ipip).keepParent = CBool(FieldEntries(3))
                        Pipelines(ipip).optPtNumber = CInt(FieldEntries(5))
                        Pipelines(ipip).maxWait = CLng(FieldEntries(7))
                    End If
                        
                        
                    If FieldEntries(2) = "Tsk" Then
                        For iSet = 0 To UBound(arr)
                            Select Case VarType(arr(iSet))
                                Case vbInteger
                                    arr(iSet) = CInt(FieldEntries(iSet * 2 + 5))
                                Case vbDouble
                                    arr(iSet) = CDbl(FieldEntries(iSet * 2 + 5))
                                Case vbBoolean
                                    arr(iSet) = CBool(FieldEntries(iSet * 2 + 5))
                                Case vbLong
                                    arr(iSet) = CLng(FieldEntries(iSet * 2 + 5))
                            End Select
                        Next iSet
                        Pipelines(ipip).addTask ArrayToTask(arr)
                    End If
                ElseIf FieldEntries(0) = "PosSet" Then
                    ipip = CInt(FieldEntries(1))
                    Me.Controls("PositionButton" & ipip).value = True
                    For iSet = 2 To UBound(FieldEntries) Step 2
                        On Error GoTo nextiSet
                        If VarType(Me.Controls("GridScan_" & FieldEntries(iSet)).value) = vbBoolean Then
                            Me.Controls("GridScan_" & FieldEntries(iSet)).value = CBool(FieldEntries(iSet + 1))
                        Else
                            Me.Controls("GridScan_" & FieldEntries(iSet)).value = CInt(FieldEntries(iSet + 1))
                        End If
nextiSet:
                    Next iSet
                End If
            End If
    Loop
    
    UpdatePipelineList PipelineConstructor.CurrentPipelineList, currPipeline
    UpdateRepetitionSettings currPipeline
    UpdateFocusEnabled
    getPeriod
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & FileName & " for AutofocusScreen settings"

   On Error GoTo 0
   Exit Sub

LoadFormSettings_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure LoadFormSettings of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub TimeOutButton_Click()
    TimeOut = TimeOutButton.value
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' START AND STOP BUTTONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StopExpButton_Click()
On Error GoTo StopExpButton_Click_Error

    ScanStop = True
    StopAcquisition

    If Pipelines(0).Grid.getNrPts > 1 And CheckDir(GlobalDataBaseName) And Pipelines(0).Grid.isRunning Then
        Pipelines(0).Grid.writePositionGridFile (GlobalDataBaseName + "positionsAfterStop.pos")
    End If
    Pipelines(0).Grid.isRunning = False

   On Error GoTo 0
   Exit Sub

StopExpButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StopExpButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub StopAfterRepButton_Click()
On Error GoTo StopAfterRepButton_Click_Error

    If StopAfterRepButton = True Then
        If Running Then
            ScanStopAfterRepetition = True
            StopAfterRepButton.BackColor = 12648447
        Else
            ScanStopAfterRepetition = False
            StopAfterRepButton.BackColor = &HE0E0E0
            StopAfterRepButton = False
        End If
    Else
        ScanStopAfterRepetition = False
        StopAfterRepButton.BackColor = &HE0E0E0
    End If

   On Error GoTo 0
   Exit Sub

StopAfterRepButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StopAfterRepButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub PauseExpButton_Click()
On Error GoTo PauseExpButton_Click_Error

    If Not Pipelines(0).Grid.isRunning Then
        ScanPause = False
        PauseExpButton.value = False
        PauseExpButton.Caption = ""
        PauseExpButton.BackColor = &HE0E0E0
    Else
        If PauseExpButton.value Then
            ScanPause = True
            PauseExpButton.Caption = "RESUME"
            PauseExpButton.BackColor = 12648447
        Else
            ScanPause = False
            PauseExpButton.Caption = ""
            PauseExpButton.BackColor = &HE0E0E0
        End If
    End If

   On Error GoTo 0
   Exit Sub

PauseExpButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure PauseExpButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

'''
' Acquire current pipeline
'''
Private Sub AcquirePipelineButton_Click()
    Dim stgPos As Vector
    Dim RepNum As Long
On Error GoTo AcquirePipelineButton_Click_Error

    resetStopFlags
    Pump = False
    If Pipelines(currPipeline).count > 0 Then
        If GlobalDataBaseName = "" Then
            MsgBox "No output folder selected! Cannot start acquisition. Click on Saving button.", VbExclamation
            Exit Sub
        End If
        If Not CheckDir(GlobalDataBaseName & "\Test") Then
            Exit Sub
        End If
        'create imaging record
        If Not GlobalRecordingDoc Is Nothing Then
            GlobalRecordingDoc.BringToTop
        End If
        NewRecordGui GlobalRecordingDoc, "IMG:" & Pipelines(currPipeline).Grid.NameGrid, ZEN, ZenV
        'create pipeline position

        Pipelines(currPipeline).Grid.initialize 1, 1, 1, 1
        Pipelines(currPipeline).Grid.setPt getCurrentPosition, True, 1, 1, 1, 1
        
        'UpdateRepetitionSettings currPipeline
        
        Pipelines(currPipeline).Grid.setAllParentPath GlobalDataBaseName & "\Test\"
        Clear_All_Files_And_SubFolders_In_Folder GlobalDataBaseName & "\Test\"
        RepNum = Pipelines(currPipeline).Repetition.number
        Pipelines(currPipeline).Repetition.number = 1
        StartPipeline CInt(currPipeline), GlobalRecordingDoc, GlobalFcsRecordingDoc, GlobalFcsData, GlobalDataBaseName & "\Test"
        Pipelines(currPipeline).Repetition.number = RepNum
        DisplayProgress ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)
    Else
        MsgBox "You need to add a task to the pipeline. Click on + button"
         For RepNum = 0 To 10
            AddJobToPipelineButton.BackColor = "&H0080FFFF&"
            SleepWithEvents (200)
            AddJobToPipelineButton.BackColor = "&H0000C000&"
        Next RepNum
        AddJobToPipelineButton.BackColor = "&H0000C000&"
    End If

   On Error GoTo 0
   Exit Sub

AcquirePipelineButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AcquirePipelineButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

'''
' Test all pipelines
'''
Private Sub TestAllPipelinesButton_Click()
    Dim stgPos As Vector
    Dim RepNum As Long
    Dim i As Integer
On Error GoTo TestAllPipelinesButton_Click_Error

    resetStopFlags
    Pump = False
    DisplayProgress ProgressLabel, "Test run for all pipelines", RGB(&HC0, &HC0, 0)
    SleepWithEvents (2000)
    
    If GlobalDataBaseName = "" Then
        MsgBox "No output folder selected! Cannot start acquisition. Click on Saving button.", VbExclamation
        GoTo Endtest
    End If
    If Not CheckDir(GlobalDataBaseName & "\Test") Then
        GoTo Endtest
    End If
    'create imaging record
    If Not GlobalRecordingDoc Is Nothing Then
        GlobalRecordingDoc.BringToTop
    End If
    StageSettings MirrorX, MirrorY, ExchangeXY
    NewRecordGui GlobalRecordingDoc, "IMG:" & Pipelines(currPipeline).Grid.NameGrid, ZEN, ZenV
    Clear_All_Files_And_SubFolders_In_Folder GlobalDataBaseName & "\Test\"
    For i = 0 To NrPipelines - 1
        If Pipelines(i).count > 0 Then
            'create pipeline position
            Pipelines(i).Grid.initialize 1, 1, 1, 1
            Pipelines(i).Grid.setPt getCurrentPosition, True, 1, 1, 1, 1
            'UpdateRepetitionSettings currPipeline
            Pipelines(i).Grid.setAllParentPath GlobalDataBaseName & "\Test\"
            RepNum = Pipelines(i).Repetition.number
            Pipelines(i).Repetition.number = 1
            StartPipeline CInt(i), GlobalRecordingDoc, GlobalFcsRecordingDoc, GlobalFcsData, GlobalDataBaseName & "\Test"
            Pipelines(i).Repetition.number = RepNum
        End If
    Next i
Endtest:
    DisplayProgress ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)
    TestedPipelines = True

   On Error GoTo 0
   Exit Sub

TestAllPipelinesButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure TestAllPipelinesButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub StartExpButton_Click()
On Error GoTo StartExpButton_Click_Error

    Pump = False
    'Do some check for consistency
    DoEvents
     'Now we're starting. This will be set to false if the stop button is pressed or if we reached the total number of repetitions.
    StartSetting
    Running = False

   On Error GoTo 0
   Exit Sub

StartExpButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StartExpButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub StartPumpExpButton_Click()
    PumpForm.Show
End Sub

''''''
'   StartSetting()
'   Setups and controls before start of experiment
'       Create list of positions for Grid or Multiposition
''''''
Public Function StartSetting() As Boolean
    Dim i As Integer
    Dim initPos As Boolean   'if False and gridsize correspond positions are taken from file positionsGrid.csv
    Dim SuccessRecenter As Boolean
    Dim Job As Variant
    Dim gridDim() As Long
    Dim pos() As Vector
    Dim posCurr As Vector   'current position
On Error GoTo StartSetting_Error

    Set FileSystem = New FileSystemObject
    
    resetStopFlags
    
    ''Create and check directory for output and log files
    SetDatabase
    If GlobalDataBaseName = "" Then
        MsgBox "No output folder selected! Cannot start acquisition. Click on Saving button.", VbExclamation
        GoTo ExitStart
    Else
        If Not CheckDir(GlobalDataBaseName) Then
            GoTo ExitStart
        End If
        GlobalAutoSaveName = GlobalDataBaseName & "\autosave"
        If Not CheckDir(GlobalAutoSaveName) Then
            GoTo ExitStart
        End If
        'initialize logFiles
        If LogFileNameBase <> "" Then
            'On Error GoTo ErrorHandleLogFile
            LogFileName = LogFileNameBase
            ErrFileName = ErrFileNameBase
            Close
            If SafeOpenTextFile(LogFileName, LogFile, FileSystem) And SafeOpenTextFile(ErrFileName, ErrFile, FileSystem) Then
                LogFile.WriteLine "% ZEN software version " & ZenV & " PipelineConstructor " & Version
                ErrFile.WriteLine "% ZEN software version " & ZenV & " PipelineConstructor " & Version
                LogFile.Close
                ErrFile.Close
                Log = True
            Else
                Log = False
            End If
        Else
            Log = False
        End If
    End If
    SetFileName
    
    'check if there is something to do
    If Pipelines(0).count = 0 Then
        MsgBox ("Nothing to do! Add at least one task to Default pipeline!")
        GoTo ExitStart
    End If
    
    ''check if pipeline has been tested
'    If Not TestedPipelines Then
'        If MsgBox("You have not tested your pipelines (press play - T button for this). Do you want to continue?", VbYesNo + VbQuestion, "PipCon") = vbNo Then
'            GoTo ExitStart
'        End If
'    End If
    
    Running = True
    StageSettings MirrorX, MirrorY, ExchangeXY
    
    'Eventually create new records
    If GlobalRecordingDoc Is Nothing Then
        NewRecordGui GlobalRecordingDoc, "IMG", ZEN, ZenV
    Else
        GlobalRecordingDoc.CloseAllWindows
        NewRecordGui GlobalRecordingDoc, "IMG", ZEN, ZenV
    End If
    
    If Not isArrayEmpty(FcsJobs) Then
        NewFcsRecordGui GlobalFcsRecordingDoc, GlobalFcsData, "FCS", ZEN, ZenV
    End If
    If Not GlobalRecordingDoc Is Nothing Then
        GlobalRecordingDoc.BringToTop
    End If
    
    posCurr = getCurrentPosition
    
    'initialze all objects
    For Each Job In ImgJobs
        Job.timeToAcquire = 0
    Next Job
    For Each Job In FcsJobs
        Job.timeToAcquire = 0
    Next Job
    
    DisplayProgress Me.ProgressLabel, "Initialize all grid positions...", RGB(0, &HC0, 0)
    
    'set all grids to 0 fo the start
    For i = 0 To UBound(Pipelines)
        Pipelines(i).Grid.initializeToZero
        Pipelines(i).Grid.wellUpperLeft = GridScan_WellUpperLeft.value
    Next i
    
    Set TimersGridCreation = Nothing
    If Not setGridFromPositionChoice(Pipelines(0).Grid, positionOption) Then
        GoTo ExitStart
    End If
    Pipelines(0).Grid.setAllParentPath GlobalDataBaseName
    'write out settings and positions
    Pipelines(0).Grid.writePositionGridFile GlobalDataBaseName & "PipelineConstructor.pos"
    SaveFormSettings GlobalDataBaseName & "PipelineConstructor.ini"
'TODO check if pump is available
    If Pump Then
        lastTimePump = CDbl(GetTickCount) * 0.001
    End If
    
    StartPipeline 0, GlobalRecordingDoc, GlobalFcsRecordingDoc, GlobalFcsData, GlobalDataBaseName, GridScan_FirstWell.value
    
ExitStart:
    LogManager.UpdateLog "End of Global pipeline", -1
    resetStopFlags
    Running = False
    DisplayProgress PipelineConstructor.ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)
    Exit Function

   On Error GoTo 0
   Exit Function

StartSetting_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StartSetting of Form PipelineConstructor at line " & Erl & " "
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' POSITIONS MANAGEMENT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''single position''
Private Sub PositionButton1_Click()
    Dim i As Integer
    If PositionButton1 Then
        enableFrame FramePositionsControl, False
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, False
        enableFrame FrameGridLoad, False
        enableFrame FrameGridOrder, False
        positionOption = 1
    End If
End Sub

''multiple position''
Private Sub PositionButton2_Click()
    Dim i As Integer
    If PositionButton2 Then
        enableFrame FramePositionsControl, True
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, False
        enableFrame FrameGridLoad, False
        enableFrame FrameGridOrder, False
        positionOption = 2
    End If
End Sub

''grid''
Private Sub PositionButton3_Click()
    Dim i As Integer
    If PositionButton3 Then
        enableFrame FramePositionsControl, True
        enableFrame FrameGridControl, True
        enableFrame FrameSubGridControl, True
        enableFrame FrameGridLoad, False
        enableFrame FrameGridOrder, True
        positionOption = 3
    End If
End Sub

''grid based on multiple positions''
Private Sub PositionButton4_Click()
    Dim i As Integer
    If PositionButton4 Then
        enableFrame FramePositionsControl, True
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, True
        enableFrame FrameGridLoad, False
        enableFrame FrameGridOrder, True
        positionOption = 4
    End If
End Sub

''grid based on multiple positions''
Private Sub PositionButton5_Click()
    Dim i As Integer
    If PositionButton5 Then
        enableFrame FramePositionsControl, False
        enableFrame FrameGridControl, False
        enableFrame FrameSubGridControl, False
        enableFrame FrameGridLoad, True
        enableFrame FrameGridOrder, True
        positionOption = 5
    End If
End Sub

Private Sub AddPositionButton_Click()
    Dim posVec As Vector
    posVec = getCurrentPosition
    AddPosition WellID.value, posVec
End Sub

Private Sub AddPosition(ID As String, pos As Vector)
    With PositionsList
        .AddItem
        .List(.ListCount - 1, 0) = .ListCount
        .List(.ListCount - 1, 1) = WellID.value
        .List(.ListCount - 1, 2) = pos.X
        .List(.ListCount - 1, 3) = pos.Y
        .List(.ListCount - 1, 4) = pos.Z
        .ListIndex = .ListCount - 1
    End With
End Sub

Private Sub MoveToPositionButton_Click()
    With PositionsList
        If .ListIndex > -1 Then
            FailSafeMoveStageXY CDbl(.List(.ListIndex, 2)), CDbl(.List(.ListIndex, 3))
            FailSafeMoveStageZ CDbl(.List(.ListIndex, 4))
        End If
    End With
End Sub

Private Sub UpdatePositionButton_Click()
    Dim posVec As Vector
    posVec = getCurrentPosition
    With PositionsList
        If .ListIndex > -1 Then
            .List(.ListIndex, 1) = WellID.value
            .List(.ListIndex, 2) = posVec.X
            .List(.ListIndex, 3) = posVec.Y
            .List(.ListIndex, 4) = posVec.Z
        End If
    End With
End Sub

Private Sub SwitchPosition_SpinUp()
    Dim i As Integer
    Dim currIndex As Integer
    With PositionsList
        currIndex = .ListIndex
        If .ListIndex > 0 Then
            MoveListboxItem PositionsList, .ListIndex, .ListIndex - 1
            For i = 0 To .ListCount - 1
                .List(i, 0) = i + 1
            Next i
            .ListIndex = currIndex - 1
        End If
    End With
End Sub

Private Sub SwitchPosition_SpinDown()
    Dim i As Integer
    Dim currIndex As Integer
    With PositionsList
        currIndex = .ListIndex
        If .ListIndex < .ListCount - 1 And .ListIndex > -1 Then
            MoveListboxItem PositionsList, .ListIndex, .ListIndex + 1
            For i = 0 To .ListCount - 1
                .List(i, 0) = i + 1
            Next i
            .ListIndex = currIndex + 1
        End If
    End With
End Sub


Private Sub RemovePositionButton_Click()
    Dim i As Integer
    With PositionsList
        If .ListIndex > -1 Then
            .RemoveItem .ListIndex
        End If
        For i = 0 To .ListCount - 1
            .List(i, 0) = i + 1
        Next i
    End With
End Sub

Private Sub PlateType_Change()
    Dim iL As Integer
    Dim iNum As Integer
    Dim MaxINum As Integer
    Dim MaxIL As Integer
    WellID.Clear
    Select Case PlateType.ListIndex
        Case 1
            MaxIL = 0
            MaxINum = 1
        Case 2
            MaxIL = 0
            MaxINum = 2
        Case 3
            MaxIL = 1
            MaxINum = 4
        Case 4
            MaxIL = 2
            MaxINum = 4
        Case 5
            MaxIL = 8
            MaxINum = 12
    End Select
    If PlateType.ListIndex > 0 Then
        For iL = 0 To MaxIL
            For iNum = 1 To MaxINum
                WellID.AddItem "" & Lett(iL) & iNum
            Next iNum
        Next iL
        WellID.ListIndex = 0
    End If
End Sub

''''
' update positions from settings in form
''''
Function setGridFromPositionChoice(locGrid As AGrid, optionPos As Integer) As Boolean
    Dim i As Integer
    Dim posVec() As Vector
    'test if positions have been defined
On Error GoTo setGridFromPositionChoice_Error

    If PositionsList.ListCount <= 0 Then
        Select Case optionPos
            Case 2
                MsgBox "For multiple positions you need to mark at least one position!", VbExclamation
                Exit Function
            Case 3
                MsgBox "For grid you need to mark one position used as reference", VbExclamation
                Exit Function
            Case 4
                MsgBox "For multiple positions + grid you need to mark positions. Main grid Positions are marked positions, subpositions are made accordingly.", VbExclamation
                Exit Function
        End Select
    End If
    
    Select Case optionPos
        Case 1 'single point
            locGrid.initialize 1, 1, 1, 1
            locGrid.setPt getCurrentPosition, True, 1, 1, 1, 1
        Case 2 'multipe points
            locGrid.initialize 1, PositionsList.ListCount, 1, 1
            For i = 0 To PositionsList.ListCount - 1
                locGrid.setPt Double2Vector(PositionsList.List(i, 2), PositionsList.List(i, 3), PositionsList.List(i, 4)), _
                True, 1, i + 1, 1, 1
            Next i
        Case 3 'grid from one point
            locGrid.makeGridFromOnePt Double2Vector(PositionsList.List(0, 2), PositionsList.List(0, 3), PositionsList.List(0, 4)), GridScan_nRow, GridScan_nColumn, _
            GridScan_nRowsub, GridScan_nColumnsub, GridScan_dRow, GridScan_dColumn, GridScan_dRowsub, GridScan_dColumnsub
        Case 4 'grid from multiple points
            ReDim posVec(0 To PositionsList.ListCount - 1)
            For i = 0 To PositionsList.ListCount - 1
                posVec(i).X = PositionsList.List(i, 2)
                posVec(i).Y = PositionsList.List(i, 3)
                posVec(i).Z = PositionsList.List(i, 4)
            Next i
            locGrid.makeGridFromManyPts posVec, 1, PositionsList.ListCount, GridScan_nRowsub, GridScan_nColumnsub, GridScan_dRowsub, GridScan_dColumnsub
        Case 5 'read from file
            If Not FileExist(GridScanPositionFile) Then
                MsgBox "Load positions from file failed. Could not find " & GridScanPositionFile
                Exit Function
            End If
            If Not locGrid.loadPositionGridFile(GridScanPositionFile) Then
                Exit Function
            End If
    End Select
    setGridFromPositionChoice = True

   On Error GoTo 0
   Exit Function

setGridFromPositionChoice_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure setGridFromPositionChoice of Form PipelineConstructor at line " & Erl & " "
End Function


''''
' load file containing coordinates of imaging positions
''''
Private Sub GridScanPositionFileButton_Click()
    Dim FSO As New FileSystemObject
    Dim Filter As String, FileName As String
    Dim Flags As Long
    Dim DefDir As String
    Dim locGrid As AGrid
    Dim gridDim() As Long
On Error GoTo GridScanPositionFileButton_Click_Error

    Set locGrid = New AGrid
    
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "position files (*.pos)" & Chr$(0) & "*.pos" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DefDir, "Select position file")
    If FileName = "" Then
        Exit Sub
    End If
    If Not FileExist(FileName) Then
        MsgBox "Load positions from file failed. File " & FileName & " does not exist"
        Exit Sub
    End If
    WorkingDir = FSO.GetParentFolderName(FileName) & "\"
    gridDim = locGrid.getGridDimFromFile(FileName)
    If Not locGrid.loadPositionGridFile(FileName) Then
        Exit Sub
    End If

    GridScanPositionFile = FileName
    UpdatePositionsListFromGrid locGrid

   On Error GoTo 0
   Exit Sub

GridScanPositionFileButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure GridScanPositionFileButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub UpdatePositionsListFromGrid(locGrid As AGrid)
    Dim index As Long
On Error GoTo UpdatePositionsListFromGrid_Error

    locGrid.setIndeces 1, 1, 1, 1
    PositionsList.Clear
    If locGrid.getNrValidPts > 100 Then
        MsgBox "Warning: Position file contains more than 100 positions!" & vbCrLf & _
        "All positions will be loaded but only the first 100 are shown in list"
    End If
    Do ''Cycle all positions defined in grid
        If locGrid.getThisValid Then
            index = index + 1
            AddPosition WellID.value, locGrid.getThisPosition
        End If
    Loop While (locGrid.nextGridPt(False) And index < 100)

   On Error GoTo 0
   Exit Sub

UpdatePositionsListFromGrid_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UpdatePositionsListFromGrid of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub SavePositionsButton_Click()
    Dim FSO As New FileSystemObject
    Dim Filter As String, FileName As String, DefDir As String
    Dim Flags As Long
    Dim locGrid As AGrid
On Error GoTo SavePositionsButton_Click_Error

    Set locGrid = New AGrid
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
    Filter = "Position file (*.pos)" & Chr$(0) & "*.pos" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    FileName = CommonDialogAPI.ShowSave(Filter, Flags, "*.pos", DefDir, "Save positions")
    DisplayProgress Me.ProgressLabel, "Saving positions..", RGB(0, &HC0, 0)
    
    If FileName <> "" Then
        If VBA.Right(FileName, 4) <> ".pos" Then
            FileName = FileName & ".pos"
        End If
    Else
        GoTo ExitSub
    End If
    WorkingDir = FSO.GetParentFolderName(FileName) & "\"
    If Not setGridFromPositionChoice(locGrid, positionOption) Then
        MsgBox "Saving of positions failed"
    End If
    If Not locGrid.writePositionGridFile(FileName) Then
         MsgBox "Saving of positions failed"
    End If
ExitSub:
    DisplayProgress Me.ProgressLabel, "Ready", RGB(&HC0, &HC0, 0)

   On Error GoTo 0
   Exit Sub

SavePositionsButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SavePositionsButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

'''
' Update GUI when reloeding the positions
'''
Private Sub GridScan_FirstWell_Change()
    If Not GridScan_FirstWell.value Then
        GridScan_FirstSubPos.value = True
    End If
End Sub


Private Sub GridScan_WellUpperLeft_Change()
    If Not GridScan_WellUpperLeft.value Then
        GridScan_WellCenter.value = True
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' PIPELINE MANAGEMENT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' if true parent position is not removed from grid
Private Sub KeepParentButton_Click()
    Pipelines(currPipeline).keepParent = keepParentButton.value
End Sub

''' max time to wait before starting subpipeline
Private Sub maxWait_Click()
    Pipelines(currPipeline).maxWait = CDbl(maxWait.value)
End Sub

''' max nr of points to wait before starting subpipeline
Private Sub optPtNumber_Click()
    Pipelines(currPipeline).optPtNumber = CInt(optPtNumber.value)
End Sub


''''
'   fills popup menu for chosing a track for post-acquisition tracking for Job with JobName
''''
Public Sub FillTrackingChannelList(tmpTask As Task)
    Dim Success As Integer
    Dim iTrack As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim Track As DsTrack
    Dim TrackOn As Boolean
On Error GoTo FillTrackingChannelList_Error

    CenterOfMassChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    If tmpTask.jobType = jobTypes.imgjob Then
        With ImgJobs(tmpTask.jobNr)
            For iTrack = 0 To .TrackNumber - 1
                Set Track = .GetRecording.TrackObjectByMultiplexOrder(iTrack, Success)
                    If .getAcquireTrack(iTrack) Then
                    For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
                        If Track.DetectionChannelObjectByIndex(c - 1, Success).Acquire Then 'if channel is activated
                            ca = ca + 1 'counter for active channels will increase by one
                            CenterOfMassChannel.AddItem Track.Name & " " & Track.DetectionChannelObjectByIndex(c - 1, Success).Name & "-T" & iTrack + 1   'entry is added to combo box to chose track for post-acquisition tracking
                            TrackOn = True
                        End If
                    Next c
                End If
            Next iTrack
        End With
    End If

   On Error GoTo 0
   Exit Sub

FillTrackingChannelList_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure FillTrackingChannelList of Form PipelineConstructor at line " & Erl & " "
End Sub
    
Private Sub TrackXY_Click()
    Dim index As Integer
On Error GoTo TrackXY_Click_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setTrackXY index, TrackXY
    End If

   On Error GoTo 0
   Exit Sub

TrackXY_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure TrackXY_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub TrackZ_Click()
    Dim index As Integer
On Error GoTo TrackZ_Click_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setTrackZ index, TrackZ
    End If

   On Error GoTo 0
   Exit Sub

TrackZ_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure TrackZ_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub CenterOfMassChannel_Click()
    Dim index As Integer
On Error GoTo CenterOfMassChannel_Click_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setTrackChannel index, CenterOfMassChannel.ListIndex
    End If

   On Error GoTo 0
   Exit Sub

CenterOfMassChannel_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure CenterOfMassChannel_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

''' update option for focusing and tracking in form according to type of job
Private Sub UpdateFocusEnabled()
    Dim index As Integer
On Error GoTo UpdateFocusEnabled_Error

    TrackingFrame.Visible = True
    index = CurrentPipelineList.ListIndex
    If index = -1 Then
        enableFrame TrackingFrame, False
        Exit Sub
    End If
    If Pipelines(currPipeline).getTask(index).jobType <> jobTypes.imgjob Then
        enableFrame TrackingFrame, False
        Exit Sub
    End If
    enableFrame TrackingFrame, True
    FocusMethod.Enabled = True
    CenterOfMassChannel.Enabled = True And (FocusMethod.ListIndex > AnalyseImage.No) And (Not FocusMethod.ListIndex = AnalyseImage.AFM) And (Not FocusMethod.ListIndex = AnalyseImage.AFM)
    TrackZ.value = Pipelines(currPipeline).getTrackZ(index)
    TrackXY.value = Pipelines(currPipeline).getTrackXY(index)
    With ImgJobs(Pipelines(currPipeline).getTask(index).jobNr)
        TrackZ.Enabled = .isZStack And (FocusMethod.ListIndex > AnalyseImage.No)
        TrackXY.Enabled = (FocusMethod.ListIndex > AnalyseImage.No) And (.Recording.ScanMode <> "ZScan") And (.Recording.ScanMode <> "Line")
    End With

   On Error GoTo 0
   Exit Sub

UpdateFocusEnabled_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UpdateFocusEnabled of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub FocusMethod_Click()
    Dim index As Integer
On Error GoTo FocusMethod_Click_Error

    index = CurrentPipelineList.ListIndex
    If index < 0 Then
        Exit Sub
    End If
    Pipelines(currPipeline).setAnalyse index, FocusMethod.ListIndex

    UpdateFocusEnabled
    If Pipelines(currPipeline).getAnalyse(index) = AnalyseImage.AFM Then
        SaveImage = True
        Pipelines(currPipeline).setSaveImage index, True
    End If

   On Error GoTo 0
   Exit Sub

FocusMethod_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure FocusMethod_Click of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub CurrentPipelineList_Click()
    Dim index As Integer
On Error GoTo CurrentPipelineList_Click_Error

    index = CurrentPipelineList.ListIndex
    getPeriod
    getZOffset
    getSaveImage
    If index > -1 And Pipelines(currPipeline).getTask(index).jobType = jobTypes.imgjob Then
        FillTrackingChannelList Pipelines(currPipeline).getTask(index)
        CenterOfMassChannel.ListIndex = Pipelines(currPipeline).getTrackChannel(index)
        FocusMethod.ListIndex = Pipelines(currPipeline).getAnalyse(index)
    Else
        CenterOfMassChannel.Clear
    End If
    enableFrame FrameTaskOptions, True
    enableFrame FramePipelineRepetitions, True
    enableFrame FramePipelineTrigger, currPipeline > 0
    CurrentPipelineList.SetFocus
    UpdateFocusEnabled

   On Error GoTo 0
   Exit Sub

CurrentPipelineList_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure CurrentPipelineList_Click of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub AddJobToPipelineButton_Click()
On Error GoTo AddJobToPipelineButton_Click_Error

    If isArrayEmpty(ImgJobs) And isArrayEmpty(FcsJobs) < 0 Then
        MsgBox "First define jobs. Press JobSetter"
        Exit Sub
    End If
    JobChoiceList.Clear
    JobChoiceFrame.Height = 126
    JobChoiceFrame.Visible = True
    JobChoiceList.SetFocus
    AddJobsToList JobChoiceList, ImgJobs
    AddJobsToList JobChoiceList, FcsJobs
    AddSwitchesToList JobChoiceList, currPipeline
    FrameTaskOptions.Visible = False

   On Error GoTo 0
   Exit Sub

AddJobToPipelineButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddJobToPipelineButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub DelJobPipelineButton_Click()
    Dim index As Integer
    Dim newIndex As Integer
On Error GoTo DelJobPipelineButton_Click_Error

    With CurrentPipelineList
        index = .ListIndex
        If index > -1 Then
            Pipelines(currPipeline).delTask index
        Else
            Exit Sub
        End If
        UpdatePipelineList CurrentPipelineList, currPipeline
        If .ListCount = 0 Then
            TrackingFrame.Visible = False
            Exit Sub
        End If
        If .ListCount - 1 >= index Then
            .Selected(index) = True
        Else
            .Selected(.ListCount - 1) = True
        End If
    End With

   On Error GoTo 0
   Exit Sub

DelJobPipelineButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure DelJobPipelineButton_Click of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub JobChoiceList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim index As Integer
    Dim indexImg As Integer
    Dim tmpTask As Task
On Error GoTo JobChoiceList_DblClick_Error

    index = JobChoiceList.ListIndex
    
    If isArrayEmpty(ImgJobs) Then
        indexImg = -1
    Else
        indexImg = UBound(ImgJobs)
    End If
    
    If index > -1 Then
        With CurrentPipelineList
                .AddItem
                .List(.ListCount - 1, 0) = .ListCount
                .List(.ListCount - 1, 1) = JobChoiceList.List(index, 0)
                .List(.ListCount - 1, 2) = JobChoiceList.List(index, 1)
        End With
        Select Case JobChoiceList.List(index, 0)
            Case "Img"
                tmpTask.jobType = jobTypes.imgjob
                tmpTask.jobNr = index
            Case "Fcs"
                tmpTask.jobType = jobTypes.fcsjob
                tmpTask.jobNr = index - (indexImg + 1)
            Case "GoTo"
                tmpTask.jobType = jobTypes.gotoPip
                tmpTask.jobNr = CInt(VBA.Right(JobChoiceList.List(index, 1), Len(JobChoiceList.List(index, 1)) - Len("trigger")))
        End Select
        tmpTask.SaveImage = True
        tmpTask.Period = 1
        Pipelines(currPipeline).addTask tmpTask
        If Pipelines(currPipeline).count = 1 Then
            Pipelines(currPipeline).Repetition.number = CInt(RepetitionNumber.value)
            RepetitionTimeUpdate currPipeline
            Pipelines(currPipeline).maxWait = CDbl(maxWait.value)
            Pipelines(currPipeline).optPtNumber = CInt(optPtNumber.value)
        End If
    End If
    JobChoiceFrame.Visible = False
    If CurrentPipelineList.ListCount > 0 Then
        If CurrentPipelineList.ListIndex < 0 Then
            CurrentPipelineList.ListIndex = 0
        End If
        TrackingFrame.Visible = True
        enableFrame FramePipelineRepetitions, True
        enableFrame FramePipelineTrigger, True
        enableFrame FrameTaskOptions, True
    End If
    FrameTaskOptions.Visible = True
    CurrentPipelineList.ListIndex = CurrentPipelineList.ListCount - 1

   On Error GoTo 0
   Exit Sub

JobChoiceList_DblClick_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure JobChoiceList_DblClick of Form PipelineConstructor at line " & Erl & " "
End Sub


Private Sub JobUpDown_SpinDown()
    Dim index As Integer
On Error GoTo JobUpDown_SpinDown_Error

    index = CurrentPipelineList.ListIndex
    If index >= 0 And index < CurrentPipelineList.ListCount - 1 Then
        Pipelines(currPipeline).swapTask index, index + 1
    Else
        Exit Sub
    End If
    UpdatePipelineList CurrentPipelineList, currPipeline
    CurrentPipelineList.Selected(index + 1) = True

   On Error GoTo 0
   Exit Sub

JobUpDown_SpinDown_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure JobUpDown_SpinDown of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub JobUpDown_SpinUp()
    Dim index As Integer
    
On Error GoTo JobUpDown_SpinUp_Error

    index = CurrentPipelineList.ListIndex
    Debug.Print index
    If index >= 1 Then
        Pipelines(currPipeline).swapTask index, index - 1
    Else
        Exit Sub
    End If
    UpdatePipelineList CurrentPipelineList, currPipeline
    CurrentPipelineList.Selected(index - 1) = True

   On Error GoTo 0
   Exit Sub

JobUpDown_SpinUp_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure JobUpDown_SpinUp of Form PipelineConstructor at line " & Erl & " "
End Sub

''' add imaging or FCS job to list
Private Sub AddJobsToList(List As ListBox, Jobs)
    Dim jobNr As Integer
    Dim prefix As String
On Error GoTo AddJobsToList_Error

    With List
        If Not isArrayEmpty(Jobs) Then
            If TypeOf Jobs(0) Is AJob Then
                prefix = "Img"
            End If
            If TypeOf Jobs(0) Is AFcsJob Then
                prefix = "Fcs"
            End If
            For jobNr = 0 To UBound(Jobs)
                .AddItem
                .List(.ListCount - 1, 0) = prefix
                .List(.ListCount - 1, 1) = Jobs(jobNr).Name
            Next jobNr
        End If
    End With

   On Error GoTo 0
   Exit Sub

AddJobsToList_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddJobsToList of Form PipelineConstructor at line " & Erl & " "
End Sub

''' add a switch from one pipeline to the other
Private Sub AddSwitchesToList(List As ListBox, indexPip As Integer)
    Dim switchModes() As String
    Dim i As Integer
On Error GoTo AddSwitchesToList_Error

    With List
        For i = 1 To UBound(Pipelines)
            If indexPip <> i Then
                .AddItem
                .List(.ListCount - 1, 0) = "GoTo"
                .List(.ListCount - 1, 1) = "trigger" & i
            End If
        Next i
    End With

   On Error GoTo 0
   Exit Sub

AddSwitchesToList_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddSwitchesToList of Form PipelineConstructor at line " & Erl & " "
End Sub


'''
' Clear List and update it according to pipeline with index
'''
Public Sub UpdatePipelineList(List As ListBox, index As Integer)
    Dim jobType As Integer
    Dim jobNr As Integer
    
    Dim i As Integer
On Error GoTo UpdatePipelineList_Error

    List.Clear
    If Pipelines(index).isEmpty Then
        Exit Sub
    End If
    Debug.Print "Counts " & Pipelines(index).count
    
    For i = 0 To Pipelines(index).count - 1
        If i > Pipelines(index).count - 1 Then
            GoTo Nexti
        End If
        jobType = Pipelines(index).getTask(i).jobType
        jobNr = Pipelines(index).getTask(i).jobNr
        'in case no img or fcs jobs delete all entries in Pipeline and move to next pipeline
        Select Case jobType
            Case jobTypes.imgjob
                If isArrayEmpty(ImgJobs) Then
                    Pipelines(index).delTask (i)
                    GoTo Nexti
                End If
            Case jobTypes.fcsjob
                If isArrayEmpty(FcsJobs) Then
                    Pipelines(index).delTask (i)
                    GoTo Nexti
                End If
        End Select
        ''Add entry in list of pipeline
        With List
            Select Case jobType
                Case jobTypes.imgjob
                    If UBound(ImgJobs) >= jobNr Then
                        .AddItem
                        .List(.ListCount - 1, 0) = .ListCount
                        .List(.ListCount - 1, 1) = "Img"
                        .List(.ListCount - 1, 2) = ImgJobs(jobNr).Name
                    Else
                        'if the corresponding jobNr has been removed remove entry in Pipeline
                        Pipelines(index).delTask (i)
                        GoTo Nexti
                    End If
                Case jobTypes.fcsjob
                    If UBound(FcsJobs) >= jobNr Then
                        .AddItem
                        .List(.ListCount - 1, 0) = .ListCount
                        .List(.ListCount - 1, 1) = "Fcs"
                        .List(.ListCount - 1, 2) = FcsJobs(jobNr).Name
                    Else
                        'if the corresponding jobNr has been removed remove entry in Pipeline
                        Pipelines(index).delTask (i)
                        GoTo Nexti
                    End If
                Case jobTypes.gotoPip
                        .AddItem
                        .List(.ListCount - 1, 0) = .ListCount
                        .List(.ListCount - 1, 1) = "GoTo"
                        .List(.ListCount - 1, 2) = "trigger" & jobNr
                End Select
        End With
Nexti:
    Next i

   On Error GoTo 0
   Exit Sub

UpdatePipelineList_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UpdatePipelineList of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub saveImage_Click()
    Dim index As Integer
On Error GoTo saveImage_Click_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setSaveImage index, SaveImage.value
    End If

   On Error GoTo 0
   Exit Sub

saveImage_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure saveImage_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub getSaveImage()
    Dim index As Integer
On Error GoTo getSaveImage_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        SaveImage.value = Pipelines(currPipeline).getSaveImage(index)
    End If

   On Error GoTo 0
   Exit Sub

getSaveImage_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure getSaveImage of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub getZOffset()
    Dim index As Integer
On Error GoTo getZOffset_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        ZOffset.value = Pipelines(currPipeline).getZOffset(index)
    End If

   On Error GoTo 0
   Exit Sub

getZOffset_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure getZOffset of Form PipelineConstructor at line " & Erl & " "
End Sub


''''
' this does not get all the changes
''''
Private Sub ZOffset_Change()
    Dim index As Integer
On Error GoTo ZOffset_Change_Error

    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        Pipelines(currPipeline).setZOffset index, ZOffset.value
    End If

   On Error GoTo 0
   Exit Sub

ZOffset_Change_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ZOffset_Change of Form PipelineConstructor at line " & Erl & " "
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' LOOPING REPETITIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RepetitionTimeUpdate(index As Integer)
On Error GoTo RepetitionTimeUpdate_Error

    If RepetitionSec.value Then
        Pipelines(index).Repetition.Time = CDbl(RepetitionTime.value)
    ElseIf RepetitionMin.value Then
        Pipelines(index).Repetition.Time = CDbl(RepetitionTime.value) * 60
    End If

   On Error GoTo 0
   Exit Sub

RepetitionTimeUpdate_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure RepetitionTimeUpdate of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub RepetitionMinSecUpdate(Min As Boolean)
On Error GoTo RepetitionMinSecUpdate_Error

    If Min Then
        RepetitionMin.BackColor = &HFF8080
        RepetitionSec.BackColor = &H8000000F
    Else
        RepetitionSec.BackColor = &HFF8080
        RepetitionMin.BackColor = &H8000000F
    End If
    RepetitionTime.MAX = 360
    RepetitionTimeUpdate (currPipeline)

   On Error GoTo 0
   Exit Sub

RepetitionMinSecUpdate_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure RepetitionMinSecUpdate of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub RepetitionTime_Click()
    RepetitionTimeUpdate (currPipeline)
End Sub

Private Sub RepetitionInterval_Click()
    Pipelines(currPipeline).Repetition.interval = RepetitionInterval.value
End Sub

Private Sub RepetitionMin_Click()
    RepetitionSec.value = Not RepetitionMin.value
    If RepetitionMin.value Then
        RepetitionMinSecUpdate (True)
    Else
        RepetitionMinSecUpdate (False)
    End If
End Sub

Private Sub RepetitionNumber_Click()
   Pipelines(currPipeline).Repetition.number = CInt(RepetitionNumber.value)
End Sub

Private Sub RepetitionNumber_Change()
    Pipelines(currPipeline).Repetition.number = CInt(RepetitionNumber.value)
End Sub

Private Sub RepetitionSec_Click()
    RepetitionMin.value = Not RepetitionSec.value
End Sub

'''
' update form from pipeline index
'''
Private Sub UpdateRepetitionSettings(index As Integer)
On Error GoTo UpdateRepetitionSettings_Error

    If Pipelines(index).Repetition.Time > 0 And ((Pipelines(index).Repetition.Time Mod 60) = 0 Or Pipelines(index).Repetition.Time > 360) Then
        RepetitionTime.value = Pipelines(index).Repetition.Time / 60
        RepetitionMin.value = True
    Else
        RepetitionTime.value = Pipelines(index).Repetition.Time
        RepetitionSec.value = True
    End If
    RepetitionNumber.value = Pipelines(index).Repetition.number
    RepetitionInterval.value = Pipelines(index).Repetition.interval

   On Error GoTo 0
   Exit Sub

UpdateRepetitionSettings_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UpdateRepetitionSettings of Form PipelineConstructor at line " & Erl & " "
End Sub

'''
' time point where to acquire an image
'''
Private Sub StartOption_Click()
    setPeriod
    Period.Enabled = False
    PeriodButton.Enabled = False
End Sub

Private Sub EndOption_Click()
    setPeriod
    Period.Enabled = False
    PeriodButton.Enabled = False
End Sub

Private Sub PeriodOption_Click()
    setPeriod
    Period.Enabled = True
    PeriodButton.Enabled = True
End Sub

Private Sub PeriodButton_SpinUp()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        If Period.value < RepetitionNumber - 1 Then
            Period.value = Period.value + 1
        Else
            MsgBox "The period of acquisition cannot be higher than number of Repetitions , i.e. " & RepetitionNumber, VbInformation
            Exit Sub
        End If
        Pipelines(currPipeline).setPeriod index, Period.value
    End If
End Sub

Private Sub PeriodButton_SpinDown()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        If Period.value > 1 Then
            Period.value = Period.value - 1
        End If
        Pipelines(currPipeline).setPeriod index, Period.value
    End If
End Sub

Private Sub setPeriod()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        If PeriodOption Then
            Pipelines(currPipeline).setPeriod index, Period.value
        End If
        If StartOption Then
            Pipelines(currPipeline).setPeriod index, 0
        End If
        If EndOption Then
            Pipelines(currPipeline).setPeriod index, -1
        End If
    End If
End Sub

Private Sub getPeriod()
    Dim index As Integer
    index = CurrentPipelineList.ListIndex
    If index > -1 Then
        With Pipelines(currPipeline)
            If .getPeriod(index) > 0 Then
                Period.value = .getPeriod(index)
                PeriodOption.value = True
                Period.Enabled = True
                
            End If
            If .getPeriod(index) = 0 Then
                StartOption.value = True
                Period.Enabled = False
            End If
            If .getPeriod(index) = -1 Then
                EndOption.value = True
                Period.Enabled = False
            End If
        End With
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' FILE OUTPUT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''
' Set output folder
''''''''''''''''''''''''''''''''''''''
Private Sub CommandButtonNewDataBase_Click()
    Dim Filter As String, FileName As String
    Dim Flags As Long
    Dim DefDir As String

On Error GoTo CommandButtonNewDataBase_Click_Error

    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE

    'Filter = "All Data (*.*)" & Chr$(0) & "*.*"
    If GlobalDataBaseName = "" Then
        DefDir = "C:\"
    Else
        DefDir = GlobalDataBaseName
    End If
    
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", DefDir, "Select output folder")
    If Len(FileName) > 3 Then
        FileName = VBA.Left(FileName, Len(FileName) - 3)
        DatabaseTextbox.value = FileName
        SetDatabase
    End If

   On Error GoTo 0
   Exit Sub

CommandButtonNewDataBase_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure CommandButtonNewDataBase_Click of Form PipelineConstructor at line " & Erl & " "
End Sub

'''''
'   Only update the outputfolder when enter is pressed. This avoids creating a folder at every keystroke
'''''
Private Sub DatabaseTextbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'this is the enter key
        SetDatabase
    End If
End Sub

Private Sub TextBoxFileName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then 'this is the enter key
        SetFileName
    End If
End Sub

'''''
'   Set global variables for files and check if we can create Outputfolder
'       [GlobalDataBaseName] Out/Global - The name of Outputfolder
'       [LogFileNameBase]    Out/Global - The name of the LogfileName
'       Log]                Out/Global - If yes results are logged
'''''
Private Sub SetDatabase()

On Error GoTo SetDatabase_Error

    GlobalDataBaseName = DatabaseTextbox.value

    If Not GlobalDataBaseName = "" Then
        If VBA.Right(GlobalDataBaseName, 1) <> "\" Then
            DatabaseTextbox.value = DatabaseTextbox.value + "\"
            GlobalDataBaseName = DatabaseTextbox.value
        End If
        On Error GoTo ErrorHandleDataBase
        If Not CheckDir(GlobalDataBaseName) Then
            Exit Sub
        End If
        AFMSettings.writeKeyToRegistry "OutputFolder", GlobalDataBaseName
        LogFileNameBase = GlobalDataBaseName & "\PipelineConstructor.log"
        ErrFileNameBase = GlobalDataBaseName & "\PipelineConstructor.err"
        If VBA.Right(GlobalDataBaseName, 1) = "\" Then
            BackSlash = ""
        Else
            BackSlash = "\"
        End If
    End If

    If LogFileNameBase <> "" Then
        On Error GoTo ErrorHandleLogFile
        LogFileName = LogFileNameBase
        ErrFileName = ErrFileNameBase
        Log = True
    Else
        Log = False
    End If
    Exit Sub
ErrorHandleDataBase:
    MsgBox "Could not create output Directory " & GlobalDataBaseName
    Exit Sub
ErrorHandleLogFile:
    MsgBox "Could not create LogFile " & LogFileName

   On Error GoTo 0
   Exit Sub

SetDatabase_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SetDatabase of Form PipelineConstructor at line " & Erl & " "
End Sub

Private Sub SetFileName()
On Error GoTo SetFileName_Error

    If TextBoxFileName.value <> "" Then
        If VBA.Right(TextBoxFileName.value, Len(FNSep)) <> FNSep Then
            TextBoxFileName.value = TextBoxFileName.value & FNSep
        End If
    End If

   On Error GoTo 0
   Exit Sub

SetFileName_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SetFileName of Form PipelineConstructor at line " & Erl & " "
End Sub

'''set global variables for file format
Private Sub fileFormatczi_Click()
#If ZENvC > 2010 Then
    imgFileFormat = eAimExportFormatCzi
    imgFileExtension = ".czi"
#Else
    MsgBox "Your ZEN version does not support czi files", VbInformation
    fileFormatlsm.value = True
#End If
End Sub

Private Sub fileFormatlsm_Click()
    imgFileFormat = eAimExportFormatLsm5
    imgFileExtension = ".lsm"
End Sub
