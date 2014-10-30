VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JobSetter 
   Caption         =   "JobSetter"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   OleObjectBlob   =   "JobSetter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "JobSetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The event manager as been disabled for the moment as it causes several crashes
Public WithEvents EventMng As EventAdmin
Attribute EventMng.VB_VarHelpID = -1



Private Sub EventMng_Ready()
    setStatus True
End Sub


Private Sub EventMng_Busy()
    setStatus False
End Sub

Private Sub FcsJobList_Click()
    Dim index As Integer
    index = FcsJobList.ListIndex
    If index = -1 Then
        Exit Sub
    End If
    On Error Resume Next
    setFcsLabels index
End Sub

Public Sub UserForm_Initialize()
    Dim strIconPath As String
    Dim lngIcon As Long
    Dim lnghWnd As Long
    ZenV = getVersionNr
    'find the version of the software
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
        MsgBox "Version is ZEN" & ZenV & " but can't find Zeiss.Micro.AIM.ApplicationInterface." & vbCrLf _
        & "Using ZEN2010 settings instead." & vbCrLf _
        & "Check if Zeiss.Micro.AIM.ApplicationInterface.dll is registered?" _
        & "See also the manual how to register a dll into windows.", VbCritical, "JobSetter Error"
        ZenV = 2010
NoError:
    End If
    
    TrackVisible False
    UpdateImgListbox ImgJobList, ImgJobs
    UpdateFcsListbox FcsJobList, FcsJobs
    Set EventMng = New EventAdmin
    EventMng.initialize
    strIconPath = Application.ProjectFilePath & "\resources\micronaut_mc.ico"
    ' Get the icon from the source
    lngIcon = ExtractIcon(0, strIconPath, 0)
    ' Get the window handle of the userform
    lnghWnd = FindWindow("ThunderDFrame", Me.Caption)
    'Set the big (32x32) and small (16x16) icons
    SendMessage lnghWnd, WM_SETICON, True, lngIcon
    SendMessage lnghWnd, WM_SETICON, False, lngIcon
    FormatUserForm Me.Caption
      
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        JobSetter.Hide
        Cancel = True
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Acquisition buttons     '''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub AcquireFcsJobButton_Click()
    Dim index As Integer
    Dim newPosition(0) As Vector

On Error GoTo AcquireFcsJobButton_Click_Error
    resetStopFlags
    index = FcsJobList.ListIndex
    If index = -1 Then
        MsgBox "FcsJob list is empty", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    'for Fcs the position for ZEN are passed in meter!! (different to Lsm5.Hardware.CpStages is in um!!)
    ' For X and Y relative position to center. For Z absolute position in meter
    newPosition(0).X = 0
    newPosition(0).Y = 0
    newPosition(0).Z = Lsm5.Hardware.CpFocus.position * 0.000001 'convert from um to meter
    NewFcsRecordGui GlobalFcsRecordingDoc, GlobalFcsData, FcsJobs(index).Name, ZEN, ZenV
    If Not GlobalFcsRecordingDoc Is Nothing Then
        GlobalFcsRecordingDoc.BringToTop
    End If
    currentFcsJob = -1
    Running = True
    EventMng.setBusy
    Application.ThrowEvent tag_Events.eEventScanStart, 0 'notify that acquisition is started

    If Not CleanFcsData(GlobalFcsRecordingDoc, GlobalFcsData) Then
        Exit Sub
    End If
    AcquireFcsJob index, FcsJobs(index), GlobalFcsRecordingDoc, GlobalFcsData, FcsJobs(index).Name, newPosition
    Application.ThrowEvent tag_Events.eEventScanStop, 0 'notify that acquisition is started
    EventMng.setReady
    Running = False

   On Error GoTo 0
   Exit Sub

AcquireFcsJobButton_Click_Error:
    Running = False
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AcquireFcsJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub


Private Sub AcquireImgJobButton_Click()
    Dim index As Integer
    
    resetStopFlags
    index = ImgJobList.ListIndex
    If index = -1 Then
        MsgBox "List is empty or no imaging job are highlighted", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Running = True
    EventMng.setBusy
    AcquireImgJob (index)
    EventMng.setReady
    Running = False
End Sub

'''
' Acquire imaging job from ImgJobs array at index
'''
Private Sub AcquireImgJob(index As Integer)
    Dim position As Vector
On Error GoTo AcquireJobIndex_Error
    If Not GlobalRecordingDoc Is Nothing Then
        GlobalRecordingDoc.BringToTop
    End If
    NewRecordGui GlobalRecordingDoc, ImgJobs(index).Name, ZEN, ZenV
    If ZenV > 2010 And Not ZEN Is Nothing Then
        Dim vo As AimImageVectorOverlay
        Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions
        If vo.GetNumberElements > 0 Then
            ZEN.GUI.Acquisition.Regions.Delete.Execute
        End If
    End If
    'start acquisition
    currentImgJob = -1
    AcquireJob index, ImgJobs(index), GlobalRecordingDoc, ImgJobs(index).Name, getCurrentPosition
   On Error GoTo 0
   Exit Sub
   
AcquireJobIndex_Error:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AcquireJobIndex of Form JobSetter at line " & Erl & " "
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Mange Imaging jobs      '''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ChangeImgJobName_Click()
    Dim index As Integer
    Dim Name As String
    index = ImgJobList.ListIndex
    If index = -1 Then
        Exit Sub
    End If
    'imgJobs(index).Name =
    
    Name = InputBox("Update name of current job", "JobSetter: Update Name", ImgJobList.List(index))
    If Name = "" Or Not UniqueListName(ImgJobList, Name) Then
        Exit Sub
    End If
    ImgJobs(index).Name = Name
    UpdateImgListbox ImgJobList, ImgJobs
    ImgJobList.Selected(index) = True
End Sub


Private Sub ChangeFcsJobName_Click()
    Dim index As Integer
    Dim Name As String
    index = FcsJobList.ListIndex
    If index = -1 Then
        Exit Sub
    End If
    'imgJobs(index).Name =
    
    Name = InputBox("Update name of current job", "JobSetter: Update Name", FcsJobList.List(index))
    If Name = "" Or Not UniqueListName(FcsJobList, Name) Then
        Exit Sub
    End If
    FcsJobs(index).Name = Name
    UpdateFcsListbox FcsJobList, FcsJobs
    FcsJobList.Selected(index) = True
End Sub

Private Sub ImgJobList_Click()
    Dim index As Integer
    index = ImgJobList.ListIndex
    If index = -1 Then
        Exit Sub
    End If
    On Error Resume Next
    setLabels index
    setTrackNames index
End Sub



Private Sub AddImgJobFromFileButton_Click()
    Dim fso As New FileSystemObject
    Dim Filter As String, fileName As String
    Dim index As Integer
    Dim Flags As Long
    Dim DefDir As String
    Dim fileNames() As String
    
    '''get filename(s) to be loaded into ZEN'''
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
#If ZENvC > 2010 Then
    Flags = OFN_LONGNAMES Or OFN_ALLOWMULTISELECT Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or _
    OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Images (*.lsm,*.czi)" & Chr$(0) & "*.lsm;*.czi" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
#Else
    Flags = OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or _
    OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Images (*.lsm)" & Chr$(0) & "*.lsm" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
#End If
    fileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DefDir, "Select file(s) to be loaded as imaging jobs")
    
    If fileName = "" Then
        Exit Sub
    End If
    
    fileNames = Split(fileName, Chr$(0))
    EventMng.setBusy
    If UBound(fileNames) = 0 Then
        AddImgJobFromFile fileNames(0)
        WorkingDir = fso.GetParentFolderName(fileNames(0)) & "\"
    Else
        For index = 1 To UBound(fileNames)
            AddImgJobFromFile fileNames(0) & "\" & fileNames(index)
            SleepWithEvents 2000
        Next index
        WorkingDir = fileNames(0) & "\"
    End If
    EventMng.setReady
End Sub


Private Sub AddImgJobFromFile(fileName As String)
    Dim fso As New FileSystemObject
    Dim JobName As String
    If Not FileExist(fileName) Then
        Exit Sub
    End If
    JobName = VBA.Split(fso.GetFileName(fileName), ".")(0)
    If Not UniqueListName(FcsJobList, JobName) Or Not UniqueListName(ImgJobList, JobName) Then
        MsgBox "Name of imaging job must be unique!", VbExclamation, "JobSetter warning"
        Exit Sub
    End If
    ImgJobList.AddItem JobName
    ImgJobList.Selected(ImgJobList.ListCount - 1) = True
    AddJob ImgJobs, ImgJobList.List(ImgJobList.ListCount - 1), getRecordingFromImageFile(fileName, ZEN), ZEN
    setLabels ImgJobList.ListCount - 1
    setTrackNames ImgJobList.ListCount - 1
End Sub
    
Private Sub SaveButton_Click()
    
    Dim fso As New FileSystemObject
    Dim index As Integer
    Dim Flags As Long
    Dim dirName As String, DefDir As String, Filter As String
    Dim answ As Integer
    If Not ImgJobList.ListCount > 0 Then
        MsgBox "No Imaging jobs defined yet!", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    
    
    If isArrayEmpty(ImgJobs) Then
        MsgBox "Sorry there has been problems in saving the jobs!", vbError, "JobSetter Error"
        ImgJobList.Clear
        Exit Sub
    End If
    answ = MsgBox("Yes: All jobs are executed and saved" & vbCrLf & "No: Highlighted job is executed and saved" & vbCrLf, VbYesNoCancel + VbQuestion, "JobSetter: Save jobs")
    If answ = vbCancel Then Exit Sub
    If answ = vbNo And ImgJobList.ListIndex < 0 Then
        MsgBox "Highlight one of the jobs to be saved", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    
    Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_EXPLORER Or OFN_NOVALIDATE
    Filter = "Images (*.lsm,*.czi)" & Chr$(0) & "*.lsm;*.czi" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*"
    If WorkingDir = "" Then
        DefDir = "C:\"
    Else
        DefDir = WorkingDir
    End If
    setStatus False
    dirName = CommonDialogAPI.ShowOpen(Filter, Flags, "*.*", DefDir, "Select output folder for jobs")
    If dirName = "" Then Exit Sub
    dirName = fso.GetParentFolderName(dirName) & "\"
    WorkingDir = dirName
#If ZENvC > 2010 Then
    If answ = vbNo Then
        AcquireImgJob ImgJobList.ListIndex
        SaveDsRecordingDoc GlobalRecordingDoc, dirName & ImgJobs(ImgJobList.ListIndex).Name & ".czi", eAimExportFormatCzi
    Else
        For index = 0 To ImgJobList.ListCount - 1
            AcquireImgJob (index)
            SaveDsRecordingDoc GlobalRecordingDoc, dirName & ImgJobs(index).Name & ".czi", eAimExportFormatCzi
        Next index
    End If
#Else
    If answ = vbNo Then
        AcquireImgJob ImgJobList.ListIndex
        SaveDsRecordingDoc GlobalRecordingDoc, dirName & ImgJobs(ImgJobList.ListIndex).Name & ".lsm", eAimExportFormatLsm5
    Else
        For index = 0 To ImgJobList.ListCount - 1
            AcquireImgJob (index)
            SaveDsRecordingDoc GlobalRecordingDoc, dirName & ImgJobs(index).Name & ".lsm", eAimExportFormatLsm5
        Next index
    End If
#End If
    setStatus True
End Sub

Private Sub StopButton_Click()
    ScanStop = True
    StopAcquisition
    
End Sub

Private Sub StopFcsButton_Click()
    ScanStop = True
    StopAcquisition
End Sub

Private Sub Track1_Click()
    TrackClick (1)
End Sub

Private Sub Track2_Click()
    TrackClick (2)
End Sub

Private Sub Track3_Click()
    TrackClick (3)
End Sub

Private Sub Track4_Click()
    TrackClick (4)
End Sub

Private Sub setStatus(value As Boolean)
    If value Then
        StatusLabel = "READY"
        StatusLabel.ForeColor = &HC000&
    Else
        StatusLabel = "BUSY"
        StatusLabel.ForeColor = &HC0&
    End If
End Sub



Private Sub TrackClick(iTrack As Integer)
    Dim index As Integer
    index = ImgJobList.ListIndex
    If index <> -1 Then
        ImgJobs(index).setAcquireTrack iTrack - 1, Me.Controls("Track" + CStr(iTrack)).value
    End If
End Sub


Private Sub TrackVisible(Visible As Boolean)
    Track1.Visible = Visible
    Track2.Visible = Visible
    Track3.Visible = Visible
    Track4.Visible = Visible
End Sub

Private Sub SetJobButton_Click()
    Dim index As Integer
    index = ImgJobList.ListIndex
    If index = -1 Then
        MsgBox "Job list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Debug.Assert (ImgJobs(index).SetJob(Lsm5.DsRecording, ZEN))
    setLabels index
    setTrackNames index
End Sub

Private Sub SetFcsJob_Click()
    Dim index As Integer
    index = FcsJobList.ListIndex
    If index = -1 Then
         MsgBox "FcsJob list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Debug.Assert (FcsJobs(index).SetJob(ZEN, ZenV))
    setFcsLabels index
End Sub

Private Sub setLabels(index As Integer)
    Dim jobDescription() As String
    jobDescription = ImgJobs(index).splittedJobDescriptor(13, ImgJobs(index).jobDescriptor)
    JobLabel1.Caption = jobDescription(0)
    JobLabel2.Caption = jobDescription(1)
End Sub

Private Sub setFcsLabels(index As Integer)
    Dim jobDescription() As String
    jobDescription = FcsJobs(index).splittedJobDescriptor(13, FcsJobs(index).jobDescriptor)
    FcsJobLabel1.Caption = jobDescription(0)
    FcsJobLabel2.Caption = jobDescription(1)
End Sub

Private Sub PutJobButton_Click()
    Dim index As Integer
    index = ImgJobList.ListIndex
    If index = -1 Then
        MsgBox "Job list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    ImgJobs(index).PutJob ZEN
    If ZenV > 2009 Then  'On 2010 it is extremely slow and the command does not wait for finishing
        Application.ThrowEvent tag_Events.eEventDsActiveRecChanged, 0
        DoEvents
    End If
End Sub


Private Sub PutFcsJob_Click()
    Dim index As Integer
    index = FcsJobList.ListIndex
    If index = -1 Then
        MsgBox "FcsJob list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    FcsJobs(index).PutJob ZEN, ZenV
End Sub

Private Sub AddFcsJobButton_Click()
    Dim i As Integer
    Dim index As Integer
    Dim Name As String
    Dim OpenForms() As Boolean
    Name = InputBox("Name of job to be created from current ZEN settings", "JobSetter: Define FcsJob name")
    'Cancel pressed
    If StrPtr(Name) = 0 Then Exit Sub

    If Name = "" Or Not UniqueListName(ImgJobList, CStr(Name)) Or Not UniqueListName(FcsJobList, CStr(Name)) Then
        MsgBox "You need to define an unique name for the imaging job!", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    OpenForms = HideShowForms(OpenForms)
    FcsJobList.AddItem CStr(Name)
    FcsJobList.Selected(FcsJobList.ListCount - 1) = True
    AddFcsJob FcsJobs, FcsJobList.List(FcsJobList.ListCount - 1), ZEN
    setFcsLabels FcsJobList.ListCount - 1
    HideShowForms OpenForms
'    ImgJobList.AddItem CStr(Name)
'    ImgJobList.Selected(ImgJobList.ListCount - 1) = True
'    AddJob ImgJobs, ImgJobList.List(ImgJobList.ListCount - 1), Lsm5.DsRecording, ZEN
'    setLabels ImgJobList.ListCount - 1
'    setTrackNames ImgJobList.ListCount - 1
'
'
'    Dim i As Integer
'
'    Dim ListEntry As Variant
'
'    If FcsJobName = "" Then
'        MsgBox "You need to specify a name for the fcs job"
'        Exit Sub
'    End If
'    If Not UniqueListName(FcsJobList, FcsJobName) Or Not UniqueListName(ImgJobList, FcsJobName) Then
'        MsgBox "Name of fcs job must be unique"
'        Exit Sub
'    End If

    'PipelineConstructor.UpdateFcsJobList
End Sub

Private Function UniqueListName(List As ListBox, JobName As String) As Boolean
    Dim ListEntry As Variant
    Debug.Print List.ListCount
    If List.ListCount > 0 Then
        For Each ListEntry In List.List
            If StrComp(ListEntry, JobName) = 0 Then
                Exit Function
            End If
        Next
    End If
    UniqueListName = True
End Function
    

Public Function HideShowForms(OpenForms() As Boolean) As Boolean()
    Dim UForm As Object
    Dim i As Integer
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
End Function

Private Sub AddJobButton_Click()
    Dim i As Integer
    Dim index As Integer
    Dim Name As String

    Name = InputBox("Name of imaging job to be created from current ZEN settings", "JobSetter: Define job name")
    'Cancel pressed
    If StrPtr(Name) = 0 Then Exit Sub

    If Name = "" Or Not UniqueListName(ImgJobList, CStr(Name)) Or Not UniqueListName(FcsJobList, CStr(Name)) Then
        MsgBox "You need to define an unique name for the imaging job!", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    ImgJobList.AddItem CStr(Name)
    ImgJobList.Selected(ImgJobList.ListCount - 1) = True
    AddJob ImgJobs, ImgJobList.List(ImgJobList.ListCount - 1), Lsm5.DsRecording, ZEN
    setLabels ImgJobList.ListCount - 1
    setTrackNames ImgJobList.ListCount - 1
End Sub

Private Sub setTrackNames(index As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim iTrack As Integer
    Dim Track As DsTrack
    Dim ChannelOK As Boolean
    Dim AcquireTrack() As Boolean
    Dim MaxTracks As Long
    MaxTracks = ImgJobs(index).Recording.GetNormalTrackCount
    AcquireTrack = ImgJobs(index).AcquireTrack
    For i = 0 To 3
        If iTrack < 5 Then
            ChannelOK = False
            Set Track = ImgJobs(index).Recording.TrackObjectByMultiplexOrder(i, 1)
            For j = 0 To Track.DataChannelCount - 1
                If Track.DataChannelObjectByIndex(j, 1).Acquire = True Then
                    ChannelOK = True
                End If
            Next j
            If ChannelOK And (Not Track.IsLambdaTrack) And (Not Track.IsBleachTrack) Then
                Me.Controls("Track" + CStr(iTrack + 1)).Visible = True
                Me.Controls("Track" + CStr(iTrack + 1)).value = AcquireTrack(iTrack)
                Me.Controls("Track" + CStr(iTrack + 1)).Caption = Track.Name
                Me.Controls("Track" + CStr(iTrack + 1)).Enabled = True
                iTrack = iTrack + 1
            Else
                Me.Controls("Track" + CStr(iTrack + 1)).Visible = False
                iTrack = iTrack + 1
            End If
        End If
    Next i
End Sub

Private Sub UpdateImgListbox(List1 As ListBox, JobArray() As AJob)
    Dim i As Integer
    If isArrayEmpty(JobArray) Then
        List1.Clear
        Exit Sub
    Else
        List1.Clear
        For i = 0 To UBound(JobArray)
            List1.AddItem JobArray(i).Name
        Next i
    End If
End Sub

Private Sub UpdateFcsListbox(List1 As ListBox, JobArray() As AFcsJob)
    Dim i As Integer
    If isArrayEmpty(JobArray) Then
        List1.Clear
        Exit Sub
    Else
        List1.Clear
        For i = 0 To UBound(JobArray)
            List1.AddItem JobArray(i).Name
        Next i
    End If
End Sub


Private Sub DeleteJobButton_Click()
    Dim index As Integer
    index = ImgJobList.ListIndex
    If index <> -1 Then
        DeleteJob ImgJobs, index, ImgJobList.List(index)
        ImgJobList.RemoveItem index
    End If
    'PipelineConstructor.UpdateImgJobList
End Sub


Public Sub AddJob(JobsV() As AJob, Name As String, Recording As DsRecording, ZEN As Object)
    If isArrayEmpty(JobsV) Then
        ReDim JobsV(0)
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) + 1)
    End If
    Set JobsV(UBound(JobsV)) = New AJob
    JobsV(UBound(JobsV)).Name = Name
    JobsV(UBound(JobsV)).SetJob Recording, ZEN
End Sub


Public Sub AddFcsJob(JobsV() As AFcsJob, Name As String, ZEN As Object)
    If isArrayEmpty(JobsV) Then
        ReDim JobsV(0)
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) + 1)
    End If
    Set JobsV(UBound(JobsV)) = New AFcsJob
    JobsV(UBound(JobsV)).Name = Name
    JobsV(UBound(JobsV)).SetJob ZEN, ZenV
End Sub



'''
' DeleteJob
'   Delete Job and decrease number of Jobs
'''
Public Sub DeleteJob(JobsV() As AJob, index As Integer, Optional Name As String = "")
    Dim i As Integer
    Dim IJob As Integer
    If isArrayEmpty(JobsV) Then
        MsgBox "Nothing to delete!", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Debug.Assert (index <= UBound(JobsV))
    If Name <> "" Then
        Debug.Assert (StrComp(JobsV(index).Name, Name) = 0)
    End If
    For i = index To UBound(JobsV) - 1
        Set JobsV(i) = JobsV(i + 1)
    Next i
    If UBound(JobsV) = 0 Then
        Erase JobsV
        TrackVisible False
        JobLabel1.Caption = ""
        JobLabel2.Caption = ""
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) - 1)
    End If
End Sub




Private Sub DeleteFcsJobButton_Click()
    Dim index As Integer
    index = FcsJobList.ListIndex
    If index <> -1 Then
        DeleteFcsJob FcsJobs, index, FcsJobList.List(index)
        FcsJobList.RemoveItem index
    End If
End Sub

Private Sub DeleteFcsJob(JobsV() As AFcsJob, index As Integer, Optional Name As String = "")
    Dim i As Integer
    Dim IJob As Integer
    If isArrayEmpty(JobsV) Then
        MsgBox "Nothing to delete", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Debug.Assert (index <= UBound(JobsV))
    If Name <> "" Then
        Debug.Assert (StrComp(JobsV(index).Name, Name) = 0)
    End If
    For i = index To UBound(JobsV) - 1
        Set JobsV(i) = JobsV(i + 1)
    Next i
    If UBound(JobsV) = 0 Then
        Erase JobsV
        FcsJobLabel1.Caption = ""
        FcsJobLabel2.Caption = ""
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) - 1)
    End If
End Sub

