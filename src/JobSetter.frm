VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JobSetter 
   Caption         =   "JobSetter"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
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
Public WithEvents EventMng As EventAdmin
Attribute EventMng.VB_VarHelpID = -1
Private Sub AcquireFcsJobButton_Click()
    Dim index As Integer
    Dim newPosition() As Vector
    ReDim newPosition(0) ' position where FCS will be done
    Dim currentPosition As Vector
    ScanStop = False
    index = FcsJobList.ListIndex
    If index = -1 Then
        MsgBox "FcsJob list is empty"
        Exit Sub
    End If
    'for Fcs the position for ZEN are passed in meter!! (different to Lsm5.Hardware.CpStages is in um!!)
    ' For X and Y relative position to center. For Z absolute position in meter
    newPosition(0).x = 0
    newPosition(0).y = 0
    newPosition(0).Z = Lsm5.Hardware.CpFocus.position * 0.000001 'convet from um to meter
    'eventually force creation of FcsRecord
    If Not GlobalFcsRecordingDoc Is Nothing Then
        GlobalFcsRecordingDoc.BringToTop
    End If
    NewFcsRecordGui GlobalFcsRecordingDoc, GlobalFcsData, FcsJobs(index).Name, ZEN, ZenV
    'this brings record to top
    FcsJobs(index).PutJob ZEN, ZenV
    Application.ThrowEvent eEventScanStart, 1
    ScanToFcs GlobalFcsRecordingDoc, GlobalFcsData
End Sub

Private Sub AcquireJobButton_Click()
    Dim index As Integer
    Dim Time As Double
    
On Error GoTo AcquireJobButton_Click_Error
    ScanStop = False
    index = ImgJobList.ListIndex

    If index = -1 Then
        MsgBox "Job list is empty"
        Exit Sub
    End If
    If Not GlobalRecordingDoc Is Nothing Then
        GlobalRecordingDoc.BringToTop
    End If
    NewRecordGui GlobalRecordingDoc, ImgJobs(index).Name, ZEN, ZenV
    If ZenV > 2010 And Not ZEN Is Nothing Then
        Dim vo As AimImageVectorOverlay
        Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions
        If vo.GetNumberElements > 0 Then
            ZEN.gui.Acquisition.Regions.Delete.Execute
        End If
    End If
    Dim position As Vector
    Lsm5.Hardware.CpStages.GetXYPosition position.x, position.y
    position.Z = Lsm5.Hardware.CpFocus.position
    Running = True
    'currentImgJob = -1
    AcquireJob index, ImgJobs(index), GlobalRecordingDoc, ImgJobs(index).Name, position
    
    'for imaging the position to image can be passed directly to AcquireJob. ZEN uses the absolute position in um
    'NewRecordGui GlobalRecordingDoc, ImgJobs(index).Name, ZEN, ZENv
    'ImgJobs(index).PutJob ZEN
    'GlobalRecordingDoc.Recording.StartScanEvent = eStartScanUser
    'Lsm5.StartScan
    'Application.ThrowEvent eEventScanStart, 1
    'ScanToImage GlobalRecordingDoc

   On Error GoTo 0
   Exit Sub

AcquireJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AcquireJobButton_Click of Form JobSetter at line " & Erl & " "

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





Private Sub StopButton_Click()
    StopAcquisition
End Sub

Private Sub StopFcsButton_Click()
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
        MsgBox "Job list is empty or you need to select one job"
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
        MsgBox "Job list is empty or you need to select one job"
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
        MsgBox "Job list is empty or you need to select one job"
        Exit Sub
    End If
    ImgJobs(index).PutJob ZEN
    If ZenV > 2010 Then  'On 2010 it is extremely slow and the command does not wait for finishing
        Application.ThrowEvent tag_Events.eEventDsActiveRecChanged, 0
        DoEvents
    End If
End Sub


Private Sub PutFcsJob_Click()
    Dim index As Integer
    index = FcsJobList.ListIndex
    If index = -1 Then
        MsgBox "Job list is empty or you need to select one job"
        Exit Sub
    End If
    FcsJobs(index).PutJob ZEN, ZenV
End Sub

Private Sub AddFcsJobButton_Click()
    Dim i As Integer
    Dim OpenForms() As Boolean
    Dim ListEntry As Variant

    If FcsJobName = "" Then
        MsgBox "You need to specify a name for the fcs job"
        Exit Sub
    End If
    If Not UniqueListName(FcsJobList, FcsJobName) Or Not UniqueListName(ImgJobList, FcsJobName) Then
        MsgBox "Name of fcs job must be unique"
        Exit Sub
    End If
    OpenForms = HideShowForms(OpenForms)
    FcsJobList.AddItem FcsJobName.value
    FcsJobList.Selected(FcsJobList.ListCount - 1) = True
    AddFcsJob FcsJobs, FcsJobList.List(FcsJobList.ListCount - 1), ZEN
    setFcsLabels FcsJobList.ListCount - 1
    HideShowForms OpenForms
    'PipelineConstructor.UpdateFcsJobList
End Sub

Private Function UniqueListName(List As ListBox, JobName As String) As Boolean
    Dim ListEntry As Variant
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
    If JobName.value = "" Then
        MsgBox "You need to specify a name for the job"
        Exit Sub
    End If
    If Not UniqueListName(FcsJobList, JobName) Or Not UniqueListName(ImgJobList, JobName) Then
        MsgBox "Name of imaging job must be unique"
        Exit Sub
    End If
    ImgJobList.AddItem JobName.value
    ImgJobList.Selected(ImgJobList.ListCount - 1) = True
    AddJob ImgJobs, ImgJobList.List(ImgJobList.ListCount - 1), Lsm5.DsRecording, ZEN
    setLabels ImgJobList.ListCount - 1
    setTrackNames ImgJobList.ListCount - 1
    'PipelineConstructor.UpdateImgJobList
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
    JobsV(UBound(JobsV)).SetJob Lsm5.DsRecording, ZEN
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
        MsgBox "Nothing to delete"
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
    'PipelineConstructor.UpdateFcsJobList
End Sub

Private Sub DeleteFcsJob(JobsV() As AFcsJob, index As Integer, Optional Name As String = "")
    Dim i As Integer
    Dim IJob As Integer
    If isArrayEmpty(JobsV) Then
        MsgBox "Nothing to delete"
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

Public Sub UserForm_Initialize()
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
        TestBool = ZEN.gui.Acquisition.EnableTimeSeries.value
        ZEN.gui.Acquisition.EnableTimeSeries.value = Not TestBool
        ZEN.gui.Acquisition.EnableTimeSeries.value = TestBool
        GoTo NoError
errorMsg:
        MsgBox "Version is ZEN" & ZenV & " but can't find Zeiss.Micro.AIM.ApplicationInterface." & vbCrLf _
        & "Using ZEN2010 settings instead." & vbCrLf _
        & "Check if Zeiss.Micro.AIM.ApplicationInterface.dll is registered?" _
        & "See also the manual how to register a dll into windows."
        ZenV = 2010

NoError:
    End If
    'Erase ImgJobs
    'Erase FcsJobs
    
    TrackVisible False
    UpdateImgListbox ImgJobList, ImgJobs
    UpdateFcsListbox FcsJobList, FcsJobs
    
    
    'PipelineConstructor.Show
End Sub
