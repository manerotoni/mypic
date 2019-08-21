VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} JobSetter 
   Caption         =   "JobSetter"
   ClientHeight    =   6396
   ClientLeft      =   48
   ClientTop       =   372
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
'---------------------------------------------------------------------------------------
' Module    : JobSetter
' Author    : Antonio Politi
' Date      : 23/10/2017
' Purpose   :
'---------------------------------------------------------------------------------------

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
On Error GoTo FcsJobList_Click_Error

    index = FcsJobList.ListIndex
    If index = -1 Then
        Exit Sub
    End If
    On Error Resume Next
    setFcsLabels index

   On Error GoTo 0
   Exit Sub

FcsJobList_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure FcsJobList_Click of Form JobSetter at line " & Erl & " "
End Sub

Public Sub UserForm_Initialize()
    Dim strIconPath As String
    Dim lngIcon As Long
    Dim lnghWnd As Long
    'find the version of the software
On Error GoTo UserForm_Initialize_Error

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
    UpdateJobListbox ImgJobList, ImgJobs
    UpdateJobListbox FcsJobList, FcsJobs
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

   On Error GoTo 0
   Exit Sub

UserForm_Initialize_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UserForm_Initialize of Form JobSetter at line " & Erl & " "
      
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error GoTo UserForm_QueryClose_Error

    If CloseMode = vbFormControlMenu Then
        JobSetter.Hide
        Cancel = True
    End If

   On Error GoTo 0
   Exit Sub

UserForm_QueryClose_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UserForm_QueryClose of Form JobSetter at line " & Erl & " "
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
    NewFcsRecordGui GlobalFcsRecordingDoc, GlobalFcsData, "FCS:" & FcsJobs(index).Name, ZEN, ZenV
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
    
On Error GoTo AcquireImgJobButton_Click_Error

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

   On Error GoTo 0
   Exit Sub

AcquireImgJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AcquireImgJobButton_Click of Form JobSetter at line " & Erl & " "
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
    NewRecordGui GlobalRecordingDoc, "IMG:" & ImgJobs(index).Name, ZEN, ZenV
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
On Error GoTo ChangeImgJobName_Click_Error

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
    UpdateJobListbox ImgJobList, ImgJobs
    ImgJobList.Selected(index) = True

   On Error GoTo 0
   Exit Sub

ChangeImgJobName_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ChangeImgJobName_Click of Form JobSetter at line " & Erl & " "
End Sub


Private Sub ChangeFcsJobName_Click()
    Dim index As Integer
    Dim Name As String
On Error GoTo ChangeFcsJobName_Click_Error

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
    UpdateJobListbox FcsJobList, FcsJobs
    FcsJobList.Selected(index) = True

   On Error GoTo 0
   Exit Sub

ChangeFcsJobName_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ChangeFcsJobName_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub ImgJobList_Click()
    Dim index As Integer
On Error GoTo ImgJobList_Click_Error

    index = ImgJobList.ListIndex
    If index = -1 Then
        Exit Sub
    End If
    On Error Resume Next
    setLabels index
    setTrackNames index

   On Error GoTo 0
   Exit Sub

ImgJobList_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ImgJobList_Click of Form JobSetter at line " & Erl & " "
End Sub



Private Sub AddImgJobFromFileButton_Click()
    Dim FSO As New FileSystemObject
    Dim Filter As String, FileName As String
    Dim index As Integer
    Dim Flags As Long
    Dim DefDir As String
    Dim fileNames() As String
    
    '''get filename(s) to be loaded into ZEN'''
On Error GoTo AddImgJobFromFileButton_Click_Error

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
    FileName = CommonDialogAPI.ShowOpen(Filter, Flags, "", DefDir, "Select file(s) to be loaded as imaging jobs")
    
    If FileName = "" Then
        Exit Sub
    End If
    
    fileNames = Split(FileName, Chr$(0))
    
    EventMng.setBusy
    If UBound(fileNames) = 0 Then
        AddImgJobFromFile fileNames(0)
        WorkingDir = FSO.GetParentFolderName(fileNames(0)) & "\"
    Else
        QuickSort fileNames, 1, UBound(fileNames)
        For index = 1 To UBound(fileNames)
            AddImgJobFromFile fileNames(0) & "\" & fileNames(index)
            SleepWithEvents 2000
        Next index
        WorkingDir = fileNames(0) & "\"
    End If
    EventMng.setReady

   On Error GoTo 0
   Exit Sub

AddImgJobFromFileButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddImgJobFromFileButton_Click of Form JobSetter at line " & Erl & " "
End Sub


Private Sub AddImgJobFromFile(FileName As String)
    Dim FSO As New FileSystemObject
    Dim JobName As String
    Dim Recording As DsRecording
On Error GoTo AddImgJobFromFile_Error

    If Not FileExist(FileName) Then
        Exit Sub
    End If
    JobName = VBA.Split(FSO.GetFileName(FileName), ".")(0)
    If Not UniqueListName(FcsJobList, JobName) Or Not UniqueListName(ImgJobList, JobName) Then
        MsgBox "Name of imaging job must be unique!", VbExclamation, "JobSetter warning"
        Exit Sub
    End If
    Recording = getRecordingFromImageFile(FileName, ZEN)
    If Recording.MultiViewAcquisition Then
        MsgBox "You can't use Multiviews in MyPic"
        Exit Sub
    End If
    ImgJobList.AddItem JobName
    ImgJobList.Selected(ImgJobList.ListCount - 1) = True
    AddJob ImgJobs, ImgJobList.List(ImgJobList.ListCount - 1), ImgJobList.ListCount - 1, Recording, ZEN
    setLabels ImgJobList.ListCount - 1
    setTrackNames ImgJobList.ListCount - 1

   On Error GoTo 0
   Exit Sub

AddImgJobFromFile_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddImgJobFromFile of Form JobSetter at line " & Erl & " "
End Sub
    
Private Sub SaveButton_Click()
    
    Dim FSO As New FileSystemObject
    Dim index As Integer
    Dim Flags As Long
    Dim dirName As String, DefDir As String, Filter As String
    Dim answ As Integer
On Error GoTo SaveButton_Click_Error

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
    dirName = FSO.GetParentFolderName(dirName) & "\"
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

   On Error GoTo 0
   Exit Sub

SaveButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveButton_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub StopButton_Click()
On Error GoTo StopButton_Click_Error

    ScanStop = True
    StopAcquisition
    ScanStop = False

   On Error GoTo 0
   Exit Sub

StopButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StopButton_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub StopFcsButton_Click()
On Error GoTo StopFcsButton_Click_Error

    ScanStop = True
    StopAcquisition
    ScanStop = False

   On Error GoTo 0
   Exit Sub

StopFcsButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure StopFcsButton_Click of Form JobSetter at line " & Erl & " "
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
''' Change Status label text '''
On Error GoTo setStatus_Error

    If value Then
        StatusLabel = "READY"
        StatusLabel.ForeColor = &HC000&
    Else
        StatusLabel = "BUSY"
        StatusLabel.ForeColor = &HC0&
    End If

   On Error GoTo 0
   Exit Sub

setStatus_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure setStatus of Form JobSetter at line " & Erl & " "
End Sub



Private Sub TrackClick(iTrack As Integer)
''' Update status of Track number iTrack if it should be acquired or not '''
    Dim index As Integer
On Error GoTo TrackClick_Error

    index = ImgJobList.ListIndex
    If index <> -1 Then
        ImgJobs(index).setAcquireTrack iTrack - 1, Me.Controls("Track" + CStr(iTrack)).value
    End If

   On Error GoTo 0
   Exit Sub

TrackClick_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure TrackClick of Form JobSetter at line " & Erl & " "
End Sub


Private Sub TrackVisible(Visible As Boolean)
''' Change Visible Status of Track in GUI '''
On Error GoTo TrackVisible_Error

    Track1.Visible = Visible
    Track2.Visible = Visible
    Track3.Visible = Visible
    Track4.Visible = Visible

   On Error GoTo 0
   Exit Sub

TrackVisible_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure TrackVisible of Form JobSetter at line " & Erl & " "
End Sub

Private Sub SetJobButton_Click()
    ''' Read imaging Job from ZEN and import into macro. ZEN->Macro button'
    Dim index As Integer
On Error GoTo SetJobButton_Click_Error
    If Lsm5.DsRecording.MultiViewAcquisition Then
        MsgBox "You can't use MultiView with MyPic. Please deselect multiview and try again."
        Exit Sub
    End If
    index = ImgJobList.ListIndex
    If index = -1 Then
        MsgBox "Job list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Debug.Assert (ImgJobs(index).SetJob(Lsm5.DsRecording, ZEN))
    setLabels index
    setTrackNames index

   On Error GoTo 0
   Exit Sub

SetJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SetJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub SetFcsJob_Click()
    ''' Read FCS Job from ZEN and import into macro. ZEN->Macro button'
    Dim index As Integer
    Dim OpenForms() As Boolean
On Error GoTo SetFcsJob_Click_Error

    index = FcsJobList.ListIndex
    If index = -1 Then
         MsgBox "FcsJob list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    OpenForms = HideShowForms(OpenForms)
    Debug.Assert (FcsJobs(index).SetJob(ZEN, ZenV))
    setFcsLabels index
    HideShowForms OpenForms

   On Error GoTo 0
   Exit Sub

SetFcsJob_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SetFcsJob_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub setLabels(index As Integer)
    ''' Update description of imaging job number index in the GUI '''
    Dim jobDescription() As String
On Error GoTo setLabels_Error
    If UBound(ImgJobs) < index Then
        Exit Sub
    End If
    
    jobDescription = ImgJobs(index).splittedJobDescriptor(13, ImgJobs(index).jobDescriptor)
    JobLabel1.Caption = jobDescription(0)
    JobLabel2.Caption = jobDescription(1)
    
   On Error GoTo 0
   Exit Sub

setLabels_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure setLabels of Form JobSetter at line " & Erl & " "
End Sub

Private Sub setFcsLabels(index As Integer)
    ''' Update description of fcs job number index in the GUI '''
    Dim jobDescription() As String

    If UBound(FcsJobs) < index Then
        Exit Sub
    End If

On Error GoTo setFcsLabels_Error

    jobDescription = FcsJobs(index).splittedJobDescriptor(13, FcsJobs(index).jobDescriptor)
    FcsJobLabel1.Caption = jobDescription(0)
    FcsJobLabel2.Caption = jobDescription(1)
   On Error GoTo 0
   Exit Sub

setFcsLabels_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure setFcsLabels of Form JobSetter at line " & Erl & " "
End Sub

Private Sub PutJobButton_Click()
    ''' Upload imaging job from macro into ZEN. Button Macro -> ZEN '''
    Dim index As Integer
On Error GoTo PutJobButton_Click_Error

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

   On Error GoTo 0
   Exit Sub

PutJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure PutJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub


Private Sub PutFcsJob_Click()
    ''' Upload FCS job from macro into ZEN. Button Macro -> ZEN '''
    Dim index As Integer
On Error GoTo PutFcsJob_Click_Error

    index = FcsJobList.ListIndex
    If index = -1 Then
        MsgBox "FcsJob list is empty or you need to select one job", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    FcsJobs(index).PutJob ZEN, ZenV

   On Error GoTo 0
   Exit Sub

PutFcsJob_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure PutFcsJob_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub AddFcsJobButton_Click()
    ''' Add new FCS job to job list. + button '''
    Dim i As Integer
    Dim index As Integer
    Dim Name As String
    Dim OpenForms() As Boolean
On Error GoTo AddFcsJobButton_Click_Error

    Name = InputBox("Name of job to be created from current ZEN settings", "JobSetter: Define FcsJob name")
    'Cancel pressed
    If StrPtr(Name) = 0 Then Exit Sub
    
    If Name = "" Or Not UniqueListName(ImgJobList, CStr(Name)) Or Not UniqueListName(FcsJobList, CStr(Name)) Then
        MsgBox "You need to define an unique name for the imaging job!", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    
    OpenForms = HideShowForms(OpenForms)
    Name = CStr(Name)
    If FcsJobList.ListCount = 0 Then
        index = 0
    Else
        index = 0
        For i = 0 To UBound(FcsJobs)
            If Name > FcsJobs(i).Name Then
                index = index + 1
            End If
        Next i
    End If
    FcsJobList.AddItem Name, index
    FcsJobList.Selected(index) = True
    AddFcsJob FcsJobs, Name, index, ZEN
    setFcsLabels index
    HideShowForms OpenForms

   On Error GoTo 0
   Exit Sub

AddFcsJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddFcsJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Function UniqueListName(List As ListBox, JobName As String) As Boolean
    Dim ListEntry As Variant
On Error GoTo UniqueListName_Error

    Debug.Print List.ListCount
    If List.ListCount > 0 Then
        For Each ListEntry In List.List
            If StrComp(ListEntry, JobName) = 0 Then
                Exit Function
            End If
        Next
    End If
    UniqueListName = True

   On Error GoTo 0
   Exit Function

UniqueListName_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UniqueListName of Form JobSetter at line " & Erl & " "
End Function
    
'''
'  HideShowForms: Hide or show different forms stored in OpenForms
'''
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
    ") in procedure HideShowForms of Form JobSetter at line " & Erl & " "
End Function

'''
' AddJobButton_Click: Create a new imaging job from ZEN and add it to list of jobs
'''
Private Sub AddJobButton_Click()
   
    Dim i As Integer
    Dim index As Integer
    Dim Name As String
    Dim Names() As String
On Error GoTo AddJobButton_Click_Error
    If Lsm5.DsRecording.MultiViewAcquisition Then
        MsgBox "You can't use MultiView with MyPic. Please deselect multiview and try again."
        Exit Sub
    End If
    
    Name = InputBox("Name of imaging job to be created from current ZEN settings", "JobSetter: Define job name")
    'Cancel pressed
    If StrPtr(Name) = 0 Then Exit Sub

    If Name = "" Or Not UniqueListName(ImgJobList, CStr(Name)) Or Not UniqueListName(FcsJobList, CStr(Name)) Then
        MsgBox "You need to define an unique name for the imaging job!", VbExclamation, "JobSetter Warning"
        Exit Sub
    End If
    Name = CStr(Name)
    If ImgJobList.ListCount = 0 Then
        index = 0
    Else
        index = 0
        For i = 0 To UBound(ImgJobs)
            If Name > ImgJobs(i).Name Then
                index = index + 1
            End If
        Next i
    End If
    'Find position where to enter the job. It should be alphabetical
    ImgJobList.AddItem Name, index
    ImgJobList.Selected(index) = True
    AddJob ImgJobs, Name, index, Lsm5.DsRecording, ZEN
    setLabels index
    setTrackNames index

   On Error GoTo 0
   Exit Sub

AddJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub

'''
' setTrackNames(index As Integer)
' Set name of imaging track with number index in GUI. Use data stored in ImgJobs list
'''
Private Sub setTrackNames(index As Integer)
   
    Dim i As Integer
    Dim j As Integer
    Dim iTrack As Integer
    Dim Track As DsTrack
    Dim ChannelOK As Boolean
    Dim AcquireTrack() As Boolean
    Dim MaxTracks As Long
On Error GoTo setTrackNames_Error
    If UBound(ImgJobs) < index Then
        Exit Sub
    End If
    MaxTracks = ImgJobs(index).Recording.GetNormalTrackCount
    AcquireTrack = ImgJobs(index).AcquireTrack
    For i = 0 To 3
        If iTrack < 5 Then
            ChannelOK = False
            Set Track = ImgJobs(index).Recording.TrackObjectByMultiplexOrder(i, 1)
            If Track Is Nothing Then
               GoTo nextstep
            End If
            
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
nextstep:
    Next i

   On Error GoTo 0
   Exit Sub

setTrackNames_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure setTrackNames of Form JobSetter at line " & Erl & " "
End Sub

Private Sub UpdateJobListbox(List1 As ListBox, JobArray)
    Dim i As Integer
On Error GoTo UpdateJobListbox_Error

    If isArrayEmpty(JobArray) Then
        List1.Clear
        Exit Sub
    Else
        List1.Clear
        For i = 0 To UBound(JobArray)
            List1.AddItem JobArray(i).Name
        Next i
    End If

   On Error GoTo 0
   Exit Sub

UpdateJobListbox_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure UpdateJobListbox of Form JobSetter at line " & Erl & " "
End Sub


'''
' DeleteJobButton_Click()
'   Remove a Imaging job from Listbox and ImgJobs Array
'''
Private Sub DeleteJobButton_Click()
    Dim index As Integer
On Error GoTo DeleteJobButton_Click_Error

    index = ImgJobList.ListIndex
    If index <> -1 Then
        DeleteJob ImgJobs, index, ImgJobList.List(index)
        ImgJobList.RemoveItem index
    End If
    'PipelineConstructor.UpdateImgJobList

   On Error GoTo 0
   Exit Sub

DeleteJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure DeleteJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub

''
' Add imaging job Recording to JobsV with Name at index base 0
''
'---------------------------------------------------------------------------------------
' Procedure : AddJob
' Purpose   : Add imaging job to an array of Jobs
' Variables :
'   JobsV()  - An Array of Imaging jobs
'   Name     - Name of new job
'   index    - index of new job
'   Recording - The ZEN imaging settings
'   ZEN       - A ZEN object (this is for compabilities ZEN2010 and >= 2011
'---------------------------------------------------------------------------------------
'
Public Sub AddJob(JobsV() As AJob, Name As String, index As Integer, Recording As DsRecording, ZEN As Object)
    Dim i As Integer
    
On Error GoTo AddJob_Error

    If isArrayEmpty(JobsV) Then
        ReDim JobsV(0)
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) + 1)
    End If
    Debug.Assert (index <= UBound(JobsV))
    If index < UBound(JobsV) Then
        If UBound(JobsV) > 0 Then
            For i = UBound(JobsV) - 1 To index Step -1
                Set JobsV(i + 1) = JobsV(i)
            Next i
        End If
    End If
    Set JobsV(index) = New AJob
    JobsV(index).Name = Name
    JobsV(index).SetJob Recording, ZEN

   On Error GoTo 0
   Exit Sub

AddJob_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddJob of Form JobSetter at line " & Erl & " "
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddFcsJob
' Purpose   : Add FCS job to an array of Jobs
' Variables :
'   JobsV()  - An Array of FCS jobs
'   Name     - Name of new job
'   index    - index of new job
'   ZEN       - A ZEN object (this is for compabilities ZEN2010 and >= 2011
'---------------------------------------------------------------------------------------
'
Public Sub AddFcsJob(JobsV() As AFcsJob, Name As String, index As Integer, ZEN As Object)
    Dim i As Integer
On Error GoTo AddFcsJob_Error

    If isArrayEmpty(JobsV) Then
        ReDim JobsV(0)
    Else
        ReDim Preserve JobsV(0 To UBound(JobsV) + 1)
    End If
    Debug.Assert (index <= UBound(JobsV))
    If index < UBound(JobsV) Then
        If UBound(JobsV) > 0 Then
            For i = UBound(JobsV) - 1 To index Step -1
                Set JobsV(i + 1) = JobsV(i)
            Next i
        End If
    End If
    Set JobsV(index) = New AFcsJob
    JobsV(index).Name = Name
    JobsV(index).SetJob ZEN, ZenV

   On Error GoTo 0
   Exit Sub

AddFcsJob_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure AddFcsJob of Form JobSetter at line " & Erl & " "
End Sub



'''

'---------------------------------------------------------------------------------------
' Procedure : DeleteJob
' Purpose   : Delete Job and decrease number of Jobs. Job is found by index. Name
'             is used to double check the correct job
' Variables :
'   JobsV()  - An Array of imaging jobs
'   index    - index of job to delete
'   Name     - Name of job to delete
'---------------------------------------------------------------------------------------
'
Public Sub DeleteJob(JobsV() As AJob, index As Integer, Optional Name As String = "")
    Dim i As Integer
    Dim IJob As Integer
On Error GoTo DeleteJob_Error

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

   On Error GoTo 0
   Exit Sub

DeleteJob_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure DeleteJob of Form JobSetter at line " & Erl & " "
End Sub




Private Sub DeleteFcsJobButton_Click()
    Dim index As Integer
On Error GoTo DeleteFcsJobButton_Click_Error

    index = FcsJobList.ListIndex
    If index <> -1 Then
        DeleteFcsJob FcsJobs, index, FcsJobList.List(index)
        FcsJobList.RemoveItem index
    End If

   On Error GoTo 0
   Exit Sub

DeleteFcsJobButton_Click_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure DeleteFcsJobButton_Click of Form JobSetter at line " & Erl & " "
End Sub

Private Sub DeleteFcsJob(JobsV() As AFcsJob, index As Integer, Optional Name As String = "")
    Dim i As Integer
    Dim IJob As Integer
On Error GoTo DeleteFcsJob_Error

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

   On Error GoTo 0
   Exit Sub

DeleteFcsJob_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure DeleteFcsJob of Form JobSetter at line " & Erl & " "
End Sub

