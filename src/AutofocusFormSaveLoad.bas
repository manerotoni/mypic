Attribute VB_Name = "AutofocusFormSaveLoad"
''''
' Module contains functions to save and load Form settings from file
'''

Option Explicit




'''''
'   SaveSettings(FileName As String)
'   SaveSettings of the UserForm AutofocusForm in file name FileName.
'   Name should correspond exactly to name used in Form
'''''
Public Sub SaveFormSettings(FileName As String)
    Dim i As Integer
    Dim iFileNum As Integer
On Error GoTo SaveFormSettings_Error

    Close
    iFileNum = FreeFile()
    Open FileName For Output As iFileNum
    
    Print #iFileNum, "% Settings for AutofocusMacro for ZEN " & ZENv & "  " & AutofocusForm.Version

    'Single MultipelocationToggle
    Print #iFileNum, "% Single Multiple "
    Print #iFileNum, "MultipleLocationToggle " & AutofocusForm.MultipleLocationToggle.Value
    Print #iFileNum, "SingleLocationToggle " & AutofocusForm.SingleLocationToggle.Value
    
    
    'Looping
    Print #iFileNum, "% GlobalRepetition "
    Print #iFileNum, "GlobalRepetitionSec " & AutofocusForm.GlobalRepetitionSec
    Print #iFileNum, "GlobalRepetitionMin " & AutofocusForm.GlobalRepetitionMin
    Print #iFileNum, "GlobalRepetitionTime " & AutofocusForm.GlobalRepetitionTime.Value
    Print #iFileNum, "GlobalRepetitionInterval " & AutofocusForm.GlobalRepetitionInterval.Value
    Print #iFileNum, "GlobalRepetitionNumber " & AutofocusForm.GlobalRepetitionNumber.Value
    
    'Output
    Print #iFileNum, "% Output "
    Print #iFileNum, "DatabaseTextbox " & AutofocusForm.DatabaseTextbox.Value
    Print #iFileNum, "TextBoxFileName " & AutofocusForm.TextBoxFileName.Value
    
    'Grid Acquisition
    Print #iFileNum, "% Grid "
    Print #iFileNum, "GridScanActive " & AutofocusForm.GridScanActive.Value
    Print #iFileNum, "GridScan_validGridDefault " & AutofocusForm.GridScan_validGridDefault.Value
    Print #iFileNum, "GridScan_nRow " & AutofocusForm.GridScan_nRow.Value
    Print #iFileNum, "GridScan_nColumn " & AutofocusForm.GridScan_nColumn.Value
    Print #iFileNum, "GridScan_dRow " & AutofocusForm.GridScan_dRow.Value
    Print #iFileNum, "GridScan_dColumn " & AutofocusForm.GridScan_dColumn.Value
    Print #iFileNum, "GridScan_refRow " & AutofocusForm.GridScan_refRow.Value
    Print #iFileNum, "GridScan_refColumn " & AutofocusForm.GridScan_refColumn.Value
    Print #iFileNum, "GridScan_nRowsub " & AutofocusForm.GridScan_nRowsub.Value
    Print #iFileNum, "GridScan_nColumnsub " & AutofocusForm.GridScan_nColumnsub.Value
    Print #iFileNum, "GridScan_dRowsub " & AutofocusForm.GridScan_dRowsub.Value
    Print #iFileNum, "GridScan_dColumnsub " & AutofocusForm.GridScan_dColumnsub.Value
    
    'Save settings of all pages
    For i = 0 To UBound(JobNames)
        SaveFormPage JobNames(i), iFileNum
    Next i
    
    'Save settings of all pages
    For i = 0 To UBound(JobFcsNames)
        SaveFormFcsPage JobFcsNames(i), iFileNum
    Next i
    Close #iFileNum

   On Error GoTo 0
   Exit Sub

SaveFormSettings_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveFormSettings of Module AutofocusFormSaveLoad at line " & Erl & " "

End Sub


''''
' SavePage(JobName As String, iFileNum As Integer)
'   Save page of specific JobName using a file specified by iFuleNum
'   TODO: Control that indeed iFileNum is a file
''''
Private Sub SaveFormPage(JobName As String, iFileNum As Integer)
    Dim i As Integer
On Error GoTo SaveFormPage_Error

    Print #iFileNum, ""
    Print #iFileNum, "% " & JobName
    Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").Value
    
    For i = 1 To 4
        Print #iFileNum, JobName & "Track" & CInt(i) & " " & _
        AutofocusForm.Controls(JobName & "Track" & CInt(i)).Value
    Next i
    
    
    Print #iFileNum, JobName & "ZOffset " & AutofocusForm.Controls(JobName & "ZOffset").Value
    Print #iFileNum, JobName & "Period " & AutofocusForm.Controls(JobName & "Period").Value
    Print #iFileNum, JobName & "TrackZ " & AutofocusForm.Controls(JobName & "TrackZ").Value
    Print #iFileNum, JobName & "TrackXY " & AutofocusForm.Controls(JobName & "TrackXY").Value
    Print #iFileNum, JobName & "CenterOfMass " & AutofocusForm.Controls(JobName & "CenterOfMass").Value
    Print #iFileNum, JobName & "CenterOfMassChannel " & AutofocusForm.Controls(JobName & "CenterOfMassChannel").Value
    Print #iFileNum, JobName & "OiaActive " & AutofocusForm.Controls(JobName & "OiaActive").Value
    Print #iFileNum, JobName & "OiaSequential " & AutofocusForm.Controls(JobName & "OiaSequential").Value
    Print #iFileNum, JobName & "OiaParallel " & AutofocusForm.Controls(JobName & "OiaParallel").Value
    Print #iFileNum, JobName & "SaveImage " & AutofocusForm.Controls(JobName & "SaveImage").Value
    
    If JobName = "Trigger1" Or JobName = "Trigger2" Then
        Print #iFileNum, JobName & "RepetitionTime " & AutofocusForm.Controls(JobName & "RepetitionTime").Value
        Print #iFileNum, JobName & "RepetitionSec " & AutofocusForm.Controls(JobName & "RepetitionSec").Value
        Print #iFileNum, JobName & "RepetitionMin " & AutofocusForm.Controls(JobName & "RepetitionMin").Value
        Print #iFileNum, JobName & "RepetitionInterval " & AutofocusForm.Controls(JobName & "RepetitionInterval").Value
        Print #iFileNum, JobName & "RepetitionNumber " & AutofocusForm.Controls(JobName & "RepetitionNumber").Value
        Print #iFileNum, JobName & "maxWait " & AutofocusForm.Controls(JobName & "maxWait").Value
        Print #iFileNum, JobName & "OptimalPtNumber " & AutofocusForm.Controls(JobName & "OptimalPtNumber").Value
        Print #iFileNum, JobName & "Autofocus " & AutofocusForm.Controls(JobName & "Autofocus").Value
        Print #iFileNum, JobName & "KeepParent " & AutofocusForm.Controls(JobName & "KeepParent").Value
    End If
    
    Print #iFileNum, ""
    Print #iFileNum, Jobs.jobDescriptorSettings(JobName)
   On Error GoTo 0
   Exit Sub

SaveFormPage_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveFormPage of Module AutofocusFormSaveLoad at line " & Erl & " "
End Sub

Public Function ControlTipText()
    Dim i As Integer
    For i = 0 To UBound(JobNames)
        JobControlTipText JobNames(i)
    Next i
End Function

'''
' Sets tip text for all pages
'''
Private Sub JobControlTipText(JobName As String)
    On Error GoTo ErrorHandle:

    AutofocusForm.Controls(JobName + "Period").ControlTipText = "Perform job " & JobName & " every xx repetitions"


    AutofocusForm.Controls(JobName + "ZOffset").ControlTipText = "Add xx to Z from previous imaging Job"
    AutofocusForm.Controls(JobName + "TrackZ").ControlTipText = "Update Z of current point with computed position"
    AutofocusForm.Controls(JobName + "TrackXY").ControlTipText = "Update XY of current point with computed position"
    AutofocusForm.Controls(JobName + "CenterOfMass").ControlTipText = "Compute new position from center of mass (done within Macro)"
    AutofocusForm.Controls(JobName + "OiaActive").ControlTipText = "If active macro listens to online image analysis"
    AutofocusForm.Controls(JobName + "OiaSequential").ControlTipText = "Macro waits for image analysis to finish. Acquire image -> OnlineImage analysis -> perform task"
    AutofocusForm.Controls(JobName + "OiaParallel").ControlTipText = "Imaging and analysis run in parallel."
    
    If JobName = "Trigger1" Or JobName = "Trigger2" Then
        AutofocusForm.Controls(JobName + "Active").ControlTipText = "Job " & JobName & " is performed only after online image analysis command"
        AutofocusForm.Controls(JobName + "OptimalPtNumber").ControlTipText = "Wait to find up to xxx positions before starting job " & JobName
        AutofocusForm.Controls(JobName + "maxWait").ControlTipText = "Wait up to xxx seconds before starting job " & JobName
        AutofocusForm.Controls(JobName + "Autofocus").ControlTipText = "Before acquiring " & JobName & " perform Job Autofocus"
        AutofocusForm.Controls(JobName + "KeepParent").ControlTipText = "If on revisit parent position from which " & JobName & " has been triggered"
    End If

    AutofocusForm.Controls(JobName + "PutJob").ControlTipText = "Put Macro acquisition settings into ZEN. Not all settings are shown in the  ZEN GUI!"
    AutofocusForm.Controls(JobName + "SetJob").ControlTipText = "Load settings from ZEN into Macro. Not all settings are shown in the  Macro GUI!"
    AutofocusForm.Controls(JobName + "Acquire").ControlTipText = "Acquire one image with settings of Job " & JobName
    Exit Sub
ErrorHandle:
    MsgBox "Error in JobControlTipText " + JobName + " " + Err.Description
End Sub

''''
'   Save page of specific JobFcs using a file specified by iFuleNum
'   TODO: Control that indeed iFileNum is a file
''''
Private Sub SaveFormFcsPage(JobName As String, iFileNum As Integer)
On Error GoTo SaveFormFcsPage_Error

    Print #iFileNum, ""
    Print #iFileNum, "% " & JobName
    Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").Value
    Print #iFileNum, JobName & "KeepParent " & AutofocusForm.Controls(JobName & "KeepParent").Value
    
    Print #iFileNum, ""
    Print #iFileNum, JobsFcs.jobDescriptorSettings(JobName)
    Exit Sub

   On Error GoTo 0
   Exit Sub

SaveFormFcsPage_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveFormFcsPage of Module AutofocusFormSaveLoad at line " & Erl & " "
End Sub

''''
'   LoadSettings(FileName As String)
'   LoadSettings of Form from FileName
''''
Public Sub LoadFormSettings(FileName As String)
    Dim iFileNum As Integer
    Dim Fields As String
    Dim JobName As String
    Dim FieldEntries() As String
    Close
    On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open FileName For Input As iFileNum
    Do While Not EOF(iFileNum)
            Line Input #iFileNum, Fields
            While Left(Fields, 1) = "%"
                Line Input #iFileNum, Fields
            Wend
            
            If Fields <> "" Then
                FieldEntries = Split(Fields, " ", 2)
                If FieldEntries(0) = "JobName" Then
                    JobName = FieldEntries(1)
                    Line Input #iFileNum, Fields
                    FieldEntries = Split(Fields, " ", 2)
                    While FieldEntries(0) <> "EndJobDef"
                        Jobs.changeJobFromDescriptor JobName, FieldEntries(0), FieldEntries(1)
                        Line Input #iFileNum, Fields
                        FieldEntries = Split(Fields, " ", 2)
                    Wend
                    'put once the job and reload it to get all the proper pixelSize according to the zoom etc
                    Jobs.putJob JobName, ZEN, True
                    Application.ThrowEvent eEventDataChanged, 0
                    Jobs.setJob JobName, Lsm5.DsRecording, ZEN
                    UpdateFormFromJob Jobs, JobName
                    UpdateJobFromForm Jobs, JobName
                End If
                If FieldEntries(0) = "JobFcsName" Then
                    JobName = FieldEntries(1)
                    Line Input #iFileNum, Fields
                    FieldEntries = Split(Fields, " ", 2)
                    While FieldEntries(0) <> "EndJobFcsDef"
                        JobsFcs.changeJobFromDescriptor JobName, FieldEntries(0), FieldEntries(1)
                        Line Input #iFileNum, Fields
                        FieldEntries = Split(Fields, " ", 2)
                    Wend
                    If JobsFcs.getLightPathConfig(JobName) <> "" Then
                        'put once the job and reload it to get all the proper pixelSize according to the zoom etc
                        JobsFcs.putJob JobName, ZEN
                        JobsFcs.setJobNoAi JobName, JobsFcs.getLightPathConfig(JobName)
                        'JobsFcs.setJob JobName, ZEN
                    End If
                        UpdateFormFromJobFcs JobsFcs, JobName
                        'UpdateJobFromForm Jobs, JobName
                End If
                On Error Resume Next
                AutofocusForm.Controls(FieldEntries(0)).Value = FieldEntries(1)
            End If
NextLine:
    Loop
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & FileName & " for AutofocusScreen settings"
End Sub


