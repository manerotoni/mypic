Attribute VB_Name = "AutofocusFormSaveLoad"
''''
' Module contains functions to save and load Form settings from file
'''

Option Explicit


''''
'   SaveSettings of the UserForm AutofocusForm in file name FileName.
'   Name should correspond exactly to name used in Form
''''
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
    Print #iFileNum, "MultipleLocationToggle " & AutofocusForm.MultipleLocationToggle.value
    Print #iFileNum, "SingleLocationToggle " & AutofocusForm.SingleLocationToggle.value
    
    
    'Looping
    Print #iFileNum, "% GlobalRepetition "
    Print #iFileNum, "GlobalRepetitionSec " & AutofocusForm.GlobalRepetitionSec
    Print #iFileNum, "GlobalRepetitionMin " & AutofocusForm.GlobalRepetitionMin
    Print #iFileNum, "GlobalRepetitionTime " & AutofocusForm.GlobalRepetitionTime.value
    Print #iFileNum, "GlobalRepetitionInterval " & AutofocusForm.GlobalRepetitionInterval.value
    Print #iFileNum, "GlobalRepetitionNumber " & AutofocusForm.GlobalRepetitionNumber.value
    
    'Output
    Print #iFileNum, "% Output "
    Print #iFileNum, "DatabaseTextbox " & AutofocusForm.DatabaseTextbox.value
    Print #iFileNum, "TextBoxFileName " & AutofocusForm.TextBoxFileName.value
    
    'Grid Acquisition
    Print #iFileNum, "% Grid "
    Print #iFileNum, "GridScanActive " & AutofocusForm.GridScanActive.value
    Print #iFileNum, "GridScan_nRow " & AutofocusForm.GridScan_nRow.value
    Print #iFileNum, "GridScan_nColumn " & AutofocusForm.GridScan_nColumn.value
    Print #iFileNum, "GridScan_dRow " & AutofocusForm.GridScan_dRow.value
    Print #iFileNum, "GridScan_dColumn " & AutofocusForm.GridScan_dColumn.value
    Print #iFileNum, "GridScan_refRow " & AutofocusForm.GridScan_refRow.value
    Print #iFileNum, "GridScan_refColumn " & AutofocusForm.GridScan_refColumn.value
    Print #iFileNum, "GridScan_nRowsub " & AutofocusForm.GridScan_nRowsub.value
    Print #iFileNum, "GridScan_nColumnsub " & AutofocusForm.GridScan_nColumnsub.value
    Print #iFileNum, "GridScan_dRowsub " & AutofocusForm.GridScan_dRowsub.value
    Print #iFileNum, "GridScan_dColumnsub " & AutofocusForm.GridScan_dColumnsub.value
    Print #iFileNum, "GridScan_SubPositionsFirst " & AutofocusForm.GridScan_SubPositionsFirst.value
    Print #iFileNum, "GridScan_WellsFirst " & AutofocusForm.GridScan_WellsFirst.value
    Print #iFileNum, "GridCurrentZPosition " & AutofocusForm.GridCurrentZposition
    Print #iFileNum, "GridMarkedZPosition " & AutofocusForm.GridMarkedZPosition
    Print #iFileNum, "GridScanPositionFile " & AutofocusForm.GridScanPositionFile
    Print #iFileNum, "GridScanValidFile " & AutofocusForm.GridScanValidFile
    
    
    'Save water pump settings
    Print #iFileNum, "% Pump "
    Print #iFileNum, "Pump_interval_time " & PumpForm.Pump_interval_time
    Print #iFileNum, "Pump_interval_distance " & PumpForm.Pump_interval_distance.value
    Print #iFileNum, "Pump_time " & PumpForm.Pump_time
    Print #iFileNum, "Pump_wait " & PumpForm.Pump_wait.value
    Print #iFileNum, "Pump_interval_Jobs " & PumpForm.Pump_interval_Jobs
    
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
'   Save page of specific JobName using a file specified by iFuleNum
'   TODO: Control that indeed iFileNum is a file
''''
Private Sub SaveFormPage(JobName As String, iFileNum As Integer)
    Dim i As Integer
On Error GoTo SaveFormPage_Error

    Print #iFileNum, ""
    Print #iFileNum, "% " & JobName
    Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").value
    
    For i = 1 To 4
        Print #iFileNum, JobName & "Track" & CInt(i) & " " & _
        AutofocusForm.Controls(JobName & "Track" & CInt(i)).value
    Next i
    
    
    Print #iFileNum, JobName & "ZOffset " & AutofocusForm.Controls(JobName & "ZOffset").value
    Print #iFileNum, JobName & "Period " & AutofocusForm.Controls(JobName & "Period").value
    Print #iFileNum, JobName & "TrackZ " & AutofocusForm.Controls(JobName & "TrackZ").value
    Print #iFileNum, JobName & "TrackXY " & AutofocusForm.Controls(JobName & "TrackXY").value
    Print #iFileNum, JobName & "FocusMethod " & AutofocusForm.Controls(JobName & "FocusMethod").value
    Print #iFileNum, JobName & "CenterOfMassChannel " & AutofocusForm.Controls(JobName & "CenterOfMassChannel").value
    Print #iFileNum, JobName & "OiaActive " & AutofocusForm.Controls(JobName & "OiaActive").value
    Print #iFileNum, JobName & "OiaSequential " & AutofocusForm.Controls(JobName & "OiaSequential").value
    Print #iFileNum, JobName & "OiaParallel " & AutofocusForm.Controls(JobName & "OiaParallel").value
    Print #iFileNum, JobName & "SaveImage " & AutofocusForm.Controls(JobName & "SaveImage").value
    Print #iFileNum, JobName & "TimeOut " & AutofocusForm.Controls(JobName & "TimeOut").value
    
    If JobName = "Trigger1" Or JobName = "Trigger2" Then
        Print #iFileNum, JobName & "RepetitionSec " & AutofocusForm.Controls(JobName & "RepetitionSec").value
        Print #iFileNum, JobName & "RepetitionMin " & AutofocusForm.Controls(JobName & "RepetitionMin").value
        Print #iFileNum, JobName & "RepetitionTime " & AutofocusForm.Controls(JobName & "RepetitionTime").value
        Print #iFileNum, JobName & "RepetitionInterval " & AutofocusForm.Controls(JobName & "RepetitionInterval").value
        Print #iFileNum, JobName & "RepetitionNumber " & AutofocusForm.Controls(JobName & "RepetitionNumber").value
        Print #iFileNum, JobName & "maxWait " & AutofocusForm.Controls(JobName & "maxWait").value
        Print #iFileNum, JobName & "OptimalPtNumber " & AutofocusForm.Controls(JobName & "OptimalPtNumber").value
        Print #iFileNum, JobName & "Autofocus " & AutofocusForm.Controls(JobName & "Autofocus").value
        Print #iFileNum, JobName & "KeepParent " & AutofocusForm.Controls(JobName & "KeepParent").value
    End If
    
    Print #iFileNum, ""
    Print #iFileNum, Jobs.jobDescriptorSettings(JobName)
   On Error GoTo 0
   Exit Sub

SaveFormPage_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveFormPage of Module AutofocusFormSaveLoad at line " & Erl & " "
End Sub


''''
'   Save page of specific JobFcs using a file specified by iFuleNum
'   TODO: Control that indeed iFileNum is a file
''''
Private Sub SaveFormFcsPage(JobName As String, iFileNum As Integer)
On Error GoTo SaveFormFcsPage_Error

    Print #iFileNum, ""
    Print #iFileNum, "% " & JobName
    Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").value
    Print #iFileNum, JobName & "ZOffset " & AutofocusForm.Controls(JobName & "ZOffset").value
    Print #iFileNum, JobName & "KeepParent " & AutofocusForm.Controls(JobName & "KeepParent").value
    Print #iFileNum, JobName & "TimeOut " & AutofocusForm.Controls(JobName & "KeepParent").value
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
                If Left(FieldEntries(0), 4) = "Pump" Then
                    On Error Resume Next
                    PumpForm.Controls(FieldEntries(0)).value = FieldEntries(1)
                ElseIf Left(FieldEntries(0), 6) <> "EndJob" Then
                    On Error Resume Next
                    AutofocusForm.Controls(FieldEntries(0)).value = FieldEntries(1)
                End If
                If Err Then
                    LogManager.UpdateErrorLog "Warning " & FieldEntries(0) & " does not exist as parameter"
                    On Error GoTo 0
                End If
            End If
NextLine:
    Loop
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & FileName & " for AutofocusScreen settings"
End Sub


