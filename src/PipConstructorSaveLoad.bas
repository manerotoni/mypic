Attribute VB_Name = "PipConstructorSaveLoad"
''''
'  Module contains functions to save and load Form settings from file
'''

Option Explicit


''''
'   SaveSettings of PipelineConstructor in file name FileName.
''''
Public Sub SaveFormSettings(fileName As String)
    Dim iTsk As Integer, ipip As Integer, iSet As Integer
    Dim tskSettings As String
    Dim iFileNum As Long
    Dim arrTsk() As Variant
    Dim tskFieldNames() As String
    Dim tsk As Task
On Error GoTo SaveFormSettings_Error
    Close
    iFileNum = FreeFile()
    Open fileName For Output As iFileNum
    tskFieldNames = TaskFieldNames
    For ipip = 0 To UBound(Pipelines)
        With Pipelines(ipip)
            Print #iFileNum, "Pip " & ipip & " Reptime " & .Repetition.Time & " RepNr " & .Repetition.number & " RepInt " & .Repetition.interval
            
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
    Close #iFileNum
    Exit Sub
SaveFormSettings_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveFormSettings of Module AutofocusFormSaveLoad at line " & Erl & " "

End Sub





'''''
''   Save page of specific JobName using a file specified by iFuleNum
''   TODO: Control that indeed iFileNum is a file
'''''
'Private Sub SaveFormPage(JobName As String, iFileNum As Integer)
'    Dim i As Integer
'On Error GoTo SaveFormPage_Error
'
'    Print #iFileNum, ""
'    Print #iFileNum, "% " & JobName
'    Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").value
'
'    For i = 1 To 4
'        Print #iFileNum, JobName & "Track" & CInt(i) & " " & _
'        AutofocusForm.Controls(JobName & "Track" & CInt(i)).value
'    Next i
'
'
'    Print #iFileNum, JobName & "ZOffset " & AutofocusForm.Controls(JobName & "ZOffset").value
'    Print #iFileNum, JobName & "Period " & AutofocusForm.Controls(JobName & "Period").value
'    Print #iFileNum, JobName & "TrackZ " & AutofocusForm.Controls(JobName & "TrackZ").value
'    Print #iFileNum, JobName & "TrackXY " & AutofocusForm.Controls(JobName & "TrackXY").value
'    Print #iFileNum, JobName & "FocusMethod " & AutofocusForm.Controls(JobName & "FocusMethod").value
'    Print #iFileNum, JobName & "CenterOfMassChannel " & AutofocusForm.Controls(JobName & "CenterOfMassChannel").value
'    Print #iFileNum, JobName & "OiaActive " & AutofocusForm.Controls(JobName & "OiaActive").value
'    Print #iFileNum, JobName & "OiaSequential " & AutofocusForm.Controls(JobName & "OiaSequential").value
'    Print #iFileNum, JobName & "OiaParallel " & AutofocusForm.Controls(JobName & "OiaParallel").value
'    Print #iFileNum, JobName & "SaveImage " & AutofocusForm.Controls(JobName & "SaveImage").value
'    Print #iFileNum, JobName & "TimeOut " & AutofocusForm.Controls(JobName & "TimeOut").value
'
'    If JobName = "Trigger1" Or JobName = "Trigger2" Then
'        Print #iFileNum, JobName & "RepetitionSec " & AutofocusForm.Controls(JobName & "RepetitionSec").value
'        Print #iFileNum, JobName & "RepetitionMin " & AutofocusForm.Controls(JobName & "RepetitionMin").value
'        Print #iFileNum, JobName & "RepetitionTime " & AutofocusForm.Controls(JobName & "RepetitionTime").value
'        Print #iFileNum, JobName & "RepetitionInterval " & AutofocusForm.Controls(JobName & "RepetitionInterval").value
'        Print #iFileNum, JobName & "RepetitionNumber " & AutofocusForm.Controls(JobName & "RepetitionNumber").value
'        Print #iFileNum, JobName & "maxWait " & AutofocusForm.Controls(JobName & "maxWait").value
'        Print #iFileNum, JobName & "OptimalPtNumber " & AutofocusForm.Controls(JobName & "OptimalPtNumber").value
'        Print #iFileNum, JobName & "Autofocus " & AutofocusForm.Controls(JobName & "Autofocus").value
'        Print #iFileNum, JobName & "KeepParent " & AutofocusForm.Controls(JobName & "KeepParent").value
'    End If
'
'    Print #iFileNum, ""
'    Print #iFileNum, Jobs.jobDescriptorSettings(JobName)
'   On Error GoTo 0
'   Exit Sub
'
'SaveFormPage_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure SaveFormPage of Module AutofocusFormSaveLoad at line " & Erl & " "
'End Sub
'

'''''
''   Save page of specific JobFcs using a file specified by iFuleNum
''   TODO: Control that indeed iFileNum is a file
'''''
'Private Sub SaveFormFcsPage(JobName As String, iFileNum As Integer)
'On Error GoTo SaveFormFcsPage_Error
'
'    Print #iFileNum, ""
'    Print #iFileNum, "% " & JobName
'    Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").value
'    Print #iFileNum, JobName & "ZOffset " & AutofocusForm.Controls(JobName & "ZOffset").value
'    Print #iFileNum, JobName & "KeepParent " & AutofocusForm.Controls(JobName & "KeepParent").value
'    Print #iFileNum, JobName & "TimeOut " & AutofocusForm.Controls(JobName & "KeepParent").value
'    Print #iFileNum, ""
'    Print #iFileNum, JobsFcs.jobDescriptorSettings(JobName)
'    Exit Sub
'
'   On Error GoTo 0
'   Exit Sub
'
'SaveFormFcsPage_Error:
'
'    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
'    ") in procedure SaveFormFcsPage of Module AutofocusFormSaveLoad at line " & Erl & " "
'End Sub

''''
'   LoadSettings(FileName As String)
'   LoadSettings of Form from FileName
''''


