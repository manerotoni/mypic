Attribute VB_Name = "AutofocusFormSaveLoad"
''''
' Module contains functions to save and load Form settings
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
    Close
    On Error GoTo ErrorHandle
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


    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "SaveFormSettings: Not able to open " & FileName & " for saving settings"
End Sub


''''
' SavePage(JobName As String, iFileNum As Integer)
'   Save page of specific JobName using a file specified by iFuleNum
'   TODO: Control that indeed iFileNum is a file
''''
Private Sub SaveFormPage(JobName As String, iFileNum As Integer)
    Dim i As Integer
    On Error GoTo ErrorHandle:
    Print #iFileNum, ""
    Print #iFileNum, "% " & JobName
    Print #iFileNum, JobName & "Period " & AutofocusForm.Controls(JobName & "Period").Value
    If JobName <> "Trigger1" And JobName <> "Trigger2" Then
        Print #iFileNum, JobName & "Active " & AutofocusForm.Controls(JobName & "Active").Value
    End If
    
    For i = 1 To 4
        Print #iFileNum, JobName & "Track" & CInt(i) & " " & _
        AutofocusForm.Controls(JobName & "Track" & CInt(i)).Value
    Next i
    
    If JobName <> "Autofocus" Then
        Print #iFileNum, JobName & "ZOffset " & AutofocusForm.Controls(JobName & "ZOffset").Value
    End If
    
    Print #iFileNum, JobName & "Period " & AutofocusForm.Controls(JobName & "Period").Value
    Print #iFileNum, JobName & "TrackZ " & AutofocusForm.Controls(JobName & "TrackZ").Value
    Print #iFileNum, JobName & "TrackXY " & AutofocusForm.Controls(JobName & "TrackXY").Value
    Print #iFileNum, JobName & "CenterOfMass " & AutofocusForm.Controls(JobName & "CenterOfMass").Value
    Print #iFileNum, JobName & "CenterOfMassChannel " & AutofocusForm.Controls(JobName & "CenterOfMassChannel").Value
    Print #iFileNum, JobName & "OiaActive " & AutofocusForm.Controls(JobName & "OiaActive").Value
    Print #iFileNum, JobName & "OiaSequential " & AutofocusForm.Controls(JobName & "OiaSequential").Value
    Print #iFileNum, JobName & "OiaParallel " & AutofocusForm.Controls(JobName & "OiaParallel").Value
    
    If JobName = "Trigger1" Or JobName = "Trigger2" Then
        Print #iFileNum, JobName & "RepetitionTime " & AutofocusForm.Controls(JobName & "RepetitionTime").Value
        Print #iFileNum, JobName & "RepetitionSec " & AutofocusForm.Controls(JobName & "RepetitionSec").Value
        Print #iFileNum, JobName & "RepetitionMin " & AutofocusForm.Controls(JobName & "RepetitionMin").Value
        Print #iFileNum, JobName & "RepetitionInterval " & AutofocusForm.Controls(JobName & "RepetitionInterval").Value
        Print #iFileNum, JobName & "RepetitionNumber " & AutofocusForm.Controls(JobName & "RepetitionNumber").Value
        Print #iFileNum, JobName & "maxWait " & AutofocusForm.Controls(JobName & "maxWait").Value
        Print #iFileNum, JobName & "OptimalPtNumber " & AutofocusForm.Controls(JobName & "OptimalPtNumber").Value
        Print #iFileNum, JobName & "Autofocus " & AutofocusForm.Controls(JobName & "Autofocus").Value
    End If
    
    Print #iFileNum, ""
    Print #iFileNum, Jobs.jobDescriptorSettings(JobName)
    Exit Sub
ErrorHandle:
    MsgBox "Error in SaveFormPage " + JobName + " " + Err.Description
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
    Dim Entries() As String
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
                    UpdateFormFromJob Jobs, JobName
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


