VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FcsJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
' A Class to manage several AFcsJob instances(commodity class in absence of inheritance in VBA)
'''
Option Explicit

'''' We can use number or a dictionary to find a Job in Jobs
Private Jobs() As AFcsJob
Private JobNames As Dictionary
Private ZoomPixelSlice(1 To 9, 1 To 3) As Double
Public ZENv As Integer


Private Function checkJobName(Name As String) As Boolean
    If JobNames.Exists(Name) Then
        checkJobName = True
    Else
        MsgNoJob Name
    End If
End Function

Private Sub MsgNoJob(Name As String)
    MsgBox "ImagingJobs: Job " & Name & " has not been defined"
End Sub

'''
'   Initialize a series of Jobs with given name
'''
Public Sub initialize(Names() As String, ZEN As Object)
    Dim i As Integer
    ZENv = getVersionNr
    For i = 0 To UBound(Names)
        AddJob CStr(Names(i)), ZEN
    Next i
End Sub

'''
' AddJob
'   Add Job and increase number of Jobs
'''
Public Sub AddJob(Name As String, ZEN As Object)
 On Error GoTo ErrorHandle:
    If JobNames Is Nothing Then
        Set JobNames = New Dictionary
    End If
    If Not JobNames.Exists(Name) Then
        If isArrayEmpty(Jobs) Then
            ReDim Jobs(0)
        Else
            ReDim Preserve Jobs(0 To UBound(Jobs) + 1)
        End If
        JobNames.Add Name, UBound(Jobs)
    End If
    Dim i As Integer
    Dim iJob As Integer
    iJob = JobNames.Item(Name)
    Set Jobs(iJob) = New AFcsJob
    Jobs(iJob).Name = Name
    'Jobs(iJob).setJob ZEN, ZENv
    Exit Sub
ErrorHandle:
    ErrorLog.UpdateLog "Error in FcsJobs.AddJob " + Name + " " + Err.Description
End Sub



'''
'   Set the job with Name from ZEN
'''
Public Function setJob(Name As String, ZEN As Object) As Boolean
    On Error GoTo ErrorHandle:
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If Jobs(JobNames.Item(Name)).setJob(ZEN, ZENv) Then
        setJob = True
    End If
    Exit Function
ErrorHandle:
    ErrorLog.UpdateLog "Error in FcsJobs.setJob " + Name + " " + Err.Description
End Function

'''
'   Set the job with Name from ZEN
'''
Public Function setJobNoAi(Name As String, ConfigName As String) As Boolean
    On Error GoTo ErrorHandle:
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If Jobs(JobNames.Item(Name)).setJobNoAi(ConfigName) Then
        setJobNoAi = True
    End If
    Exit Function
ErrorHandle:
    ErrorLog.UpdateLog "Error in FcsJobs.setJobNoAi " + Name + " " + Err.Description
End Function

'''
'   Put the job with Name from Macro into ZEN
'''
Public Function putJob(Name As String, ZEN As Object) As Boolean
    On Error GoTo ErrorHandle:
    Dim iJob As Integer
    If Not checkJobName(Name) Then
        Exit Function
    End If
    If Jobs(JobNames.Item(Name)).putJob(ZEN, ZENv) Then
        putJob = True
    End If
    Exit Function
ErrorHandle:
    MsgBox "Error in FcsJobs.putJob " + Name + " " + Err.Description
 End Function

''''
' Create a string with the description of a Job
''''
Public Function jobDescriptor(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    jobDescriptor = Jobs(JobNames.Item(Name)).jobDescriptor
End Function


''''
' Create a string with the description of a Job
''''
Public Function splittedJobDescriptor(Name As String, MaxLines As Integer) As String()
    If Not checkJobName(Name) Then
        Exit Function
    End If
    splittedJobDescriptor = Jobs(JobNames.Item(Name)).splittedJobDescriptor(MaxLines, jobDescriptor(Name))
End Function

Public Function getLightPathConfig(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    getLightPathConfig = Jobs(JobNames.Item(Name)).LightPathConfig
End Function

''''
' Create a string with the description of a Job to be used as settings
''''
Public Function jobDescriptorSettings(Name As String) As String
    If Not checkJobName(Name) Then
        Exit Function
    End If
    jobDescriptorSettings = Jobs(JobNames.Item(Name)).jobDescriptorSettings
End Function

Public Function changeJobFromDescriptor(Name As String, Descriptor As String, Value As Variant)
    If Not checkJobName(Name) Then
        Exit Function
    End If
    changeJobFromDescriptor = Jobs(JobNames.Item(Name)).changeJobFromDescriptor(Descriptor, Value)
End Function

