VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagingRepetitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' class to manage several ARepetition
'''
Option Explicit
Private Repetitions As New Collection 'ARepetition to be changed
Private pZENv As Integer
Private RepetitionNames As Dictionary


Private Sub MsgNoRepetition(Name As String)
    MsgBox "ImagingRepetitions: Repetition " & Name & " has not been defined"
End Sub

Public Function checkRepetitionNames(Name As String) As Boolean
    If RepetitionNames.Exists(Name) Then
        checkRepetitionNames = True
    Else
        MsgNoRepetition Name
    End If
End Function


'''
'   Initialize a series of RepetitionJobs with given name
'''
Public Sub initialize(names() As String, RepetitionTimes() As Double, RepetitionNumbers() As Integer, RepetitionIntervals() As Boolean)

    Dim Name As Variant
    Dim i As Integer
    Dim MaxNames As Integer
        
    If isArrayEmpty(names) Then
        MsgBox "initialize ImagingRepetitions class: provide a non-empty array"
        Exit Sub
    End If
    Set Repetitions = New Collection
    MaxNames = UBound(names)
    If UBound(RepetitionTimes) <> MaxNames Or UBound(RepetitionNumbers) <> MaxNames Or _
        UBound(RepetitionIntervals) <> MaxNames Then
        MsgBox "initialize ImagingRepetitions class: all arrays should have the same length"
        Exit Sub
    End If
    
    For i = 0 To MaxNames
        AddRepetition names(i), RepetitionTimes(i), RepetitionNumbers(i), RepetitionIntervals(i)
    Next i

End Sub


'''
'   Initialize a series of RepetitionJobs with given name
'''
Public Sub AddRepetition(Name As String, RepetitionTime As Double, RepetitionNumber As Integer, RepetitionInterval As Boolean)
    Dim Rep As ARepetition
    Dim i As Integer
    Set Rep = New ARepetition
    
    Rep.initialize Name, RepetitionTime, RepetitionNumber, RepetitionInterval
    Repetitions.Add Rep, Name
    
End Sub

'''
'  Get current index of repetition
'''
Public Function getIndex(Name As String) As Integer
    getIndex = Repetitions(Name).index
End Function

'''
'   Reset everything to new values
'''
Public Function Reset(Name As String, RepetitionTime As Double, RepetitionNumber As Integer, RepetitionInterval As Boolean)
    Repetitions(Name).index = 0
    Repetitions(Name).update RepetitionTime, RepetitionNumber, RepetitionInterval
End Function

'''
'   Just reset the index
'''
Public Function resetIndex(Name As String)
    Repetitions(Name).index = 0
End Function

'''
'   Update setting of Job
'''
Public Function update(Name As String, RepetitionTime As Double, RepetitionNumber As Integer, RepetitionInterval As Boolean)
    ' this will not upate the index
    Reps.update Name, RepetitionTime, RepetitionNumber, RepetitionInterval
End Function

'''
'   Switch to new Repetition
'''
Public Function nextRep(Name As String) As Boolean
    nextRep = Repetitions(Name).nextRep
End Function

'''
'   Compute waiting time
'''
Public Function wait(Name As String) As Double
    wait = Repetitions(Name).wait
End Function

'''
'   Check if new timeStart
'''
Public Sub updateTimeStart(Name As String)
    Repetitions(Name).updateTimeStart
End Sub

Public Sub setInterval(Name As String, value As Boolean)
    Repetitions(Name).interval = value
End Sub

Public Function thisIndex(Name As String) As Integer
    thisIndex = Repetitions(Name).index
End Function


'''
'
'''
Public Sub setRepetitionTime(Name As String, value As Double)
    Repetitions(Name).Time = value
End Sub

'''
'set the number of repetitions
'''
Public Sub setRepetitionNumber(Name As String, value As Integer)
    Repetitions(Name).number = value
End Sub

'''
'get the number of repetitions
'''
Public Function getRepetitionNumber(Name As String) As Integer
    getRepetitionNumber = Repetitions(Name).number
End Function


Public Function thisSuffix(Name As String) As String
    thisSuffix = Repetitions(Name).thisSuffix
End Function

Public Function suffix(Name As String, value As Integer) As String
    suffix = Repetitions(Name).suffix(value)
End Function

