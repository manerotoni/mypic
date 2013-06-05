VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagingRepetitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Repetitions As New Collection
Private pZENv As Integer



'''
'   Initialize a series of RepetitionJobs with given name
'''
Public Sub initialize(Names() As String, RepetitionTimes() As Double, RepetitionNumbers() As Integer, RepetitionIntervals() As Boolean)

    Dim Name As Variant
    Dim i As Integer
    Dim MaxNames As Integer
        
    If isArrayEmpty(Names) Then
        MsgBox "initialize ImagingRepetitions class: provide a non-empty array"
        Exit Sub
    End If
    Set Repetitions = New Collection
    MaxNames = UBound(Names)
    If UBound(RepetitionTimes) <> MaxName Or UBound(RepetitionNumbers) <> MaxName Or _
        UBound(RepetitionInterval) <> MaxName Then
        MsgBox "initialize ImagingRepetitions class: all arrays should have the same length"
        Exit Sub
    End If
    
    For i = 0 To MaxNames
        AddRepetition Names(i), RepetitionTimes(i), RepetitionNumbers(i), RepetitionIntervals(i)
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
Public Function reset(Name As String, RepetitionTime As Double, RepetitionNumber As Integer, RepetitionInterval As Boolean)
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

Public Sub setInterval(Name As String, Value As Boolean)
    Repetitions(Name).interval = Value
End Sub

'''
'
'''
Public Sub setRepetitionTime(Name As String, Value As Double)
    Repetitions(Name).time = Value
End Sub

'''
'
'''
Public Sub setRepetitionNumber(Name As String, Value As Integer)
    Repetitions(Name).number = Value
End Sub


Public Function thisSuffix(Name As String) As String
    thisSuffix = Repetitions(Name).thisSuffix
End Function

Public Function suffix(Name As String, Value As Integer) As String
    suffix = Repetitions(Name).suffix(Value)
End Function

