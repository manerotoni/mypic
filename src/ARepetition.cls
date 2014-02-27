VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARepetition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
' A Class to manage repetitions
'''
Option Explicit
Public Name As String            ' this is should belong to the upper class but to see what is inside you need the name
Public number As Integer        'Number of repetitions
Public index As Integer          'currentindex of repetition
Public Time As Double         ' Time between repetition
Private TimeStart As Double     'Time start of current repetition this will be updated if Interval is used
Public interval As Boolean
Const maxZeros = 4

'''
' Update index of Repetition and store timeStart
'''
Public Function nextRep() As Boolean
    If index < number Then
        index = index + 1
        TimeStart = CDbl(GetTickCount) * 0.001
        nextRep = True
    End If
End Function

Public Sub updateTimeStart()
    If Not interval Then
        TimeStart = CDbl(GetTickCount) * 0.001
    End If
End Sub

Public Function thisSuffix() As String
    Dim nrZero As Integer
    nrZero = maxZeros - Len(CStr(index))
    thisSuffix = FNSep + "T" + ZeroString(nrZero) + CStr(index)
End Function


Public Function suffix(value As Integer) As String
    Dim nrZero As Integer
    nrZero = maxZeros - Len(CStr(value))
    suffix = FNSep + "T" + ZeroString(nrZero) + CStr(value)
End Function


'''
'   Calculate that remains to wait
'''
Public Function wait() As Double
    wait = Time + TimeStart - CDbl(GetTickCount) * 0.001
End Function

'''
' Constructor initialize the class
'''
Public Function initialize(repetitionName As String, RepetitionTime As Double, RepetitionNumber As Integer, RepetitionInterval As Boolean)
    index = 0
    number = RepetitionNumber
    Name = repetitionName
    Time = RepetitionTime
    interval = RepetitionInterval
End Function

'''''''''''
''''TYPES''
'''''''''''
'Public Type RepetitionType     'Contains Number and Time interval for repetition of acquisition protocol
'    Number As Integer   'Number of repetitions
'    Time As Double      'Interval between repetitions
'    Interval As Boolean ' If Interval is True than one computes interval between first and second image = Time other wise
'End Type
'
