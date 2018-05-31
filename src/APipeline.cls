VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APipeline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private TaskArray() As Task
Public Repetition As ARepetition
Public Grid As AGrid
Private pKeepParent As Boolean
Private pOptPtNumber As Integer
Private pMaxWait As Double

Public Property Get keepParent() As Boolean
    keepParent = pKeepParent
End Property

Public Property Let keepParent(value As Boolean)
    pKeepParent = value
End Property

Public Property Get optPtNumber() As Integer
    optPtNumber = pOptPtNumber
End Property

Public Property Let optPtNumber(value As Integer)
    pOptPtNumber = value
End Property

Public Property Get maxWait() As Double
    maxWait = pMaxWait
End Property

Public Property Let maxWait(value As Double)
    pMaxWait = value
End Property



Public Sub addTask(TaskIn As Task)
    On Error GoTo emptyArray
    ReDim Preserve TaskArray(UBound(TaskArray) + 1)
    TaskArray(UBound(TaskArray)) = TaskIn
    Exit Sub
emptyArray:
    ReDim TaskArray(0)
    TaskArray(UBound(TaskArray)) = TaskIn
End Sub

Public Sub swapTask(i As Integer, j As Integer)
    Dim tmpTask As Task
    If i <= UBound(TaskArray) And j <= UBound(TaskArray) Then
        tmpTask = TaskArray(j)
        TaskArray(j) = TaskArray(i)
        TaskArray(i) = tmpTask
    End If
End Sub

Public Sub updateTask(index As Integer, TaskIn As Task)
    If index <= UBound(TaskArray) Then
        TaskArray(i) = TaskIn
    End If
End Sub

Public Sub delAllTasks()
    Erase TaskArray
End Sub

Public Sub delTask(index As Integer)
    If Me.count > 0 Then
        If index <= UBound(TaskArray) Then
           For i = index To UBound(TaskArray) - 1
                TaskArray(i) = TaskArray(i + 1)
            Next i
            If UBound(TaskArray) = 0 Then
                Erase TaskArray
            Else
                ReDim Preserve TaskArray(0 To UBound(TaskArray) - 1)
            End If
        End If
    End If
End Sub

''
' nr tasks
''
Public Function count() As Integer
    If Me.isEmpty() Then
        count = 0
    Else
        count = UBound(TaskArray) + 1
    End If
End Function

Public Function getTask(index As Integer) As Task
    If Me.isEmpty() Then
        Exit Function
    End If
    getTask = TaskArray(index)
End Function



Public Function isEmpty() As Boolean
    
    On Error GoTo isTrue
    If UBound(TaskArray) >= 0 Then
        isEmpty = False
        Exit Function
    End If
isTrue:
    isEmpty = True
End Function

Public Function getTrackXY(index As Integer) As Boolean
    getTrackXY = TaskArray(index).TrackXY
End Function

Public Sub setTrackXY(index As Integer, Track As Boolean)
    TaskArray(index).TrackXY = Track
End Sub


Public Function getTrackZ(index As Integer) As Boolean
    getTrackZ = TaskArray(index).TrackZ
End Function

Public Sub setTrackZ(index As Integer, Track As Boolean)
    TaskArray(index).TrackZ = Track
End Sub


Public Function getPeriod(index As Integer) As Integer
    getPeriod = TaskArray(index).Period
End Function

Public Sub setPeriod(index As Integer, Period As Integer)
    TaskArray(index).Period = Period
End Sub

Public Function getSaveImage(index As Integer) As Boolean
    getSaveImage = TaskArray(index).SaveImage
End Function

Public Sub setSaveImage(index As Integer, save As Boolean)
    TaskArray(index).SaveImage = save
End Sub

Public Function getAnalyse(index As Integer) As Integer
    getAnalyse = TaskArray(index).analyse
End Function

Public Sub setAnalyse(index As Integer, analyse As Integer)
    TaskArray(index).analyse = analyse
End Sub


Public Function getTrackChannel(index As Integer) As Integer
    getTrackChannel = TaskArray(index).TrackChannel
End Function

Public Sub setTrackChannel(index As Integer, channel As Integer)
    TaskArray(index).TrackChannel = channel
End Sub

Public Function getZOffset(index As Integer) As Double
    getZOffset = TaskArray(index).ZOffset
End Function

Public Sub setZOffset(index As Integer, Offset As Double)
    TaskArray(index).ZOffset = Offset
End Sub

