Attribute VB_Name = "Functions"
'''
' Some utility functions
''''
Option Explicit

''' A vector
Public Type Vector
  X As Double
  Y As Double
  Z As Double
End Type

Public Type WellPoint
    pos As Vector
    well As String
End Type





Public Function ProcessEvents(ByVal EventNr As Long, ByVal ObjName As String, ByVal PropertyNr As Long, ByVal Param As Variant)
    If EventNr = eEventScanEnd Then
        While Lsm5.ExternalDsObject.ScanController.IsGrabbing
            SleepWithEvents (200)
        Wend
        PipelineConstructor.EventMng.setReady True
    End If
    
    If EventNr = eEventScanStart Then
        PipelineConstructor.EventMng.setBusy True
    End If

    If EventNr = eEventScanStop Then
        'this events can be triggered from within a program
        PipelineConstructor.EventMng.setReady True
    End If
    
    ''This is specific to see if the FCS system is free
    If EventNr = ePropertyEventShutters And Param = 2 Then
        If InStr(ObjName, "FCS") Then
            Dim FcsControl As AimFcsController
            Set FcsControl = Application.Fcs
            While FcsControl.IsAcquisitionRunning(1)
                SleepWithEvents (200)
            Wend
            PipelineConstructor.EventMng.setReady
        End If
    End If
        ''This is specific to see if the FCS system is still acquiring
    If EventNr = ePropertyEventShutters And Param = 1 Then
        If InStr(ObjName, "FCS") Then
            PipelineConstructor.EventMng.setBusy 1
        End If
    End If
End Function

''''''''''''''''''''''''
''' Vector operations'''
''' To do: overload?''''
''''''''''''''''''''''''
Public Function Double2Vector(X As Double, Y As Double, Z As Double) As Vector
    Double2Vector.X = X
    Double2Vector.Y = Y
    Double2Vector.Z = Z
End Function

Public Function Vector2Array(vec As Vector) As Vector()
    Dim vecA(0) As Vector
    vecA(0) = vec
    Vector2Array = vecA
End Function


Public Function Vector2Double(vec As Vector) As Double()
    Dim vec2D(3) As Double
    vec2D(0) = vec.X
    vec2D(1) = vec.Y
    vec2D(2) = vec.Z
    Vector2Double = vec2D
End Function

Public Function diffVector(vec1 As Vector, vec2 As Vector) As Vector
    diffVector.X = vec1.X - vec2.X
    diffVector.Y = vec1.Y - vec2.Y
    diffVector.Z = vec1.Z - vec2.Z
End Function

Public Function sumVector(vec1 As Vector, vec2 As Vector) As Vector
    sumVector.X = vec1.X + vec2.X
    sumVector.Y = vec1.Y + vec2.Y
    sumVector.Z = vec1.Z + vec2.Z
End Function

Public Function normVector2D(vec As Vector) As Double
    normVector2D = Sqr(vec.X ^ 2 + vec.Y ^ 2)
End Function

Public Function normVector3D(vec As Vector) As Double
    normVector3D = Sqr(vec.X ^ 2 + vec.Y ^ 2 + vec.Z ^ 2)
End Function

Public Function scaleVector(vec As Vector, Alpha As Double) As Vector
    scaleVector.X = vec.X * Alpha
    scaleVector.Y = vec.Y * Alpha
    scaleVector.Z = vec.Z * Alpha
End Function

Public Function scaleVectorList(vec() As Vector, Alpha As Double) As Vector()
    Dim outVec() As Vector
    ReDim outVec(0 To UBound(vec))
    Dim i As Integer
    For i = 0 To UBound(vec)
        outVec(i) = scaleVector(vec(i), Alpha)
    Next i
    scaleVectorList = outVec
End Function



'''
' Create a ; separated string of the elements in a vector list
'''
Public Function VectorList2String(vec() As Vector, Optional Rnd = 2) As String()
    Dim i As Integer
    Dim OutString(0 To 2) As String
    OutString(0) = "" & Round(vec(0).X, Rnd)
    OutString(1) = "" & Round(vec(0).Y, Rnd)
    OutString(2) = "" & Round(vec(0).Z, Rnd)
    If UBound(vec) > 0 Then
        For i = 1 To UBound(vec)
            OutString(0) = OutString(0) & "; " & Round(vec(i).X, Rnd)
            OutString(1) = OutString(1) & "; " & Round(vec(i).Y, Rnd)
            OutString(2) = OutString(2) & "; " & Round(vec(i).Z, Rnd)
        Next i
    End If
    VectorList2String = OutString
End Function



'''Starts the form
Public Sub PipelineConstructor_Setup()
    ZenV = getVersionNr
    'find the version of the software
#If ZENvC >= 2012 Then
    If ZenV < 2012 Then
        LogManager.UpdateErrorLog "ZENvC Compiler constant is set to 2012 but your ZEN version is " & ZenV & vbCrLf & _
        "Edit project properties in the VBA editor by right clicking on project name and modify conditional compiler arguments ZENvC to your ZEN version"
    End If
#Else
    If ZenV >= 2012 Then
        LogManager.UpdateErrorLog "ZENvC Compiler constant is not set for 2012 or higher but your ZEN version is 2012 or higher." & vbCrLf & _
        "Edit project properties in the VBA editor by right clicking on project name and modify conditional compiler arguments ZENvC to your ZEN version"
    End If
#End If
    PipelineConstructor.Show
End Sub

'''
'   Display progress in bottom labal of AutofocusForm
'''
Public Sub DisplayProgress(Label1 As Label, State As String, Color As Long)       'Used to display in the progress bar what the macro is doing
    If (Color & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Then
       Label1.ForeColor = 0
    Else
       Label1.ForeColor = &HFFFFFF
    End If
    Label1.BackColor = Color
    Label1.Caption = State
    DoEvents
End Sub

'''
' Shifts position of element in list
'''
Public Sub MoveListboxItem(List1 As ListBox, CurrentIndex As Integer, newIndex As Integer)
    Dim strItem() As String
    Dim i As Integer
    With List1
        If CurrentIndex > -1 And CurrentIndex < .ColumnCount And newIndex > -1 And newIndex < .ColumnCount Then
            ReDim strItem(0 To .ColumnCount - 1)
            For i = 0 To .ColumnCount - 1
                strItem(i) = List1.List(CurrentIndex, i)
            Next i
            .RemoveItem CurrentIndex
            .AddItem strItem(0), newIndex
            For i = 1 To .ColumnCount - 1
                .List(newIndex, i) = strItem(i)
            Next i
        End If
    End With
End Sub

'''
' Set all elements in frame to enabled = value
''''
Public Sub enableFrame(AFrame As Frame, value As Boolean)
    Dim i As Integer
    For i = 0 To AFrame.Controls.count - 1
        AFrame.Controls.item(i).Enabled = value
    Next i
    If value Then
        AFrame.ForeColor = "&H80000012"
    Else
         AFrame.ForeColor = "&H8000000A"
    End If
End Sub

'''
' compute a weighted mean of the positiions of an array
'''
Public Function weightedMean(values() As Variant, imin As Long, imax As Long, Optional threshL As Double = 0) As Double
    Dim sum As Variant
    Dim weight As Variant
    Dim minV As Variant
    Dim MaxV As Variant
    Dim thresh As Double
    Dim i As Long
    sum = 0
    weight = 0
    minV = MINA(values, imin)
    MaxV = MAXA(values, imax)
    If threshL < 0 Or threshL > 1 Then
        threshL = 0
    End If
    thresh = minV + (MaxV - minV) * threshL
    For i = LBound(values) To UBound(values)
        sum = sum + Positive(values(i) - thresh)
        weight = weight + i * Positive(values(i) - thresh)
    Next i

    If sum > 0 Then
        weightedMean = weight / sum
        
    Else
        'if sum is 0 then mean is in the center
        weightedMean = (UBound(values) - LBound(values)) / 2 + LBound(values)
        imax = (UBound(values) - LBound(values)) / 2 + LBound(values)
        imin = (UBound(values) - LBound(values)) / 2 + LBound(values)
    End If
End Function

''''
' Set negative values to 0
'''''
Public Function Positive(value As Variant) As Variant
    If value < 0 Then
        Positive = 0
    Else
        Positive = value
    End If
End Function

''
' Calculate MIN of two values
'''
Public Function Min(value1 As Variant, value2 As Variant) As Variant
    If value1 > value2 Then
        Min = value2
    Else
        Min = value1
    End If
End Function


''
' Calculate MIN of two values
'''
Public Function MAX(value1 As Variant, value2 As Variant) As Variant
    If value1 < value2 Then
        MAX = value2
    Else
        MAX = value1
    End If
End Function


''
' Calculate MIN of Array
'''
Public Function MINA(values() As Variant, Optional imin As Long) As Variant
    Dim minLocal As Variant
    Dim i As Integer
    minLocal = values(0)
    imin = 0
    For i = LBound(values) To UBound(values)
        minLocal = Min(values(i), minLocal)
        If minLocal = values(i) Then
            imin = i
        End If
    Next i
    MINA = minLocal
End Function

''
' Calculate MAX of Array
'''
Public Function MAXA(values() As Variant, Optional imax As Long) As Variant
    Dim maxLocal As Variant
    Dim i As Long
    maxLocal = values(0)
    imax = 0
    For i = LBound(values) To UBound(values)
        maxLocal = MAX(values(i), maxLocal)
        If maxLocal = values(i) Then
            imax = i
        End If
    Next i
    MAXA = maxLocal
End Function



'''''
'  isArrayEmpty(parArray As Variant) As Boolean
'  Returns false if not an array or dynamic array that has not been initialised (ReDim) or has been erased (Erase)
'''''
Public Function isArrayEmpty(parArray As Variant) As Boolean
    If IsArray(parArray) = False Then isArrayEmpty = True
    On Error Resume Next
    If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False
End Function

'''''
'  isArrayEmpty(parArray As Variant) As Boolean
'  Returns false if not an array or dynamic array that has not been initialised (ReDim) or has been erased (Erase)
'''''
Public Function isPosArrayEmpty(parArray() As Vector) As Boolean
    On Error Resume Next
    If UBound(parArray) < LBound(parArray) Then isPosArrayEmpty = True: Exit Function Else: isPosArrayEmpty = False
End Function

''''
' Check if key is in collection
''''
Public Function InCollection(Col As Collection, Key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = Col.item(Key)
    errNumber = CLng(Err.number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

''''
' return index of List entry that has been selected. Return -1 if no entry is selected
''''
Public Function selectedListIndex(List As ListBox) As Long
    Dim i As Long
    
    If List.ListIndex = -1 Then
        selectedListIndex = -1
        Exit Function
    End If
    For i = 0 To List.ListCount - 1
        If List.Selected(i) Then
            selectedListIndex = i
            Exit Function
        End If
    Next i
    
End Function

Public Sub QuickSort(ByRef Field() As String, ByVal LB As Long, ByVal UB As Long)
    Dim P1 As Long, P2 As Long, Ref As String, TEMP As String

    P1 = LB
    P2 = UB
    Ref = Field((P1 + P2) / 2)

    Do
        Do While (Field(P1) < Ref)
            P1 = P1 + 1
        Loop

        Do While (Field(P2) > Ref)
            P2 = P2 - 1
        Loop

        If P1 <= P2 Then
            TEMP = Field(P1)
            Field(P1) = Field(P2)
            Field(P2) = TEMP

            P1 = P1 + 1
            P2 = P2 - 1
        End If
    Loop Until (P1 > P2)

    If LB < P2 Then Call QuickSort(Field, LB, P2)
    If P1 < UB Then Call QuickSort(Field, P1, UB)
End Sub
