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

''''''''''''''''''''''''
''' Vector operations'''
''' To do: overload?''''
''''''''''''''''''''''''
Public Function Double2Vector(X As Double, Y As Double, Z As Double) As Vector
    Double2Vector.X = X
    Double2Vector.Y = Y
    Double2Vector.Z = Z
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


'''Starts the form
Public Sub Autofocus_Setup()
        AutofocusForm.Show
End Sub
'''
'   Display progress in bottom labal of AutofocusForm
'''
Public Sub DisplayProgress(State As String, Color As Long)       'Used to display in the progress bar what the macro is doing
    If (Color & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Then
        AutofocusForm.ProgressLabel.ForeColor = 0
    Else
        AutofocusForm.ProgressLabel.ForeColor = &HFFFFFF
    End If
    AutofocusForm.ProgressLabel.BackColor = Color
    AutofocusForm.ProgressLabel.Caption = State
    DoEvents
End Sub


'''
' compute a weighted mean of the positiions of an array
'''
Public Function weightedMean(values() As Variant) As Double
    Dim sum As Variant
    Dim weight As Variant
    Dim MIN As Variant
    Dim i As Integer
    sum = 0
    weight = 0
    MIN = MINA(values)
    For i = LBound(values) To UBound(values)
        sum = sum + (values(i) - MIN)
        weight = weight + i * (values(i) - MIN)
    Next i
    'if sum is 0
    If sum > 0 Then
        weightedMean = weight / sum
    Else
        ' then mean is in the center
        weightedMean = (UBound(values) - LBound(values)) / 2 + LBound(values)
    End If
End Function


''
' Calculate MIN of two values
'''
Public Function MIN(value1 As Variant, value2 As Variant) As Variant
    If value1 > value2 Then
        MIN = value2
    Else
        MIN = value1
    End If
End Function


''
' Calculate MIN of two values
'''
Public Function Max(value1 As Variant, value2 As Variant) As Variant
    If value1 < value2 Then
        Max = value2
    Else
        Max = value1
    End If
End Function


''
' Calculate MIN of Array
'''
Public Function MINA(values() As Variant) As Variant
    Dim minLocal As Variant
    Dim i As Integer
    minLocal = values(0)
    For i = LBound(values) To UBound(values)
        minLocal = MIN(values(i), minLocal)
    Next i
    MINA = minLocal
End Function

''
' Calculate MIN of Array
'''
Public Function MAXA(values() As Variant) As Variant
    Dim maxLocal As Variant
    Dim i As Integer
    maxLocal = values(0)
    For i = LBound(values) To UBound(values)
        maxLocal = Max(values(i), maxLocal)
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
    var = Col.Item(Key)
    errNumber = CLng(Err.number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

'''''
'   FServerFromDescription(strName As String, StrPath As String, ExecName As String) As Boolean
'   TODO: What is this?
'''''
Function FServerFromDescription(strName As String, StrPath As String, ExecName As String) As Boolean
    Dim lngResult As Long
    Dim strTmp As String
    Dim hKeyServer As Long
    Dim strBuffer As String
    Dim cb As Long
    Dim i As Integer
     
    FServerFromDescription = False
    
    strTmp = VBA.Space(255)
    strTmp = strName + "\CLSID"
    lngResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTmp, 0&, KEY_READ, hKeyServer)
    
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
    strBuffer = VBA.Space(255)
    cb = Len(strBuffer)
    
    lngResult = RegQueryValueEx(hKeyServer, "", 0&, REG_SZ, ByVal strBuffer, cb)
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
    
    lngResult = RegCloseKey(hKeyServer)
    strTmp = VBA.Space(255)
    strTmp = "CLSID\" + Strings.Left(strBuffer, cb - 1) + "\LocalServer32"
    strBuffer = VBA.Space(255)
    cb = Len(strBuffer)
    lngResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTmp, 0&, KEY_READ, hKeyServer)
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
        
    lngResult = RegQueryValueEx(hKeyServer, "", 0&, REG_SZ, ByVal strBuffer, cb)
    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
    StrPath = Strings.Left(strBuffer, cb - 1)
    ExecName = StrPath
    lngResult = RegCloseKey(hKeyServer)
    
    i = Len(StrPath)
    
    Do Until (i = 0)
        If (VBA.Mid(StrPath, i, 1) = "\") Then
            StrPath = Strings.Left(StrPath, i - 1)
            FServerFromDescription = True
            Exit Do
        End If
        i = i - 1
    Loop

error_exit:
    If (Not hKeyServer = 0) Then lngResult = RegCloseKey(hKeyServer)

End Function

