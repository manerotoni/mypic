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

Public Function scaleVector(vec As Vector, alpha As Double) As Vector
    scaleVector.X = vec.X * alpha
    scaleVector.Y = vec.Y * alpha
    scaleVector.Z = vec.Z * alpha
End Function


Public Function scaleVectorList(vec() As Vector, alpha As Double) As Vector()
    Dim outVec() As Vector
    ReDim outVec(0 To UBound(vec))
    Dim i As Integer
    For i = 0 To UBound(vec)
        outVec(i) = scaleVector(vec(i), alpha)
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
Public Sub Autofocus_Setup()
    ZENv = getVersionNr
    'find the version of the software
#If (ZENvC >= 2012) Then
    If ZENv < 2012 Then
        MsgBox "ZENvC Compiler constant is set to 2012 but your ZEN version is below ZEN2012." & vbCrLf & _
        "Edit project properties in the VBA editor by right clicking on project name and modify conditional compiler arguments ZENvC to your ZEN version"
        Exit Sub
    End If
#Else
    If ZENv >= 2012 Then
        MsgBox "ZENvC Compiler constant is not to 2012 or higher but your ZEN version is 2012 or higher." & vbCrLf & _
        "Edit project properties in the VBA editor by right clicking on project name and modify conditional compiler arguments ZENvC to your ZEN version"
        Exit Sub
    End If
#End If
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
Public Function weightedMean(values() As Variant, imin As Long, imax As Long, Optional threshL As Double = 0) As Double
    Dim sum As Variant
    Dim weight As Variant
    Dim minV As Variant
    Dim maxV As Variant
    Dim thresh As Double
    Dim i As Long
    sum = 0
    weight = 0
    minV = MINA(values, imin)
    maxV = MAXA(values, imax)
    If threshL < 0 Or threshL > 1 Then
        threshL = 0
    End If
    thresh = minV + (maxV - minV) * threshL
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
        minLocal = MIN(values(i), minLocal)
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

''''''
''   FServerFromDescription(strName As String, StrPath As String, ExecName As String) As Boolean
''   TODO: What is this?
''''''
'Function FServerFromDescription(strName As String, StrPath As String, ExecName As String) As Boolean
'    Dim lngResult As Long
'    Dim strTmp As String
'    Dim hKeyServer As Long
'    Dim strBuffer As String
'    Dim cb As Long
'    Dim i As Integer
'
'    FServerFromDescription = False
'
'    strTmp = VBA.Space(255)
'    strTmp = strName + "\CLSID"
'    lngResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTmp, 0&, KEY_READ, hKeyServer)
'
'    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
'    strBuffer = VBA.Space(255)
'    cb = Len(strBuffer)
'
'    lngResult = RegQueryValueEx(hKeyServer, "", 0&, REG_SZ, ByVal strBuffer, cb)
'    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
'
'    lngResult = RegCloseKey(hKeyServer)
'    strTmp = VBA.Space(255)
'    strTmp = "CLSID\" + Strings.Left(strBuffer, cb - 1) + "\LocalServer32"
'    strBuffer = VBA.Space(255)
'    cb = Len(strBuffer)
'    lngResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTmp, 0&, KEY_READ, hKeyServer)
'    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
'
'    lngResult = RegQueryValueEx(hKeyServer, "", 0&, REG_SZ, ByVal strBuffer, cb)
'    If (Not lngResult = ERROR_SUCCESS) Then GoTo error_exit
'    StrPath = Strings.Left(strBuffer, cb - 1)
'    ExecName = StrPath
'    lngResult = RegCloseKey(hKeyServer)
'
'    i = Len(StrPath)
'
'    Do Until (i = 0)
'        If (VBA.Mid(StrPath, i, 1) = "\") Then
'            StrPath = Strings.Left(StrPath, i - 1)
'            FServerFromDescription = True
'            Exit Do
'        End If
'        i = i - 1
'    Loop
'
'error_exit:
'    If (Not hKeyServer = 0) Then lngResult = RegCloseKey(hKeyServer)
'
'End Function

