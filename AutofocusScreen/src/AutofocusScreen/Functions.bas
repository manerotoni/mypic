Attribute VB_Name = "Functions"
Option Explicit

Public Sub Autofocus_Setup()
        AutofocusForm.Show
End Sub

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
Public Function InCollection(Col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = Col.Item(key)
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
'   TODO: What is this
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

