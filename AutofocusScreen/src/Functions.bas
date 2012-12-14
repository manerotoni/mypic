Attribute VB_Name = "Functions"
Option Explicit


'''''''''
'Introduce minimize button for Macro window
''''''
Private Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
 
Private Declare Function GetWindowLongA Lib "user32" _
(ByVal hWnd As Long, _
ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLongA Lib "user32" _
(ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Sub FormatUserForm(UserFormCaption As String)
     
    Dim hWnd            As Long
    Dim exLong          As Long
     
    hWnd = FindWindowA(vbNullString, UserFormCaption)
    exLong = GetWindowLongA(hWnd, -16)
    If (exLong And &H20000) = 0 Then
        SetWindowLongA hWnd, -16, exLong Or &H20000
    Else
    End If
     
End Sub
''''''''


'''''
'   FileExist(ByVal Pathname)
'   Check if file is present or not
'''''
Public Function FileExist(ByVal Pathname) As Boolean
    If (Dir(Pathname) = "") Then
        FileExist = False
     Else
        FileExist = True
     End If
End Function

'''''
'   FileName(iPosition As Integer, iSubposition As Integer, iRepetition As Integer ) As String
'   Returns string by concatanating well, and sublocation and timepoint
'       [iPosition] In - Well or large grid position
'       [iSubPosition] In - Meandering grid position
'       [iRepetition]  In - Timepoint/repetition
'''''
Public Function FileName(iPosition As Long, iSubposition As Long, iRepetition As Integer) As String
    'convert numbers into a string
    Dim name As String
    Dim nrZero As Integer
    Dim maxZeros As Integer
    maxZeros = 4
    name = ""
    nrZero = maxZeros - Len(CStr(iPosition))
    name = name + "W" + ZeroString(nrZero) + CStr(iPosition) + "_"
    nrZero = maxZeros - Len(Chr(iSubposition))
    name = name + "P" + ZeroString(nrZero) + CStr(iSubposition) + "_"
    nrZero = maxZeros - Len(Chr(iRepetition))
    name = name + "T" + ZeroString(nrZero) + CStr(iRepetition)
    FileName = name
End Function

'''''
'   ZeroString(NrofZeros As Integer) As String
'   Returns a string of zeros
'       [NrofZeros] In - Length of string
'''''
Private Function ZeroString(NrofZeros As Integer) As String
    'convert numbers into a string
    Dim i As Integer
    Dim name As String
    name = ""
    For i = 1 To NrofZeros
        name = name + "0"
    Next i
    ZeroString = name
End Function

'''''
'   Range() As Double
'   Returs maximal range of Objective movement in um
'''''
Public Function Range() As Double
    Dim RevolverPosition As Long
    RevolverPosition = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
    If RevolverPosition >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(RevolverPosition) * 1000# ' the # is a double declaration
    Else
        Range = 0#
    End If
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

