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

Public Function Range() As Double
    Dim RevolverPosition As Long
    RevolverPosition = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
    If RevolverPosition >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(RevolverPosition) * 1000# ' the # is a double declaration
    Else
        Range = 0#
    End If
End Function


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

