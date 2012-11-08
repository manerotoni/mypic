Attribute VB_Name = "Functions"

Public Function Range() As Double
    Dim RevolverPosition As Long
    RevolverPosition = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
    If RevolverPosition >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(RevolverPosition) * 1000#
    Else
        Range = 0#
    End If
End Function


Public Function GetGlobalZZero(SetZeroMarked As Boolean, ZeroChanged As Boolean)
    Dim Count As Integer
    Dim idx As Long
    
    Dim XPos As Double
    Dim YPos As Double
    Dim ZPos As Double
    Dim x1 As Double
    Dim Y1 As Double
    Dim res1 As Integer

    Dim Success As Boolean
    Dim result As Long
    Dim Positions As Long
    ZeroChanged = False
    If GlobalIsStage Then
        Positions = Lsm5.Hardware.CpStages.MarkCount
        ZPos = CpFocus.Position
        result = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ(0, 0, 0)
        If result <> (Positions + 1) Then
            result = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ(10, 10, 0)
        End If
        result = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).GetMarkZ(Positions, GlobalXZero, GlobalYZero, GlobalZZero)
        res1 = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).ClearMark(Positions)
        
        Positions = GlobalPositionsStage
        If SetZeroMarked Then
            GlobalZZeroMarked = GlobalZZero
        Else
            If Positions <= 1 Then
                Positions = 1
                If GlobalZZeroMarked <> GlobalZZero Then
                    GlobalZZeroMarked = GlobalZZero
                    ZeroChanged = True
                End If
                
            Else
                If GlobalZZeroMarked <> GlobalZZero Then
                    For idx = 1 To Positions
                        GlobalZpos(idx) = GlobalZpos(idx) + GlobalZZeroMarked - GlobalZZero
                    Next idx
                    GlobalZZeroMarked = GlobalZZero
                    ZeroChanged = True
    
                End If
            End If
        End If
    Else
        GlobalZZero = 0
        Positions = 0
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

Public Function PubFuncOverWriteZ() As Boolean
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
If GlobalZmapAquired = True Then
    Msg = "Do You Want to overwrite Z-Values?"
    Style = VbYesNo + VbQuestion + VbDefaultButton2   ' Define buttons.
    Title = "ZValues"  ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then ' User chose Yes.
        PubFuncOverWriteZ = True
    Else
        PubFuncOverWriteZ = False
        GlobalZposOld() = GlobalZpos()
       GlobalLocationsOrderOld() = GlobalLocationsOrder()
    End If
 Else
    PubFuncOverWriteZ = True
End If

End Function
