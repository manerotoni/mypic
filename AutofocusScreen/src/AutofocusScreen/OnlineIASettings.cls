VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OnlineIASettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''
' OiaSettings: A class for Keeping track of settings that can be read from the Registry or from files
'''''
Option Explicit
Public Settings As Dictionary
Private Keys() As String
''keys used for codeMic
Private codeMic As Dictionary
''keys used for codeOia
Private codeOia As Dictionary

'''
' A set of functions to IN/OUT parameters for Onlineimageanalisys this should be a class!!!
''''

''''
' Set Registry to default values
'''
Public Sub initializeDefault()
    On Error GoTo errMsg
'This variable contains the keys for the OnlineImageanalysis
    ReDim Keys(0 To 14)
    Keys(0) = "codeMic"
    Keys(1) = "codeOia"
    Keys(2) = "fileAnalyzed"
    Keys(3) = "filePath"
    Keys(4) = "X"
    Keys(5) = "Y"
    Keys(6) = "Z"
    Keys(7) = "fcsX"
    Keys(8) = "fcsY"
    Keys(9) = "fcsZ"
    Keys(10) = "deltaZ"
    Keys(11) = "roiType"
    Keys(12) = "roiAim"
    Keys(13) = "roiX"
    Keys(14) = "roiY"
    
    Set codeMic = New Dictionary
    codeMic.Add "wait", "wait for image analysis"
    codeMic.Add "", "(empty char) wait for image analysis"
    codeMic.Add "error", "error from image analysis"
    codeMic.Add "timeExpired", "time for image analysis has expired"
    codeMic.Add "nothing", "Do nothing"
    codeMic.Add "focus", "read X, Y, and Z  and change focus"
    codeMic.Add "trigger1", "perform job Trigger1 at position X, Y, and Z"
    codeMic.Add "trigger2", "perform job Trigger2 at position X, Y, and Z"
    
    Set codeOia = New Dictionary
    codeOia.Add "nothing", "online image analysis does nothing"
    codeOia.Add "", "(empty char) online image analysis does nothing"
    codeOia.Add "newImage", "a new image is there"
    
    Set Settings = New Dictionary
    Dim Key As Variant
    For Each Key In Keys
        Settings.Add Key, ""
    Next Key
    Exit Sub
errMsg:
    ErrorLog.UpdateLog "OnlineIASettings.initializeDefault error " + Err.Description
End Sub

Public Function getSettings(Key As Variant) As Variant
    Dim LogTxt As String
    Dim i  As Integer
    If Settings.Exists(Key) Then
        getSettings = Settings.Item(Key)
        Exit Function
    End If
    LogTxt = "OnlineIASettings.getSettings: Failed to get " + Key + " from OiaSettings. Possible values are:" & vbCrLf
    LogTxt = LogTxt & availableKeys
    
    MsgBox LogTxt
    ErrorLog.UpdateLog (LogTxt)
    
End Function

Public Function availableKeys() As String
    Dim i As Integer
    Dim LogTxt As String
    LogTxt = ""
    For i = 0 To UBound(Keys)
        LogTxt = LogTxt & Keys(i) & ", "
    Next i
    availableKeys = LogTxt
End Function

Public Function availableCodeMic() As String
    Dim codeKeys() As Variant
    Dim Key As Variant
    codeKeys = codeMic.Keys
    Dim LogTxt As String
    LogTxt = ""
    For Each Key In codeKeys
       LogTxt = LogTxt & Key & ", "
    Next Key
    availableCodeMic = LogTxt
End Function

Public Function availableCodeOia() As String
    Dim codeKeys() As Variant
    Dim Key As Variant
    codeKeys = codeOia.Keys
    Dim LogTxt As String
    LogTxt = ""
    For Each Key In codeKeys
       LogTxt = LogTxt & Key & ", "
    Next Key
    availableCodeOia = LogTxt
End Function

'''
' Check existence of Key and whether Item is consistent with it
''
Public Function checkKeyItem(Key As Variant, Value As Variant)
    Dim errorMsg As String
    If Settings.Exists(Key) Then
        Settings.Item(Key) = Value
        Select Case Key
            Case "codeOia"
                checkKeyItem = checkCodeOia(Value)
                Exit Function
            Case "codeMic"
                checkKeyItem = checkCodeMic(Value)
                Exit Function
            Case Else
                checkKeyItem = True
        End Select
    Else
        errorMsg = "OnlineIASettings error: key " & CStr(Key) & " is not a correct key for registry. Values are " & vbCrLf
        errorMsg = errorMsg & availableKeys
        MsgBox errorMsg
        ErrorLog.UpdateLog errorMsg
    End If
End Function

Public Function checkCodeOia(Value As Variant) As Boolean
    Dim errorMsg As String
    Dim Key As Variant
    Dim codeKeys() As Variant
    codeKeys = codeOia.Keys
    For Each Key In codeKeys
        If Value = Key Then
            checkCodeOia = True
            Exit Function
        End If
    Next Key
    errorMsg = "OnlineIASettings error: Wrong codeOia " & Me.getSettings("codeOia") & " in Registry. Values are " & vbCrLf
    errorMsg = errorMsg & availableCodeOia
    MsgBox errorMsg
    ErrorLog.UpdateLog errorMsg
End Function



Public Function checkCodeMic(Value As Variant) As Boolean
    Dim errorMsg As String
    Dim Key As Variant
    Dim codeKeys() As Variant
    codeKeys = codeMic.Keys
    For Each Key In codeKeys
        If Settings.Item("codeMic") = Key Then
            checkCodeMic = True
            Exit Function
        End If
    Next Key
    errorMsg = "OnlineIASettings error: Wrong codeMic = " & Settings.Item("codeMic") & " in Registry values are " & vbCrLf
    errorMsg = errorMsg & availableCodeMic
    MsgBox errorMsg
    ErrorLog.UpdateLog errorMsg
End Function


''''
' Set Registry to default values
'''
Public Sub initialize(KeysIn() As String)
'This variable contains the keys for the OnlineImageanalysis
    Dim i As Integer
    ReDim Keys(UBound(KeysIn))
    For i = 0 To UBound(KeysIn)
        Keys(i) = KeysIn(i)
    Next i
    Set Settings = New Dictionary
End Sub

'''
'   Initialize Keys as default when it is empty
'''
Private Sub checkExistKeys()
    If isArrayEmpty(Keys) Then
        initializeDefault
    End If
End Sub


'''
' Default registry values is always empty
'''
Public Sub resetRegistry()
    checkExistKeys
    Dim i As Integer
    For i = LBound(Keys) To UBound(Keys)
        SaveSetting "OnlineImageAnalysis", "macro", Keys(i), ""
    Next i
End Sub

''''
'   ReadOiaSettingsFromRegistry(Settings As Collection, Keys() As String)
'       Read Registry using the keys stored in Keys and create a new dictionary
''''
Public Sub readFromRegistry()
    checkExistKeys
    Dim Key As Variant
    For Each Key In Keys
        If Settings.Exists(Key) Then
            Settings.Item(Key) = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:=Key)
        Else
            Settings.Add Key, GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:=Key)
        End If
    Next Key
End Sub




''''
'   writeToRegistry()
'       Write settings with keys defined in Keys to Registry
''''

Public Sub writeToRegistry()
    checkExistKeys
    Dim Key As Variant
    For Each Key In Keys
        If Settings.Exists(Key) Then
            SaveSetting "OnlineImageAnalysis", "macro", Key, Settings.Item(Key)
        End If
    Next Key
End Sub

''''
'   writeKeyToRegistry()
'       Write a specific value in Key of registry
''''
Public Sub writeKeyToRegistry(Key As Variant, Value As Variant)
    Dim Success As Boolean
    checkExistKeys
    
    If checkKeyItem(Key, Value) Then
        Settings.Item(Key) = Value
        SaveSetting "OnlineImageAnalysis", "macro", Key, Value
    Else
        MsgBox "OnlineIASettings.writeKeyToRegistry: no key " & Key & " available. Possible Keys are:" & vbCrLf & _
        availableKeys
        ErrorLog.UpdateLog "OnlineIASettings.writeKeyToRegistry: no key " & Key & " available. Possible Keys are:" & vbCrLf & _
        availableKeys
    End If
End Sub

''''
'   ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
'       Read FileName and store key and paramter into Settings
''''
Public Sub readFromFile(FileName As String)
    checkExistKeys
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    Dim Entries() As String
    Close
    On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open FileName For Input As iFileNum
    Do While Not EOF(iFileNum)
            Line Input #iFileNum, Fields
            While Left(Fields, 1) = "%" 'this are comments
                Line Input #iFileNum, Fields
            Wend
            FieldEntries = Split(Fields, " ", 2)
            Settings.Add FieldEntries(0), FieldEntries(1)
    Loop
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & FileName & " for OiaSettings"
End Sub

''''
'   WriteOiaSettingsToFile(Settings As Dictionary, FileName As String)
'   Write Settings to file with FileName
''''
Public Sub writeToFile(FileName As String)
    checkExistKeys
    Dim i As Integer
    Dim iFileNum As Integer
    Dim Key As Variant
    Close
    'On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open FileName For Output As iFileNum
    
    For Each Key In Settings.Keys
        Print #iFileNum, Key & " " & Settings.Item(Key)
    Next Key
    Close
ErrorHandle:
End Sub

''''
'   Parse FileName to get name of SettingFile
'   It is assumed that FileName = something_Txxx.lsm
'   ToDO: better parsing
'''
Public Function settingsFileName(FileName As String) As String
    settingsFileName = Left(FileName, Len(FileName) - 9) & "_oia.txt"
End Function


''''
'   Parse rois from settings
'   This should be put in the Jobdefinition as a Roi is associated to a Job
''''
Public Function getRois(Rois() As Roi) As Boolean
    checkExistKeys
    Dim RoiOut() As Roi
    Dim XRois()  As String 'the string containinig all X-positions of a ROI
    Dim YRois()  As String 'the string containinig all Y-positions of a ROI
    Dim x() As String ' string containing X pos of a single ROI
    Dim y() As String ' String containing y pos of a single ROI
    Dim XD()  As Double 'the double array containinig the X-positions
    Dim YD()  As Double 'the double array containinig the Y-positions
    Dim i As Integer
    Dim iRoi As Integer
    Dim roiType() As String
    Dim roiAim() As String
    If Settings.Item("roiType") = "" Then
        Exit Function
    End If
    roiType = Split(Settings.Item("roiType"), ";")
    If Settings.Item("roiAim") = "" Or Settings.Item("roiX") = "" Or Settings.Item("roiY") = "" Then
        MsgBox ("GetRoisFromRegistry: For each roi you need to define roiType, roiAim, roiX, and roiY!" + vbCrLf + "roiType1 ; roiType2; etc." & vbCrLf & _
        "roiX1_roi1, roiX2_roi1; roiX1_roi2, roiX2_roi2, roiX3_roi2; etc. Coordinates in pixels")
        Exit Function
    End If
    
    roiAim = Split(Settings.Item("roiAim"), ";")
    XRois() = Split(Settings.Item("roiX"), ";")
    YRois() = Split(Settings.Item("roiY"), ";")
    
    If UBound(roiType) <> UBound(roiAim) And UBound(roiType) <> UBound(XRois) And UBound(roiType) <> UBound(YRois) Then
        MsgBox ("GetRoisFromRegistry: Number of Rois and coordinates need to correspond." + vbCrLf + "roiType1 ; roiType2; etc." & vbCrLf & _
        "roiX1_roi1, roiX2_roi1; roiX1_roi2, roiX2_roi2, roiX3_roi2; etc. Coorindates in pixels")
        Exit Function
    End If
    ReDim Rois(0 To UBound(roiType))
    For iRoi = 0 To UBound(roiType)
        x() = Split(XRois(iRoi), ",")
        y() = Split(YRois(iRoi), ",")
        If isArrayEmpty(x) Or isArrayEmpty(y) Then
            MsgBox "GetRoisFromRegistry: No coordinates found in registry foir roix and roiy (in pixel)"
            Exit Function
        End If
        ReDim XD(UBound(x))
        ReDim YD(UBound(y))
        For i = 0 To UBound(x)
            XD(i) = CDbl(x(i))
            YD(i) = CDbl(y(i))
        Next i
        Rois(iRoi).setRoi roiType(iRoi), roiAim(iRoi), XD, YD
        If Not Rois(iRoi).roiConsistency Then
            Exit Function
        End If
    Next iRoi
    getRois = True
End Function


'''
'   GetPositionsFromSettings(Settings As Dictionary, StgPos() As Vector) As Boolean
'   Settings contains all settings for OnlineImageanalysis
'   StgPos containes the coordinates
'   StgPos.X, StgPos.Y: defined 0,0 at upper left corner
'   StgPos.Z: defined 0 for the central slice
'   units are converted afterwards as it depends on the type of Job (default unit is px)
''''
Public Function getPositions(StgPos() As Vector, currentPosition As Vector) As Boolean
    checkExistKeys
    ' store postion from windows registry in array
    Dim locX()  As String 'the string containinig the X-positions
    Dim locY()  As String 'the string containinig the Y-positions
    Dim locZ() As String  'the string containinig the Z-positions
    Dim i As Integer
    
    If Settings.Exists("X") Then
        locX() = Split(Settings.Item("X"), ",")
    End If
    
    If isArrayEmpty(locX) Then
        Exit Function
    End If
    
    If Settings.Exists("Y") Then
        locY() = Split(Settings.Item("Y"), ",")
    End If
    
    If isArrayEmpty(locY) Then
        Exit Function
    End If
    
    If UBound(locX) <> UBound(locY) Then
        MsgBox ("OnlineIASettings.getPositions: nr of values in registry for X, Y are not the same, separate the values with comma!")
        ErrorLog.UpdateLog "OnlineIASettings.getPositions: nr of values in registry for X, Y are not the same, separate the values with comma!"
        Exit Function
    End If
    
    If Settings.Exists("Z") Then
        locZ() = Split(Settings.Item("Z"), ",")
    End If
    
    If isArrayEmpty(locZ) Then 'Z has not been set. We use the currentvalue as default
        ReDim locZ(UBound(locX))
        For i = 0 To UBound(locX)
            locZ(i) = currentPosition.Z
        Next i
    End If
        
    If UBound(locZ) <> UBound(locX) Then 'Z has not been set for all positions
        MsgBox ("OnlineIASettings.getPositions: nr of values in registry for offsetX, offsetZ are not the same, separate the values with comma!")
        ErrorLog.UpdateLog "OnlineIASettings.getPositions: nr of values in registry for offsetX, offsetZ are not the same, separate the values with comma!"
        Exit Function
    End If
    
'    If Settings.Exists("deltaZ") Then
'        locDeltaZ() = Split(Settings.Item("deltaZ"), ",")
'    End If
'
'    If isArrayEmpty(locDeltaZ) Then
'        ReDim locDeltaZ(UBound(locX)) 'deltaZ has not been set. We use a default values
'        For i = 0 To UBound(locDeltaZ)
'            locDeltaZ(i) = -1
'        Next i
'    End If
'
'    If UBound(locDeltaZ) <> UBound(locX) Then 'deltaZ has not been set for all positions
'        MsgBox ("StorePositionsFromRegistry: nr of values in registry for deltaz, z are not the same, separate the values with comma!")
'        Exit Function
'    End If
    
    ' Convert values to Double
    ReDim StgPos(0 To UBound(locX))
    For i = 0 To UBound(locX)
        StgPos(i).x = CDbl(locX(i))
        StgPos(i).y = CDbl(locY(i))
        StgPos(i).Z = CDbl(locZ(i))
    Next i
    getPositions = True
End Function



