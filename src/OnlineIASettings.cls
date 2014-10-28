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
' the seetings are use for the interaction with an external online image analysis program
'''''

Option Explicit
'contains all settings from registry
Public Settings As Dictionary
Private Keys() As String
Private Keys_HowTo As Dictionary
''keys used for codeMic
Private subKey_codeMic_HowTo As Dictionary
''keys used for codeOia
Private subKey_codeOia_HowTo As Dictionary

Private subKey_roiAim_HowTo As Dictionary
Private subKey_roiType_HowTo As Dictionary

'''
' A set of functions to IN/OUT parameters for Onlineimageanalisys this should be a class!!!
''''

''''
' initiate dictionary for standard key values
'''
Public Sub initializeDefault()
    On Error GoTo errMsg
'This variable contains the keys for the OnlineImageanalysis
    ReDim Keys(0 To 18)
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
    Keys(15) = "OutputFolder"
    Keys(16) = "errorMsg"
    Keys(17) = "codePump"
    Keys(18) = "prefix" 'prefix to be added to filenames
    
    Set subKey_roiAim_HowTo = New Dictionary
    subKey_roiAim_HowTo.Add "", "(empty char) do nothing with roi"
    subKey_roiAim_HowTo.Add "nothing", "do nothing with roi"
    subKey_roiAim_HowTo.Add "bleach", "bleach and analyse roi"
    subKey_roiAim_HowTo.Add "acquisition", "acquire roi"
    subKey_roiAim_HowTo.Add "acquisitionBleach", "acquire and bleach"
    subKey_roiAim_HowTo.Add "acquisitionAnalyse", "acquire and analyse"
    subKey_roiAim_HowTo.Add "bleachAnalyse", "bleach and analyse"

    
    Set subKey_roiType_HowTo = New Dictionary
    subKey_roiType_HowTo.Add "", "(empty char) no type"
    subKey_roiType_HowTo.Add "circle", "Define 2 points (in px) roiX=center_X, point_on_circle_X roiY=center_Y, point_on_circle_Y"
    subKey_roiType_HowTo.Add "rectangle", "Define 2 points (in px) roiX=upper_left_X, lower_right_X roiY=upper_left_Y, lower_right_Y"
    subKey_roiType_HowTo.Add "polyline", "Define at least 3 points (in px) roiX=point1_X, point2_X ... roiY=point1_Y, point2_Y..."
    subKey_roiType_HowTo.Add "ellipse", "Define 3 points (in px) roiX=center_X, axis1_X, axis2_X roiY= center_Y, axis1_Y, axis2_Y"
    
    Set Keys_HowTo = New Dictionary
    Keys_HowTo.Add Keys(0), "task to be executed by the microscope"
    Keys_HowTo.Add Keys(1), "message for the online image analysis (Oia)"
    Keys_HowTo.Add Keys(2), "file that has been analyzed by Oia (optional)"
    Keys_HowTo.Add Keys(3), "filepath of current image"
    Keys_HowTo.Add Keys(4), "X position stage. Multiple positions: X1; X2; X3; ..."
    Keys_HowTo.Add Keys(5), "Y position stage. Multiple positions: Y1; Y2; Y3; ..."
    Keys_HowTo.Add Keys(6), "Z position stage. Empty char: keep current Z-position. Multiple positions: Z1; Z2; Z3; ..."
    Keys_HowTo.Add Keys(7), "X position for fcs. Multiple positions: X1; X2; X3; ..."
    Keys_HowTo.Add Keys(8), "Y position for fcs. Multiple positions: Y1; Y2; Y3; ..."
    Keys_HowTo.Add Keys(9), "Z position for fcs. Empty char: keep current Z-position. Multiple positions: Z1; Y2; Y3; ..."
    Keys_HowTo.Add Keys(10), "number of Z-slices for current job (not in use)"
    Keys_HowTo.Add Keys(11), "type of roi. Multiple rois: roiType1; roiTyp2"
    Keys_HowTo.Add Keys(12), "aim of roi. Multiple rois: roiAim1; roiAim2"
    Keys_HowTo.Add Keys(13), "X positions for roi. Multiple rois: X1roi1, X2roi1; X1roi1, X2roi2 ..."
    Keys_HowTo.Add Keys(14), "Y positions for roi. Multiple rois: Y1roi1, Y2roi1; Y1roi1, Y2roi2 ..."
    Keys_HowTo.Add Keys(15), "main directory where images are stored"
    Keys_HowTo.Add Keys(16), "error message coming from Oia"
    Keys_HowTo.Add Keys(17), "time (ms) to activate water pump"
    Keys_HowTo.Add Keys(18), "prefix that is added to filename"
    
    Set subKey_codeMic_HowTo = New Dictionary
    subKey_codeMic_HowTo.Add "", "(empty char) wait for image analysis"
    subKey_codeMic_HowTo.Add "wait", "wait for image analysis"
    subKey_codeMic_HowTo.Add "error", "error from image analysis"
    subKey_codeMic_HowTo.Add "timeExpired", "time for image analysis has expired"
    subKey_codeMic_HowTo.Add "nothing", "Do nothing"
    subKey_codeMic_HowTo.Add "focus", "read X, Y, and Z  and change focus"
    subKey_codeMic_HowTo.Add "trigger1", "perform job Trigger1 at position X, Y, and Z"
    subKey_codeMic_HowTo.Add "trigger2", "perform job Trigger2 at position X, Y, and Z"
    subKey_codeMic_HowTo.Add "fcs1", "perform job Fcs1 at position fcsX, fcsY, and fcsZ"
    
    Set subKey_codeOia_HowTo = New Dictionary
    subKey_codeOia_HowTo.Add "", "(empty char) online image analysis does nothing"
    subKey_codeOia_HowTo.Add "nothing", "online image analysis does nothing"
    subKey_codeOia_HowTo.Add "newImage", "a new image is there"
    
    Set Settings = New Dictionary
    Dim Key As Variant
    For Each Key In Keys
        Settings.Add Key, ""
    Next Key
    Exit Sub
errMsg:
    LogManager.UpdateErrorLog "OnlineIASettings.initializeDefault error " & Err.Description
End Sub

'''
' This get settings from stored one not directly from registry
'''
Public Function getSettings(Key As Variant) As Variant
    Dim LogTxt As String
    Dim i  As Integer
    If Settings.Exists(Key) Then
        getSettings = Settings.item(Key)
        Exit Function
    End If
    LogTxt = "OnlineIASettings.getSettings: Failed to get " + Key + " from OiaSettings. Possible values are:" & vbCrLf
    LogTxt = LogTxt & available_subKey(Keys_HowTo)
    
    MsgBox LogTxt
    LogManager.UpdateErrorLog LogTxt
End Function


Public Function available_subKey(subKey_HowTo As Dictionary) As String
    Dim TabSp  As String
    Dim codeKeys() As Variant
    Dim Key As Variant
    codeKeys = subKey_HowTo.Keys
    Dim LogTxt As String
    For Each Key In codeKeys
        If Len(CStr(Key)) > 6 Then
            TabSp = "" & vbTab
        Else
            TabSp = "" & vbTab & vbTab
        End If
        LogTxt = LogTxt & "   " & Key & " " & TabSp & subKey_HowTo.item(Key) & vbCrLf
    Next Key
    available_subKey = LogTxt
End Function

Public Function createKeyReport() As String
    Dim report As String
    report = "Windows registry keys to interact with Online image analysis" & vbCrLf & _
    "location is: HKCU\SOFTWARE\VB and VBA Program Settings\OnlineImageAnalysis\macro" & vbCrLf & _
    "All coordinates must be in pixels. (X, Y, Z) = (0,0,0) is the upper left corner, first slice." & vbCrLf & _
    "Use dot notation for subpixel (e.g. 2.5) !" & vbCrLf & _
    "------------------------" & vbCrLf & _
    "Principal Keys and meaning:" & vbCrLf
    report = report & available_subKey(Keys_HowTo)
    report = report & "------------------------" & vbCrLf
    report = report & "codeMic values: several commands separated with ;" & vbCrLf
    report = report & available_subKey(subKey_codeMic_HowTo)
    report = report & "------------------------" & vbCrLf
    report = report & "codeOia values:" & vbCrLf
    report = report & available_subKey(subKey_codeOia_HowTo)
    report = report & "------------------------" & vbCrLf
    report = report & "roiType values:" & vbCrLf
    report = report & available_subKey(subKey_roiType_HowTo)
    report = report & "------------------------" & vbCrLf
    report = report & "roiAim values:" & vbCrLf
    report = report & available_subKey(subKey_roiAim_HowTo)
    
    createKeyReport = report
End Function


'''
' Check existence of Key and whether Item is consistent with it
''
Public Function checkKeyItem(Key As Variant, value As Variant)
    Dim errorMsg As String
    If Settings.Exists(Key) Then
        Settings.item(Key) = value
        Select Case Key
            Case "codeOia"
                checkKeyItem = checkSubkey(CStr(Key), value, subKey_codeOia_HowTo)
                Exit Function
            Case "codeMic"
                checkKeyItem = checkSubkey(CStr(Key), value, subKey_codeMic_HowTo)
                Exit Function
            Case "roiAim"
                checkKeyItem = checkSubkey(CStr(Key), value, subKey_roiAim_HowTo)
            Case "roiType"
                checkKeyItem = checkSubkey(CStr(Key), value, subKey_roiType_HowTo)
                Exit Function
            Case Else
                checkKeyItem = True
        End Select
    Else
        errorMsg = "OnlineIASettings error: key " & CStr(Key) & " is not a correct key for registry. Values are " & vbCrLf
        errorMsg = errorMsg & available_subKey(Keys_HowTo)
        MsgBox errorMsg
        LogManager.UpdateErrorLog errorMsg
    End If
End Function

Public Function checkSubkey(KeyName As String, value As Variant, subKey_HowTo As Dictionary) As Boolean
    Dim errorMsg As String
    Dim Key As Variant
    Dim code As Variant
    Dim codeArr() As String
    Dim item As Variant
    codeArr = Split(value, ";")
    For Each code In codeArr
        code = Replace(code, " ", "")
        On Error GoTo noKey
        If Not subKey_HowTo.Exists(code) Then
            GoTo noKey
        End If
    Next code
    checkSubkey = True
    Exit Function
noKey:
    errorMsg = "OnlineIASettings error: Wrong " & KeyName & " = " & CStr(code) _
    & " in Registry. Values are " & vbCrLf
    errorMsg = errorMsg & available_subKey(subKey_HowTo)
    MsgBox errorMsg
    LogManager.UpdateErrorLog errorMsg
End Function



'''
'   check if Keys has been initialized
'''
Private Sub checkExistKeys()
    If isArrayEmpty(Keys) Then
        initializeDefault
    End If
End Sub

Public Sub deleteKeys()
    On Error Resume Next
    DeleteSetting "OnlineImageAnalysis", "macro"
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
            Settings.item(Key) = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:=Key)
        Else
            Settings.Add Key, GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:=Key)
        End If
        checkKeyItem Key, Settings.item(Key)
    Next Key
End Sub

''''
'   ReadOiaSettingsFromRegistry(Settings As Collection, Keys() As String)
'       Read Registry using the keys stored in Keys and create a new dictionary
''''
Public Function readKeyFromRegistry(Key As Variant) As Variant
    Dim KeyValue As Variant
    checkExistKeys
    KeyValue = GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:=Key)
    If Settings.Exists(Key) Then
        Settings.item(Key) = KeyValue
    Else
        Settings.Add Key, KeyValue
    End If
    checkKeyItem Key, Settings.item(Key)
    readKeyFromRegistry = KeyValue
End Function


''''
'   writeToRegistry()
'       Write settings with keys defined in Keys to Registry
''''

Public Sub writeToRegistry()
    checkExistKeys
    Dim Key As Variant
    For Each Key In Keys
        If Settings.Exists(Key) Then
            SaveSetting "OnlineImageAnalysis", "macro", Key, Settings.item(Key)
        End If
    Next Key
End Sub

''''
'   writeKeyToRegistry()
'       Write a specific value in Key of registry
''''
Public Sub writeKeyToRegistry(Key As Variant, value As Variant)
    Dim Success As Boolean
    checkExistKeys
    
    If checkKeyItem(Key, value) Then
        Settings.item(Key) = value
        SaveSetting "OnlineImageAnalysis", "macro", Key, value
    End If
End Sub



''''
'   ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
'       Read FileName and store key and paramter into Settings
''''
Public Sub ReadFromFile(fileName As String)
    checkExistKeys
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    Dim Entries() As String
    Close
    On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open fileName For Input As iFileNum
    Do While Not EOF(iFileNum)
            Line Input #iFileNum, Fields
            While VBA.Left(Fields, 1) = "%" 'this are comments
                Line Input #iFileNum, Fields
            Wend
            FieldEntries = Split(Fields, " ", 2)
            Settings.Add FieldEntries(0), FieldEntries(1)
    Loop
    Close #iFileNum
    Exit Sub
ErrorHandle:
    MsgBox "Not able to read " & fileName & " for OiaSettings"
End Sub

''''
'   WriteOiaSettingsToFile(Settings As Dictionary, FileName As String)
'   Write Settings to file with FileName
''''
Public Sub writeToFile(fileName As String)
    checkExistKeys
    Dim i As Integer
    Dim iFileNum As Integer
    Dim Key As Variant
    Close
    'On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open fileName For Output As iFileNum
    
    For Each Key In Settings.Keys
        Print #iFileNum, Key & " " & Settings.item(Key)
    Next Key
    Close
ErrorHandle:
End Sub

''''
'   Parse FileName to get name of SettingFile
'   It is assumed that FileName = something_Txxx.lsm
'   ToDO: better parsing
'''
Public Function settingsFileName(fileName As String) As String
    settingsFileName = VBA.Left(fileName, Len(fileName) - 9) & "_oia.txt"
End Function


''''
'   Parse rois from settings
''''
Public Function getRois(Rois() As Roi) As Boolean
    checkExistKeys
    Dim RoiOut() As Roi
    Dim XRois()  As String 'the string containinig all X-positions of a ROI
    Dim YRois()  As String 'the string containinig all Y-positions of a ROI
    Dim X() As String ' string containing X pos of a single ROI
    Dim Y() As String ' String containing y pos of a single ROI
    Dim XD()  As Double 'the double array containinig the X-positions
    Dim YD()  As Double 'the double array containinig the Y-positions
    Dim i As Integer
    Dim iRoi As Integer
    Dim roiType() As String
    Dim roiAim() As String
    If Settings.item("roiType") = "" Then
        Exit Function
    End If
    roiType = Split(Replace(Settings.item("roiType"), " ", ""), ";")
    If Settings.item("roiAim") = "" Or Settings.item("roiX") = "" Or Settings.item("roiY") = "" Then
        LogManager.UpdateErrorLog "OnlineIASettings.getRois: For each roi you need to define roiType, roiAim, roiX, and roiY! Different rois are separated by semicolumns. " + vbCrLf + "roiType1 ; roiType2; etc." & vbCrLf & _
        "roiX1_roi1, roiX2_roi1; roiX1_roi2, roiX2_roi2, roiX3_roi2; etc. Coordinates in pixels"
        Exit Function
    End If
    
    roiAim = Split(Replace(Settings.item("roiAim"), " ", ""), ";")
    XRois() = Split(Replace(Settings.item("roiX"), " ", ""), ";")
    YRois() = Split(Replace(Settings.item("roiY"), " ", ""), ";")
    
    If UBound(roiType) <> UBound(roiAim) And UBound(roiType) <> UBound(XRois) And UBound(roiType) <> UBound(YRois) Then
       LogManager.UpdateErrorLog "OnlineIASettings.getRois: Number of Rois and coordinates need to correspond. Different rois are separated by semicolumns. " + vbCrLf + "roiType1 ; roiType2; etc." & vbCrLf & _
        "roiX1_roi1, roiX2_roi1; roiX1_roi2, roiX2_roi2, roiX3_roi2; etc. Coorindates in pixels"
        Exit Function
    End If
    ReDim Rois(0 To UBound(roiType))
    For iRoi = 0 To UBound(roiType)
        X() = Split(XRois(iRoi), ",")
        Y() = Split(YRois(iRoi), ",")
        If isArrayEmpty(X) Or isArrayEmpty(Y) Then
            MsgBox "GetRoisFromRegistry: No coordinates found in registry foir roix and roiy (in pixel)"
            Exit Function
        End If
        ReDim XD(UBound(X))
        ReDim YD(UBound(Y))
        For i = 0 To UBound(X)
            XD(i) = CDbl(X(i))
            YD(i) = CDbl(Y(i))
        Next i
        Set Rois(iRoi) = New Roi
        Rois(iRoi).setRoi roiType(iRoi), roiAim(iRoi), XD, YD
        If Not Rois(iRoi).roiConsistency Then
            Exit Function
        End If
    Next iRoi
    getRois = True
End Function


'''
'   getPositions(stgPos() As Vector, currentPosition As Vector) As Boolean
'   Read position in pixel from registry
'   stgPos containes the coordinates
'   stgPos.X, stgPos.Y: defined 0,0 at upper left corner
'   stgPos.Z: defined 0 for first slice
'   units are converted afterwards as it depends on the type of Job (default input unit is px)
'   currentPosition refers to center of image (changes depending on size of image)
''''
Public Function getPositions(stgPos() As Vector, currentPosition As Vector) As Boolean
    checkExistKeys
    ' store postion from windows registry in array
    Dim locX()  As String 'the string containinig the X-positions
    Dim locY()  As String 'the string containinig the Y-positions
    Dim locZ() As String  'the string containinig the Z-positions
    Dim i As Integer
    
    If Settings.Exists("X") Then
        locX() = Split(Settings.item("X"), ";")
    End If
    
    If isArrayEmpty(locX) Then
        Exit Function
    End If
    
    If Settings.Exists("Y") Then
        locY() = Split(Settings.item("Y"), ";")
    End If
    
    If isArrayEmpty(locY) Then
        Exit Function
    End If
    
    If UBound(locX) <> UBound(locY) Then
        LogManager.UpdateErrorLog "OnlineIASettings.getPositions: nr of values in registry for X, Y are not the same, separate the values with semicolumn " _
        & vbCrLf & "X=X1; X2; ..." & vbCrLf & "Y = Y1; Y2; Y3...."
        Exit Function
    End If
    
    If Settings.Exists("Z") Then
        locZ() = Split(Settings.item("Z"), ";")
    End If
    
    If isArrayEmpty(locZ) Then 'Z has not been set. We use the currentvalue as default
        ReDim locZ(UBound(locX))
        For i = 0 To UBound(locX)
            locZ(i) = currentPosition.Z
        Next i
    End If
        
    If UBound(locZ) <> UBound(locX) Then 'Z has not been set for all positions
        LogManager.UpdateErrorLog "OnlineIASettings.getPositions: nr of values in registry for X, Y, Z are not the same, separate the values with semicolumn " _
        & vbCrLf & "X=X1; X2; ... " & vbCrLf & "Y = Y1; Y2; Y3.... " & vbCrLf & "Z = Z1; Z2; Z3.... "
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
    ReDim stgPos(0 To UBound(locX))
    For i = 0 To UBound(locX)
        stgPos(i).X = CDbl(locX(i))
        stgPos(i).Y = CDbl(locY(i))
        stgPos(i).Z = CDbl(locZ(i))
    Next i
    getPositions = True
End Function

'''
'   getFcsPositions(StgPos() As Vector) As Boolean
'   fcsPos array where to load fcsPositions and currentPosition: current stagePosition
'   fcsPos.X, fcsPos.Y: defined 0,0 at upper left corner
'   fcsPos.Z: defined 0 for bottom slice
'   currentPosition should refer to center of the image (this depends on size X, Y and Z for specific job)
'   units are now in pixels. For Fcs these needs to be converted and placed with respect to center of image!
''''
Public Function getFcsPositions(fcsPos() As Vector, currentPosition As Vector) As Boolean
    checkExistKeys
    ' store postion from windows registry in array
    Dim locX()  As String 'the string containinig the X-positions
    Dim locY()  As String 'the string containinig the Y-positions
    Dim locZ() As String  'the string containinig the Z-positions
    Dim i As Integer
    
    If Settings.Exists("fcsX") Then
        locX() = Split(Settings.item("fcsX"), ";")
    End If
    
    If isArrayEmpty(locX) Then
        Exit Function
    End If
    
    If Settings.Exists("fcsY") Then
        locY() = Split(Settings.item("fcsY"), ";")
    End If
    
    If isArrayEmpty(locY) Then
        Exit Function
    End If
    
    If UBound(locX) <> UBound(locY) Then
        LogManager.UpdateErrorLog "OnlineIASettings.getFcsPositions: nr of values in registry for fcsX, fcsY are not the same, separate the values with semicolumn " _
        & vbCrLf & "fcsX=X1; X2; ..." & vbCrLf & "Y = Y1; Y2; Y3...."
        Exit Function
    End If
    
    If Settings.Exists("fcsZ") Then
        locZ() = Split(Settings.item("fcsZ"), ";")
    End If
    
    If isArrayEmpty(locZ) Then 'Z has not been set. We use the currentvalue as default
        ReDim locZ(UBound(locX))
        For i = 0 To UBound(locX)
            locZ(i) = currentPosition.Z
        Next i
    End If
        
    If UBound(locZ) <> UBound(locX) Then 'Z has not been set for all positions
        LogManager.UpdateErrorLog "OnlineIASettings.getFcsPositions: nr of values in registry for fcsX, fcsY, fcsZ are not the same, separate the values with semicolumn " _
        & vbCrLf & "fcsX=X1; X2; ... " & vbCrLf & "fcsY = Y1; Y2; Y3.... " & vbCrLf & "fcsZ = Z1; Z2; Z3.... "
        Exit Function
    End If
    
    ' Convert values to Double
    ReDim stgPos(0 To UBound(locX))
    For i = 0 To UBound(locX)
        fcsPos(i).X = CDbl(locX(i))
        fcsPos(i).Y = CDbl(locY(i))
        fcsPos(i).Z = CDbl(locZ(i))
    Next i
    getFcsPositions = True
End Function



