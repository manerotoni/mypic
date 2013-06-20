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
''keys used for codeIn
Private codeIn As Dictionary
''keys used for codeOut
Private codeOut As Dictionary

'''
' A set of functions to IN/OUT parameters for Onlineimageanalisys this should be a class!!!
''''

''''
' Set Registry to default values
'''
Public Sub initializeDefault()
'This variable contains the keys for the OnlineImageanalysis
    ReDim Keys(0 To 11)
    Keys(0) = "codeIn"
    Keys(1) = "codeOut"
    Keys(2) = "fileAnalyzed"
    Keys(3) = "filePath"
    Keys(4) = "X"
    Keys(5) = "Y"
    Keys(6) = "Z"
    Keys(7) = "deltaZ"
    Keys(8) = "roiType"
    Keys(9) = "roiAim"
    Keys(10) = "roiX"
    Keys(11) = "roiY"
    
    Set codeIn = New Dictionary
    codeIn.Add "wait", "wait for image analysis"
    codeIn.Add "", "(empty char) wait for image analysis"
    codeIn.Add "error", "error from image analysis"
    codeIn.Add "timeExpired", "time for image analysis has expired"
    codeIn.Add "nothing", "Do nothing"
    codeIn.Add "focus", "read X, Y, and Z  and change focus"
    codeIn.Add "Trigger1", "perform job Trigger1 at position X, Y, and Z"
    codeIn.Add "Trigger2", "perform job Trigger2 at position X, Y, and Z"
    
    Set codeOut = New Dictionary
    codeOut.Add "", "(empty char) online image analysis does nothing"
    codeOut.Add "newImage", "a new image is there"
    
    Set Settings = New Dictionary
    Dim Key As Variant
    For Each Key In Keys
        Settings.Add Key, ""
    Next Key
End Sub

Public Function getSettings(Key As Variant) As Variant
    Dim Log As String
    Dim i  As Integer
    If Settings.Exists(Key) Then
        getSettings = Settings.Item(Key)
        Exit Function
    End If
    Log = "OnlineIASettings.getSettings: Failed to get " + Key + " from OiaSettings. Possible values are:" & vbCrLf
    Log = Log & availableKeys
    
    MsgBox Log
    'ErrorLog.UpdateLog (Log)
End Function

Public Function availableKeys() As String
    Dim i As Integer
    Dim Log As String
    Log = ""
    For i = 0 To UBound(Keys)
        Log = Log & Keys(i) & ", "
    Next i
    availableKeys = Log
End Function

Public Function availableCodeIn() As String
    Dim codeKeys() As Variant
    Dim Key As Variant
    codeKeys = codeIn.Keys
    Dim Log As String
    Log = ""
    For Each Key In codeKeys
       Log = Log & Key & ", "
    Next Key
    availableCodeIn = Log
End Function

Public Function availableCodeOut() As String
    Dim codeKeys() As Variant
    Dim Key As Variant
    codeKeys = codeOut.Keys
    Dim Log As String
    Log = ""
    For Each Key In codeKeys
       Log = Log & Key & ", "
    Next Key
    availableCodeOut = Log
End Function

Public Function checkCodeOut() As Boolean
    Dim errorMsg As String
    Dim Key As Variant
    Dim codeKeys() As Variant
    codeKeys = codeOut.Keys
    For Each Key In codeKeys
        If Me.getSettings("codeOut") = Key Then
            checkCodeOut = True
            Exit Function
        End If
    Next Key
    errorMsg = "Wrong codeOut = " & Me.getSettings("codeOut") & " in Registry. Values are " & vbCrLf
    For Each Key In codeKeys
        errorMsg = errorMsg & Key & ": " & codeOut.Item(Key) & vbCrLf
    Next Key
    MsgBox errorMsg
End Function



Public Function checkCodeIn() As Boolean
    Dim errorMsg As String
    Dim Key As Variant
    Dim codeInKeys() As Variant
    codeInKeys = codeIn.Keys
    For Each Key In codeInKeys
        If Settings.Item("codeIn") = Key Then
            checkCodeIn = True
            Exit Function
        End If
    Next Key
    errorMsg = "Wrong codeIn= " & Settings.Item("codeIn") & " in Registry values are " & vbCrLf
    For Each Key In codeInKeys
        errorMsg = errorMsg & Key & ": " & codeIn.Item(Key) & vbCrLf
    Next Key
    MsgBox errorMsg
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

Private Sub checkKeys()
    If isArrayEmpty(Keys) Then
        initializeDefault
    End If
End Sub




'''
' Default registry values is always empty
'''

Public Sub resetRegistry()
    checkKeys
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
    checkKeys
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
'   WriteOiaSettingsToRegistry(Settings As Dictionary, Keys() As String)
'       Write settings with keys defined in Keys to Registry
''''

Public Sub writeToRegistry()
    checkKeys
    Dim Key As Variant
    For Each Key In Keys
        If Settings.Exists(Key) Then
            SaveSetting "OnlineImageAnalysis", "macro", Key, Settings.Item(Key)
        End If
    Next Key
    
End Sub

Public Sub writeKeyToRegistry(Key As Variant, Value As Variant)
    checkKeys
    If Settings.Exists(Key) Then
        If Key = "codeIn" Then
            checkCodeIn
        End If
        If Key = "codeOut" Then
            checkCodeOut
        End If
        SaveSetting "OnlineImageAnalysis", "macro", Key, Value
    Else
        MsgBox "OnlineIASettings.writeKeyToRegistry: no key " & Key & " available. Possible Keys are:" & vbCrLf & _
        availableKeys
    End If
End Sub

''''
'   ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
'       Read FileName and store key and paramter into Settings
''''
Public Sub readFromFile(FileName As String)
    checkKeys
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
    checkKeys
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
    checkKeys
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
Public Function getPositions(StgPos() As Vector) As Boolean
    checkKeys
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
        Exit Function
    End If
    
    If Settings.Exists("Z") Then
        locZ() = Split(Settings.Item("Z"), ",")
    End If
    
    If isArrayEmpty(locZ) Then 'ZOffset has not been set. We use a default values
        ReDim locZ(UBound(locX))
        For i = 0 To UBound(locX)
            locZ(i) = 0
        Next i
    End If
        
    If UBound(locZ) <> UBound(locX) Then 'Z has not been set for all positions
        MsgBox ("OnlineIASettings.getPositions: nr of values in registry for offsetX, offsetZ are not the same, separate the values with comma!")
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
        StgPos(i).X = CDbl(locX(i))
        StgPos(i).Y = CDbl(locY(i))
        StgPos(i).Z = CDbl(locZ(i))
    Next i
    getPositions = True
End Function



