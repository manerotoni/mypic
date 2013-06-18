Attribute VB_Name = "OnlineImageAnalysisIO"
Public Sub testOiaSettings()
    Dim Settings As OnlineIASettings
    Set Settings = New OnlineIASettings
    
End Sub
''''
'' A set of functions to IN/OUT parameters for Onlineimageanalisys this should be a class!!!
'''''
'
'''''
'' Set Registry to default values
''''
'Public Sub SetRegistryKeyNames(Keys() As String)
''This variable contains the keys for the OnlineImageanalysis
'    ReDim Keys(11)
'    Keys(0) = "code"
'    Keys(1) = "fileAnalyzed"
'    Keys(2) = "filePath"
'    Keys(3) = "X"
'    Keys(4) = "Y"
'    Keys(5) = "Z"
'    Keys(6) = "deltaZ"
'    Keys(7) = "roiType"
'    Keys(8) = "roiAim"
'    Keys(9) = "roiX"
'    Keys(10) = "roiY"
'    Keys(11) = "unit"
'End Sub
'
''''
'' Default registry value is always empty
''''
'Public Sub resetRegistry(Keys() As String)
'    Dim i As Integer
'    For i = LBound(Keys) To UBound(Keys)
'        SaveSetting "OnlineImageAnalysis", "macro", Keys(i), ""
'    Next i
'End Sub
'
'''''
''   ReadOiaSettingsFromRegistry(Settings As Collection, Keys() As String)
''       Read Registry using the keys stored in Keys and create a new dictionary
'''''
'Public Sub ReadOiaSettingsFromRegistry(Settings As Dictionary, Keys() As String)
'    Set Settings = New Dictionary
'    Dim key As Variant
'    For Each key In Keys
'        Settings.Add key, GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:=key)
'    Next key
'End Sub
'
'''''
''   WriteOiaSettingsToRegistry(Settings As Dictionary, Keys() As String)
''       Write settings with keys defined in Keys to Registry
'''''
'
'Public Sub writeOiaSettingsToRegistry(Settings As Dictionary, Keys() As String)
'    Dim key As Variant
'    For Each key In Keys
'        If Settings.Exists(key) Then
'            SaveSetting "OnlineImageAnalysis", "macro", key, Settings.Item(key)
'        End If
'    Next key
'End Sub
'
'''''
''   ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
''       Read FileName and store key and paramter into Settings
'''''
'Public Sub ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
'    Set Settings = New Dictionary
'    Dim iFileNum As Integer
'    Dim Fields As String
'    Dim FieldEntries() As String
'    Dim Entries() As String
'    Close
'    On Error GoTo ErrorHandle
'    iFileNum = FreeFile()
'    Open FileName For Input As iFileNum
'    Do While Not EOF(iFileNum)
'            Line Input #iFileNum, Fields
'            While Left(Fields, 1) = "%" 'this are comments
'                Line Input #iFileNum, Fields
'            Wend
'            FieldEntries = Split(Fields, " ", 2)
'            Settings.Add FieldEntries(0), FieldEntries(1)
'    Loop
'    Close #iFileNum
'    Exit Sub
'ErrorHandle:
'    MsgBox "Not able to read " & FileName & " for OiaSettings"
'End Sub
'
'''''
''   WriteOiaSettingsToFile(Settings As Dictionary, FileName As String)
''   Write Settings to file with FileName
'''''
'Public Sub WriteOiaSettingsToFile(Settings As Dictionary, FileName As String)
'    Dim i As Integer
'    Dim iFileNum As Integer
'    Dim key As Variant
'    Close
'    'On Error GoTo ErrorHandle
'    iFileNum = FreeFile()
'    Open FileName For Output As iFileNum
'
'    For Each key In Settings.Keys
'        Print #iFileNum, key & " " & Settings.Item(key)
'    Next key
'    Close
'ErrorHandle:
'End Sub
'
'''''
''   Parse FileName to get name of SettingFile
''   It is assumed that FileName = something_Txxx.lsm
''   ToDO: better parsing
''''
'Public Function OiaSettingsFileName(FileName As String) As String
'    OiaSettingFileName = Left(FileName, Len(FileName) - 9) & "_oia.txt"
'End Function
'
'
'''''
''   Parse rois from settings
''   This should be put in the Jobdefinition as a Roi is associated to a Job
'''''
'Public Function GetRoisFromSettings(OiaSettings As Dictionary) As Roi()
'    Dim RoiOut() As Roi
'    Dim XRois()  As String 'the string containinig all X-positions of a ROI
'    Dim YRois()  As String 'the string containinig all Y-positions of a ROI
'    Dim X() As String ' string containing X pos of a single ROI
'    Dim Y() As String ' String containing y pos of a single ROI
'    Dim XD()  As Double 'the double array containinig the X-positions
'    Dim YD()  As Double 'the double array containinig the Y-positions
'    Dim i As Integer
'    Dim iRoi As Integer
'    If OiaSettings.Item("roiType") = "" Then
'        Exit Function
'    End If
'
'    roiType = Split(OiaSettings.Item("roiType"), ";")
'
'    If OiaSettings.Item("roiAim") = "" Or OiaSettings.Item("roiX") = "" Or OiaSettings.Item("roiY") = "" Then
'        MsgBox ("GetRoisFromRegistry: For each roi you need to define roiType, roiAim, roiX, and roiY!" + vbCrLf + "roiType1 ; roiType2; etc." & vbCrLf & _
'        "roiX1_roi1, roiX2_roi1; roiX1_roi2, roiX2_roi2, roiX3_roi2; etc. Coordinates in pixels")
'        Exit Function
'    End If
'
'    roiAim = Split(OiaSettings.Item("roiAim"), ";")
'    XRois() = Split(OiaSettings.Item("roiX"), ";")
'    YRois() = Split(OiaSettings.Item("roiY"), ";")
'
'    If UBound(roiType) <> UBound(roiAim) And UBound(roiType) <> UBound(XRoi) And UBound(roiType) <> UBound(YRoi) Then
'        MsgBox ("GetRoisFromRegistry: Number of Rois and coordinates need to correspond." + vbCrLf + "roiType1 ; roiType2; etc." & vbCrLf & _
'        "roiX1_roi1, roiX2_roi1; roiX1_roi2, roiX2_roi2, roiX3_roi2; etc. Coorindates in pixels")
'        Exit Function
'    End If
'    ReDim RoiOut(0 To UBound(roiType))
'    For iRoi = 0 To UBound(roiType)
'        X() = Split(XRois(iRoi), ",")
'        Y() = Split(YRois(iRoi), ",")
'        If isArrayEmpty(X) Or isArrayEmpty(Y) Then
'            MsgBox "GetRoisFromRegistry: No coordinates found in registry foir roix and roiy (in pixel)"
'            Exit Function
'        End If
'        ReDim XD(UBound(X))
'        ReDim YD(UBound(Y))
'        For i = 0 To UBound(X)
'            XD(i) = CDbl(X(i))
'            YD(i) = CDbl(Y(i))
'        Next i
'        RoiOut(iRoi).setRoi roiType(iRoi), roiAim(iRoi), CDbl(X), CDbl(YK)
'        If Not RoiOut(iRoi).roiConsitency Then
'            Exit Function
'        End If
'    Next iRoi
'    GetRoisFromSettings = RoiOut
'End Function
'
'
''''
''   GetPositionsFromSettings(OiaSettings As Dictionary, X() As Double, Y() As Double, Z() As Double, deltaZ() As Integer) As Boolean
''   OiaSettings contains all settings for OnlineImageanalysis (see OiaKeys)
''   StgPos containes the coordinates
''   StgPos.X, StgPos.Y: defined 0,0 at upper left corner
''   StgPos.Z: defined 0 at central slice
''   units are converted afterwards as it depends on the type of Job (default unit is px)
'''''
'Public Function GetPositionsFromSettings(OiaSettings As Dictionary, StgPos() As Vector) As Boolean
'
'    ' store postion from windows registry in array
'    Dim locX()  As String 'the string containinig the X-positions
'    Dim locY()  As String 'the string containinig the Y-positions
'    Dim locZ() As String  'the string containinig the Z-positions
'    Dim i As Integer
''    If OiaSettings.Exists("unit") Then
''        If OiaSettings.Item("unit") = "um" Or OiaSettings.Item("unit") = Chr(181) & "m" Then      'has correct pixelSize of um
''            pixelSizeXY = 1
''            pixelSizeZ = 1
''        ElseIf OiaSettings.Item("unit") = "px" Or OiaSettings.Item("unit") = "" Then
''            pixelSizeXY = Jobs.GetSampleSpacing(JobName) * 1000000
''            pixelSizeZ = Jobs.GetFrameSpacing(JobName)
''        Else
''            MsgBox "GetPositionsFromSettings: Do not understand unit " & OiaSettings.Item("unit") & ". Possible value for registry entry unit are um or px!"
''            Exit Function
''        End If
''    Else
''        pixelSizeXY = Jobs.GetSampleSpacing(JobName) * 1000000
''        pixelSizeZ = Jobs.GetFrameSpacing(JobName)
''    End If
'
'    If OiaSettings.Exists("X") Then
'        locX() = Split(OiaSettings.Item("X"), ",")
'    End If
'
'    If isArrayEmpty(locX) Then
'        Exit Function
'    End If
'
'    If OiaSettings.Exists("Y") Then
'        locY() = Split(OiaSettings.Item("Y"), ",")
'    End If
'
'    If isArrayEmpty(locY) Then
'        Exit Function
'    End If
'
'    If UBound(locX) <> UBound(locY) Then
'        MsgBox ("StorePositionsFromRegistry: nr of values in registry for X, Y are not the same, separate the values with comma!")
'        Exit Function
'    End If
'
'    If OiaSettings.Exists("Z") Then
'        locZ() = Split(OiaSettings.Item("Z"), ",")
'    End If
'
'    If isArrayEmpty(locZ) Then 'ZOffset has not been set. We use a default values
'        ReDim locZ(UBound(locX))
'        For i = 0 To UBound(locX)
'            locZ(i) = 0
'        Next i
'    End If
'
'    If UBound(locZ) <> UBound(locX) Then 'Z has not been set for all positions
'        MsgBox ("StorePositionsFromRegistry: nr of values in registry for offsetX, offsetZ are not the same, separate the values with comma!")
'        Exit Function
'    End If
'
'    If OiaSettings.Exists("deltaZ") Then
'        locDeltaZ() = Split(OiaSettings.Item("deltaZ"), ",")
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
'
'    ' Convert values to Double
'    ReDim StgPos(0 To UBound(locX))
'    For i = 0 To UBound(locX)
'        StgPos(i).X = CDbl(locX(i))
'        StgPos(i).Y = CDbl(locY(i))
'        StgPos(i).Z = CDbl(locZ(i))
'    Next i
'    GetPositionsFromSettings = True
'End Function
'
