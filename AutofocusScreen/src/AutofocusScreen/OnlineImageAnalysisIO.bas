Attribute VB_Name = "OnlineImageAnalysisIO"
'''
' A set of functions to IN/OUT parameters for Onlineimageanalisi
''''

''''
' Set Registry to default values
'''
Public Sub ResetRegistry()
    SaveSetting "OnlineImageAnalysis", "macro", "fileAnalyzed", ""
    SaveSetting "OnlineImageAnalysis", "macro", "filePath", ""
    SaveSetting "OnlineImageAnalysis", "macro", "code", "nothing"
    SaveSetting "OnlineImageAnalysis", "macro", "X", ""
    SaveSetting "OnlineImageAnalysis", "macro", "Y", ""
    SaveSetting "OnlineImageAnalysis", "macro", "Z", ""
    SaveSetting "OnlineImageAnalysis", "macro", "deltaZ", ""
    SaveSetting "OnlineImageAnalysis", "macro", "unit", "px"
    SaveSetting "OnlineImageAnalysis", "macro", "roiType", ""
    SaveSetting "OnlineImageAnalysis", "macro", "roiAim", ""
    SaveSetting "OnlineImageAnalysis", "macro", "roiX", ""
    SaveSetting "OnlineImageAnalysis", "macro", "roiY", ""
End Sub

''''
'   ReadOiaSettingsFromRegistry(Settings As Collection, Keys() As String)
'       Read Registry using the keys stored in Keys and create a new dictionary
''''
Public Sub ReadOiaSettingsFromRegistry(Settings As Dictionary, Keys() As String)
    Set Settings = New Dictionary
    Dim key As Variant
    For Each key In Keys
        Settings.Add key, GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:=key)
    Next key
End Sub

''''
'   WriteOiaSettingsToRegistry(Settings As Dictionary, Keys() As String)
'       Write settings with keys defined in Keys to Registry
''''

Public Sub WriteOiaSettingsToRegistry(Settings As Dictionary, Keys() As String)
    Dim key As Variant
    For Each key In Keys
        If Settings.Exists(key) Then
            SaveSetting "OnlineImageAnalysis", "macro", key, Settings.Item(key)
        End If
    Next key
End Sub

''''
'   ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
'       Read FileName and store key and paramter into Settings
''''
Public Sub ReadOiaSettingsFromFile(Settings As Dictionary, FileName As String)
    Set Settings = New Dictionary
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
Public Sub WriteOiaSettingsToFile(Settings As Dictionary, FileName As String)
    Dim i As Integer
    Dim iFileNum As Integer
    Dim key As Variant
    Close
    'On Error GoTo ErrorHandle
    iFileNum = FreeFile()
    Open FileName For Output As iFileNum
    
    For Each key In Settings.Keys
        Print #iFileNum, key & " " & Settings.Item(key)
    Next key
    Close
ErrorHandle:
End Sub

''''
'   Parse FileName to get name of SettingFile
'   It is assumed that FileName = something_Txxx.lsm
'   ToDO: better parsing
'''
Public Function OiaSettingFileName(FileName As String) As String
    OiaSettingFileName = Left(FileName, Len(FileName) - 9) & "_oia.txt"
End Function


''''
'   Read Rois from registry and create them
'   ToDo: create several roi!!
''''
Public Function GetRoisFromSettings(OiaSettings As Dictionary) As Boolean
    DisplayProgress "CreateRoisFromRegistry", RGB(0, &HC0, 0)
    Dim RoiType As String
    Dim RoiAim As String
    Dim XKnot()  As String 'the string containinig the X-positions
    Dim YKnot()  As String 'the string containinig the Y-positions
    Dim XKnotD()  As Double 'the double array containinig the X-positions
    Dim YKnotD()  As Double 'the double array containinig the Y-positions
    
    Dim i As Integer
    
    XKnot() = Split(GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="roix"), ",")
    YKnot() = Split(GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="roiy"), ",")
    If isArrayEmpty(XKnot) Or isArrayEmpty(YKnot) Then
        MsgBox "CreateRoisFromRegistry: No coordinates found in registry roix and roiy (in pixel)"
        Exit Function
    End If
    
    If UBound(XKnot) <> UBound(YKnot) Then 'Z position as not been set
        MsgBox "CreateRoisFromRegistry: Equal number of coordinates x and y"
        Exit Function
    End If
    
    RoiType = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="roitype")
    If RoiType = "" Then
        MsgBox "CreateRoisFromRegistry: You have to specify the type of roi: circle, rectangle, polyline, or ellipse"
        Exit Function
    End If

    RoiAim = GetSetting(appname:="OnlineImageAnalysis", section:="macro", key:="roiaim")
    If RoiAim = "" Then
        MsgBox "CreateRoisFromRegistry: You have to specify the aim of roi: bleach, acquire, or analyse"
        Exit Function
    End If
    
    
    ReDim XKnotD(UBound(XKnot))
    ReDim YKnotD(UBound(YKnot))
    For i = 0 To UBound(XKnot)
        XKnotD(i) = CDbl(XKnot(i))
        YKnotD(i) = CDbl(YKnot(i))
    Next i
    
    If MakeVectorElement(RoiType, XKnotD, YKnotD, RoiAim) Then
        CreateRoisFromRegistry = True
        Exit Function
    Else
        Exit Function
    End If
End Function


'''
'   StorePositionsFromRegistry(ByVal Xref As Double, ByVal Yref As Double, ByVal Zref As Double, X() As Double, Y() As Double, _
'   Z() As Double, DeltaZ() As Double) As Boolean
''''
Public Function GetPositionsFromSettings(JobName As String, OiaSettings As Dictionary, X() As Double, Y() As Double, Z() As Double, deltaZ() As Integer) As Boolean
    
    ' store postion from windows registry in array
    Dim Xoffset()  As String 'the string containinig the X-positions
    Dim Yoffset()  As String 'the string containinig the Y-positions
    Dim ZOffset() As String  'the string containinig the Z-positions
    Dim locDeltaZ() As String
    Dim defaultDeltaZ As Double
    Dim Xnew As Double
    Dim Ynew As Double
    Dim Znew As Double
    Dim pixelSizeXY As Double
    Dim pixelSizeZ As Double
    Dim unit As String
    Dim LowBound As Integer
    Dim i As Integer
    If OiaSettings.Exists("unit") Then
        If OiaSettings.Item("unit") = "um" Or OiaSettings.Item("unit") = Chr(181) & "m" Then      'has correct pixelSize of um
            pixelSizeXY = 1
            pixelSizeZ = 1
        ElseIf OiaSettings.Item("unit") = "px" Or OiaSettings.Item("unit") = "" Then
            pixelSizeXY = Jobs.GetSampleSpacing(JobName) * 1000000
            pixelSizeZ = Jobs.GetFrameSpacing(JobName)
        Else
            MsgBox "GetPositionsFromSettings: Do not understand unit " & OiaSettings.Item("unit") & ". Possible value for registry entry unit are um or px!"
            Exit Function
        End If
    Else
        pixelSizeXY = Jobs.GetSampleSpacing(JobName) * 1000000
        pixelSizeZ = Jobs.GetFrameSpacing(JobName)
    End If
    
    If OiaSettings.Exists("X") Then
        Xoffset() = Split(OiaSettings.Item("X"), ",")
    End If
    
    If isArrayEmpty(Xoffset) Then
        Exit Function
    End If
    
    If OiaSettings.Exists("Y") Then
        Yoffset() = Split(OiaSettings.Item("Y"), ",")
    End If
    
    If isArrayEmpty(Yoffset) Then
        Exit Function
    End If
    
    If UBound(Xoffset) <> UBound(Yoffset) Then
        MsgBox ("StorePositionsFromRegistry: nr of values in registry for offsetX, offsetY are not the same, separate the values with comma!")
        Exit Function
    End If
    
    If OiaSettings.Exists("Z") Then
        ZOffset() = Split(OiaSettings.Item("Z"), ",")
    End If
    
    If isArrayEmpty(ZOffset) Then 'ZOffset has not been set. We use a default values
        ReDim ZOffset(UBound(Xoffset))
        For i = 0 To UBound(Xoffset)
            ZOffset(i) = 0
        Next i
    End If
        
    If UBound(ZOffset) <> UBound(Xoffset) Then 'Z has not been set for all positions
        MsgBox ("StorePositionsFromRegistry: nr of values in registry for offsetX, offsetZ are not the same, separate the values with comma!")
        Exit Function
    End If
    
    If OiaSettings.Exists("deltaZ") Then
        locDeltaZ() = Split(OiaSettings.Item("deltaZ"), ",")
    End If
    
    If isArrayEmpty(locDeltaZ) Then
        ReDim locDeltaZ(UBound(Xoffset)) 'deltaZ has not been set. We use a default values
        For i = 0 To UBound(locDeltaZ)
            locDeltaZ(i) = -1
        Next i
    End If
    
    If UBound(locDeltaZ) <> UBound(Xoffset) Then 'deltaZ has not been set for all positions
        MsgBox ("StorePositionsFromRegistry: nr of values in registry for deltaz, z are not the same, separate the values with comma!")
        Exit Function
    End If
    

    ' Add values to existing array
    If isArrayEmpty(X) Then
        LowBound = 0
        ReDim X(0 To UBound(Xoffset))
        ReDim Y(0 To UBound(Xoffset))
        ReDim Z(0 To UBound(Xoffset))
        ReDim deltaZ(0 To UBound(Xoffset))
    Else
        LowBound = UBound(X) + 1
        ReDim Preserve X(0 To UBound(X) + UBound(Xoffset))
        ReDim Preserve Y(0 To UBound(Y) + UBound(Yoffset))
        ReDim Preserve Z(0 To UBound(Z) + UBound(ZOffset))
        ReDim Preserve deltaZ(0 To UBound(deltaZ) + UBound(locDeltaZ))
    End If
    
    For i = 0 To UBound(Xoffset)
        Xnew = Xref
        Ynew = Yref
        Znew = Zref
        ComputeShiftedCoordinates CDbl(Xoffset(i)) * pixelSizeXY, CDbl(Yoffset(i)) * pixelSizeXY, CDbl(ZOffset(i)) * pixelSizeZ, Xnew, Ynew, Znew
        X(LowBound + i) = Xnew  ' this needs to be unified with computing internal AF
        Y(LowBound + i) = Ynew  ' needs to be unified with computing internal AF
        Z(LowBound + i) = Znew
        deltaZ(LowBound + i) = CDbl(locDeltaZ(i)) * pixelSizeZ
    Next i
    DisplayProgress "StorePositionsFromRegistry - Position stored", RGB(0, &HC0, 0)
    
End Function

