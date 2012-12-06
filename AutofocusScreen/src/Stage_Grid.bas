Attribute VB_Name = "Stage_Grid"
Public GlobalPositionsStage As Long 'is used
Public GlobalXpos() As Double ' not really used check
Public GlobalYpos() As Double ' not really used check
Public GlobalZpos() As Double ' not really used check
Public GlobalLocationsName() As String  ' not really used check
Public GlobalLocationsNameOld() As String ' not really used check
Public GlobalZposOld() As Double ' not really used check
Public GlobalXposOld() As Double ' not really used check
Public GlobalYposOld() As Double ' not really used check
Public dsDoc As DsRecordingDoc
Public Stage As CpStages ' has been defined twice
Public GettingZmap As Boolean ' not really used
Public idpos As Long

Public x As Long ' does it make sense to have it as global variable?
Public y As Long ' does it make sense to have it as global variable?
    



Public GlobalDeActivatedLocations() As Boolean
Public GlobalLocationsOrder() As Long
Public GlobalLocationsOrderOld() As Long



Public Const ZBacklash = -50 'Has to do with the movements of the focus wheel that are "better" if they are long enough.


Public Sub MakeBlankImage(DestImage As DsRecordingDoc, _
BitsPerSample As Long, bpp As Long, Visible As Boolean, _
ImgName As String, TimeSeries As Boolean, NumberScans As Long, XPixels As Long, YPixels As Long, Channels As Long)
Dim i As Long

Dim Success As Integer
Dim DataChannel As DsDataChannel
Dim DataChannelIndex As Long
Dim Track As DsTrack
Dim TrackIndex As Long
Dim lpReOpenBuff As OFSTRUCT
Dim lpRootPathName As String
Dim lpSectorsPerCluster As Long
Dim lpBytesPerSector As Long
Dim lpNumberOfFreeClusters As Long
Dim lpTotalNumberOfClusters As Long
Dim lSpace As Long
Dim lFreeSpace As Double
Dim fSize As Double
Dim hFile As Long
Dim zIndex As Long
Dim TimeIndex As Long
Dim channel As Long
Dim SourceChannel As Long
Dim NumberChannels As Long
Dim DestStackNumber As Long
Dim TimeStampIndex As Long
Dim indxArr() As Long
Dim NumberOfSelected As Long
Dim NumberOfStacks As Long

'Dim Channels() As String
Dim ReturnValue As Boolean
Dim OK As Boolean
Dim scnline As Variant
Dim spl As Long
Dim Tnum As Long
Dim newTime As Double
Dim myDate As Date
Dim myDate1 As Date
Dim newTime1 As Double
Dim myTime As Date
Dim OldImage As Object
Dim ImageType As Long   'ImageType=1 Non Lambda Stack, ImageType=2 Lambda Stack
                
    If TimeSeries Then
        Set DestImage = Lsm5.MakeNewImageDocument(XPixels, _
        YPixels, 1, NumberScans, _
        Channels, bpp, Visible)
    Else
        Set DestImage = Lsm5.MakeNewImageDocument(XPixels, _
        YPixels, 1, 1, _
        Channels, bpp, Visible)
    End If
    If (DestImage Is Nothing) Then
        MsgBox "Cannot Create New Window!", VbExclamation
        Exit Sub
    End If
    If TimeSeries Then
        DestImage.Recording.TimeSeries = True
    Else
        DestImage.Recording.TimeSeries = False
    End If

    DestImage.SetTitle ImgName
            
    
Finish:

End Sub


Public Sub ReadLoc(x As Double, y As Double)
    Dim cnt As Long
    
    cnt = 0
    On Error GoTo retry
retry:
    If cnt > 1000 Then GoTo Finish
    cnt = cnt + 1
    x = Lsm5.Hardware.CpStages.PositionX
    y = Lsm5.Hardware.CpStages.PositionY
Finish:
End Sub

Public Sub CoordinateConversion(bExchangeXY As Boolean, bMirrorX As Boolean, bMirrorY As Boolean)
    Dim bLSM As Boolean
    Dim bLIVE As Boolean
    Dim bCamera As Boolean
    Dim lsystem As Long
'    If GlobalSystemVersion = 32 Then
'        Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisState bExchangeXY, bMirrorX, bMirrorY
'    ElseIf GlobalSystemVersion > 32 Then
        UsedDevices40 bLSM, bLIVE, bCamera
        If bLSM Then
            lsystem = 0
        ElseIf bLIVE Then
            lsystem = 1
        ElseIf bCamera Then
            lsystem = 3
        End If
'    End If
    Lsm5.ExternalCpObject.pHardwareObjects.GetImageAxisStateS lsystem, bExchangeXY, bMirrorX, bMirrorY

End Sub

'''''
'   UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
'   Ask which system is the macro runnning on
'       [bLSM]  In/Out - True if LSM system
'       [bLive] In/Out - True for LIVE system
'       [bCamera] In/Out - True if Camera is used
''''
Public Sub UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
    Dim Scancontroller As AimScanController
    Dim TrackParameters As AimTrackParameters
    Dim Size As Long
    Dim lTrack As Long
    Dim eDeviceMode As Long
    
    bLSM = False
    bLIVE = False
    bCamera = False
    Set Scancontroller = Lsm5.ExternalDsObject.Scancontroller
    Set TrackParameters = Scancontroller.TrackParameters
    If TrackParameters Is Nothing Then Exit Sub
    Size = TrackParameters.GetTrackArraySize
    For lTrack = 0 To Size - 1
            eDeviceMode = TrackParameters.TrackDeviceMode(lTrack)
            Select Case eDeviceMode
                Case eAimDeviceModeLSM
                    bLSM = True
                   
                Case eAimDeviceModeLSM_ChannelMode
                    bLSM = True
                   
                Case eAimDeviceModeLSM_NDD
                    bLSM = True
                    
                Case eAimDeviceModeLSM_DD
                    bLSM = True
                   
                Case eAimDeviceModeSpectralImager
                    bLSM = True
                    Exit Sub
                
                Case eAimDeviceModeRtScanner
                    bLIVE = True
                    Exit Sub
                
                Case eAimDeviceModeCamera1
                    bCamera = True
                    Exit Sub
                
            End Select
    Next lTrack
End Sub
