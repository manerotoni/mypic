Attribute VB_Name = "TestCode"
Option Explicit

Private Sub ChangeAutosaveSettingsImaging()
'    '**************************************
'    'Recorded: 02/09/2013
'    'Description:
'    '**************************************
'    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
'    Set ZEN = Application.ApplicationInterface
'    Dim ZENopt  As Zeiss_Micro_AIM_ApplicationInterface.Options
'    Set ZENopt = ZEN.gui.Fcs.Options
'    ZEN.SetString "Scan.AutoSaveDiretory", "C:\FolderName\Test2"
'
'    'ZEN.SetListStringSelected "Scan.AutoSaveDirectory", "C:\FolderName\Test\"
'    'ZEN.SetString "Scan.AutoSaveDir", "C:\FolderName\Test\"
'    ZEN.SetString "Scan.AutoSaveName", "Tmp2"
'    Debug.Print ZENopt.AutoSaveDirectory.Value
'    ZEN.SetSelected "Scan.AutoSaveDimensions:0:Sub-directory", True
'    ZEN.SetSelected "Scan.AutoSaveDimensions:0:SeparateFiles", False
'    ZEN.SetListEntrySelected "Scan.AutoSaveFormat", 1
''    Lsm5.Options.AutoSaveBaseName = "Tmp"
''    Lsm5.Options.AutoSaveDatabaseName = "C:\FolderName\Tmp"
''    Debug.Print "Hello " & Lsm5.Options.AutoSaveBaseName
''    Debug.Print "Hello2 " & Lsm5.Options.AutoSaveDatabaseName
'    'ZEN.SetSelected "Scan.AutoSaveDimensions:1:SeparateFiles", False
'    ZEN.SetSelected "Scan.AutoSaveDimensions:0:Sub-directory", True
'    ZEN.SetSelected "Scan.AutoSaveDimensions:0:SeparateFiles", False
    
End Sub

Private Sub test1()
    Debug.Print Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing
    Debug.Print Lsm5.DsRecording.SampleSpacing
    'ScanToImage Lsm5.DsRecordingActiveDocObject
    'Set Lsm5.DsRecordingActiveDocObject = Lsm5.StartScan
'    Set viewerGuiServer = Lsm5.viewerGuiServer
End Sub

Private Sub test()
  
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.Scancontroller.AcquisitionRegions
    
    Debug.Print vo.GetNumberElements
    Debug.Print vo.ElementAcquisitionFlags(0)
    Debug.Print "number of Knots " & vo.GetElementNumberKnots(0)
    '' use Ctrl+G to display immediate window
    'Debug.Print "A Line"
    'WScript.StdOut
End Sub
'Sub ExportVBAFiles()
'  Dim pVBAProject As Lsm5.L
'  Dim vbComp As VBComponent  'VBA module, form, etc...
'  Dim strDocPath As String   'Current document path
'  Dim strSavePath As String  'Path to save the exported files to
'
'  ' strSavePath will be the pathname of the document with a _VBACode suffix
'  ' If you want to export the code for Normal instead, change the following
'  ' line to:
'  ' strDocPath = Application.Templates.Item(0)
'  strDocPath = Application.Templates.Item(Application.Templates.Count - 1)
'
'  strSavePath = Left(strDocPath, Len(strDocPath) - 4)
'  strSavePath = strSavePath & "_VBACode"
'
'  ' If this folder doesn't exist, create it
'  If Dir(strSavePath, vbDirectory) = "" Then
'    MkDir strSavePath
'  End If
'
'  ' Get the VBA project
'  ' If you want to export code for Normal instead, paste this macro into
'  ' ThisDocument in the Normal VBA project and change the following line to:
'  ' Set pVBAProject = ThisDocument.VBProject
'  Set pVBAProject = Application.Document.VBProject
'
'  ' Loop through all the components (modules, forms, etc) in the VBA project
'  For Each vbComp In pVBAProject.VBComponents
'    Select Case vbComp.Type
'    Case vbext_ct_StdModule
'      vbComp.Export strSavePath & "\" & vbComp.name & ".bas"
'    Case vbext_ct_Document, vbext_ct_ClassModule
'      ' ThisDocument and class modules
'      vbComp.Export strSavePath & "\" & vbComp.name & ".cls"
'    Case vbext_ct_MSForm
'      vbComp.Export strSavePath & "\" & vbComp.name & ".frm"
'    Case Else
'      vbComp.Export strSavePath & "\" & vbComp.name
'    End Select
'  Next
'    MsgBox "VBA files have been exported to: " & strSavePath
'End Sub
'

'''''
'   Test CODE
'   DisplayAmplifierDescriptions()
'''''
Private Sub DisplayAmplifierDescriptions()
    Dim Track As DsTrack
    Dim Success As Integer
  '  Dim amp As CpAmplifiers
 '   Set amp = Lsm5.Hardware.CpAmplifiers
    
'    Lsm5.Hardware.CpAmplifiers.Summary
        
    'MsgBox "Amp:" + Lsm5.Hardware.CpAmplifiers.name + CStr(Lsm5.Hardware.CpAmplifiers.Summary)
    
    Dim channel As DsDetectionChannel
    
    Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, Success)
    Set channel = Track.DetectionChannelObjectByIndex(0, Success)

    channel.DetectorGain = 300
    MsgBox "Detector 0: " + CStr(channel.Name) + " " + CStr(channel.DetectorGain)
    channel.DetectorGain = 500
    MsgBox "Detector 0: " + CStr(channel.Name) + " " + CStr(channel.DetectorGain)
                        
    
    'If Track.Acquire Then 'if track is activated for acquisition
    '    For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
    '                Set Channel = Track.DetectionChannelObjectByIndex(c - 1, success)
    '                If Channel.Acquire Then 'if channel is activated
    'MsgBox "Det: " + CStr(Lsm5.DsRecording.DetectionChannelOfActiveOrder.name)
    
    'Set channel = Lsm5.Hardware.
    
    'MsgBox "Amp:" + Lsm5.DsDetectionChannel.name
    
    'If (Lsm5.Hardware.CpPmts.Select(1)) Then
    '    MsgBox "Amp:" + CStr(Lsm5.Hardware.CpPmts.DetectorType) + " " + CStr(Lsm5.Hardware.CpPmts.DetectorType)
    'End If
End Sub
Private Sub setPositions()
    '**************************************
    'Recorded: 24/10/2013
    'Description:
    '**************************************
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
     
    ZEN.gui.Acquisition.Positions.currentPosition.ByIndex = 0

    ZEN.gui.Acquisition.Positions.currentPosition.ByIndex = 1

    ZEN.gui.Acquisition.Positions.Remove.Execute

    ZEN.gui.Acquisition.Positions.Add.Execute
    Dim Positions As Long
    Positions = Lsm5.DsRecording.MultiPositionArraySize
End Sub