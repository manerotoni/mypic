Attribute VB_Name = "MicroscopeIO"
''''
' Module with functions for controlling stage, starts and stop scan
'''''

Option Explicit
Option Base 0
Public SystemVersion As String

Public Declare Function GetInputState Lib "user32" () As Long ' Check if mouse or keyboard has been pushed

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function RegOpenKeyEx _
    Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long

Public Declare Function RegCloseKey _
    Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegQueryValueEx _
    Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.



Public imgFileFormat As enumAimExportFormat
Public imgFileExtension As String


Public Const PrecZ = 2                     'precision of Z passed for stage movements i.e. Z = Round(Z, PrecZ)
Public Const PrecXY = 2                    'precision of X and Y passed for stage movements

Public ZBacklash  As Double           'ToDo: is it still recquired?.
                                           'Has to do with the movements of the focus wheel that are "better"
                                            'if they are long enough. For amoment a test did not gave significant differences This is required for ZEN2010
Public ZENv As Integer            'String variable indicating the version of ZEN used 2010 ir 2011 (2012)
Public ZEN As Object             'Object containing Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface (for ZEN > 2011)

''''''''''''''''''''''
'''GLOBAL VARIABLES'''
''''''''''''''''''''''

Public ScanStop As Boolean      'if TRUE current recording is stopped
Public ScanPause As Boolean     'if TRUE current recording is paused
Public Running As Boolean       'TRUE when system is running (e.g. after start)
Public GlobalDataBaseName As String   'Name of output folder
Public TrackNumber As Integer    'number of available tracks

Public OverwriteFiles As Boolean  'if TRUE we do not overwrite files (not active anymore)
Public FocusMapPresent As Boolean 'if TRUE we have a focus map (not active anymore)
Public GlobalCorrectionOffset As Double 'not anymore used

'''
'RecordingDoc used globally for imaging
'''
Public GlobalRecordingDoc As DsRecordingDoc

'''
'FcsData used globally
'''
Public GlobalFcsData As AimFcsData
'''
'RecordingDoc used globally for Fcs
'''
Public GlobalFcsRecordingDoc As DsRecordingDoc


Const PauseGrabbing = 50 'pause for polling the whether scan/fcscontroller are acquiring. A high value makes more errors!


'''
' Returns version number (ZEN2010, etc.)
'''
Public Function getVersionNr() As Integer
    Dim VersionNr As Long
    VersionNr = CLng(Left(Lsm5.Info.VersionIs, 1))
    Select Case VersionNr
        Case 6:
            getVersionNr = 2010
        Case 7:
            getVersionNr = 2011
        Case 8:
            getVersionNr = 2012
    End Select
    
    If VersionNr > 8 Then
        MsgBox "Don't understand the version of ZEN used. Set to ZEN2012"
    End If
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

''''
' Stop all running FCS jobs and imaging jobs
''''
Public Function StopAcquisition()
    Lsm5.StopScan
    If Lsm5.Info.IsFCS Then
        Dim FcsControl As AimFcsController
        Set FcsControl = Fcs
        FcsControl.StopAcquisitionAndWait
    End If
    DoEvents
End Function

'''''
'   ScanToImage ( RecordingDoc As DsRecordingDoc) As Boolean
'   scan overwrite the same image, even with several z-slices
'''''
Public Function ScanToImage(RecordingDoc As DsRecordingDoc) As Boolean
    On Error GoTo ErrorHandle:
    Dim ProgressFifo As IAimProgressFifo ' this shows how far you are with the acquisition image ( the blue bar at the bottom). The usage of it makes the macro quite slow
    Dim AcquisitionController As AimScanController
    Dim treenode As Object
    Dim Time As Double
    'Dim gui As Object
    'Set gui = Lsm5.ViewerGuiServer not recquired anymore
    
    Time = Timer
    If RecordingDoc Is Nothing Then
        Exit Function
    End If
    Set treenode = RecordingDoc.RecordingDocument.image(0, True)
    'Set treenode = Lsm5.NewDocument this will create a new document we want to use the same document
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    AcquisitionController.DestinationImage(0) = treenode 'EngelImageToHechtImage(GlobalSingleImage).Image(0, True)
    AcquisitionController.DestinationImage(1) = Nothing
    Set ProgressFifo = AcquisitionController.DestinationImage(0)
    Lsm5.tools.CheckLockControllers True
    AcquisitionController.StartGrab eGrabModeSingle
    'Set RecordingDoc = Lsm5.StartScan this does not overwrite
    If Not ProgressFifo Is Nothing Then ProgressFifo.Append AcquisitionController
    'Debug.Print "ScanToImage part1 " & Round(Timer - Time, 3)
    Sleep (PauseGrabbing)
    Time = Timer
    While AcquisitionController.IsGrabbing
        Sleep (PauseGrabbing) ' the timing makes the different whether we release the system or not often enough. funny enough a small value is better
        DoEvents
        If ScanStop Then
            Exit Function
        End If
    Wend
    'Debug.Print "ScanToImage properAcq " & Round(Timer - Time, 3)
    ScanToImage = True
    Exit Function
ErrorHandle:
    LogManager.UpdateErrorLog "Error in ScanToImage for image " _
    & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & " " & Err.Description
End Function

''''
' Start Fcs Measurment
''''
Public Function ScanToFcs(FcsData As AimFcsData) As Boolean
    On Error GoTo ErrorHandle
    Dim FcsControl As AimFcsController
    Set FcsControl = Fcs
    If FcsData Is Nothing Then
      Exit Function
    End If
    FcsControl.StartMeasurement FcsData
    Sleep (PauseGrabbing)
    While FcsControl.IsAcquisitionRunning(1)
        Sleep (PauseGrabbing)
        If ScanStop Then
            Exit Function
        End If
        DoEvents
    Wend
    ScanToFcs = True
    Exit Function
ErrorHandle:
    LogManager.UpdateErrorLog "Error in ScanToFcs from image " & GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath") & Err.Description
End Function

'''''
'   Set the FCS controller and data stuff
'''''
Private Sub Initialize_Controller()
    Set FcsControl = Fcs 'member of Lsm5VBAProject
    Set viewerGuiServer = Lsm5.viewerGuiServer
    Set FcsPositions = FcsControl.SamplePositionParameters
    viewerGuiServer.FcsSelectLsmImagePositions = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''
''''''''' Creates NewRecords'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'''
' Creates New DsRecordingDoc and a entry in the experiment Tree (works only with ZENv > 2010)
'   RecordingDoc [In/Out] - A document. If it exists and ForceCreation = False then only the name will be changed
'   Name                  - Name of the document (Tab-name)
'   ForceCreation         - force creation of a new document and entry in the experiment tree
''''
Public Function NewRecord(RecordingDoc As DsRecordingDoc, Name As String, Optional ForceCreation As Boolean = False) As Boolean
    
    On Error GoTo ErrorHandle:
    Dim node As AimExperimentTreeNode
    If RecordingDoc Is Nothing Or ForceCreation Then
        'for version > 2011 you could also specify containers this is not used here
        'Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
        'node.type = eExperimentTeeeNodeTypeLsm
        'viewerGuiServer.InsertExperimentTreeNode node, True, Container (this last option does not exist for ZEN<2011)
        Set node = Lsm5.NewDocument
        node.type = eExperimentTeeeNodeTypeLsm
        Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
        While RecordingDoc.IsBusy
            Sleep (Pause)
            DoEvents
        Wend
    End If
    RecordingDoc.SetTitle Name
    NewRecord = True
    Exit Function
    
ErrorHandle:
    LogManager.UpdateErrorLog " Error in NewRecord " & Name & Err.Description
    
End Function



''
' Check if document exists and if it is loaded in the GUI. Otherwise creates a new one.
''
Public Function NewRecordGui(RecordingDoc As DsRecordingDoc, Name As String, ZEN As Object, ZENv As Integer) As Boolean

    If ZENv > 2010 Then
        NewRecordGui = NewRecordGuiAi(RecordingDoc, Name, ZEN)
    Else
        NewRecordGui = NewRecord(RecordingDoc, Name, False) ' no idea how to check the name of documents in ZENv < 2011
    End If
    
End Function


''''
'  Check if Name exists in GUI
'  recquires ZEN_Micro_AIM_ApplicationInterface
''''
Public Function NewRecordGuiAi(RecordingDoc As DsRecordingDoc, Name As String, ZEN As Object) As Boolean

    On Error GoTo ErrorHandle
    
    If Not ZEN Is Nothing Then
        If Not NewRecord(RecordingDoc, Name, False) Then
            Exit Function
        End If
        'leave some time to set the name
        Sleep (1000)
        If ZEN.gui.Document.ItemCount > 0 Then
            ZEN.gui.Document.ByName = Name
            If ZEN.gui.Document.Name.Value <> Name Then
                If Not NewRecord(RecordingDoc, Name, True) Then
                    Exit Function
                End If
                ZEN.gui.Document.ByName = Name
            End If
        Else
            If Not NewRecord(RecordingDoc, Name, True) Then
                Exit Function
            End If
        End If
        NewRecordGuiAi = True
    Else
        MsgBox "Error: NewRecordGuiAi. Tried to use ZEN_Micro_AIM_ApplicationInterface but no ZEN objet has been initialized"
        LogManager.UpdateErrorLog "Error: NewRecordGuiAi. Tried to use ZEN_Micro_AIM_ApplicationInterface but no ZEN objet has been initialized"
    End If
    Exit Function
ErrorHandle:
    LogManager.UpdateErrorLog "Error in NewRecordGuiAi " & Name & " " & Err.Description
End Function



''''
'   Create a new FCSData record
''''
Public Function NewFcsRecord(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, Name As String, Optional ForceCreation As Boolean = False) As Boolean
    Dim MaxDataSets As Integer
    Dim i As Integer
    On Error GoTo ErrorHandle:
    Dim node As AimExperimentTreeNode
    If FcsData Is Nothing Or RecordingDoc Is Nothing Or ForceCreation Then
        'for version > 2011 you could also specify containers this is not used here
        'Set viewerGuiServer = Lsm5.viewerGuiServer
        'Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
        'node.type = eExperimentTeeeNodeTypeConfoCor
        'viewerGuiServer.InsertExperimentTreeNode node, True, Container (this last option does not exist for ZEN<2011)
        Set node = Lsm5.NewDocument
        node.type = eExperimentTeeeNodeTypeConfoCor
        Set FcsData = node.FcsData
        FcsData.Name = Name
        Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
        While RecordingDoc.IsBusy
            Sleep (Pause)
            DoEvents
        Wend
    End If

    RecordingDoc.SetTitle Name
    NewFcsRecord = True
    Exit Function

ErrorHandle:
    LogManager.UpdateErrorLog "Error in NewFcsRecord " & Name & " " & Err.Description
End Function

'''
' Remove all existing Data from FcsData. This is recquired if you want only to save new data
'''
Public Function CleanFcsData(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData) As Boolean
    Dim MaxDataSets As Integer
    Dim i As Integer
    If FcsData Is Nothing Or RecordingDoc Is Nothing Then
        GoTo NoRecord
    End If
    MaxDataSets = FcsData.DataSets - 1
    If MaxDataSets >= 0 Then
        For i = MaxDataSets To 0 Step -1
            FcsData.Remove (i)
        Next i
    End If
    CleanFcsData = True
    Exit Function
NoRecord:
    LogManager.UpdateErrorLog "CleanFcsRecord: Found no active record for FCS!"
End Function


''
' Check if document exists and if it is loaded in the GUI. Otherwise creates a new one.
''
Public Function NewFcsRecordGui(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, Name As String, ZEN As Object, ZENv As Integer) As Boolean

    If ZENv > 2010 Then
        NewFcsRecordGui = NewFcsRecordGuiAi(RecordingDoc, FcsData, Name, ZEN)
    Else
        NewFcsRecordGui = NewFcsRecord(RecordingDoc, FcsData, Name, False)  ' no idea how to check the name of documents in ZENv < 2011
    End If
    
End Function


''''
'  Check if Name exists in GUI
'  recquires ZEN_Micro_AIM_ApplicationInterface
''''
Public Function NewFcsRecordGuiAi(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, Name As String, ZEN As Object) As Boolean
    
    On Error GoTo ErrorHandle
    
    If Not ZEN Is Nothing Then
        If Not NewFcsRecord(RecordingDoc, FcsData, Name, False) Then
            Exit Function
        End If
        'leave some time to set the name
        Sleep (1000)
        If ZEN.gui.Document.ItemCount > 0 Then
            ZEN.gui.Document.ByName = Name
            If ZEN.gui.Document.Name.Value <> Name Then
                If Not NewFcsRecord(RecordingDoc, FcsData, Name, True) Then
                    Exit Function
                End If
                ZEN.gui.Document.ByName = Name
            End If
        Else
            If Not NewFcsRecord(RecordingDoc, FcsData, Name, True) Then
                Exit Function
            End If
        End If
        NewFcsRecordGuiAi = True
    Else
        MsgBox "Error: NewFcsRecordGuiAi. Tried to use ZEN_Micro_AIM_ApplicationInterface but no ZEN objet has been initialized"
        LogManager.UpdateErrorLog "Error: NewFcsRecordGuiAi. Tried to use ZEN_Micro_AIM_ApplicationInterface but no ZEN objet has been initialized"
    End If
    Exit Function
ErrorHandle:
    LogManager.UpdateErrorLog "Error in NewFcsRecordGuiAi " & Name & Err.Description
End Function

''''
' SaveFcsMeasurment to File
''''
Public Function SaveFcsMeasurement(FcsData As AimFcsData, FileName As String) As Boolean
    
    If FcsData Is Nothing Then
        MsgBox "No Fcs Recording to Save"
        Exit Function
    End If
    ' Write to file
    Dim writer As AimFcsFileWrite
    Set writer = Lsm5.CreateObject("AimFcsFile.Write")
    writer.FileName = FileName
    writer.FileWriteType = eFcsFileWriteTypeAll
    writer.format = eFcsFileFormatConfoCor3WithRawData
    writer.Source = FcsData
    writer.Run
    Sleep (1000)
    'write twice to be sure
    If Not writer.DestinationFilesExist(FileName) Then
        writer.FileName = FileName
        writer.FileWriteType = eFcsFileWriteTypeAll
        writer.format = eFcsFileFormatConfoCor3WithRawData
        writer.Source = FcsData
        writer.Run
    Else
        
    End If
    SaveFcsMeasurement = True
End Function

Public Sub SaveFcsPositionList(sFile As String, positionsPx() As Vector)
    On Error GoTo ErrorHandle
    Close
    Dim iFileNum As Integer
    Dim i As Long
    Dim PosX As Double
    Dim PosY As Double
    Dim PosZ As Double
    iFileNum = FreeFile()
    Open sFile For Output As iFileNum
    Print #iFileNum, "%X Y Z (um): ZEN Fcs position convention 0, 0  is center of image, Z is absolute coordinate"
    For i = 0 To GetFcsPositionListLength - 1
        GetFcsPosition PosX, PosY, PosZ, i
        Print #iFileNum, Round(PosX * 1000000, PrecXY) & " " & Round(PosY * 1000000, PrecXY) & " " & Round(PosZ * 1000000, PrecXY)
    Next i
    On Error GoTo ErrorHandle2:
    Print #iFileNum, "%X Y Z (px). Imaging convention 0,0,0 is upper left corner bottom slice"
    For i = 0 To UBound(positionsPx)
        Print #iFileNum, positionsPx(i).X & " " & positionsPx(i).Y & " " & positionsPx(i).Z
    Next i

    Close
    Exit Sub
ErrorHandle:
    LogManager.UpdateErrorLog "SaveFcsPositionList Can't write " & sFile & " for the FcsPositions"
    Exit Sub
ErrorHandle2:
    Close
    LogManager.UpdateErrorLog "positionsPx not assigned"
End Sub

'''''
'   SystemVersionOffset()
'   Calculate an offset added to z-stack changes
'       [GlobalCorrectionOffset] Global Out - Offset added to shift in zStack
'   TODO: Do we still need it. Only for Axioskop does the Offset change
'''''
Public Sub SystemVersionOffset(Optional Tmp As Boolean) ' tmp is a hack to hide function from menu
    SystemVersion = Lsm5.Info.VersionIs
    If StrComp(SystemVersion, "2.8", vbBinaryCompare) >= 0 Then
        If Lsm5.Info.IsAxioskop Then

        ElseIf Lsm5.Info.IsAxioplan Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioplan2 Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioplan2i Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioVert Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxiovert100M Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxiovert200M Then
            GlobalCorrectionOffset = 0
        Else
            GlobalCorrectionOffset = 0
        End If
    Else
        If Lsm5.Info.IsAxioskop Then

        ElseIf Lsm5.Info.IsAxioplan Then
            GlobalCorrectionOffset = 0
        ElseIf Lsm5.Info.IsAxioVert Then
            GlobalCorrectionOffset = 0
        Else
            GlobalCorrectionOffset = 0
        End If
    End If
End Sub



''''' ' this should move to function
'   FailSafeMoveStage(Optional Mark As Integer = 0)
'   Moves stage and wait till it is finished
'       [x] In - x-position
'       [y] In - y-position
'''''
Public Function FailSafeMoveStageXY(X As Double, Y As Double) As Boolean
    
    FailSafeMoveStageXY = False


    Lsm5.Hardware.CpStages.SetXYPosition X, Y
    'TODO Check this
    Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
        Sleep (200)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                ScanStop = True
                Exit Function
            End If
        End If
    Loop
    
    FailSafeMoveStageXY = True
    
End Function


'''''
'   FailSafeMoveStageZ(z As Double)
'   Moves focus and wait till it is finished
'       [z] In - z-position )
'''''
Public Function FailSafeMoveStageZ(Z As Double) As Boolean
    FailSafeMoveStageZ = False
    If ZBacklash <> 0 Then
        Lsm5.Hardware.CpFocus.position = Z - ZBacklash ' move at correct position
        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
            Sleep (20)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    FailSafeMoveStageZ = False
                    Exit Function
                End If
            End If
        Loop
    End If
    Lsm5.Hardware.CpFocus.position = Z  ' move at correct position
    Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy Or Lsm5.Hardware.CpFocus.IsBusy
        Sleep (20)
        If GetInputState() <> 0 Then
            DoEvents
            If ScanStop Then
                FailSafeMoveStageZ = False
                Exit Function
            End If
        End If
    Loop

    FailSafeMoveStageZ = True
End Function


'''''
'   MoveToNextLocation(Optional Mark As Integer = 0)
'   Moves to next location as set in the stage (mark)
'   Default will cycle through all positions sequentially starting from actual position
'       [Mark] In - Number of position where to move.
'''''
Public Function MoveToNextLocation(Optional Mark As Integer = 0) As Boolean
        Dim MarkCount As Long
        Dim count As Long
        Dim idx As Long
        Dim dX As Double
        Dim dY As Double
        Dim dZ As Double
        Dim i As Integer
        Lsm5.Hardware.CpStages.MarkMoveToZ (Mark)
        'Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ (0)  'old code Moves to the first location marked in the stage control. How to move to next point?
        ' the points were deleted and readded at the end of list in the Acquisition function
        'TODO: Check code
        Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
            Sleep (100)
            If GetInputState() <> 0 Then
                DoEvents
                If ScanStop Then
                    Exit Function
                End If
            End If
        Loop
        MoveToNextLocation = True
End Function




''''
'   WaitForRecentering(Z As Double, Success As Boolean) As Boolean
'   calls the microscope specific WaitForRecentering
'''
Public Function WaitForRecentering(Z As Double, Optional Success As Boolean = False, Optional ZENv As Integer = 2011) As Boolean
    If ZENv = 2010 Then
        If Not WaitForRecentering2010(Z, Success) Then
            Exit Function
        End If
    End If
    If ZENv > 2010 Then
        If Not WaitForRecentering2011(Z, Success) Then
            Exit Function
        End If
    End If
    WaitForRecentering = True
End Function



''''
'   WaitForRecentering2010(Z As Double) As Boolean
'   Helping function to check if after acquisition focus returns to its correct position
'       [Z] - is value where the central slice should be.
'   Additional remarks: Lsm5.Hardware.CpFocus.Position is not updated correctly after acquisition (CpFocus needs to return to working position) on the other hand
'   Lsm5.DsRecording.Sample0Z keeps track correctly of the position
'''
Public Function WaitForRecentering2010(Z As Double, Optional Success As Boolean = False) As Boolean
    Dim Cnt As Integer
    Dim MaxCnt As Integer
    MaxCnt = 6
    Cnt = 0
    ' Wait up to 4 sec for centering
    ' position central slice is Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position (or the real actual position)
    ' this waits for central slice at Z
    Dim pos As Double
    Dim Sample0Z As Double
    pos = Lsm5.Hardware.CpFocus.position
    'in this case stage has bene moved
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        While Round(pos, 1) <> Round(Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + pos - Z
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
            GoTo FailedWaiting
        End If
'        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 + _
'            pos - Z, 1) And Cnt < 10 'this is slow and not required and makes it slow
'            Sleep (400)
'            DoEvents
'            Cnt = Cnt + 1
'            If ScanStop Then
'                Exit Function
'            End If
'        Wend
    Else
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + _
            pos - Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + pos - Z
            GoTo FailedWaiting
        End If
    End If
    Success = True
    WaitForRecentering2010 = True
    Exit Function
FailedWaiting:
    DoEvents
    Success = False
    WaitForRecentering2010 = True
End Function

''''
'   WaitForRecentering2011(Z As Double, Success As Boolean) As Boolean
'   Helping function to check if after acquisition focus returns to its correct position
'       [Z] - is value where the central slice should be.
'       [Success] - Tells if central slide has been found before maximal number of iterations
'   Additional remarks: Lsm5.Hardware.CpFocus.Position is not updated correctly after acquisition (CpFocus needs to return to working position) on the other hand
'   Lsm5.DsRecording.Sample0Z keeps track correctly of the position
'''
Public Function WaitForRecentering2011(Z As Double, Optional Success As Boolean = False) As Boolean
    Dim Cnt As Integer
    Dim MaxCnt As Integer
    MaxCnt = 6
    Cnt = 0
    ' Wait up to 4 sec for centering
    ' Note pculiarity of centering
    ' position central slice is Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position (or the real actual position)
    ' this waits for central slice at Z
    Dim pos As Double
    Dim Sample0Z As Double
    pos = Lsm5.Hardware.CpFocus.position
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        While Round(pos, 1) <> Round(Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        
        If Cnt = MaxCnt Then
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + pos - Z
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
            GoTo FailedWaiting
        End If
        
    Else
    
        While Round(Lsm5.DsRecording.Sample0Z, 1) <> Round(Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + _
            pos - Z, 1) And Cnt < MaxCnt
            Sleep (400)
            DoEvents
            Cnt = Cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        
        If Cnt = MaxCnt Then
            Success = False
            Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + pos - Z
            GoTo FailedWaiting
        End If
    End If
    DoEvents
    Success = True
    WaitForRecentering2011 = True
    Exit Function
FailedWaiting:
    DoEvents
    Success = False
    WaitForRecentering2011 = True
End Function

''''
'   Recenter(Z As Double)
'   Sets the central slice. This slice is then maintained even when framespacing is changing.
'       [Z]     - Absolute position of central slice
'   position central slice is Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position
''''
Public Function Recenter_pre(Z As Double, Optional Success As Boolean = False, Optional ZENv As Integer = 2011) As Boolean
    If Not Recenter(Z, ZENv) Then
        Exit Function
    End If
    If Not WaitForRecentering(Z, Success, ZENv) Then
        Exit Function
    End If
    If Not ScanStop Then
        Recenter_pre = True
    End If
End Function

Public Function Recenter_post(Z As Double, Optional Success As Boolean = False, Optional ZENv As Integer = 2011) As Boolean
    If Not WaitForRecentering(Z, Success, ZENv) Then
        Exit Function
    End If
    
    If Not ScanStop Then
        Recenter_post = True
    End If
End Function

Public Function Recenter(Z As Double, Optional ZENv As Integer = 2011) As Boolean
    Dim i As Integer
    If ZENv = 2010 Then
        For i = 1 To 1
            If Not Recenter2010(Z) Then
                Exit Function
            End If
            Sleep (200)
        Next i
    End If
    If ZENv > 2010 Then
        If Not Recenter2011(Z) Then
            Exit Function
        End If
    End If
    Recenter = True
End Function

Public Function Recenter2010(Z As Double) As Boolean
    Dim MoveStage As Boolean
    Dim pos As Double
    Dim Sample0Z As Double
    pos = Lsm5.Hardware.CpFocus.position
    MoveStage = True ' this is the only difference to 2011 version
    
    If Lsm5.DsRecording.SpecialScanMode = "ZScanner" Or (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Then
        MoveStage = True
    End If
    Dim Tmp As Integer
    
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + pos - Z
    Sleep (100)
    DoEvents
    If MoveStage Then
        If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
        End If
    End If
    DoEvents
    Recenter2010 = True
End Function

Public Function Recenter2011(Z As Double) As Boolean
    Dim MoveStage As Boolean
    Dim framesPerStack As Integer
    Dim pos As Double
    pos = Lsm5.Hardware.CpFocus.position
    MoveStage = False 'only move stage when required
    
    If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then
        MoveStage = True
    End If
        
    'Center slide
    Lsm5.DsRecording.Sample0Z = Lsm5.DsRecording.frameSpacing * (Lsm5.DsRecording.framesPerStack - 1) / 2 + pos - Z
    Sleep (100)
    DoEvents
    If MoveStage Then
        If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
        End If
    End If
    DoEvents
    'this messes around with the slice number. Don't use it
    'ZEN.gui.Acquisition.ZStack.CenterPositionZ.Value = Z

    Recenter2011 = True
End Function








''''
' Compute the centerofmass of image stored in RecordingDoc return values according
'   Use channel with name TrackingChannel
''''
Public Function MassCenter(RecordingDoc As DsRecordingDoc, TrackingChannel As String) As Vector

    On Error GoTo ErrorHandle:
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp")
    Dim Match As MatchCollection
    
    Dim scrline As Variant
    Dim spl As Long
    Dim bpp As Long
    Dim ColMax As Long
    Dim LineMax As Long
    Dim FrameMax As Integer
    Dim pixelSize As Double
    Dim frameSpacing As Double
    Dim IntLine() As Variant
    Dim IntCol() As Variant
    Dim IntFrame() As Variant
    
    Dim channel As Integer
    Dim Frame As Long
    Dim Line As Long
    Dim Col As Long
    Dim MinColValue As Single
    Dim minLineValue As Single
    Dim minFrameValue As Single
    Dim MaxColValue As Single
    Dim MaxLineValue As Single
    Dim MaxframeValue As Single
    Dim LineSum As Double
    Dim LineWeight As Single
    Dim MidLine As Single
    Dim Threshold As Single
    Dim LineValue As Single
    Dim PosValue As Single
    Dim ColSum As Single
    Dim ColWeight As Single
    Dim MidCol As Single
    Dim ColValue As Single
    Dim FrameSum As Single
    Dim FrameWeight As Single
    Dim MidFrame As Single
    Dim FrameValue As Single
    
   
    DoEvents
    
    
    'Find the channel to track
    Dim Rec As DsRecordingDoc
    Dim FoundChannel As Boolean
    FoundChannel = False
    RegEx.Pattern = "(\w+) (\w+\d+-\w+\d+)"
    Dim name_channel As String
    If RegEx.Test(TrackingChannel) Then
        Set Match = RegEx.Execute(TrackingChannel)
        name_channel = Match.Item(0).SubMatches.Item(1)
    End If
    Dim name_channelA() As String
    name_channelA = Split(name_channel, "-")
    For channel = 0 To RecordingDoc.GetDimensionChannels - 1
        Debug.Print "Channel Names " & RecordingDoc.ChannelName(channel)
        If RecordingDoc.ChannelName(channel) = name_channelA(0) & "-" & name_channelA(1) Then ' old Code: Left(TrackingChannel,4)
            FoundChannel = True
            Exit For
        End If
    Next channel
    
    If Not FoundChannel Then
        For channel = 0 To RecordingDoc.GetDimensionChannels - 1 ' this is true when only one track is acquired
            If RecordingDoc.ChannelName(channel) = name_channelA(0) Then ' old Code: Left(TrackingChannel,4)
                FoundChannel = True
                Exit For
            End If
        Next channel
    End If
    
    If Not FoundChannel Then
        LogManager.UpdateErrorLog " MassCenter Was not able to find channel: " & TrackingChannel & " for tracking in " & _
        GetSetting(appname:="OnlineImageAnalysis", section:="macro", Key:="filePath")
        Exit Function
    End If


    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
    ColMax = RecordingDoc.Recording.SamplesPerLine
    LineMax = RecordingDoc.Recording.LinesPerFrame
    
    If RecordingDoc.Recording.ScanMode = "ZScan" Then
        LineMax = 1
    End If
    If RecordingDoc.Recording.ScanMode = "ZScan" Or RecordingDoc.Recording.ScanMode = "Stack" Then
        FrameMax = RecordingDoc.Recording.framesPerStack
    Else
        FrameMax = 1
    End If
    
     
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    ReDim IntLine(LineMax - 1)
    ReDim IntCol(ColMax - 1)
    ReDim IntFrame(FrameMax - 1)
        

    
    'Compute center of mass
    'Reads the pixel values and fills the tables with the projected (integrated) pixels values in the three directions
    ' Intline  => Y: is the sum along X and along Z
    ' IntCol   => X : is the sum along Y and Z
    ' IntFrame => Z : is the sum along X and Y
    
    For Frame = 0 To FrameMax - 1
        For Line = 0 To LineMax - 1
            bpp = 0 ' bytes per pixel (this will be changed by ScanLine
            ' spl samples per line (will be changed by scal line)
            scrline = RecordingDoc.ScanLine(channel, 0, Frame, Line, spl, bpp)  'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 0 To ColMax - 1             'Now I'm scanning all the pixels in the line
                IntLine(Line) = IntLine(Line) + scrline(Col)
                IntCol(Col) = IntCol(Col) + scrline(Col)
                IntFrame(Frame) = IntFrame(Frame) + scrline(Col)
            Next Col
        Next Line
    Next Frame
    
    'no thresholding for the moment
    'compute center of mass
    
    'First it finds the minimum and maximum projected (integrated) pixel values in the 3 dimensions
    MassCenter.Y = weightedMean(IntLine)
    MassCenter.X = weightedMean(IntCol)
    MassCenter.Z = weightedMean(IntFrame)
'    Dim Max As Single
'    Max = MAXA(IntFrame)
'    For Frame = 0 To FrameMax - 1
'        If IntFrame(Frame) = Max Then
'            Exit For
'        End If
'    Next Frame
'    MassCenter.Z = Frame
    Exit Function
ErrorHandle:
    MsgBox ("Error in MicroscopeIO.MassCenter " + TrackingChannel + " " + Err.Description)
    ScanStop = True
End Function



''''''
' SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String) As Boolean
' Copied and adapted from MultiTimeSeries macro
''''''
Public Function SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String, FileFormat As enumAimExportFormat) As Boolean
    Dim Export As AimImageExport
    Dim image As AimImageMemory
    Dim Error As AimError
    Dim Planes As Long
    Dim Plane As Long
    Dim positions As Long
    Dim Horizontal As enumAimImportExportCoordinate
    Dim Vertical As enumAimImportExportCoordinate

    On Error GoTo Done

    'Set Image = EngelImageToHechtImage(Document).Image(0, True)
    If Not Document Is Nothing Then
        Set image = Document.RecordingDocument.image(0, True)
    End If
    
    Set Export = Lsm5.CreateObject("AimImageImportExport.Export.4.5")
    'Set Export = New AimImageExport
    Export.FileName = FileName
    Export.format = FileFormat
    Export.StartExport image, image
    Set Error = Export
    Error.LastErrorMessage
    
    Planes = 1
    Export.GetPlaneDimensions Horizontal, Vertical
    If Document.Recording.MultiPositionAcquisition Then
        positions = Document.Recording.MultiPositionArraySize
    End If
    If positions = 0 Then
        positions = 1
    End If
    
    Select Case Vertical
        Case eAimImportExportCoordinateY:
             Planes = image.GetDimensionZ * image.GetDimensionT * positions
        Case eAimImportExportCoordinateZ:
             Planes = image.GetDimensionT
    End Select
    'TODO check. what happens here with Export.ExportPlane Nothing why Nothing (thumbnails)
    For Plane = 0 To Planes - 1
        If GetInputState() <> 0 Then
            DoEvents
             If ScanStop Then
                Export.FinishExport
                Exit Function
            End If
        End If
        Export.ExportPlane Nothing
    Next Plane
    Export.FinishExport
    SaveDsRecordingDoc = True
    Exit Function
    
Done:
    MsgBox "Check Temporary Files Folder! Cannot Save Temporary File(s)!"
    ScanStop = True
    Export.FinishExport
    StopAcquisition
End Function



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






''''''
''
''''''
'Public Function SubImagingWorkFlowFcs(FcsData As AimFcsData, HighResArrayX() As Double, _
' HighResArrayY() As Double, HighResArrayZ() As Double, fileDir As String, FileName As String, Optional pixelSizeXY As Double = 0, Optional pixelSizeZ As Double = 0) As Boolean
'
'
'    Dim i As Long
'    ClearFcsPositionList
'    Dim Recording As DsRecordingDoc
'    For i = 1 To UBound(HighResArrayX)
'        SetFcsPosition HighResArrayX(i), HighResArrayY(i), HighResArrayZ(i), i - 1
'    Next i
'
'    If Not CheckDir(fileDir) Then
'        Exit Function
'    End If
'
'    NewFcsRecord FcsData, "fcs_" & FileName, 1
'
'    If Not FcsMeasurement(FcsData) Then
'        Exit Function
'    End If
'
'    Set Recording = Lsm5.DsRecordingActiveDocObject
'    Recording.SetTitle "fcs_" & FileName
'
'    SaveFcsPositionList fileDir & FileName & "_fcsPos.txt", pixelSizeXY, pixelSizeZ
'    'save measurement
'    SaveFcsMeasurement FcsData, fileDir & FileName & ".fcs"
'
'End Function



'''''
'' TODO: Why not use Lsm5.StartScan?
'''''
'Public Sub ScanToImageOld(RecordingDoc As DsRecordingDoc) ' new routine to scan overwrite the same image, even with several z-slices
'   ' Dim AcquisitionController As AimAcquisitionController40.AimScanController 'now public
'    Dim image As AimImage
'
'    If Not RecordingDoc Is Nothing Then
'        Set image = RecordingDoc.RecordingDocument.image(0, True)
'
'        If Not image Is Nothing Then
'            Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
'            AcquisitionController.DestinationImage(0) = image
'            AcquisitionController.DestinationImage(1) = Nothing
'            AcquisitionController.StartGrab eGrabModeSingle
'        End If
'    End If
'
'End Sub



