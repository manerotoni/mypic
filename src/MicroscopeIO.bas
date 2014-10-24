Attribute VB_Name = "MicroscopeIO"
''''
' Module with functions for controlling stage, starts and stop scan, creating documents and saving images
'''''

Option Explicit
Option Base 0
Public Const DebugCode = True
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
Public Const TolPx = 0.0001                'Tolerance for computed img px value

Public ZBacklash  As Double           'ToDo: is it still recquired?.
                                           'Has to do with the movements of the focus wheel that are "better"
                                            'if they are long enough. For amoment a test did not gave significant differences This is required for ZEN2010
Public ZSafeDown As Double
Public ZenV As Integer            'String variable indicating the version of ZEN used 2010 ir 2011 (2012)
Public ZEN As Object             'Object containing Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface (for ZEN > 2011)

''''''''''''''''''''''
'''GLOBAL VARIABLES'''
''''''''''''''''''''''

Public ScanStop As Boolean      'if TRUE current recording is stopped
Public ScanPause As Boolean     'if TRUE current recording is paused
Public Running As Boolean       'TRUE when system is running (e.g. after start)
Public GlobalDataBaseName As String   'Name of output folder
Public TrackNumber As Integer    'number of available tracks

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

Public Const fcsTimeOverhead = 5000 'This is roughly the time for switching from imaging GaSP to FCS_APD (could be longer or shorter depending on the protocol)
Public Const PauseGrabbing = 50 'pause for polling the whether scan/fcscontroller are acquiring. A high value makes more errors!
Public PauseEndAcquisition As Double 'A workaround to avoid errors in FCS/imaging. Does not seem to work. Disabled

'''
' Returns version number (ZEN2010, etc.)
'''
Public Function getVersionNr() As Integer
    Dim VersionNr As Long
    VersionNr = CLng(VBA.Left(Lsm5.Info.VersionIs, 1))
    Select Case VersionNr
        Case Is <= 6:
            getVersionNr = 2010
        Case 7:
            getVersionNr = 2011
        Case Is >= 8:
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


''''''
''   remove all vector elements
''''''
Public Function ClearVectorElements() As Boolean
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions
    vo.Cleanup
End Function


''''
' Check if system is busy in some of its actions
''''
Public Function isReady(Optional Time As Double = 0.1) As Boolean
    Dim BusyFcs As Boolean
    Dim message As String
    Dim AimScanCalibration As AimScanCalibration
    Set AimScanCalibration = Lsm5.ExternalDsObject.ScanController

    If Lsm5.Info.IsFCS Then
        Dim FcsControl As AimFcsController
        Set FcsControl = Fcs
        BusyFcs = FcsControl.IsAcquisitionRunning(1)
    End If
    isReady = CInt(AimScanCalibration.GetIsHardwareBusy(True, Time, message) Or BusyFcs Or Lsm5.Info.IsAnyHardwareBusy)
    isReady = Not isReady
End Function

'''
' Sleep for a certain time and perform DoEvents inbetween. WaitTime is in milliseconds
'''
Public Sub SleepWithEvents(WaitTime As Double)
    Dim i As Long
    Dim cycles As Long
    cycles = Round(WaitTime / PauseGrabbing)
    For i = 0 To cycles
        Sleep (PauseGrabbing)
        DoEvents
    Next i
End Sub

'''''
'   ScanToImage (RecordingDoc As DsRecordingDoc) As Boolean
'   scan overwrite the same image, even with several z-slices
'''''
Public Function ScanToImage(RecordingDoc As DsRecordingDoc, Optional TimeOut As Double = -1) As Boolean


    Dim Time As Double
    Dim ProgressFifo As IAimProgressFifo ' this shows how far you are with the acquisition image ( the blue bar at the bottom). The usage of it makes the macro quite slow
    Dim AcquisitionController As AimScanController
    Dim treenode As Object
    Dim iTry As Integer

    iTry = 1
    'Procedure is completely executed 3 times in case of error. RecordingDoc.IsBusy is less (not at all?) error prone
RepeatScanToImage:
    On Error GoTo ErrorScanToImage
    'Dim gui As Object
    'Set gui = Lsm5.ViewerGuiServer not recquired anymore
    If RecordingDoc Is Nothing Then
        Exit Function
    End If
    Set treenode = RecordingDoc.RecordingDocument.image(0, True)
    'Set treenode = Lsm5.NewDocument this will create a new document we want to use the same document
    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
    AcquisitionController.DestinationImage(0) = treenode 'EngelImageToHechtImage(GlobalSingleImage).Image(0, True)
    AcquisitionController.DestinationImage(1) = Nothing
    Set ProgressFifo = AcquisitionController.DestinationImage(0)
    Lsm5.tools.CheckLockControllers True
    
    AcquisitionController.StartGrab eGrabModeSingle
    Time = Timer
    'Set RecordingDoc = Lsm5.StartScan this does not overwrite
    If Not ProgressFifo Is Nothing Then ProgressFifo.Append AcquisitionController
    'Debug.Print "ScanToImage part1 " & Round(Timer - Time, 3)
'    Sleep (PauseGrabbing)
'    'While AcquisitionController.isGrabbing this command seems to hang quite frequently
    While RecordingDoc.IsBusy
        Sleep (PauseGrabbing)
        DoEvents
        If ScanStop Then
            Exit Function
        End If
        If TimeOut > 0 And (Timer - Time > TimeOut) Then
            LogManager.UpdateErrorLog "TimeOut of image acquisition  after " & TimeOut & " sec"
            StopAcquisition
            GoTo ExitWhile
        End If
    Wend
ExitWhile:
    ScanToImage = True
    On Error GoTo 0
    Exit Function
ErrorScanToImage:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ScanToImage of Module MicroscopeIO at line " & Erl & ". Try " & iTry & "/3"
    StopAcquisition
    Sleep (500)
    DoEvents
    iTry = 1 + iTry
    If iTry <= 3 Then
        Err.Clear
        Resume RepeatScanToImage
    End If
End Function

''''
' Start Fcs Measurment
''''
Public Function ScanToFcs(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, Optional TimeOut As Double = -1) As Boolean
On Error GoTo ErrorScanToFcs
    Dim iTry As Integer
    Dim FcsControl As AimFcsController
    Dim Time As Double

    iTry = 1
    'Procedure is completely executed 3 times in case of error. RecordingDoc.IsBusy is less (not at all?) error prone
RepeatScanToFcs:
    Set FcsControl = Fcs
    If FcsData Is Nothing Then
      Exit Function
    End If
    FcsControl.StartMeasurement FcsData
    Time = Timer
    'this is the minimal time it takes + some extra time for the hardware to switch
    With FcsControl.AcquisitionParameters
        If FcsControl.SamplePositionParameters.PositionListSize = 0 Then
            SleepWithEvents (.MeasurementTime * .MeasurementRepeat * 1000 + fcsTimeOverhead)
        Else
            SleepWithEvents (.MeasurementTime * .MeasurementRepeat * FcsControl.SamplePositionParameters.PositionListSize * 1000 + fcsTimeOverhead)
        End If
    End With
    'check if for sure we are finished
    While FcsControl.IsAcquisitionRunning(1)
        SleepWithEvents (PauseGrabbing)
        If ScanStop Then
            Exit Function
        End If
        If TimeOut > 0 And (Timer - Time > TimeOut) Then
            LogManager.UpdateErrorLog "TimeOut of Fcs acquisition  after " & TimeOut & " sec"
            StopAcquisition
            GoTo ExitWhile
        End If
    Wend
ExitWhile:
    StopAcquisition
    ScanToFcs = True
    Exit Function
ErrorScanToFcs:
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure ScanToFcs of Module MicroscopeIO at line " & Erl & ". Try " & iTry & "/3"
    StopAcquisition
    Sleep (500)
    DoEvents
    iTry = 1 + iTry
    If iTry <= 3 Then
        Err.Clear
        Resume RepeatScanToFcs
    End If
    '    While RecordingDoc.IsBusy ' this does not wait so it is useless
'        Sleep (PauseGrabbing)
'        If ScanStop Then
'           Exit Function
'        End If
'    Wend
End Function


'''''
'   Set the FCS controller and data stuff
'''''
Private Sub Initialize_Controller()
    Set FcsControl = Fcs 'member of Lsm5VBAProject
    Set ViewerGuiServer = Lsm5.ViewerGuiServer
    Set FcsPositions = FcsControl.SamplePositionParameters
    ViewerGuiServer.FcsSelectLsmImagePositions = True
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
    Dim Node As AimExperimentTreeNode
    If RecordingDoc Is Nothing Or ForceCreation Then
        'for version > 2011 you could also specify containers this is not used here
        'Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
        'node.type = eExperimentTeeeNodeTypeLsm
        'viewerGuiServer.InsertExperimentTreeNode node, True, Container (this last option does not exist for ZEN<2011)
        Set Node = Lsm5.NewDocument
        Node.Type = eExperimentTeeeNodeTypeLsm
        Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
        While RecordingDoc.IsBusy
            SleepWithEvents (100)
        Wend
    End If
    RecordingDoc.SetTitle Name
    'this does not help in ZEN2010
    ' Application.ThrowEvent tag_Events. .eEventRecordingNameChanged, 0
    
    NewRecord = True
    Exit Function
    
ErrorHandle:
    LogManager.UpdateErrorLog " Error in NewRecord " & Name & Err.Description
    
End Function



''
' Check if document exists and if it is loaded in the GUI. Otherwise creates a new one.
''
Public Function NewRecordGui(RecordingDoc As DsRecordingDoc, Name As String, ZEN As Object, ZenV As Integer) As Boolean

    If ZenV > 2010 Then
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
        SleepWithEvents (1000)
        If ZEN.GUI.Document.ItemCount > 0 Then
            ZEN.GUI.Document.ByName = Name
            If ZEN.GUI.Document.Name.value <> Name Then
                If Not NewRecord(RecordingDoc, Name, True) Then
                    Exit Function
                End If
                ZEN.GUI.Document.ByName = Name
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
    Dim Node As AimExperimentTreeNode
    If FcsData Is Nothing Or RecordingDoc Is Nothing Or ForceCreation Then
        'for version > 2011 you could also specify containers this is not used here
        'Set viewerGuiServer = Lsm5.viewerGuiServer
        'Set node = Lsm5.CreateObject("AimExperiment.TreeNode")
        'node.type = eExperimentTeeeNodeTypeConfoCor
        'viewerGuiServer.InsertExperimentTreeNode node, True, Container (this last option does not exist for ZEN<2011)
        Set Node = Lsm5.NewDocument
        Node.Type = eExperimentTeeeNodeTypeConfoCor
        Set FcsData = Node.FcsData
        Set FcsData = New AimFcsData
        FcsData.Name = Name
        Set RecordingDoc = Lsm5.DsRecordingActiveDocObject
        While RecordingDoc.IsBusy
            SleepWithEvents (100)
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
Public Function NewFcsRecordGui(RecordingDoc As DsRecordingDoc, FcsData As AimFcsData, Name As String, ZEN As Object, ZenV As Integer) As Boolean

    If ZenV > 2010 Then
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
        SleepWithEvents (1000)
        If ZEN.GUI.Document.ItemCount > 0 Then
            ZEN.GUI.Document.ByName = Name
            If ZEN.GUI.Document.Name.value <> Name Then
                If Not NewFcsRecord(RecordingDoc, FcsData, Name, True) Then
                    Exit Function
                End If
                ZEN.GUI.Document.ByName = Name
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
Public Function SaveFcsMeasurement(FcsData As AimFcsData, fileName As String) As Boolean
    
    If FcsData Is Nothing Then
        MsgBox "No Fcs Recording to Save"
        Exit Function
    End If
    ' Write to file
    Dim writer As AimFcsFileWrite
    Set writer = Lsm5.CreateObject("AimFcsFile.Write")
    writer.fileName = fileName
    writer.FileWriteType = eFcsFileWriteTypeAll
    writer.format = eFcsFileFormatConfoCor3WithRawData
    writer.Source = FcsData
    writer.Run
    Sleep (1000)
    'write twice to be sure
    If Not writer.DestinationFilesExist(fileName) Then
        writer.fileName = fileName
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
        getFcsPosition PosX, PosY, PosZ, i
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
    Close
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



'''''
'   Moves stage and wait till it is finished. Repeat the process twice if precision is not achieved
'       [x] In - x-position
'       [y] In - y-position
'''''
Public Function FailSafeMoveStageXY(X As Double, Y As Double) As Boolean
    Dim CurrentX As Double
    Dim CurrentY As Double
    Dim WaitTime As Integer
    Dim Prec As Double
    Dim Trial As Integer
    
    Prec = 1 'um (Thorsten Lenser uses 1 um)
    Trial = 1
SetPosition:
    WaitTime = 0
    Lsm5.Hardware.CpStages.GetXYPosition CurrentX, CurrentY
    Lsm5.Hardware.CpStages.SetXYPosition X, Y
    Do While (Abs(CurrentX - X) > Prec) Or (Abs(CurrentY - Y) > Prec) Or Lsm5.Hardware.CpStages.IsBusy
        SleepWithEvents (100)
        Lsm5.Hardware.CpStages.GetXYPosition CurrentX, CurrentY
        WaitTime = WaitTime + 1
        If ScanStop Then
            ScanStop = True
            Exit Function
        End If
        If WaitTime > 50 And Not Lsm5.Hardware.CpStages.IsBusy Then
            LogManager.UpdateWarningLog " FailSafeMoveStageXY did not reach the precision of " & Prec _
            & " um  within " & WaitTime * 100 & " ms on trial " & Trial & ". Goal position is XY: " & X & " " & Y & " reached XY: " & CurrentX _
            & " " & CurrentY
            Exit Do
        End If
    Loop
    
    '''Try a second time if it failed
    If (Abs(CurrentX - X) > Prec) Or (Abs(CurrentY - Y) > Prec) And Trial < 2 Then
        Trial = Trial + 1
        GoTo SetPosition
    End If
    FailSafeMoveStageXY = True
End Function


'''''
'   Wrapper to run with ZBacklash (generally ZBacklash is zero) or not
'''''
Public Function FailSafeMoveStageZ(Z As Double) As Boolean
    If ZBacklash <> 0 Then
        If Not FailSafeMoveStageZExec(Z - ZBacklash) Then
            Exit Function
        End If
    End If
    FailSafeMoveStageZ = FailSafeMoveStageZExec(Z)
End Function


'''''
'   Moves focus and wait till it is finished
'       [z] In - z-position in um
'''''
Public Function FailSafeMoveStageZExec(Z As Double) As Boolean
    Dim CurrentZ As Double
    Dim WaitTime As Integer
    Dim Prec As Double
    Dim Trial As Integer
    Prec = 0.2 'Used in MultitimeZEN2012 Thoresten Lenser
    Trial = 1
    
SetPosition:
    WaitTime = 0
    CurrentZ = Lsm5.Hardware.CpFocus.position
    Lsm5.Hardware.CpFocus.position = Z
    
    Do While (Abs(CurrentZ - Z) > Prec) Or Lsm5.Hardware.CpFocus.IsBusy
        SleepWithEvents (100)
        CurrentZ = Lsm5.Hardware.CpFocus.position
        WaitTime = WaitTime + 1
        If ScanStop Then
            Exit Function
        End If
        If WaitTime > 50 And Not Lsm5.Hardware.CpFocus.IsBusy Then
            LogManager.UpdateErrorLog "Warning: FocusZMovement did not reach the precision of " & Prec _
            & "um  within " & WaitTime * 100 & " ms on trial  " & Trial & ". Goal position is " & Z & " reached " & CurrentZ
            Exit Do
        End If
    Loop
    
    ''second round
    If (Abs(CurrentZ - Z) > Prec) And Trial < 2 Then
        Trial = Trial + 1
        GoTo SetPosition
    End If
    SleepWithEvents (500) 'With 500 it works
    FailSafeMoveStageZExec = True
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


'---------------------------------------------------------------------------------------
' Procedure : getMarkedStagePosition
' Purpose   : get positions in marked stage gui
'---------------------------------------------------------------------------------------
'
Public Function getMarkedStagePosition() As Vector()
    Dim MarkCount As Long
    Dim pos() As Vector
    Dim i As Long
On Error GoTo getMarkedStagePosition_Error

    MarkCount = Lsm5.Hardware.CpStages.MarkCount
    If MarkCount >= 1 Then
        ReDim pos(MarkCount - 1)
        For i = 0 To MarkCount - 1
            Lsm5.Hardware.CpStages.MarkGetZ i, pos(i).X, pos(i).Y, pos(i).Z
        Next i
    End If
    getMarkedStagePosition = pos
   On Error GoTo 0
   Exit Function

getMarkedStagePosition_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure getMarkedStagePosition of Module MicroscopeIO at line " & Erl & " "
End Function


'---------------------------------------------------------------------------------------
' Procedure : setMarkedStagePosition
' Purpose   : set positions in marked stage gui from Vector pos
'---------------------------------------------------------------------------------------
'
Public Sub setMarkedStagePosition(pos() As Vector)
    Dim MarkCount As Long
    Dim i As Long
    On Error GoTo noPos
    If UBound(pos) >= 0 Then
On Error GoTo getMarkedStagePosition_Error
        Lsm5.Hardware.CpStages.MarkClearAll
        SleepWithEvents (500)
        Application.ThrowEvent ePropertyEventStage, 0 'normally not recquired maybe it helps for the update
        For i = 0 To UBound(pos)
            Lsm5.Hardware.CpStages.MarkAddZ pos(i).X, pos(i).Y, pos(i).Z
            SleepWithEvents (100)
        Next i
        SleepWithEvents (500)
        Debug.Print "Marked Positions " & Lsm5.Hardware.CpStages.MarkCount
        If UBound(pos) + 1 > Lsm5.Hardware.CpStages.MarkCount Then
            For i = 0 To UBound(pos)
                Lsm5.Hardware.CpStages.MarkAddZ pos(i).X, pos(i).Y, pos(i).Z
            Next i
        End If
    End If
    Application.ThrowEvent ePropertyEventStage, 0 'normally not recquired maybe it helps for the update
   On Error GoTo 0
   Exit Sub

getMarkedStagePosition_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure getMarkedStagePosition of Module MicroscopeIO at line " & Erl & " "
    Exit Sub
noPos:
    
End Sub

''''
'   WaitForRecentering(Z As Double, Success As Boolean) As Boolean
'   calls the microscope specific WaitForRecentering
'''
Public Function WaitForRecentering(Z As Double, Optional Success As Boolean = False, Optional ZenV As Integer = 2011, Optional Reset As Boolean) As Boolean
    If ZenV = 2010 Then
        If Not WaitForRecentering2010(Z, Success, Reset) Then
            Exit Function
        End If
    End If
    If ZenV > 2010 Then
        If Not WaitForRecentering2011(Z, Success, Reset) Then
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
Public Function WaitForRecentering2010(Z As Double, Optional Success As Boolean = False, Optional Reset As Boolean = True) As Boolean
    WaitForRecentering2010 = WaitForRecentering2011(Z, Success, Reset)
End Function

#If (ZENvC >= 2012) Then
''''
'   WaitForRecentering2011(Z As Double, Success As Boolean) As Boolean
'   Helping function to check if after acquisition focus returns to its correct position
'       [Z] - is value where the central slice should be.
'       [Success] - Tells if central slide has been found before maximal number of iterations
'   Additional remarks: Lsm5.Hardware.CpFocus.Position is not updated correctly after acquisition (CpFocus needs to return to working position) on the other hand
'   Lsm5.DsRecording.Sample0Z keeps track correctly of the position
'''
Public Function WaitForRecentering2011(Z As Double, Optional Success As Boolean = False, Optional Reset As Boolean = False) As Boolean
    Dim cnt As Integer
    Dim MaxCnt As Integer
    Dim pos As Double
    Dim ZOffset As Double
    Dim ReferenceZ As Double
    Dim Prec As Double
    'Dim ScanController As AimScanController
    'Set ScanController = Lsm5.ExternalDsObject.ScanController
        
    Prec = 0.01
    MaxCnt = 6
    cnt = 0
    
    'ScanController.LockAll True ''The lock command may be recquired to properly pass command (without it it seems to work too, leave it out for the moment)'''
    pos = Lsm5.Hardware.CpFocus.position
    ZOffset = getHalfZRange(Lsm5.DsRecording) + pos - Z
    ReferenceZ = Lsm5.DsRecording.ReferenceZ
    
    If isZStack(Lsm5.DsRecording) Then
        Debug.Print "WaitForRecentering ZOffset start " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        While (Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec Or Abs(Z - Lsm5.DsRecording.ReferenceZ) > Prec) And cnt < MaxCnt
            If Reset Then
                'ScanController.LockAll False
                Lsm5.DsRecording.Sample0Z = ZOffset
                Lsm5.DsRecording.ReferenceZ = Z
                'ScanController.LockAll True
            End If
            SleepWithEvents (400)
            cnt = cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        Debug.Print "WaitForRecentering ZOffset " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        If cnt > 0 Then
            LogManager.UpdateWarningLog " Warning " & CurrentFileName & " waitForRecentering recquired " & cnt & " rounds. Reset = " & Reset & ". If False warning happened at end of imaging"
        End If
        'ScanController.LockAll False
        If cnt = MaxCnt And Reset Then
            Lsm5.DsRecording.Sample0Z = ZOffset
            Lsm5.DsRecording.ReferenceZ = Z
            GoTo FailedWaiting
        End If
    End If
    'ScanController.LockAll False
    DoEvents
    Success = True
    WaitForRecentering2011 = True
    Exit Function
FailedWaiting:
    DoEvents
    Success = False
    LogManager.UpdateWarningLog " Warning: " & CurrentFileName & " waitForRecentering forced recentering of the stack"
    WaitForRecentering2011 = True
End Function
#Else
''''
'   WaitForRecentering2011(Z As Double, Success As Boolean) As Boolean
'   Helping function to check if after acquisition focus returns to its correct position
'       [Z] - is value where the central slice should be.
'       [Success] - Tells if central slide has been found before maximal number of iterations
'   Additional remarks: Lsm5.Hardware.CpFocus.Position is not updated correctly after acquisition (CpFocus needs to return to working position) on the other hand
'   Lsm5.DsRecording.Sample0Z keeps track correctly of the position
'''
Public Function WaitForRecentering2011(Z As Double, Optional Success As Boolean = False, Optional Reset As Boolean = False) As Boolean
    Dim cnt As Integer
    Dim MaxCnt As Integer
    Dim pos As Double
    Dim ZOffset As Double
    Dim Prec As Double
    'Dim ScanController As AimScanController
    'Set ScanController = Lsm5.ExternalDsObject.ScanController
        
    Prec = 0.01
    MaxCnt = 6
    cnt = 0
    
    'ScanController.LockAll True ''The lock command may be recquired to properly pass command (without it it seems to work too, leave it out for the moment)'''
    pos = Lsm5.Hardware.CpFocus.position
    ZOffset = getHalfZRange(Lsm5.DsRecording) + pos - Z
    
    If isZStack(Lsm5.DsRecording) Then
        Debug.Print "WaitForRecentering ZOffset start " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        While Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec And cnt < MaxCnt
            If Reset Then
                'ScanController.LockAll False
                Lsm5.DsRecording.Sample0Z = ZOffset
                'ScanController.LockAll True
            End If
            SleepWithEvents (400)
            cnt = cnt + 1
            If ScanStop Then
                Exit Function
            End If
        Wend
        Debug.Print "WaitForRecentering ZOffset " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        If cnt > 0 Then
            LogManager.UpdateWarningLog " Warning " & CurrentFileName & " waitForRecentering recquired " & cnt & " rounds. Reset = " & Reset & ". If False warning happened at end of imaging"
        End If
        'ScanController.LockAll False
        If cnt = MaxCnt And Reset Then
            Lsm5.DsRecording.Sample0Z = ZOffset
            GoTo FailedWaiting
        End If
    End If
    'ScanController.LockAll False
    DoEvents
    Success = True
    WaitForRecentering2011 = True
    Exit Function
FailedWaiting:
    DoEvents
    Success = False
    LogManager.UpdateWarningLog " Warning: " & CurrentFileName & " waitForRecentering forced recentering of the stack"
    WaitForRecentering2011 = True
End Function
#End If


''''
'   Recenter(Z As Double)
'   Sets the central slice. This slice is then maintained even when framespacing is changing.
'       [Z]     - Absolute position of central slice
'   position central slice is Z = Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position
''''
Public Function Recenter_pre(Z As Double, Optional Success As Boolean = False, Optional ZenV As Integer = 2011) As Boolean
    If Not Recenter(Z, ZenV) Then
        Exit Function
    End If
    If Not WaitForRecentering(Z, Success, ZenV) Then
        Exit Function
    End If
    If Not ScanStop Then
        Recenter_pre = True
    End If
End Function

Public Function Recenter_post(Z As Double, Optional Success As Boolean = False, Optional ZenV As Integer = 2011, Optional Reset As Boolean = True) As Boolean
    If Not WaitForRecentering(Z, Success, ZenV, Reset) Then
        Exit Function
    End If
    
    If Not ScanStop Then
        Recenter_post = True
    End If
End Function

Public Function Recenter(Z As Double, Optional ZenV As Integer = 2011) As Boolean
    Dim i As Integer
    If ZenV = 2010 Then
        For i = 1 To 1
            If Not Recenter2010(Z) Then
                Exit Function
            End If
            Sleep (200)
        Next i
    End If
    If ZenV > 2010 Then
        If Not Recenter2011(Z) Then
            Exit Function
        End If
    End If
    Recenter = True
End Function

Public Function Recenter2010(Z As Double) As Boolean
    Recenter2010 = Recenter2011(Z)
End Function

#If ZENvC >= 2012 Then 'because of DsRecording.RecenterZ
Public Function Recenter2011(Z As Double) As Boolean
    Dim pos As Double
    Dim ZOffset As Double
    Dim Prec As Double
    Dim count As Integer
    'Dim ScanController As AimScanController
    'Set ScanController = Lsm5.ExternalDsObject.ScanController
    count = 0
    Prec = 0.001
    

    'ScanController.LockAll True ''The lock command may be recquired to properly pass command (without it it seems to work too)'''
    pos = Lsm5.Hardware.CpFocus.position
    'ScanController.LockAll False
    'If Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then 'Move at the start alwayes if piezo, for ZEN 2012 the best strategy is always to move at the start
    If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
        If Not FailSafeMoveStageZ(Z) Then
            GoTo EndOfFun
        End If
    End If
    pos = Z
    'End If
    ''Only recenter central slice if we have a ZStack
    If isZStack(Lsm5.DsRecording) Then
        ZOffset = getHalfZRange(Lsm5.DsRecording) + pos - Z
        Debug.Print "Recenter at start " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z) & " " & Lsm5.DsRecording.ReferenceZ - Z
        'ScanController.LockAll False
        Lsm5.DsRecording.Sample0Z = ZOffset
        Lsm5.DsRecording.ReferenceZ = Z  'this does not exist for previous Zen versions for ZEN2012 this is absolutely recquired
        'ScanController.LockAll True
        SleepWithEvents (500) 'with 500 it works this makes the imaging slower but more reliable
        Debug.Print "Recenter Offset after 500 ms " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        While count < 3 And (Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec Or Abs(Lsm5.DsRecording.ReferenceZ - Z > Prec))
            LogManager.UpdateLog " Warning: " & CurrentFileName & " Recenter. Problem in settings ZStack on round " & count + 1 & _
            ".  Sample0Z_diff: " & ZOffset - Lsm5.DsRecording.Sample0Z & " ReferenceZ_diff: " & Z - Lsm5.DsRecording.ReferenceZ
            While Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec Or Abs(Lsm5.DsRecording.ReferenceZ - Z > Prec)
                'ScanController.LockAll False
                Lsm5.DsRecording.Sample0Z = ZOffset
                Lsm5.DsRecording.ReferenceZ = Z
                'ScanController.LockAll True
                SleepWithEvents (100)
            Wend
            'ScanController.LockAll False
            SleepWithEvents (200)
            If ScanStop Then
                GoTo EndOfFun
            End If
            count = count + 1
        Wend
        If (Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec) Or (Abs(Lsm5.DsRecording.ReferenceZ - Z) > Prec) Then
            LogManager.UpdateWarningLog " Warning: " & CurrentFileName & " Recenter. Problem in settings ZStack. Sample0Z_diff: " & ZOffset - Lsm5.DsRecording.Sample0Z & " ReferenceZ_diff: " & Z - Lsm5.DsRecording.ReferenceZ
        End If
    End If
    'ScanController.LockAll False
    Recenter2011 = True
    Exit Function
EndOfFun:
    'ScanController.LockAll False
End Function

#Else
Public Function Recenter2011(Z As Double) As Boolean
    
    Dim pos As Double
    Dim ZOffset As Double
    Dim Prec As Double
    Dim count As Integer
    'Dim ScanController As AimScanController
    'Set ScanController = Lsm5.ExternalDsObject.ScanController
    count = 0
    Prec = 0.001
    

    'ScanController.LockAll True ''The lock command may be recquired to properly pass command (without it it seems to work too)'''
    pos = Lsm5.Hardware.CpFocus.position
    'ScanController.LockAll False ''The lock command may be recquired to properly pass command (without it it seems to work too)'''
    
    If Lsm5.DsRecording.SpecialScanMode = "ZScanner" Then 'Move at the start always if piezo
        If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
            If Not FailSafeMoveStageZ(Z) Then
                Exit Function
            End If
        End If
        pos = Z
    End If

    ''Only recenter central slice if we have a ZStack
    If isZStack(Lsm5.DsRecording) Then
        ZOffset = getHalfZRange(Lsm5.DsRecording) + pos - Z
        Debug.Print "Recenter at start " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        'Not clear if this is recquired
        'ScanController.LockAll False
        Lsm5.DsRecording.Sample0Z = ZOffset
        'ScanController.LockAll True
        SleepWithEvents (500) 'with 500 it works
        Debug.Print "Recenter Offset after 500 ms " & Abs(ZOffset - Lsm5.DsRecording.Sample0Z)
        While count < 3 And Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec
            LogManager.UpdateLog " Warning: " & CurrentFileName & " Recenter2011. Problem in settings ZStack on round " & count + 1 & _
            ".  Sample0Z_diff: " & ZOffset - Lsm5.DsRecording.Sample0Z
            While Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec
                'ScanController.LockAll False
                Lsm5.DsRecording.Sample0Z = ZOffset
                'ScanController.LockAll True
                SleepWithEvents (100)
            Wend
            'ScanController.LockAll False
            SleepWithEvents (200)
            If ScanStop Then
                Exit Function
            End If
            count = count + 1
        Wend
        
        If (Abs(ZOffset - Lsm5.DsRecording.Sample0Z) > Prec) Then
            LogManager.UpdateWarningLog " Warning: " & CurrentFileName & " Recenter2011. Problem in settings ZStack. Sample0Z_diff: " & ZOffset - Lsm5.DsRecording.Sample0Z
        End If
    End If
    'ScanController.LockAll False
    'Always move at the end if we did not move before
    If Round(pos, PrecZ) <> Round(Z, PrecZ) Then ' move only if necessary
        If Not FailSafeMoveStageZ(Z) Then
            Exit Function
        End If
    End If
    
    Recenter2011 = True
    
End Function
#End If


''''
' Compute the centerofmass of image stored in RecordingDoc return values according
'   Use channel with name TrackingChannel
''''
Public Function MassCenter(RecordingDoc As DsRecordingDoc, TrackingChannel As Integer, method As Integer) As Vector

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
    Dim frameSpacing As Double
    Dim IntLine() As Variant
    Dim IntCol() As Variant
    Dim IntFrame() As Variant
    
    Dim Frame As Long
    Dim Line As Long
    Dim Col As Long
    Dim XMinMax(1) As Long
    Dim YMinMax(1) As Long
    Dim ZMinMax(1) As Long
    Dim thresh As Double
    DoEvents
       
    'Find the channel to track

    Debug.Print RecordingDoc.ChannelName(TrackingChannel)
    
   
    If TrackingChannel > RecordingDoc.GetDimensionChannels - 1 Or TrackingChannel < -1 Then
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
            scrline = RecordingDoc.ScanLine(TrackingChannel, 0, Frame, Line, spl, bpp)  'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 0 To ColMax - 1             'Now I'm scanning all the pixels in the line
                IntLine(Line) = IntLine(Line) + scrline(Col)
                IntCol(Col) = IntCol(Col) + scrline(Col)
                IntFrame(Frame) = IntFrame(Frame) + scrline(Col)
            Next Col
        Next Line
    Next Frame
    Select Case method
        Case 1 'Center of mass (thr)
            thresh = 0.8
        Case 2 'Center of mass
            thresh = 0
        Case 3 'Peak
            thresh = 0
        Case Else
            GoTo WrongMethod
    End Select
            
            
    'compute center of mass, threshold by 80% the image
    MassCenter.X = weightedMean(IntCol, XMinMax(0), XMinMax(1), thresh)
    MassCenter.Y = weightedMean(IntLine, YMinMax(0), YMinMax(1), thresh)
    MassCenter.Z = weightedMean(IntFrame, ZMinMax(0), ZMinMax(1), thresh)
    If method = 3 Then
        MassCenter.X = XMinMax(1)
        MassCenter.Y = YMinMax(1)
        MassCenter.Z = ZMinMax(1)
    End If
    On Error GoTo 0
    Exit Function
ErrorHandle:
    MsgBox ("Error in MicroscopeIO.MassCenter " + TrackingChannel + " " + Err.Description)
    ScanStop = True
    Exit Function
WrongMethod:
    MsgBox ("Method " & method & " for computing focus is not known")
    ScanStop = True
End Function



''''''
' SaveDsRecordingDoc(Document As DsRecordingDoc, FileName As String) As Boolean
' Copied and adapted from MultiTimeSeries macro
''''''
Public Function SaveDsRecordingDoc(Document As DsRecordingDoc, fileName As String, FileFormat As enumAimExportFormat) As Boolean
    Dim Export As AimImageExport
    Dim image As AimImageMemory
    Dim error As AimError
    Dim Planes As Long
    Dim Plane As Long
    Dim Positions As Long
    Dim Horizontal As enumAimImportExportCoordinate
    Dim Vertical As enumAimImportExportCoordinate

On Error GoTo SaveDsRecordingDoc_Error
    'Set Image = EngelImageToHechtImage(Document).Image(0, True)
    If Not Document Is Nothing Then
        Set image = Document.RecordingDocument.image(0, True)
    End If
    
    Set Export = Lsm5.CreateObject("AimImageImportExport.Export.4.5")
    'Set Export = New AimImageExport
    Export.fileName = fileName
    Export.format = FileFormat
    Export.StartExport image, image
    Set error = Export
    error.LastErrorMessage
    
    Planes = 1
    Export.GetPlaneDimensions Horizontal, Vertical
    If Document.Recording.MultiPositionAcquisition Then
        Positions = Document.Recording.MultiPositionArraySize
    End If
    If Positions = 0 Then
        Positions = 1
    End If
    
    Select Case Vertical
        Case eAimImportExportCoordinateY:
             Planes = image.GetDimensionZ * image.GetDimensionT * Positions
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
   On Error GoTo 0
   Exit Function

SaveDsRecordingDoc_Error:
    ScanStop = True
    Export.FinishExport
    StopAcquisition
    MsgBox "Check Temporary Files Folder! Cannot Save Temporary File(s)!"
    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure SaveDsRecordingDoc of Module MicroscopeIO at line " & Erl & " "
End Function

'''
' Check if we have a ZStack
''''
Public Function isZStack(ARecording As DsRecording) As Boolean
    Dim ScanMode As String
    ScanMode = ARecording.ScanMode
    If ScanMode = "ZScan" Or ScanMode = "Stack" Then
        isZStack = True
    End If
End Function


'''
' Compute half the size of the ZRange
'''
Public Function getHalfZRange(ARecording As DsRecording) As Double
    If isZStack(ARecording) Then
        getHalfZRange = ARecording.frameSpacing * (ARecording.framesPerStack - 1) / 2
    Else
        getHalfZRange = 0
    End If
End Function


'''''
'   UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
'   Ask which system is the macro runnning on
'       [bLSM]  In/Out - True if LSM system
'       [bLive] In/Out - True for LIVE system
'       [bCamera] In/Out - True if Camera is used
''''
Public Sub UsedDevices40(bLSM As Boolean, bLIVE As Boolean, bCamera As Boolean)
    Dim ScanController As AimScanController
    Dim TrackParameters As AimTrackParameters
    Dim Size As Long
    Dim lTrack As Long
    Dim eDeviceMode As Long

    bLSM = False
    bLIVE = False
    bCamera = False
    Set ScanController = Lsm5.ExternalDsObject.ScanController
    Set TrackParameters = ScanController.TrackParameters
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

'''''
''   Autofocus_MoveAcquisition_HRZ(ZOffset As Double) may be interesting to introduce it back for speed reasons
''   Allow to use HRZ for Move Z-stage (not used at the moment)
'''''
'Public Sub Autofocus_MoveAcquisition_HRZ(ZOffset As Double)
'    Dim NoZStack As Boolean
'    Const ZBacklash = -50
'    Dim ZFocus As Double
'    Dim Zbefore As Double
'    Dim X As Double
'    Dim Y As Double
'
'    AutofocusForm.RestoreAcquisitionParameters
'
'    Set GlobalBackupRecording = Nothing
'    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
'    DoEvents
'
'    NoZStack = True
'    If GlobalAcquisitionRecording.ScanMode = "ZScan" Or GlobalAcquisitionRecording.ScanMode = "Stack" Then  'Looks if a Z-Stack is going to be acquired
'        NoZStack = False
'    End If
'
'    'Moving to the correct position in Z
'    If AutofocusForm.AutofocusHRZ.Value And NoZStack Then                                            'If using HRZ for autofocusing and there is no Zstack for image acquisition
'        Lsm5.Hardware.CpHrz.Stepsize = 0.2
'        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
'        DoEvents
'     '   ZFocus = Lsm5.Hardware.CpHrz.Position + ZShift - ZOffset
'
'     'Defines the new focus position as the actual position plus the shift and goes back to the object position (that's why you need the offset)
'
'        ZFocus = Lsm5.Hardware.CpHrz.position + ZOffset + ZShift
'
'        Lsm5.Hardware.CpHrz.position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the HRZ can acces the position, and also to wait it is done...
'
'        DoEvents
'
'    Else                                        'either there is a Z stack for image acquisition or we're using the focuswheel for autofocussing
'        If AutofocusForm.AutofocusHRZ.Value Then                             ' Now I'm not sure with the signs and... I some point I just tried random combinations...
'            ZFocus = Lsm5.Hardware.CpHrz.position - ZOffset - ZShift '         'ZBefore corresponds to the position where the focuswheel was before doing anything. Zshift is the calculated shift
'        Else                                    'If the HRZ is not calibrated the Z shift might be wrong
'            ZFocus = Zbefore + ZShift
'        End If
'
'        Lsm5.Hardware.CpHrz.position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'    End If
'
'    'Moving to the correct position in X and Y
'
'    If AutofocusForm.ScanFrameToggle Then
'        If AutofocusForm.AutofocusTrackXY Then
'            X = Lsm5.Hardware.CpStages.PositionX - XShift  'the fact that it is "-" in this line and "+" in the next line  probably has to do with where the XY of the origin is set (top right corner and not botom left, I think)
'            Y = Lsm5.Hardware.CpStages.PositionY - YShift
'            Success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(X, Y)
'        End If
'
'        Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            If ScanStop Then
'                Lsm5.StopScan
'                AutofocusForm.StopAcquisition
'                DisplayProgress "Stopped", RGB(&HC0, 0, 0)
'                Exit Sub
'            End If
'            DoEvents
'            Sleep (5)
'        Loop
'    End If
'
'
'    DisplayProgress "Autofocus 14", RGB(0, &HC0, 0)
'    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
'    DoEvents
'    DisplayProgress "Autofocus 15", RGB(0, &HC0, 0)
'End Sub





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



