VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutofocusForm 
   Caption         =   "AutofocusScreen"
   ClientHeight    =   13065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   OleObjectBlob   =   "AutofocusForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "AutofocusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Private Sub CheckBoxInnactivateAutofocus_Click()    ' changes look of inactivated buttom if checked and verfies that the posacquisition Z tracking is not activated if autofocusing is reactivated
                                                    
    If CheckBoxInnactivateAutofocus.Value = False Then
              CheckBoxInnactivateAutofocus.BackColor = &HFFFFFF
     
    Else
        CheckBoxInnactivateAutofocus.BackColor = 33023
    End If
End Sub



Private Sub CheckBoxLowZoom_Click()
'There is nothing to do when clicking on this
End Sub








Private Sub CommandButton1_Click()

    Dim dblTask As Double
    Dim MacroPath As String
    Dim Mypath As String
    Dim MyPathPDF As String
    
    Dim bslash As String
    Dim success As Integer
    Dim pos As Integer
    Dim Start As Integer
    Dim Count As Long
    Dim ProjName As String
    Dim indx As Integer
    Dim AcrobatObject As Object
    Dim AcrobatViewer As Object
    Dim OK As Boolean
    Dim StrPath As String
    Dim ExecName As String
    
       
    
    Count = ProjectCount()
    For indx = 0 To Count - 1
        MacroPath = ProjectPath(indx, success)
        ProjName = ProjectTitle(indx, success)
        If StrComp(ProjName, GlobalProjectName, vbTextCompare) = 0 Then
            Start = 1
            bslash = "\"
            pos = Start
            Do While pos > 0
                pos = InStr(Start, MacroPath, bslash)
                If pos > 0 Then
                    Start = pos + 1
                End If
            Loop
            Mypath = Strings.Left(MacroPath, Start - 1)
            MyPathPDF = Mypath + HelpNamePDF

            OK = False
            On Error GoTo RTFhelp
            OK = FServerFromDescription("AcroExch.Document", StrPath, ExecName)
            dblTask = Shell(ExecName + " " + MyPathPDF, vbNormalFocus)
            
RTFhelp:
            If Not OK Then
                MsgBox "Install Acrobat Viewer!"
            End If
            Exit For
        End If
    Next indx
End Sub

Private Sub CommandButtonStoreApply_Click()
 StoreApplyForm.Show 0
 End Sub































Private Sub ScanLineToggle_Click()
    ScanFrameToggle.Value = Not ScanLineToggle.Value 'if ScanFrame is true ScanLine is false (you can only chose one of them)
    FrameAutofocussing = ScanFrameToggle.Value 'if ScanFrame is true than FrameAutofocusing (boolean variable) will be set true as well
    FrameSizeLabel.Visible = ScanFrameToggle.Value 'FrameSize Label is only displayed if ScanFrame is activated
    BSliderFrameSize.Visible = ScanFrameToggle.Value 'FrameSize Slider is only displayed if ScanFrame is activated
    BSliderScanSpeed.Visible = ScanLineToggle.Value
    ScanSpeedLabel.Visible = ScanLineToggle.Value
End Sub

Private Sub ScanFrameToggle_Click()
    ScanLineToggle.Value = Not ScanFrameToggle.Value 'if ScanLine is chosen, ScanFrame will be unchecked
    FrameAutofocussing = ScanFrameToggle.Value 'if ScanFrame is true than FrameAutofocusing (boolean variable) will be set true
    FrameSizeLabel.Visible = ScanFrameToggle.Value
    
    BSliderScanSpeed.Visible = ScanLineToggle.Value
    ScanSpeedLabel.Visible = ScanLineToggle.Value

'         If SystemName = "LSM" Then
'
'            BSliderFrameSize.ValueEditable = True
'             BSliderFrameSize.Min = 16
'            BSliderFrameSize.Max = 1024
'            BSliderFrameSize.Step = 128
'            BSliderFrameSize.StepSmall = 4
'           BSliderFrameSize.ValueDisplay = True
'
'        ElseIf SystemName = "LIVE" Then
'
'
'            BSliderFrameSize.ValueEditable = False
'            BSliderFrameSize.Min = 128
'            BSliderFrameSize.Max = 1024
'            BSliderFrameSize.Step = 128
'            BSliderFrameSize.StepSmall = 128
'            BSliderFrameSize.Value = 128
'
'        End If
   
  BSliderFrameSize.Visible = ScanFrameToggle.Value
  
End Sub

Private Sub SetFocusButton_Click()
    AutofocusForm.GetBlockValues                                             'Updates the parameters value for BlockZRange, BlockZStep....
    SetFocus BlockZRange, BlockZStep, BlockLowZoom, BlockHighSpeed, BlockZOffset  ' Performs the scan in Z (line or Frame, to find the offset value
End Sub



Private Sub AutofocusButton_Click()
    Dim AutofocusDoc As DsRecordingDoc
    Try = 1
    AutofocusForm.GetBlockValues 'Updates the parameters value for BlockZRange, BlockZStep..
    DisplayProgress "Autofocus 0", RGB(0, &HC0, 0)
    Lsm5.StopScan
    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
    StoreAquisitionParameters
    Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
    If ScanStop = True Then
        GoTo Abort
    End If
    'check if Z shift makes sense
    CheckRefControl BlockZRange
    If CheckBoxMoveHRZ.Value = True Then
    Autofocus_MoveAquisition_HRZ BlockLowZoom, BlockZOffset
    Else
    Autofocus_MoveAquisition BlockLowZoom, BlockZOffset
    End If
    If ScanStop = True Then
        GoTo Abort
    End If
    'DoAutofocus BlockZOffset, BlockZRange, BlockZStep, BlockHRZ, BlockLowZoom, BlockHighSpeed  ' Performs the scan in Z (line or Frame, to find the offset value

    ActivateAcquisitionTrack
    If IsAcquisitionTrackSelected And IsAutofocusTrackSelected Then
        Sleep (20)
        Set AutofocusDoc = Lsm5.StartScan
    Do While AutofocusDoc.IsBusy                                  ' Waiting untill the image acquisition is done
        If ScanStop Then
            Lsm5.StopScan
            GoTo Abort
        End If
        DoEvents
        Sleep (10)
    Loop
        Lsm5.tools.WaitForScanEnd False, 20       'Wait untill the scan is finnished, the waiting time is 20s. This could be too short in some instances
    Else
     GoTo Abort
    End If
Abort:
If Not (GlobalBackupRecording Is Nothing) Then
    RestoreAquisitionParameters
    Set GlobalBackupRecording = Nothing
   Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents                                'Finnish everything which had started
    'ActivateAcquisitionTrack                'Activates the tracks for image acquisition
End If
    If ScanStop = True Then
        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
        ScanStop = False
    Else
        DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
    End If
End Sub

Private Sub StartBleachButton_Click()
    Dim success As Integer
    AutomaticBleaching = False
    If LocationTracking And TrackingChannelString = "" Then
        MsgBox ("Select a channel for tracking, or uncheck the tracking button")
        Exit Sub
    End If
    If MultipleLocation And Lsm5.Hardware.CpStages.MarkCount < 1 Then
        MsgBox ("Select at least one location in the stage control window, or uncheck the multiple location button")
        Exit Sub
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No Database selected ! Cannot start acquisition.")
        Exit Sub
    End If
    Set Track = Lsm5.DsRecording.TrackObjectBleach(success)
    If success Then
        If Track.BleachPositionZ <> 0 Then
            MsgBox ("This macro does not enable to bleach at a different Z position. Please uncheck the corresponding box in the Bleach Control Window")
            Exit Sub
        End If
        If Lsm5.IsValidBleachRoi Then
            If (Track.BleachScanNumber + 1) > BlockRepetitions Then
                MsgBox ("Not enough repetitions to bleach")
                Exit Sub
            End If
            FillBleachTable
            AutomaticBleaching = True
            Track.UseBleachParameters = True
        Else
            MsgBox ("A bleaching ROI needs to be defined to start the macro in the bleaching mode")
            Exit Sub
        End If
    Else
        MsgBox ("A bleach track needs to be defined to start the macro in the bleaching mode")
        Exit Sub
    End If
    StartAcquisition
End Sub

Private Sub FillBleachTable()                              'Fils a table for the macro to know when the bleaches have to occur. This works for FRAPs (and FLIPS if working with LSM 3.2)
    Set Track = Lsm5.DsRecording.TrackObjectBleach(success)
    If success Then
        ReDim BleachTable(BlockRepetitions)               'The bleach table contains as many timepoints as blockrepetitions
    'When working with the Lsm 2.8, remove all this test, except the one indicated line
        If Track.EnableBleachRepeat Then
            For i = Track.BleachScanNumber + 1 To BlockRepetitions Step Track.BleachRepeat
                BleachTable(i) = True
            Next
        Else
            BleachTable(Track.BleachScanNumber + 1) = True  'This is the only line to be kept when working with the Lsm 2.8
        End If
    End If
End Sub

Private Sub StartButton_Click()
    Try = 1
    AutomaticBleaching = False                                  'We do not do FRAps or FLIPS in this case. Bleaches can still be done with the "ExtraBleach" button.
    If LocationTracking And TrackingChannelString = "" Then
        MsgBox ("Select a channel for tracking, or uncheck the tracking button")
        Exit Sub
    End If
    If MultipleLocation And Lsm5.Hardware.CpStages.MarkCount < 1 Then
        MsgBox ("Select at least one location in the stage control window, or uncheck the multiple location button")
        Exit Sub
    End If
    If GlobalDataBaseName = "" Then
        MsgBox ("No Database selected ! Cannot start acquisition.")
        Exit Sub
    End If
    
    StartAcquisition 'This is the main function of the macro
    
End Sub

Private Sub StartAcquisition()
     Dim rettime, difftime As Double
    Dim GlobalPrvTime As Double
    Dim Location As Integer
    Dim LocationNumber As Integer
    Dim LocationName As String
    Dim name As String
    Dim tilename As String
    Dim x As Double
    Dim XCor As Double
    Dim y As Double
    Dim YCor As Double
    Dim z As Double
    Dim ZCor As Double
    Dim ZBacklash As Double                 'I forgot to initialize this to -50
    Dim success As Integer
    Dim RelativeLocation As Integer
    Dim StitchImage As RecordingDocument
    Dim ScanImage As RecordingDocument
    Dim ImageCopy As New AimImageCopy
    Dim Progress As AimProgress
    Dim Scancontroller As AimScanController
    Dim TileDatabaseName As String
    Dim NameLength As Integer
    Dim Myname As String
    Dim Mypath As String
    Dim TileXOld As Integer
    Dim r As Integer
    
    
    
            '''''''''''''''''''''''''create stiching database''''''''''''''''''''
TileX = AutofocusForm.TextBoxTileX.Value
TileY = AutofocusForm.TextBoxTileY.Value

If TileX > 1 Or TileY > 1 Then
            
            Start = 1
            bslash = "\"
            pos = Start
            Do While pos > 0
                pos = InStr(Start, GlobalDataBaseName, bslash)
                If pos > 0 Then
                    Start = pos + 1
                End If
            Loop
            Mypath = Strings.Left(GlobalDataBaseName, Start - 1)
            NameLength = Len(GlobalDataBaseName)
            Myname = Strings.Right(GlobalDataBaseName, NameLength - Start + 1)
            NameLength = Len(Myname)
            Myname = Strings.Left(Myname, NameLength - 4)
            TileDatabaseName = Mypath & Myname & "_tile.mdb"
            Lsm5.NewDatabase (TileDatabaseName)
            TileDatabaseName = TileDatabaseName & "\" & Myname & "_tile.mdb"
End If
''''''''''''''''''''end create stiching database
    
    
    InitializeStageProperties
    SetStageSpeed 9, True
    
    GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
    If MultipleLocation Or Grid Then
    PutStagePositionsInArray
    End If
    
    
            
    RepetitionNumber = 1
    Running = True                              'Now we're starting. This will be set to false if the sopt button is pressed or if we reached the total number of repetitions.
    StartButton.Enabled = False
    StartBleachButton.Enabled = False
    CloseButton.Enabled = False
    ReinitializeButton.Enabled = False
   'SetFocusButton.Enabled = False
   'AutofocusButton.Enabled = False
   If TileX > 1 Or TileY > 1 Then
    Set Scancontroller = Lsm5.ExternalDsObject.Scancontroller
   End If
    
    If MultipleLocation Or Grid Then                    'Defines the Location Number parameter
        LocationNumber = Lsm5.Hardware.CpStages.MarkCount       'Counts the locations stored in the Stage control window from the LSM
    Else
        LocationNumber = 1                                      'If using the single location you do not have to mark it in the stage control window.
    End If
    If LocationTracking Or FrameAutofocussing Then
'        Excel.Application.Visible = True                               'The Excel stuff is to store the XYZ position of the cells at each time point
'        Set PositionData = Excel.Workbooks.Add
'        For Location = 1 To LocationNumber
'            PositionData.Sheets.Add
'            PositionData.ActiveSheet.name = "Location " & Location
'            PositionData.ActiveSheet.Columns("A:A").Select
'            Selection.NumberFormat = "m/d/yyyy h:mm:ss"
'            PositionData.ActiveSheet.Cells(1, 1) = "Time"
'            PositionData.ActiveSheet.Cells(1, 2) = "X (µm)"
'            PositionData.ActiveSheet.Cells(1, 3) = "Y (µm)"
'            PositionData.ActiveSheet.Cells(1, 4) = "Z (µm)"
'            PositionData.ActiveSheet.Cells(1, 6) = "Time delay"
'            PositionData.ActiveSheet.Columns("F:F").Select
'            Selection.NumberFormat = "[h]:mm:ss"
'            PositionData.ActiveSheet.Cells(1, 7) = "Total Distance (µm)"
'        Next Location
    End If
    
    
    Do While Running                                    'As long as the macro is running we're in this loop
    If CheckBoxZMap.Value Then
        Location = 1
        Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ (0)  'Moves to the first location marked in the stage control
                Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
                    If ScanStop Then        'now when we're waiting for things to happen we allow the user to stop the macro
                        Lsm5.StopScan
                        StopAcquisition
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        Exit Sub
                    End If
                    DoEvents
                    Sleep (5)
                Loop
   
        AutofocusForm.GetBlockValues                                'Updates the parameters value for BlockZRange, BlockZStep..
        Lsm5.StopScan
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
        DoEvents
        StoreAquisitionParameters
        Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
        'check if Z shift makes sense
        CheckRefControl BlockZRange
        'Caluclate new z Position, Store Z in Array
'        If Grid Or MultipleLocation Then
'        Lsm5.Hardware.CpStages.MarkClearAll
'        For idpos = 1 To GlobalPositionsStage
'            GlobalZpos(idpos) = GlobalZpos(idpos) + ZShift
'            Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
'        Next idpos
'        Else
'        GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
'            For idpos = 1 To GlobalPositionsStage
'            success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).GetMarkZ(0, x, y, z)
'           success = Lsm5.Hardware.CpStages.MarkClear(0)
'            z = z + ZShift
'            Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ x, y, z
'        Next idpos
'        End If
        Sleep (100)
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
                 DoEvents
         Sleep (100)
                  
        RestoreAquisitionParameters
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
                 DoEvents
    End If
    
    If Not (LocationTracking Or FrameAutofocussing) Then
        If Grid Or MultipleLocation Then
            Lsm5.Hardware.CpStages.MarkClearAll
            For idpos = 1 To GlobalPositionsStage
                GlobalZpos(idpos) = GlobalZpos(idpos) - ZShift
                Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
            Next idpos
            Else
            GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
                For idpos = 1 To GlobalPositionsStage
                success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).GetMarkZ(0, x, y, z)
               success = Lsm5.Hardware.CpStages.MarkClear(0)
                z = z + ZShift
                Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ x, y, z
            Next idpos
            End If
    End If
    
    
        For Location = 1 To LocationNumber              'This loops all the locations (only one if Single location is selected
            If MultipleLocation Or Grid Then
                Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ (0)  'Moves to the first location marked in the stage control
                Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
                    If ScanStop Then        'now when we're waiting for things to happen we allow the user to stop the macro
                        Lsm5.StopScan
                        StopAcquisition
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        Exit Sub
                    End If
                    DoEvents
                    Sleep (5)
                Loop
            End If
    
            If CheckBoxInnactivateAutofocus Then  ' Looking if needs to perform an autofocus
                ZShift = 0
            Else
                
                If CheckBoxZMap.Value Then GoTo AfterAutofocus
                  
                'MsgBox ("doing AF")
                 AutofocusForm.GetBlockValues
                 DisplayProgress "Autofocus 0", RGB(0, &HC0, 0)
                 Lsm5.StopScan
                 Lsm5Vba.Application.ThrowEvent eRootReuse, 0
                 DoEvents
                 StoreAquisitionParameters
                 Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
                 CheckRefControl BlockZRange
                 
                 Autofocus_MoveAquisition BlockLowZoom, BlockZOffset
                ' DoAutofocus BlockZOffset, BlockZRange, BlockZStep, BlockHRZ, BlockLowZoom, BlockHighSpeed
            End If
          
    
AfterAutofocus:
           AutofocusForm.ActivateAcquisitionTrack
           Sleep (100) 'laser hast to swtich from standby to on
            If Not IsAcquisitionTrackSelected Then      'An additional control....
                StopAcquisition
                MsgBox "No track selected for Acquisition! Cannot Acquire!"
                DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
                Exit Sub
            End If
    
            If BleachTable(RepetitionNumber) = True Then                                            'Check if we're performing a bleach before image acquisition
                Set Track = Lsm5.DsRecording.TrackObjectBleach(success)
                If success Then
                    DisplayProgress "Bleaching", &HFF00FF
                    DoEvents
                    Track.UseBleachParameters = True            'Bleach parameters are lasers lines, bleach iterations... stored in the bleach control window
'                    BleachStartTable(RepetitionNumber) = Lsm5.ExternalCpObject.pHardwareObjects.pScanController.GetDspTime
                    BleachStartTable(RepetitionNumber) = GetTickCount       'Get the time right before bleach to store this in the image metadata
                    Lsm5.Bleach 0
                    Lsm5.tools.WaitForScanEnd False, 1                      'Waits for the end of the bleach during one second, I think this is not long enough
                    BleachStopTable(RepetitionNumber) = GetTickCount       'Get the time right after bleach to store this in the image metadata
'                    BleachStopTable(RepetitionNumber) = Lsm5.ExternalCpObject.pHardwareObjects.pScanController.GetDspTime
                    Track.UseBleachParameters = False  'switch off the bleaching
                Else
                    MsgBox ("Could not set bleach track. Did not bleach.")
                End If
                If Location = LocationNumber Then   'Alowas again to do an extrableach at the en
                    ExtraBleachButton.Caption = "Bleach"
                    ExtraBleachButton.BackColor = &H8000000F
                End If
            End If
    
            DisplayProgress "Acquiring location " & Location & ", repetition " & RepetitionNumber, RGB(&HC0, &HC0, 0)  'Now we're going to do the acquisition
            Lsm5.DsRecording.TimeSeries = True              'This is for the concatenation I think: we're doing a timeseries with one timepoint. I'm not sure why is the reason for this
            Lsm5.DsRecording.StacksPerRecord = 1
            
  
            Sleep (100)
            'Set Track = Lsm5.DsRecording.TrackObjectByIndex(3, Success)
            
            
         If TileX > 1 Or TileY > 1 Then
            
            If Lsm5.DsRecording.ScanMode = "Stack" Then
                Set ScanImage = Lsm5.ExternalDsObject.MakeNewImageDocument(Lsm5.DsRecording.RtRegionWidth, _
                                                               Lsm5.DsRecording.RtRegionWidth, _
                                                               Lsm5.DsRecording.FramesPerStack, _
                                                               Lsm5.DsRecording.StacksPerRecord, _
                                                               Lsm5.DsRecording.NumberOfChannels, _
                                                               Bytesperpixel, _
                                                               0)
            Else
                Set ScanImage = Lsm5.ExternalDsObject.MakeNewImageDocument(Lsm5.DsRecording.RtRegionWidth, _
                                                               Lsm5.DsRecording.RtRegionWidth, _
                                                               1, _
                                                               Lsm5.DsRecording.StacksPerRecord, _
                                                               Lsm5.DsRecording.NumberOfChannels, _
                                                               Bytesperpixel, _
                                                               0)
            End If
        
                   
            If ScanImage Is Nothing Then Exit Sub
                
            ScanImage.NeverAgainScanToTheImage
            
            'MsgBox "taking the tile imgae"
            Scancontroller.DestinationImage(0) = ScanImage.Image(0, True)
            Scancontroller.InitializeDestinationImages eGrabModeSingle
            Scancontroller.StartGrab eGrabModeSingle
            While Scancontroller.IsGrabbing
                DoEvents
                Sleep (10)
                If ScanStop Then
                   Lsm5.StopScan
                   StopAcquisition
                   DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                   Exit Sub
                End If
            Wend
            
        Else ' NO TILING
            
            ' HERE THE TILING IMAGE IS TAKEN!!!!
            Set RecordingDoc = Lsm5.StartScan()         'Start image acquisition
          
              End If
          
          
            If RepetitionNumber = 1 Then
                StartTime = GetTickCount            'Get the time when the acquisition was started
            End If
            If MultipleLocation Or Grid Then                'Sets the name of the image to store in the database
                LocationName = "_L" & Location
            Else
                LocationName = ""
            End If
            If Grid Then
                LocationName = "_" & GlobalLocationsName(Location) & LocationName
            End If
            
            If TileX > 1 Or TileY > 1 Then
                name = GlobalFileName & LocationName & "_R" & RepetitionNumber
                ScanImage.SetTitle name
'            Do While ScanImage.IsBusy                        'Wait the end of the scan
'               If ScanStop Then
'                    Lsm5.StopScan
'                    StopAcquisition
'                    DisplayProgress "Stopped", RGB(&HC0, 0, 0)
'                    Exit Sub
'               End If
'               DoEvents
'               Sleep (5)
'            Loop
            Else
                name = GlobalFileName & LocationName & "_R" & RepetitionNumber
                RecordingDoc.SetTitle name
                Do While RecordingDoc.IsBusy                        'Wait the end of the scan
                   If ScanStop Then
                        Lsm5.StopScan
                        StopAcquisition
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        Exit Sub
                   End If
                   DoEvents
                   Sleep (5)
                Loop
            End If
            Lsm5.tools.WaitForScanEnd False, 10
            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''stitching
        If TileX > 1 Or TileY > 1 Then    'step into stitching routine if you have more than 1 frame at each location
            RelativeLocation = Location Mod (TileX * TileY)
            If RelativeLocation = 0 Then RelativeLocation = (TileX * TileY)
            
            If Lsm5.DsRecording.TrackObjectByIndex(0, success).DataChannelObjectByIndex(0, success).BitsPerSample > 8 Then
                Bytesperpixel = 2
            Else
                Bytesperpixel = 1
            End If
         
        
                If RelativeLocation = 1 Then 'at each first frame of a new tile group define a new image
            If AreStageCoordinateExchanged Then
                If Lsm5.DsRecording.ScanMode = "Stack" Then
                    Set StitchImage = Lsm5.ExternalDsObject.MakeNewImageDocument(CLng(Lsm5.DsRecording.RtRegionWidth * TileY), _
                                                                     CLng(Lsm5.DsRecording.RtRegionHeight * TileX), _
                                                                     Lsm5.DsRecording.FramesPerStack, _
                                                                     1, _
                                                                     Lsm5.DsRecording.NumberOfChannels, _
                                                                     Bytesperpixel, _
                                                                     1)
                  Else
                    Set StitchImage = Lsm5.ExternalDsObject.MakeNewImageDocument(CLng(Lsm5.DsRecording.RtRegionWidth * TileY), _
                                                                     CLng(Lsm5.DsRecording.RtRegionHeight * TileX), _
                                                                     1, _
                                                                     1, _
                                                                     Lsm5.DsRecording.NumberOfChannels, _
                                                                     Bytesperpixel, _
                                                                     1)
                    End If
        Else
                 If Lsm5.DsRecording.ScanMode = "Stack" Then
                    Set StitchImage = Lsm5.ExternalDsObject.MakeNewImageDocument(CLng(Lsm5.DsRecording.RtRegionWidth * TileX), _
                                                                     CLng(Lsm5.DsRecording.RtRegionHeight * TileY), _
                                                                     Lsm5.DsRecording.FramesPerStack, _
                                                                     1, _
                                                                     Lsm5.DsRecording.NumberOfChannels, _
                                                                     Bytesperpixel, _
                                                                     1)
                  Else
                    Set StitchImage = Lsm5.ExternalDsObject.MakeNewImageDocument(CLng(Lsm5.DsRecording.RtRegionWidth * TileX), _
                                                                     CLng(Lsm5.DsRecording.RtRegionHeight * TileY), _
                                                                     1, _
                                                                     1, _
                                                                     Lsm5.DsRecording.NumberOfChannels, _
                                                                     Bytesperpixel, _
                                                                     1)
                    End If
            End If
                                                                     '''''overlap is still missing
                                                                     
                    If StitchImage Is Nothing Then Exit Sub
                    
                
                End If
              
                StitchImage.NeverAgainScanToTheImage
            
       ' ImageCopy.SourceImage = RecordingDoc.Image(0, False)
  
        ImageCopy.SourceImage = ScanImage.Image(0, False)
        ImageCopy.DestinationImage = StitchImage.Image(0, False)
        
'        If RelativeLocation Mod TileY = 0 Then
'        ImageCopy.DestinationY = 0
'        Else
If AreStageCoordinateExchanged Then
If RelativeLocation = 1 Then r = 1
 ImageCopy.DestinationX = (TileY - r) * Lsm5.DsRecording.RtRegionWidth
 If RelativeLocation Mod TileX = 0 Then r = r + 1
'If RelativeLocation Mod TileX = 0 Then
'            ImageCopy.DestinationX = 0
'        Else
'
' ImageCopy.DestinationX = CLng(Abs(1 - Int((RelativeLocation - 1) / TileX)) * Lsm5.DsRecording.RtRegionWidth)
'
'      End If
       If RelativeLocation Mod TileX = 0 Then
            ImageCopy.DestinationY = 0
        Else
            ImageCopy.DestinationY = CLng((TileX - (RelativeLocation Mod TileX)) * Lsm5.DsRecording.RtRegionWidth)
        End If
Else
        ImageCopy.DestinationY = CLng(Int((RelativeLocation - 1) / TileX) * Lsm5.DsRecording.RtRegionWidth)
        
'        End If
       If RelativeLocation Mod TileX = 0 Then
            ImageCopy.DestinationX = 0
        Else
            ImageCopy.DestinationX = CLng(Abs((RelativeLocation Mod TileX) - TileX) * Lsm5.DsRecording.RtRegionWidth)
        End If
 End If
        
        If RelativeLocation = 1 Then
            ImageCopy.ImageParameterCopyFlags = eAimImageParameterCopyAll
            StitchImage.SetVoxelSizeX CLng(Lsm5.DsRecording.RtRegionWidth * TileX)
            StitchImage.SetVoxelSizeY CLng(Lsm5.DsRecording.RtRegionHeight * TileY)
            
        Else
            ImageCopy.ImageParameterCopyFlags = 0
        End If
            
        ImageCopy.Start
        Set Progress = ImageCopy
        
        While Not Progress.Ready
           DoEvents
           Sleep (10)
           If ScanStop Then Exit Sub
        Wend
            
            
     If RelativeLocation = TileX * TileY Then
     
        tilename = "Tile_" & GlobalLocationsName(Location) & "_L" & (Location / RelativeLocation) & "_R" & RepetitionNumber
        StitchImage.SetTitle tilename
        
        
       
         GlobalImageIndex(RepetitionNumber) = StitchImage.SaveToDatabase(TileDatabaseName, tilename)
         StitchImage.CloseAllWindows
    End If
    End If
''''''''''''''''''''''''''''''''''''''''''''''''end stitching
            
            
            
            If BleachStartTable(RepetitionNumber) > 0 Then          'If a bleach was performed we add the information to the image metadata
                Lsm5.DsRecordingActiveDocObject.AddEvent (BleachStartTable(RepetitionNumber) - StartTime) / 1000, eEventTypeBleachStart, "Bleach Start"
                Lsm5.DsRecordingActiveDocObject.AddEvent (BleachStopTable(RepetitionNumber) - StartTime) / 1000, eEventTypeBleachStop, "Bleach End"
            End If
            'Now we save the image
            
            If TileX > 1 Or TileY > 1 Then
                GlobalImageIndex(RepetitionNumber) = ScanImage.SaveToDatabase(GlobalDataBaseName, name)
            Else
                GlobalImageIndex(RepetitionNumber) = RecordingDoc.SaveToDatabase(GlobalDataBaseName, name)      'This is a strange way to call the function to save images, but it works
            End If
            If ScanStop Then
                Lsm5.StopScan
                StopAcquisition
                DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                Exit Sub
            End If
            
            If LocationTracking Or FrameAutofocussing Then
'                PositionData.Sheets("Location " & Location).Select
'                PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 1) = CDate(Lsm5.DsRecordingActiveDocObject.Recording.Sample0Time)
'                PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 2) = Lsm5.Hardware.CpStages.PositionX
'                PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 3) = Lsm5.Hardware.CpStages.PositionY
'                PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 4) = Lsm5.Hardware.CpFocus.Position
'                PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 6) = PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 1) - PositionData.ActiveSheet.Cells(2, 1)
'                If RepetitionNumber > 1 Then
'                    PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 7) = PositionData.ActiveSheet.Cells(RepetitionNumber, 7) + Sqr((PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 2) - PositionData.ActiveSheet.Cells(RepetitionNumber, 2)) ^ 2 + (PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 3) - PositionData.ActiveSheet.Cells(RepetitionNumber, 3)) ^ 2)
'                Else
'                    PositionData.ActiveSheet.Cells(RepetitionNumber + 1, 7) = 0
'                End If
            End If
            
            'This is if we're doing some postacquisition tracking
            If LocationTracking Then
                DisplayProgress "Analysing the new position of location " & Location, &H80FF&
                DoEvents
                MassCenter ("Tracking")
                XCor = XMass
                YCor = -YMass
                If TrackZ Then
                    ZCor = ZMass
                Else
                    If BlockHRZ Then
                        ZCor = 0
'                        Success = Lsm5.Hardware.CpHrz.Leveling
                    Else
                        ZCor = 0
                    End If
                End If
                '''''changed
            If AreStageCoordinateExchanged Then
            XCor = YMass
            YCor = XMass
            End If
            '''changed
            Else
                XCor = 0
                YCor = 0
                If BlockHRZ Then
                    ZCor = 0
                    success = Lsm5.Hardware.CpHrz.Leveling   'This I think puts the HRZ to its resting position, and moves the focuswheel correspondingly
                Else
                    ZCor = 0
                End If
            End If
            
            x = Lsm5.Hardware.CpStages.PositionX - XCor                     'Records the current X,Y,Z positions
            y = Lsm5.Hardware.CpStages.PositionY - YCor
            z = Lsm5.Hardware.CpFocus.Position + ZCor
            
            If MultipleLocation Or Grid Then
                success = Lsm5.Hardware.CpStages.MarkClear(0)                   'Deletes the first mark location in the stage control (the current one)
                                                                                'This deletion and new addition of the location was necessary to change the X, Y and Z properties of that location. I did not know how to do it otherwise
                Lsm5.Hardware.CpStages.MinMarkDistance = 0.1                    'Put a very small mark distance to make it possible to have two cells coming close together. This parameter can be cahnged with the macro but is not accessible from the main software !
                While Lsm5.Hardware.CpStages.MarkGetIndex(x, y) <> -1
                    x = x + 0.1
                    y = y + 0.1
                Wend
                success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ(x, y, z) 'Adds the location again,at the end of the list
                
                Lsm5.Hardware.CpStages.MinMarkDistance = 10                     'Put back the minimal marking distance to its default value
                'test if this is working
                Do While Lsm5.Info.IsAnyHardwareBusy
                    Sleep (20)
                    DoEvents
                Loop

                If Location < LocationNumber Then   'Close the image window if other images will be taken at other locations
                    If TileX > 1 Or TileY > 1 Then
                        If ScanImage.IsValid Then
                            ScanImage.CloseAllWindows
                        End If
                    Else
                        If RecordingDoc.IsValid Then
                            RecordingDoc.CloseAllWindows
                        End If
                    End If
                End If
            Else                                                        'In the single location case with postacquisition tracking one still has to move to the new focus before next acquisition
                Lsm5.Hardware.CpFocus.Position = z + ZBacklash          'Note that ZBacklash was not initialized to -50
                Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                    Sleep (20)
                    DoEvents
                Loop
                Lsm5.Hardware.CpFocus.Position = z
                Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                    Sleep (20)
                    DoEvents
                Loop
                If LocationTracking Then                                'In the single location case one also neess to correct for the XY movements if location tracking is activated
                    success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
                    Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
                        If ScanStop Then
                            Lsm5.StopScan
                            StopAcquisition
                            DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                            Exit Sub
                        End If
                        DoEvents
                        Sleep (5)
                    Loop
                End If
            End If
        Next Location           'Now all the locations have been visited

        If AutomaticBleaching Then FillBleachTable     ' Updating the bleaching table before the next acquisitions, just in case there were changes n the bleaching window
        
        If (RepetitionNumber < BlockRepetitions) Then
            GlobalPrvTime = CDbl(GetTickCount) * 0.001      'Parameters for the waiting. This I took over from the initial Zeiss Macro
            rettime = GlobalPrvTime
            difftime = rettime - GlobalPrvTime
            Do While (difftime <= BlockTimeDelay) And Not (BleachTable(RepetitionNumber + 1) = True)        'This loops define the waiting delay before going back to the first location
                If ExtraBleach Then                                 'Modifies the bleaching table to do an Extrableach for al locatins at the next repetition
                    ExtraBleach = False
                    BleachTable(RepetitionNumber + 1) = True
                End If
                If ScanPause = True Then
                    Pause
                End If
                If ScanStop Then
                    StopAcquisition
                    DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                    Exit Sub
                End If
                DisplayProgress "Waiting " & CStr(CInt(BlockTimeDelay - difftime)) + " s before scanning repetition  " & (RepetitionNumber + 1), RGB(&HC0, &HC0, 0)
                DoEvents
                Sleep (10)
                rettime = CDbl(GetTickCount) * 0.001
                difftime = rettime - GlobalPrvTime
            Loop
            GlobalPrvTime = rettime
            
            If TileX > 1 Or TileY > 1 Then
                
                If ScanImage.IsValid Then
                    ScanImage.CloseAllWindows
                End If
                
                
            Else
                If RecordingDoc.IsValid Then
                    RecordingDoc.CloseAllWindows
                End If
            End If
        Else
            Running = False
        End If
        RepetitionNumber = RepetitionNumber + 1
'        TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    Loop
    StopAcquisition
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
End Sub

Public Sub MassCenter(Context As String)
     Dim scrline As Variant
    Dim spl As Long
    Dim bpp As Long
    Dim ColMax As Long
    Dim LineMax As Long
Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
    DoEvents
    'Gets the dimensions of the image in X (Columns), Y (lines) and Z (Frames)
   If FrameAutofocussing And SystemName = "LIVE" Then ' binning only with LIVE device
   ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
    LineMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionHeight
   Else
   If SystemName = "LIVE" Then
    ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth
    LineMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionHeight
    ElseIf SystemName = "LSM" Then
             ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
             LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
            Else
                MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
                Exit Sub
            End If
    
    
    
     End If
    If Lsm5.DsRecordingActiveDocObject.Recording.ScanMode = "ZScan" Or Lsm5.DsRecordingActiveDocObject.Recording.ScanMode = "Stack" Then
    
        FrameNumber = Lsm5.DsRecordingActiveDocObject.Recording.FramesPerStack
    Debug.Print Lsm5.DsRecordingActiveDocObject.Recording.ScanMode
    Else
    Debug.Print Lsm5.DsRecordingActiveDocObject.Recording.ScanMode
        FrameNumber = 1
    End If
    'Gets the pixel size
    PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
    'Gets the distance between frames in Z
    FrameSpacing = Lsm5.DsRecordingActiveDocObject.Recording.FrameSpacing
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    ReDim Intline(LineMax) As Long
    ReDim IntCol(ColMax) As Long
    ReDim IntFrame(FrameNumber) As Long

    'Select the image channel on which to do the calculations
    If Context = "Autofocus" Then       'Takes the first channel in the context of preacquisition focussing
        Channel = 0
    ElseIf Context = "Tracking" Then    'Takes the channle selected in the pop-up menue when doing postacquisition tracking
        For Channel = 0 To Lsm5.DsRecordingActiveDocObject.GetDimensionChannels - 1
            If Lsm5.DsRecordingActiveDocObject.ChannelName(Channel) = Left(TrackingChannelString, 3) Then
                Exit For
            ElseIf Lsm5.DsRecordingActiveDocObject.ChannelName(Channel) = TrackingChannelString Then
             Exit For
            Else
            End If
        Next Channel
    End If
    
   'lineMax = 1

    'Reads the pixel values and fills the tables with the projected (integrated) pixels values in the three directions
    For Frame = 1 To FrameNumber
        For line = 1 To LineMax
            bpp = 0
           
            scrline = Lsm5.DsRecordingActiveDocObject.ScanLine(Channel, 0, Frame - 1, line - 1, spl, bpp) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
            For Col = 1 To ColMax               'Now I'm scanning all the pixels in the line
                Intline(line - 1) = Intline(line - 1) + scrline(Col - 1)
                IntCol(Col - 1) = IntCol(Col - 1) + scrline(Col - 1)
                IntFrame(Frame - 1) = IntFrame(Frame - 1) + scrline(Col - 1)
            Next Col
        Next line
    Next Frame
    
    'First it finds the minimum and maximum porjected (integrated) pixel values in the 3 dimensions
    MinColValue = 4095 * LineMax * FrameNumber           'The maximum values are initially set to the maximum possible value
    minLineValue = 4095 * ColMax * FrameNumber
    minFrameValue = 4095 * LineMax * ColMax
    MaxColValue = 0                                     'The maximun values are initialliy set to 0
    MaxLineValue = 0
    MaxframeValue = 0
    For line = 1 To LineMax
        If Intline(line - 1) < minLineValue Then
            minLineValue = Intline(line - 1)
        End If
        If Intline(line - 1) > MaxLineValue Then
            MaxLineValue = Intline(line - 1)
        End If
    Next line
    For Col = 1 To ColMax
        If IntCol(Col - 1) < MinColValue Then
            MinColValue = IntCol(Col - 1)
        End If
        If IntCol(Col - 1) > MaxColValue Then
            MaxColValue = IntCol(Col - 1)
        End If
    Next Col
    For Frame = 1 To FrameNumber
        If IntFrame(Frame - 1) < minFrameValue Then
            minFrameValue = IntFrame(Frame - 1)
        End If
        If IntFrame(Frame - 1) > MaxframeValue Then
            MaxframeValue = IntFrame(Frame - 1)
        End If
    Next Frame

    'Calculates the threshold values. It is set to an arbitrary value of the minimum projected value plus 20% of the difference between the minimum and the maximum projected value.
    'Then calculates the center of mass
    LineSum = 0
    LineWeight = 0
    MidLine = (LineMax + 1) / 2
    If CheckBoxRefControl.Value = True Then
        If (MaxframeValue - minFrameValue) < minFrameValue * 0.5 Or MaxframeValue = 0 Then NoReflectionSignal = True
    End If
    Threshold = minLineValue + (MaxLineValue - minLineValue) * 0.8         'Threshold calculation
    For line = 1 To LineMax
        LineValue = Intline(line - 1) - Threshold                           'Subtracs the threshold
        PosValue = LineValue + Abs(LineValue)                               'Makes sure that the value is positive or zero. If LineValue is negative, the Posvalue = 0; if Line value is positive, then Posvalue = 2*LineValue
        LineWeight = LineWeight + (PixelSize * (line - MidLine)) * PosValue 'Calculates the weight of the Thresholded projected pixel values according to their position relative to the center of the image and sums them up
        LineSum = LineSum + PosValue                                        'Calculates the sum of the thresholded pixel values
    Next line
    If LineSum = 0 Then
        YMass = 0
    Else
        YMass = LineWeight / LineSum                                       'Normalizes the weights to get the center of mass
    End If

    ColSum = 0
    ColWeight = 0
    MidCol = (ColMax + 1) / 2
    Threshold = MinColValue + (MaxColValue - MinColValue) * 0.8
    For Col = 1 To ColMax
        ColValue = IntCol(Col - 1) - Threshold
        PosValue = ColValue + Abs(ColValue)
        ColWeight = ColWeight + (PixelSize * (Col - MidCol)) * PosValue
        ColSum = ColSum + PosValue
    Next Col
    If ColSum = 0 Then
        XMass = 0
    Else
        XMass = ColWeight / ColSum
    End If

    FrameSum = 0
    FrameWeight = 0
    MidFrame = (FrameNumber + 1) / 2
    Threshold = minFrameValue + (MaxframeValue - minFrameValue) * 0.8
    For Frame = 1 To FrameNumber
        FrameValue = IntFrame(Frame - 1) - Threshold
        PosValue = FrameValue + Abs(FrameValue)
        FrameWeight = FrameWeight + (FrameSpacing * (Frame - MidFrame)) * PosValue
        FrameSum = FrameSum + PosValue
    Next Frame
    If FrameSum = 0 Then
        ZMass = 0
    Else
        ZMass = FrameWeight / FrameSum
        End If
        
End Sub


Private Sub StopButton_Click()
   
        ScanStop = True
        DisplayProgress "Restore Settings", RGB(&HC0, 0, 0)
   
End Sub

Public Sub StopAcquisition()
    Dim FileName As String
    Running = False
    ScanStop = False
    RepetitionNumber = 1
    ScanPause = False
    PauseButton.Caption = "Pause"
    PauseButton.BackColor = &H8000000F
    ExtraBleach = False
    ExtraBleachButton.Caption = "Bleach"
    ExtraBleachButton.BackColor = &H8000000F
    ReDim BleachTable(BlockRepetitions)
    ReDim BleachStartTable(BlockRepetitions)
    ReDim BleachStopTable(BlockRepetitions)
    CloseButton.Enabled = True
    ReinitializeButton.Enabled = True
    StartButton.Enabled = True
    StartBleachButton.Enabled = True
    If LocationTracking Or FrameAutofocussing Then
'        For i = 1 To PositionData.Sheets.count
'            PositionData.Sheets.Item(i).Select
'            Cells.Select
'            Selection.Columns.AutoFit
'        Next i
'        FileName = Left(DataBaseLabel, Len(DataBaseLabel) - 4) & ".xls"
'        PositionData.SaveAs FileName:=FileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
'        PositionData.Close
'        Excel.Application.Quit
    End If
End Sub



Private Sub PauseButton_Click()
    If Running Then
        If ScanPause = False Then
            ScanPause = True
            PauseButton.Caption = "Resume"
            PauseButton.BackColor = 12648447
        Else
            ScanPause = False
            PauseButton.Caption = "Pause"
            PauseButton.BackColor = &H8000000F
        End If
    Else
        MsgBox "The acquisition has not started yet or is already finished. Cannot pause."
    End If
End Sub

Public Sub Pause()
    Dim rettime As Double
    Dim GlobalPrvTime As Double
    SetFocusButton.Enabled = True
    AutofocusButton.Enabled = True
    GlobalPrvTime = CDbl(GetTickCount) * 0.001
    rettime = GlobalPrvTime
    difftime = rettime - GlobalPrvTime
    Do While True
        DisplayProgress "Pause " & CStr(CInt(difftime)) & " s", RGB(&HC0, &HC0, 0)
        If ScanStop Then
            Exit Sub
        End If
        If ScanPause = False Then
            SetFocusButton.Enabled = False
            AutofocusButton.Enabled = False
            Exit Sub
        End If
        DoEvents
        Sleep (20)
        rettime = CDbl(GetTickCount) * 0.001
        difftime = rettime - GlobalPrvTime
    Loop
End Sub


Private Sub ExtraBleachButton_Click()
    If Running Then
        ExtraBleach = True
        ExtraBleachButton.Caption = "Will Bleach"
        ExtraBleachButton.BackColor = 12648447
    Else
        MsgBox "The acquisition has not started yet or is already finished. Cannot bleach."
    End If
End Sub

Private Sub StripeScanToggle_Click()
 If MultipleLocationToggle.Value = True Then MultipleLocationToggle.Value = Not StripeScanToggle.Value
    If SingleLocationToggle.Value = True Then SingleLocationToggle.Value = Not StripeScanToggle.Value
    If GridToggle.Value = True Then GridToggle.Value = Not StripeScanToggle.Value
    GridObjectsandVarialbles True
    CheckBoxMeander.Visible = False
    CheckBoxScannAll.Visible = True
End Sub

Private Sub GridToggle_Click()
    GridToggle.Value = True
    If MultipleLocationToggle.Value = True Then MultipleLocationToggle.Value = Not GridToggle.Value
    If SingleLocationToggle.Value = True Then SingleLocationToggle.Value = Not GridToggle.Value
    If StripeScanToggle.Value = True Then StripeScanToggle.Value = Not GridToggle.Value
    GridObjectsandVarialbles True
    CheckBoxScannAll.Visible = False
    StartBleachButton.Visible = False
    ExtraBleachButton.Visible = False
End Sub

Private Sub SingleLocationToggle_Click()
    SingleLocationToggle.Value = True
   If MultipleLocationToggle.Value = True Then MultipleLocationToggle.Value = Not SingleLocationToggle.Value
    If StripeScanToggle.Value = True Then StripeScanToggle.Value = Not SingleLocationToggle.Value
    If GridToggle.Value = True Then GridToggle.Value = Not SingleLocationToggle.Value
    GridObjectsandVarialbles False
    CheckBoxScannAll.Visible = False
    Label15.Caption = "Define Locations Using Stage and Focus Corntrol Dialog!"
  '  StartBleachButton.Visible = True
 '   ExtraBleachButton.Visible = True
    Frame15.Visible = False
End Sub

Private Sub MultipleLocationToggle_Click()
    MultipleLocationToggle.Value = True
    If SingleLocationToggle.Value = True Then SingleLocationToggle.Value = Not MultipleLocationToggle.Value
    If GridToggle.Value = True Then GridToggle.Value = Not MultipleLocationToggle.Value
    If StripeScanToggle.Value = True Then StripeScanToggle.Value = Not MultipleLocationToggle.Value
    GridObjectsandVarialbles False
    Label15.Caption = "Define Locations Using Stage and Focus Corntrol Dialog!"
    CheckBoxScannAll.Visible = False
   ' ZMapButton.Left = 12
   ' ZMapButton.Top = 258
    ZMapButton.Visible = True
   ' CheckBoxZMap.Left = 80
   ' CheckBoxZMap.Top = 258
    CheckBoxZMap.Visible = True
    StartBleachButton.Visible = False
  '  ExtraBleachButton.Visible = True
    Frame15.Visible = True
    TextBoxTileX.Visible = True
    TextBoxTileY.Visible = True
    Tileframe.Visible = True
    Label17.Visible = True
    Label18.Visible = True
    Label20.Visible = True
    CreateLocationsButton.Visible = True
    TextBoxOverlap.Visible = True
End Sub
Private Sub GridObjectsandVarialbles(Activate As Boolean)
   ' ZMapButton.Left = 198.05
   ' ZMapButton.Top = 306
    ZMapButton.Visible = Activate
    CreateLocationsButton.Visible = Activate
    CommandButtonRemove.Visible = Activate
    CommandButtonGrid.Visible = Activate
    CommandButtonStoreApply.Visible = Activate
    TextBoxYGrid.Visible = Activate
    TextBoxXGrid.Visible = Activate
    TextBoxYStep.Visible = Activate
    TextBoxXStep.Visible = Activate
    Tileframe.Visible = Activate
    Frame16.Visible = Activate
    Frame15.Visible = Activate
    Label1.Visible = Activate
    Label2.Visible = Activate
    Label3.Visible = Activate
    Label4.Visible = Activate
    Label5.Visible = Activate
   ' Label16.Visible = Activate
    Label7.Visible = Activate
    Label17.Visible = Activate
    Label18.Visible = Activate
    Label20.Visible = Activate
    TextBoxOverlap.Visible = Activate
    TextBoxTileX.Visible = Activate
    TextBoxTileY.Visible = Activate
   ' CheckBoxKeepSteps.Visible = Activate
    CheckBoxMeander.Visible = Activate
  '  CheckBoxZMap.Left = 132
   ' CheckBoxZMap.Top = 288
    CheckBoxZMap.Visible = Activate
    'LabelGrid.Visible = Activate
    Label15.Visible = Not Activate
    Grid = GridToggle.Value
    MultipleLocation = MultipleLocationToggle.Value ' Sets the MultipleLocation Boolean to False
   
End Sub



Public Sub AutoFindTracks()
    Dim i, j As Integer
    Dim ChannelOK As Boolean
    Dim DataChannel As DsDataChannel
    Dim Color As Long
    Dim ConfiguredTracks As Integer

    
    OptionButtonTrack1.Visible = False
    OptionButtonTrack1.Enabled = False
    OptionButtonTrack1.Value = False
    CheckBoxTrack1.Visible = False
    CheckBoxTrack1.Enabled = False
    CheckBoxTrack1.Value = False
    
    OptionButtonTrack2.Visible = False
    OptionButtonTrack2.Enabled = False
    OptionButtonTrack2.Value = False
    CheckBoxTrack2.Visible = False
    CheckBoxTrack2.Enabled = False
    CheckBoxTrack2.Value = False
    
    OptionButtonTrack3.Visible = False
    OptionButtonTrack3.Enabled = False
    OptionButtonTrack3.Value = False
    CheckBoxTrack3.Visible = False
    CheckBoxTrack3.Enabled = False
    CheckBoxTrack3.Value = False
    
    OptionButtonTrack4.Visible = False
    OptionButtonTrack4.Enabled = False
    OptionButtonTrack4.Value = False
    CheckBoxTrack4.Visible = False
    CheckBoxTrack4.Enabled = False
    CheckBoxTrack4.Value = False
    
    ConfiguredTracks = Lsm5.DsRecording.TrackCount
    ChannelOK = False
    GoodTracks = 0
    
'The next line and the following "if" should be removed when working with the LSM 2.8 software (where the lambda mode is not defined)
    Set Track = Lsm5.DsRecording.TrackObjectLambda(success)
    If success Then
        If Track.Acquire Then
            MsgBox ("This macro does not work in the Lambda Mode. Please switch to the Channel Mode and reinitialize the Macro.")
            Exit Sub
        End If
    End If
            
    For i = 1 To ConfiguredTracks
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, success)
        TrackName = Track.name
        j = 0
'In the next line remove "Or Track.IslambdaTrack" when working with the LSM 2.8 software
        If Not (Track.IsBleachTrack Or Track.IsLambdaTrack) Then
            Do While (Not ChannelOK) And (j < Track.DataChannelCount)
                Set DataChannel = Track.DataChannelObjectByIndex(j, success)
                If DataChannel.Acquire = True Then ChannelOK = True
                Color = DataChannel.ColorRef
                j = j + 1
            Loop
            If ChannelOK Then
                If Not Track.IsRatioTrack Then
                    GoodTracks = GoodTracks + 1
                    If GoodTracks = 5 Then
                        MsgBox ("This Macro only accepts 4 different tracks")
                    End If
                    If GoodTracks = 1 Then
                        OptionButtonTrack1.Visible = True
                        OptionButtonTrack1.Caption = TrackName
                        OptionButtonTrack1.Enabled = True
                        OptionButtonTrack1.BackColor = Color
                        CheckBoxTrack1.Visible = True
                        CheckBoxTrack1.Caption = TrackName
                        CheckBoxTrack1.Enabled = True
                        CheckBoxTrack1.BackColor = Color
                    End If
                    If GoodTracks = 2 Then
                        OptionButtonTrack2.Visible = True
                        OptionButtonTrack2.Caption = TrackName
                        OptionButtonTrack2.Enabled = True
                        OptionButtonTrack2.BackColor = Color
                        CheckBoxTrack2.Visible = True
                        CheckBoxTrack2.Caption = TrackName
                        CheckBoxTrack2.Enabled = True
                        CheckBoxTrack2.BackColor = Color
                    End If
                    If GoodTracks = 3 Then
                        OptionButtonTrack3.Visible = True
                        OptionButtonTrack3.Caption = TrackName
                        OptionButtonTrack3.Enabled = True
                        OptionButtonTrack3.BackColor = Color
                        CheckBoxTrack3.Visible = True
                        CheckBoxTrack3.Caption = TrackName
                        CheckBoxTrack3.Enabled = True
                        CheckBoxTrack3.BackColor = Color
                    End If
                    If GoodTracks = 4 Then
                        OptionButtonTrack4.Visible = True
                        OptionButtonTrack4.Caption = TrackName
                        OptionButtonTrack4.Enabled = True
                        OptionButtonTrack4.BackColor = Color
                        CheckBoxTrack4.Visible = True
                        CheckBoxTrack4.Caption = TrackName
                        CheckBoxTrack4.Enabled = True
                        CheckBoxTrack4.BackColor = Color
                    End If
                Else
                    MsgBox ("This macro does not allow to use a Ratio Channel. The Ratio Channel will thus be disabled.")
                    For j = 0 To Track.DataChannelCount - 1
                        Set DataChannel = Track.DataChannelObjectByIndex(j, success)
                        DataChannel.Acquire = False
                    Next
                End If
                ChannelOK = False
            End If
        End If
    Next
    If GoodTracks < 4 Then
        TrackNumber = GoodTracks
    Else
        TrackNumber = 4
    End If
End Sub



Private Sub BSliderZoffset_Change()
    'Tests whether chosen Offset is less or equal to half of the working distance of the objective but why can't it be bigger??
    Dim Position As Long 'gets the postion of the actual objective revolver by number
    Dim Range As Double 'contains value of working distance in um
    If flgUserChange Then '??? What is the sense of flgUserChange
        Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
        If Position >= 0 Then ' ??? is it possible that Revolver Position has another value
            Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000# ' ??? why is there a # behind that number if range is already defined as double
                                                                                            ' in which unit is working distance read out and why multiplication with 1000
        Else
            Range = 0#
        End If
        If Abs(BSliderZOffset.Value) > Range * 0.9 Then
            BSliderZOffset.Value = 0
            MsgBox "Zoffset has to be less than the working distance of the objective: " + CStr(Range) + " um"
        End If
    End If
End Sub




Private Sub BSliderZRange_Change()    ' It should be possible to change the limit of the range to bigger values than half of the working distance
    Dim Position As Long
    Dim Range As Double
    If flgUserChange Then
        Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
        If Position >= 0 Then
            Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
        Else
            Range = 0#
        End If
        If BSliderZRange.Value > Range * 0.9 Then
            BSliderZRange.Value = Range * 0.9
            MsgBox "ZRange has to be less or equal to the working distance of the objective: " + CStr(Range) + " um"
        End If
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub


Private Sub CloseButton_Click()
    AutoStore
'    Excel.Application.DisplayAlerts = False
'    Excel.Application.Quit
    End
End Sub


Private Sub ReInitializeButton_Click()
    Re_Initialize
End Sub


Private Sub CommandButtonStore_Click()
    StoreApply.Show
End Sub






Private Sub UserForm_Initialize()           ' This contained some initialization  that have then been deleted or moved to Re_Start
    Re_Start
End Sub


Private Sub Re_Start()                      'Initializations that need to be performed only at the first start of the Macro
    Dim delay As Single
    Dim standType As String
    Dim Count As Long
    Dim ImageDatabase As DsGuidedModeDatabase
    Dim i As Long
    Dim MruList As DsMruList
    Dim cnt As Long
    Dim lpReOpenBuff As OFSTRUCT
    Dim wStyle As Long
    Dim lpRootPathName As String
    Dim lpSectorsPerCluster As Long
    Dim lpBytesPerSector As Long
    Dim lpNumberOfFreeClusters As Long
    Dim lpTotalNumberOfClusters As Long
    Dim lSpace As Long
    Dim lFreeSpace As Double
    Dim fSize As Double
    Dim hFile As Long
   
    
    Set tools = Lsm5.tools
    GlobalMacroKey = "Autofocus"
   
'    bRunning = False
 '   LbStatus = "inactive"
    flgUserChange = True
    delay = 1
    flgEvent = 7
    flg = 0
    Lsm5.StopScan
    Wait (delay)
    TimerUnit = 1
    CommandTimeSec.BackColor = &HFF8080
    BlockRepetitions = 1
    ReDim Preserve GlobalImageIndex(BlockRepetitions)
    ScanLineToggle.Value = True
    SingleLocationToggle.Value = True
    Label15.Caption = "Define Locations Using Stage and Focus Corntrol Dialog!!"
    GlobalProject = "AutofocusScreen1.7"
    GlobalProjectName = GlobalProject + ".lvb"
    HelpNamePDF = "AutofocusScreen_help.pdf"
    Re_Initialize ' Continues the initialization process
  
     
    
End Sub


Public Sub Re_Initialize()                  'Initializations that need to be performed only when clicling the "initialize" button
    Dim delay As Single
    Dim standType As String
    Dim Count As Long
     Dim bLSM As Boolean
    Dim bLIVE As Boolean
    Dim bCamera As Boolean
    
'    StopAcquisition
'    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
    AutoFindTracks
  
     BSliderZOffset.Value = 0
    BSliderZRange.Value = 80
    BSliderZStep.Value = 0.1
    TextBoxXGrid.Value = 3
    TextBoxYGrid.Value = 2
    TextBoxXStep.Value = -1125
    TextBoxYStep.Value = 1125
    BSliderScanSpeed = 1000
    BSliderRepetitions = 300
    BSliderTime = 1
    
    CheckBoxLowZoom = False
    CheckBoxInnactivateAutofocus = False
    PubSearchScan = False
    NoReflectionSignal = False
    PubSentStageGrid = False
    GlobalZmapAquired = False
    
    
    If Lsm5.Hardware.CpHrz.Exist("HRZ") = True Then     'Check if an HRZ is available. If not the "HRZ checkbox is not available.
        CheckBoxHRZ.Visible = True
        CheckBoxHRZ.Value = True
    Else
'   I take this out, because at the LMS we have an HRZ but the LSM Software does not give the right signal for that, but now you can can use the HRZ
'   CheckBoxHRZ.Visible = False
'   CheckBoxHRZ.Value = False
    End If
    
    ScanLineToggle.Value = True
    SingleLocationToggle.Value = True
    UsedDevices40 bLSM, bLIVE, bCamera
    If bLSM Then
            SystemName = "LSM"
            CheckBoxHighSpeed.Value = True
            CheckBoxHighSpeed.Visible = True
            CheckBoxHighSpeed.Top = 48
            CheckBoxLowZoom.Top = 71.35
            CheckBoxHRZ.Top = 90.95
            CheckBoxRefControl.Top = 110.6
     
             BSliderFrameSize.Min = 16
            BSliderFrameSize.Max = 1024
            BSliderFrameSize.Step = 8
            BSliderFrameSize.StepSmall = 4
          
           
            Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
        ElseIf bLIVE Then
            SystemName = "LIVE"
            CheckBoxHighSpeed.Value = False
            CheckBoxHighSpeed.Visible = False
            CheckBoxHRZ.Top = 85
            CheckBoxRefControl.Top = 108
            CheckBoxLowZoom.Top = 60
'            BSliderFrameSize.ValueEditable = False
            BSliderFrameSize.Min = 128
            BSliderFrameSize.Max = 1024
            BSliderFrameSize.Step = 128
            BSliderFrameSize.StepSmall = 128
          
            Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents
            
          
            
        ElseIf bCamera Then
            SystemName = "Camera"
        End If
    AutofocusForm.Caption = GlobalProject + " for " + SystemName
      
End Sub

Private Sub CreditButton_Click()
    CreditForm.Show
End Sub

Private Sub TrackingToggle_Click()                                          ' Sets the parameters for postacquisition tracking
    LocationTracking = TrackingToggle.Value
    ComboBoxTrackingChannel.Visible = TrackingToggle.Value
    FillTrackingChannelList
    CheckBoxTrackZ.Visible = TrackingToggle.Value
    If Lsm5.DsRecording.ScanMode = "Stack" Then
        CheckBoxTrackZ.Enabled = True
    Else
        CheckBoxTrackZ.Enabled = False
        CheckBoxTrackZ.Value = False
    End If
End Sub


'fills popup menu for chosing a track for post-acquisition tracking in ScanLine mode
Private Sub FillTrackingChannelList()
    Dim t As Integer
    Dim c As Integer
    Dim ca As Integer
    Dim Channel As DsDetectionChannel

    ActivateAcquisitionTrack 'will set IsAcquisitionTrack selected true if a valid track is selected for acquisition, and "marks the track in the Zeiss config window
    
    ReDim ActiveChannels(Lsm5.Constants.MaxActiveChannels)  'ActiveChannels is a dynamic array (variable size), ReDim defines array size required next
                                                            'Array size is (MaxActiveChannels gets) the total max number of active channels in all tracks
    ComboBoxTrackingChannel.Clear 'Content of popup menu for chosing track for post-acquisition tracking is deleted
    ca = 0
    
    If IsAcquisitionTrackSelected Then 'IsAcquisitionTrackSelected is True if one channel is activated in tracks 1-4
        For t = 1 To TrackNumber 'This loop goes through all tracks and will collect all activated channels to display them in popup menu
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(t - 1, success) 'goes through all defined tracks
            If Track.Acquire Then 'if track is activated for acquisition
                For c = 1 To Track.DetectionChannelCount 'for every detection channel of track
                    Set Channel = Track.DetectionChannelObjectByIndex(c - 1, success)
                    If Channel.Acquire Then 'if channel is activated
                        ca = ca + 1 'counter for active channels will increase by one
                        ComboBoxTrackingChannel.AddItem Track.name & " " & Channel.name 'entry is added to combo box to chose track for post-acquisition tracking
                        ActiveChannels(ca) = Channel.name & "-T" & Track.MultiplexOrder + 1  'adds entry to ActiveChannel Array with name of channel + name of track
                    End If
                Next c
            End If
        Next t
        
        ComboBoxTrackingChannel.Value = ComboBoxTrackingChannel.List(0) 'initially displayed text in popup menu is a blank line (first channel is 1).
    End If
End Sub
Private Sub ComboBoxTrackingChannel_Change()        'Sets the name of the channel for PostAcquisition tracking.
    TrackingChannelString = ActiveChannels(ComboBoxTrackingChannel.ListIndex + 1)
End Sub
Private Sub CheckBoxTrackZ_Click()
    TrackZ = CheckBoxTrackZ.Value
    If CheckBoxTrackZ.Value = True Then
        CheckBoxInnactivateAutofocus.Value = True                  'If posacquisition Z-tracking is activated, it is necessary to deactivate autofocussing
        CheckBoxInnactivateAutofocus.BackColor = 33023
        CheckBoxTrackZ.BackColor = 33023
    Else
        CheckBoxTrackZ.BackColor = &H8000000F
    End If
End Sub

















Private Sub CommandTimeMin_Click()
    TimerUnit = 60
    BSliderTime.Max = 60                        'When workings with minutes the maximum delay that can be set with the slider is 1 hour
    BSliderTime.Value = BlockTimeDelay / 60
    CommandTimeMin.BackColor = &HFF8080
    CommandTimeSec.BackColor = &H8000000F
End Sub

Private Sub CommandTimeSec_Click()
    TimerUnit = 1
    BSliderTime.Max = 180                       'When workings with seconds the maximum delay that can be set with the slider is 3 minutes
    BSliderTime.Value = BlockTimeDelay
    CommandTimeSec.BackColor = &HFF8080
    CommandTimeMin.BackColor = &H8000000F
End Sub









Private Sub BSliderTime_Click()
   

    BlockTimeDelay = BSliderTime.Value * TimerUnit                      'BlockTimedelay gets the value of the slider in seconds
    
End Sub





Private Sub BSliderRepetitions_Change()
    If Not Running Then
        BlockRepetitions = BSliderRepetitions.Value
    ElseIf Not (BSliderRepetitions.Value <= (RepetitionNumber + 1)) Then
        BlockRepetitions = BSliderRepetitions.Value
    Else
        BSliderRepetitions.Value = RepetitionNumber + 1
        BlockRepetitions = BSliderRepetitions.Value
    End If
    
    ReDim Preserve GlobalImageIndex(BlockRepetitions)           'The global image index I'm not sure how this is working.
    ReDim Preserve BleachTable(BlockRepetitions)                'BleachTable defines when bleaching will have to occur
    If AutomaticBleaching Then FillBleachTable                  'Reads the parameters defined in the Bleach control window of the main software
    ReDim Preserve BleachStartTable(BlockRepetitions)           'This is to store the timepoints when the bleaches started. Preserve is to keep the timepoints if the slider is moved during an experiment
    ReDim Preserve BleachStopTable(BlockRepetitions)            'This is to store the timepoints when the bleaches stopped. Preserve is to keep the timepoints if the slider is moved during an experiment
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub




Private Sub TextBoxFileName_Change()
    GlobalFileName = TextBoxFileName.Value
End Sub

Private Sub CommandButtonSelectDataBase_Click()
    Dim lpReOpenBuff As OFSTRUCT
    Dim wStyle As Long
    Dim hFile As Long
    Dim flgUserChangeSaved As Boolean
    flgUserChangeSaved = flgUserChange
    
    flgUserChange = False

'Common Dialog is used to open folders and files in windows
    CommonDialog.FileName = GlobalDataBaseName                      'remembers which was the latest Database that was opened
    CommonDialog.Filter = "Database files ( *.mdb ) |*.mdb"         'filter to only display database files
    CommonDialog.ShowOpen
    hFile = OpenFile(CommonDialog.FileName, lpReOpenBuff, wStyle)
    If hFile <> -1 Then
        GlobalDataBaseName = CommonDialog.FileName                  'Store the path of the database in the GlobalDatabaseName variable
        DataBaseLabel.Caption = CommonDialog.FileName
    Else
        MsgBox "Selected file does not exist"
    End If
    flgUserChange = flgUserChangeSaved
End Sub

Private Sub CommandButtonNewDataBase_Click()   'Creates a new database
    Lsm5.NewDatabase (NewDatabase)              'Directly opens the LSM window to create a new database
    Lsm5.CloseAllDatabaseWindows                'Strange that this is there and not before the previous line...
    GlobalDataBaseName = Lsm5.MruDatabases.name(0)      'Write the name of the database in a varialbe (used afterwards for saving to the right database)
    DataBaseLabel.Caption = Lsm5.MruDatabases.name(0)   'Indicates the name of the databse for the user to check
End Sub
 



Public Sub ActivateAutofocusTrack(HighSpeed As Boolean)
    Dim i As Integer

    IsAutofocusTrackSelected = False
    For i = 1 To TrackNumber
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, success)
        If i = 1 Then
            If OptionButtonTrack1.Value = True Then
                Track.Acquire = 1
                IsAutofocusTrackSelected = True
                AutofocusTrack = i - 1
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 2 Then
            If OptionButtonTrack2.Value = True Then
                Track.Acquire = 1
                IsAutofocusTrackSelected = True
                AutofocusTrack = i - 1
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 3 Then
            If OptionButtonTrack3.Value = True Then
                Track.Acquire = 1
                IsAutofocusTrackSelected = True
                AutofocusTrack = i - 1
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 4 Then
            If OptionButtonTrack4.Value = True Then
                Track.Acquire = 1
                IsAutofocusTrackSelected = True
                AutofocusTrack = i - 1
           Else
                Track.Acquire = 0
            End If
        End If
    Next i
    If HighSpeed Then
        Track.SamplingNumber = 1
    End If
End Sub


Public Sub ActivateAcquisitionTrack()
    Dim i As Integer

    'this loop goes through all tracks; it will check for actual track in loop whether corresponding checkbox is activated
    'if checkbox of one of tracks is selected IsAcquisitionTrack will be set true
    'is this so complicated to be sure that if one track is chosen this track is a track that is defined in track list ???
    IsAcquisitionTrackSelected = False
    For i = 1 To TrackNumber 'TrackNumber is maximum 4 or less (see definition with GoodTracks)
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(i - 1, success) 'choses track corresponding to track number
        If i = 1 Then
            If CheckBoxTrack1.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 2 Then
            If CheckBoxTrack2.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 3 Then
            If CheckBoxTrack3.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
        If i = 4 Then
            If CheckBoxTrack4.Value = True Then
                Track.Acquire = 1
                IsAcquisitionTrackSelected = True
            Else
                Track.Acquire = 0
            End If
        End If
    Next i
End Sub


' I could delete this procedure if I do not add any help button
'Private Sub HelpButton_Click()
'    Dim dblTask As Double
'    Dim MacroPath As String
'    Dim MyPath As String
'    Dim bslash As String
'    Dim Success As Integer
'    Dim pos As Integer
'    Dim Start As Integer
'    Dim count As Long
'    Dim ProjName As String
'    Dim indx As Integer
'
'    count = ProjectCount()
'    For indx = 0 To count - 1
'        MacroPath = ProjectPath(indx, Success)
'        ProjName = ProjectTitle(indx, Success)
'        If StrComp(ProjName, GlobalProjectName, vbTextCompare) = 0 Then
'            Start = 1
'            bslash = "\"
'            pos = Start
'            Do While pos > 0
'                pos = InStr(Start, MacroPath, bslash)
'                If pos > 0 Then
'                    Start = pos + 1
'                End If
'            Loop
'            MyPath = Left(MacroPath, Start - 1)
'            MyPath = MyPath + GlobalHelpName
'            dblTask = Shell("C:\Program Files\Windows NT\Accessories\wordpad.exe " + MyPath, vbNormalFocus)
'            Exit For
'        End If
'    Next indx
'End Sub

Sub Wait(PauseTime As Single)
    Dim Start As Single
    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
       DoEvents    ' Yield to other processes.
       'Lsm5.DsRecording.StartScanTriggerIn
    Loop
End Sub



Public Sub SetFocus(ZRange As Double, ZStep As Double, LowZoom As Boolean, HighSpeed As Boolean, Zoffset As Double)
    Dim SpeedCopy As Double
    Dim ZoomXCopy As Double
    Dim ZoomYCopy As Double
    Dim SamplesPerLineCopy As Long
    Dim LinesPerFrameCopy As Long
    Dim ScanModeCopy As String
    Dim SpecialScanModeCopy As String

    Dim Range As Double
    Dim Position As Long
  
    Dim MyRecording As DsRecording

    Dim Tnum As Long
    Dim i As Long
    Dim success As Integer
    Dim NewPicture As DsRecordingDoc
    Dim Pixel As Long
    Dim scrline As Variant
    Dim PxlArray() As Long
    Dim spl As Long
    Dim bpp As Long
    Dim IntensityStr As String
    Dim ChNumber As Long
    Dim Channel As Long
    Dim LongRange As Long
    Dim PxlMax As Long
    Dim PxlTot As Long
    Dim LineMax As Long
    Dim StackSize As Double
    Dim SavedSampling As Long
    Dim key As String
    Dim line As Long
    Dim lT As Long
    Dim NoFrames As Long
    Dim SystemVersion As String
    Dim Speed As Long
    Dim MaxSpeed As Long
    
    
  Zbefore = Lsm5.Hardware.CpFocus.Position
    DisplayProgress "Get Offset Value", RGB(0, &HC0, 0)             'Gives information to the user
    Lsm5.StopScan                                                   'Just in case some scanning wa working
   Lsm5Vba.Application.ThrowEvent eRootReuse, 0                   'Was there in the initial Zeiss macro, but it seems notnecessary
    DoEvents
    
    ZAuto = 0                                                       'I do not know why is this Z Auto there. I believe it is obsolete
   ' ZBacklash = -50 'Has to do with the movements of the focus wheel that are "better" if they are long enough.
    
    StoreAquisitionParameters
    
    
    ActivateAutofocusTrack HighSpeed                                'Sets the track for autofocussing (i.e. "selects" the track in the Zeiss config window )
    If Not IsAutofocusTrackSelected Then                                'The variable IsAutofocusTrackSelected has been updated in the ActivateAutofocausTrack function
        MsgBox "No track selected for Autofocus! Cannot Autofocus!"
        Exit Sub
    End If
  
    Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition       'Verifies that the working distnce is OK. Comes from the initial Zeiss autofocussing macro
    If Position >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
    Else
        Range = 0#
    End If
    If Range = 0 Then
        MsgBox "Objective's working distance not defined! Cannot Autofocus!"
        Exit Sub
    End If
    If ZRange > Range * 0.9 Then
        ZRange = Range * 0.9
    End If
    If Abs(Zoffset) > Range * 0.9 Then                   'The offset has to be within half of the working distance. May want to change this when working with large samples in Z
        Zoffset = 0
    End If

    SystemVersionOffset
    
     AutofocusForm.AutofocusSetting HRZ, BlockHighSpeed
     Lsm5.DsRecording.FrameSpacing = ZStep
    NoFrames = CLng(ZRange / ZStep) + 1                     'Calculates the number of frames per stack. Clng converts it to a long and rounds up the fraction
    Lsm5.DsRecording.FramesPerStack = NoFrames
    If NoFrames > 2048 Then                                 'overwrites the userdefined value if too many frames have been defined by the user
        NoFrames = 2048
    End If
    If Not HRZ Then
    Lsm5.DsRecording.Sample0Z = ZStep * NoFrames / 2
    End If                                                    'Distance of the actual focus to the first Z position of the image (or line) to acquire in the stack.
                                                            'I think this is only valid for the focus wheel and not the HRZ
    
    If Zoffset <= Range * 0.9 Then
       Lsm5.Hardware.CpFocus.Position = Zbefore + Zoffset + GlobalCorrectionOffset + ZBacklash 'Move down 50um (=ZBacklash) below the position of the offset
      Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
            Sleep (20)  '20ms
           DoEvents
       Loop
       Lsm5.Hardware.CpFocus.Position = Zbefore + Zoffset + GlobalCorrectionOffset            'Moves up to the position of the offset
       Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
           Sleep (20)
           DoEvents
        Loop
   End If
If Not FrameAutofocussing Then
 Lsm5.DsRecording.ScanMode = "ZScan"
 If Not HRZ Then
 Lsm5.DsRecording.SpecialScanMode = "FocusStep"
 End If
 End If
    Set NewPicture = Lsm5.StartScan                             'Starts the image acquisition for autofocussing
    Do While NewPicture.IsBusy                                  ' Waiting untill the image acquisition is done
        If ScanStop Then
            Lsm5.StopScan
            GoTo Abort
        End If
        DoEvents
        Sleep (10)
    Loop
    Lsm5.tools.WaitForScanEnd False, 40                        'This looks redoundant with the previous, but I had trried to remove it and had problems. It's better to have 2 contols than none !
 
    AutofocusForm.MassCenter ("Autofocus")                                    'Calculates the mass center in 3 dimensions
    XShift = XMass
    YShift = -YMass
    ZShift = ZMass
    
        'check if Z shift makes sense
        CheckRefControl BlockZRange
        
 If Zoffset <= Range * 0.9 Then
       Lsm5.Hardware.CpFocus.Position = Zbefore + GlobalCorrectionOffset + ZBacklash  'Move down 50um (=ZBacklash) below the position of the offset
      Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
            Sleep (20)  '20ms
           DoEvents
       Loop
       Lsm5.Hardware.CpFocus.Position = Zbefore + GlobalCorrectionOffset             'Moves up to the position of the offset
       Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy                 'Waits that the objective movement is finished, code from the original macro
           Sleep (20)
           DoEvents
        Loop
   End If

    If HRZ Then                             'The HRZ and the focus wheel are acquiring Z stacks in opposite directions
        Zoffset = -ZShift + Zoffset
    Else
        Zoffset = ZShift + Zoffset
    End If
    BSliderZOffset.Value = Zoffset          'Writes the calculated value in the offset value
Abort:
   RestoreAquisitionParameters
    Set GlobalBackupRecording = Nothing
   Lsm5Vba.Application.ThrowEvent eRootReuse, 0
    DoEvents                                'Finnish everything which had started
    'ActivateAcquisitionTrack                'Activates the tracks for image acquisition
    
    If ScanStop = True Then
        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
        ScanStop = False
    Else
        DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
    End If
End Sub


Public Sub SetBlockValues()
    Dim Position As Long
    Dim Range As Double
 
    CheckBoxHighSpeed.Value = BlockHighSpeed
    CheckBoxLowZoom.Value = BlockLowZoom
    CheckBoxHRZ.Value = BlockHRZ
    Position = Lsm5.Hardware.CpObjectiveRevolver.RevolverPosition
    If Position >= 0 Then
        Range = Lsm5.Hardware.CpObjectiveRevolver.FreeWorkingDistance(Position) * 1000#
    Else
        Range = 0#
    End If
    If BlockZRange > Range * 0.9 Then
        BlockZRange = Range * 0.9
    End If
    If Abs(BlockZOffset) > Range * 0.9 Then
        BlockZOffset = 0
    End If
    BSliderZOffset.Value = BlockZOffset
    BSliderZRange.Value = BlockZRange
    BSliderZStep.Value = BlockZStep

End Sub


Public Sub GetBlockValues()
   
    BlockHighSpeed = CheckBoxHighSpeed.Value
    BlockLowZoom = CheckBoxLowZoom.Value
    HRZ = CheckBoxHRZ.Value
    BlockZOffset = BSliderZOffset.Value
    BlockZRange = BSliderZRange.Value
    BlockZStep = BSliderZStep.Value
End Sub



Private Function TimeDisplay(Value As Double) As String         'Calculates the String to display in a "user frindly format". Value is in seconds
    Dim Hour, Min As Integer
    Dim Sec As Double

    Hour = Int(Value / 3600)                                        'calculates number of full hours                           '
    Min = Int(Value / 60) - (60 * Hour)                             'calculates number of left minutes
    Sec = (Fix((Value - (60 * Min) - (3600 * Hour)) * 100)) / 100   'calculates the number of left seconds
    If (Hour = 0) And (Min = 0) Then                                'Defines a "user friendly" string to display the time
        TimeDisplay = Sec & " sec"
    ElseIf (Hour = 0) And (Sec = 0) Then
        TimeDisplay = Min & " min"
    ElseIf (Hour = 0) Then
        TimeDisplay = Min & " min " & Sec
    Else
        TimeDisplay = Hour & " h " & Min
    End If
End Function


Public Function AcquisitionTime() As Double
    Dim Track1Speed, Track2Speed, Track3Speed, Track4Speed As Double
    Dim Pixels As Long
    Dim FrameNumber As Integer
    Dim ScanDirection As Integer
    Dim i As Integer
   
    Track1Speed = 0
    Track2Speed = 0
    Track3Speed = 0
    Track4Speed = 0
    If CheckBoxTrack1.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, success)
        Track1Speed = Track.SampleObservationTime
    End If
    If CheckBoxTrack2.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, success)
        Track2Speed = Track.SampleObservationTime
    End If
    If CheckBoxTrack3.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(2, success)
        Track3Speed = Track.SampleObservationTime
    End If
    If CheckBoxTrack4.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(3, success)
        Track4Speed = Track.SampleObservationTime
    End If
    Pixels = Lsm5.DsRecording.LinesPerFrame * Lsm5.DsRecording.SamplesPerLine
    FrameNumber = Lsm5.DsRecording.FramesPerStack
    If Lsm5.DsRecording.ScanDirection = 0 Then
        ScanDirection = 1
    Else
        ScanDirection = 2
    End If
    AcquisitionTime = (Track1Speed + Track2Speed + Track3Speed + Track4Speed) * Pixels * FrameNumber / ScanDirection * 3.3485
End Function



Private Sub CheckBoxTrack1_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub

Private Sub CheckBoxTrack2_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub

Private Sub CheckBoxTrack3_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub

Private Sub CheckBoxTrack4_Change()
'    AcquisitionTimeFrame.Caption = TimeDisplay(AcquisitionTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
    FillTrackingChannelList
End Sub



Public Function AutofocusTime() As Double
    Dim Speed As Double
    Dim Pixels As Long
    Dim FrameNumber As Integer
    Dim ScanDirection As Integer
    Dim i As Integer

    Speed = 0
    If CheckBoxHighSpeed.Value = True Then
        Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, success)
        Speed = 1.76 * 10 ^ -6
    Else
        If OptionButtonTrack1.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, success)
            Speed = Track.SampleObservationTime
        End If
        If OptionButtonTrack2.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, success)
            Speed = Track.SampleObservationTime
        End If
        If OptionButtonTrack3.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, success)
            Speed = Track.SampleObservationTime
        End If
        If OptionButtonTrack4.Value = True Then
            Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(1, success)
            Speed = Track.SampleObservationTime
        End If
    End If
    Pixels = 512
    AutofocusForm.GetBlockValues
    FrameNumber = CLng(BlockZRange / BlockZStep) + 1
    If Lsm5.DsRecording.ScanDirection = 0 Then
        ScanDirection = 1
    Else
        ScanDirection = 2
    End If
    If CheckBoxHRZ.Value = True Then
        AutofocusTime = Speed * Pixels * FrameNumber * 3.3485 + 4.9
    Else
        AutofocusTime = Speed * Pixels * FrameNumber / ScanDirection * 3.3485 + 4.9
    End If
End Function



Public Sub CheckBoxHRZ_Change() 'I was trying to display the time needed for autofocus, single image acquisition and total time of the experiments, but I gave and and commented out those functions
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Public Sub CheckBoxHighSpeed_Change()  'I was trying to display the time needed for autofocus, single image acquisition and total time of the experiments, but I gave and and commented out those functions
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub BSliderZStep_Change()  'I was trying to display the time needed for autofocus, single image acquisition and total time of the experiments, but I gave and and commented out those functions
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack1_Click()
    If OptionButtonTrack1.Value Then 'if track 1 checked others are not autofocus track but false
        OptionButtonTrack2.Value = Not OptionButtonTrack1.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack1.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack1.Value
        CheckAutofocusTrack (1) 'sets SelectedTrack to 1, see below
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack2_Click()
    If OptionButtonTrack2.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack2.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack2.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack2.Value
        CheckAutofocusTrack (2)
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack3_Click()
    If OptionButtonTrack3.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack3.Value
        OptionButtonTrack2.Value = Not OptionButtonTrack3.Value
        OptionButtonTrack4.Value = Not OptionButtonTrack3.Value
        CheckAutofocusTrack (3)
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub

Private Sub OptionButtonTrack4_Click()
    If OptionButtonTrack4.Value Then
        OptionButtonTrack1.Value = Not OptionButtonTrack4.Value
        OptionButtonTrack2.Value = Not OptionButtonTrack4.Value
        OptionButtonTrack3.Value = Not OptionButtonTrack4.Value
        CheckAutofocusTrack (4)
    End If
'    AutofocusTimeFrame.Caption = TimeDisplay(AutofocusTime)
'    TotalTimeLeftFrame.Caption = TimeDisplay(TotalTimeLeft)
End Sub


'this Function checks whether the track that was selected for autofocusing only contains a single channel (alternetivly defines one of the checked channels)
'and finds the name of the autofocusing channel
Private Sub CheckAutofocusTrack(SelectedTrack)
    Dim Track As DsTrack 'a new track is defined
    Dim DataChannel As DsDataChannel    'a new interface to a data channel is defined
                                        'contains channel dependend parameters of the
                                        'scan memory/display/calculation of scan data during scan operation
    Dim ActiveChannelNumber As Integer
    Dim AutofocusChannel As String
    
    Set Track = Lsm5.DsRecording.TrackObjectByMultiplexOrder(SelectedTrack - 1, success)
        'gets the track object by multiplexorder which starts with 0
        'since selected track starts with 1 (see CheckAutofocusTrack (n)), 1 has to be substracted
        
    'the following loop will count the number of activated channels in the track chosen for autofocusing
    ActiveChannelNumber = 0
    
    For j = 0 To Track.DataChannelCount - 1 'gets number of channels that are potentially activatable in track
        Set DataChannel = Track.DataChannelObjectByIndex(j, success) 'data channel corresponding to loop index is analysed
        If DataChannel.Acquire = True Then  'checks whether the data channel corresponding to loop index is activated
            ActiveChannelNumber = ActiveChannelNumber + 1 'counts the number of activated channels
            If ActiveChannelNumber = 1 Then AutofocusChannel = DataChannel.name 'Gets the name of the first activated channel
        End If
    Next
    
    If ActiveChannelNumber > 1 Then 'if more than one channel is activated...
        MsgBox ("The track you selected has more than one active Channel. " & AutofocusChannel & " will be used to calculate autofocus parameters.")
    End If
End Sub



Public Function TotalTimeLeft() As Double
    Dim Speed As Double
    Dim Pixels As Long
    Dim FrameNumber As Integer
    Dim ScanDirection As Integer
    Dim i As Integer
    TotalTimeLeft = (AcquisitionTime + AutofocusTime + BlockTimeDelay) * (BlockRepetitions - RepetitionNumber + 1) - BlockTimeDelay
End Function





Public Sub AutofocusSetting(HRZ As Boolean, HighSpeed As Boolean)
If LowZoom Then                                         'Changes the zoom if necessary
        Lsm5.DsRecording.ZoomX = 1
        Lsm5.DsRecording.ZoomY = 1
    End If
        
    Lsm5.DsRecording.TimeSeries = False                     'Disable the timeseries, because autofocussing is juste one image at one timepoint.
    
    If FrameAutofocussing Then                              'Setting the way the Stage is going to move in Z, plus speed and number of pixels
        Lsm5.DsRecording.ScanMode = "Stack"                 'This is defining to acquire a Z stack of Z-Y images
        If HRZ Then
            Lsm5.DsRecording.SpecialScanMode = "ZScanner"
        Else
' !!!!!!!!!!!! potential error source  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' improvement
    If SystemName = "LSM" Then
    Lsm5.DsRecording.SpecialScanMode = "FocusStep"

'Lsm5.DsRecording.FrameSpacing = ZStep
'    NoFrames = CLng(ZRange / ZStep) + 1
'    Lsm5.DsRecording.FramesPerStack = NoFrames
'    If NoFrames > 2048 Then
'        NoFrames = 2048
'    End If
    Lsm5.DsRecording.Sample0Z = ZStep * NoFrames / 2
    Else
            Lsm5.DsRecording.SpecialScanMode = "OnTheFly"
             Lsm5.DsRecording.FramesPerStack = 1201
            Lsm5.DsRecording.Sample0Z = Range / 2
            Lsm5.DsRecording.FrameSpacing = Range / 1200
            Sleep (100)
        End If
        End If
        If HighSpeed Then
            Lsm5.DsRecording.ScanDirection = 1                  'If Highspeed is selected it uses the bidirectionnal scanning
        End If
        If SystemName = "LIVE" Then
            Lsm5.DsRecording.RtRegionWidth = BSliderFrameSize.Value 'If doing frame autofocussing it uses the userdefined frame size
            Lsm5.DsRecording.RtBinning = 512 / BSliderFrameSize.Value
            Lsm5.DsRecording.RtRegionHeight = BSliderFrameSize.Value
        ElseIf SystemName = "LSM" Then
            Lsm5.DsRecording.SamplesPerLine = BSliderFrameSize.Value  'If doing frame autofocussing it uses the userdefined frame size
            'Lsm5.DsRecording.RtBinning = 4
            Lsm5.DsRecording.LinesPerFrame = BSliderFrameSize.Value
        Else
            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
        Exit Sub
        End If
    Else
        Lsm5.DsRecording.ScanMode = "ZScan"                     'This is defining to acquire a single X-Z image, like what is done with the "Range" button in the LSM ScanControl window
        If HRZ Then
            Lsm5.DsRecording.SpecialScanMode = "ZScanner"
           If SystemName = "LIVE" Then
                Lsm5.DsRecording.RtLinePeriod = 1 / BSliderScanSpeed.Value
                Lsm5.DsRecording.RtRegionWidth = 512
                Lsm5.DsRecording.RtRegionHeight = 1
            ElseIf SystemName = "LSM" Then
                Lsm5.DsRecording.SamplesPerLine = 256
                Lsm5.DsRecording.LinesPerFrame = 1
            Else
                MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
                Exit Sub
            End If
           
        Else
'
            Lsm5.DsRecording.SpecialScanMode = "FocusStep"
            'Lsm5.DsRecording.SpecialScanMode = "OnTheFly" ' I this does not work om old LSM, use "FocusStep"
            
            
       
        End If
         If SystemName = "LIVE" Then
            Lsm5.DsRecording.RtLinePeriod = 1 / BSliderScanSpeed.Value
            Lsm5.DsRecording.RtRegionWidth = 512
            Lsm5.DsRecording.RtRegionHeight = 1
            
       
         ElseIf SystemName = "LSM" Then
                Lsm5.DsRecording.SamplesPerLine = 256
                Lsm5.DsRecording.LinesPerFrame = 1
        Else
                MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
                Exit Sub
        End If
        End If
     Sleep (100)
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''Grid Definition'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextBoxXgrid_Change()
            If IsNumeric(TextBoxXGrid.Value) Then
            GlobalXGrid = TextBoxXGrid.Value
        Else
             MsgBox "Please enter the number of columns!"
        End If
        
        ReDim GlobalDeActivatedLocations(GlobalMaximumPositions, GlobalXGrid * GlobalYGrid)
        ReDim GlobalLocationsOrder(GlobalMaximumPositions, GlobalXGrid * GlobalYGrid)
        GlobalOrderChanged = True
    End Sub

Private Sub TextBoxYGrid_Change()
         If IsNumeric(TextBoxYGrid.Value) Then
            GlobalYGrid = TextBoxYGrid.Value
        Else
             MsgBox "Please enter the number of rows!"
        End If
        
        
        ReDim GlobalDeActivatedLocations(GlobalMaximumPositions, GlobalXGrid * GlobalYGrid)
        ReDim GlobalLocationsOrder(GlobalMaximumPositions, GlobalXGrid * GlobalYGrid)
        GlobalOrderChanged = True

End Sub
    

Private Sub TextBoxXStep_Change()
If IsNumeric(TextBoxXStep.Value) Then
            GlobalXStep = TextBoxXStep.Value
        Else
             MsgBox "Please enter the horizontal distance between two neighbouring locations!"
        End If
    
End Sub


Private Sub TextBoxYStep_Change()

If IsNumeric(TextBoxYStep.Value) Then
            GlobalYStep = TextBoxYStep.Value
        Else
             MsgBox "Please enter the vertical distance between two neighbouring locations!"
        End If

   
End Sub
Private Sub CommandButtonGrid_Click()
    ShowGrid
End Sub


Public Sub ShowGrid()
Dim BitsPerSample As Long
Dim bpp As Long
Dim ImgName As String
Dim LsmMath As New LsmVectorMath
Dim SpareArray() As Single
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim xIndx As Long
Dim yIndx As Long
Dim Start As Long
Dim iX As Long
Dim iY As Long
Dim PlaneSize As Long
    If Not GlobalGridImage Is Nothing Then
        If Not GlobalGridImage.IsValid Then
            Set GlobalGridImage = Nothing
        Else
             GlobalGridImage.CloseAllWindows
             Set GlobalGridImage = Nothing
            
        End If
    End If
    If GlobalGridImage Is Nothing Then
        BitsPerSample = 12
        bpp = 2
        XPixels = 1024
        YPixels = 1024
        XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
        YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
        If XGroup <= YGroup Then
            YGroup = XGroup
        Else
            XGroup = YGroup
        End If
        XPixels = XGroup * (3 * GlobalXGrid + 1)
        YPixels = YGroup * (3 * GlobalYGrid + 1)
        MakeBlankImage GlobalGridImage, BitsPerSample, bpp, True, ImgName, False, 1, XPixels, YPixels, 3
        GlobalGridImage.SetTitle "Location Grid"
        GlobalGridImage.NeverAgainScanToTheImage
        RedrawGrid GlobalGridImage
        GlobalGridImage.VectorOverlay.LineWidth = 1
        GlobalGridImage.VectorOverlay.Color = RGB(0, 255, 0)
        GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
        If GlobalReferencePoints = 2 Then
            DrawCrossGrid GlobalGridX1, GlobalGridY1
            DrawCrossGrid GlobalGridX2, GlobalGridY2
        ElseIf GlobalReferencePoints = 1 Then
            DrawCrossGrid GlobalGridX1, GlobalGridY1
        End If
        
        GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowNoButtonMouseMoveEvent, 1
        GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLButtonMouseMoveEvent, 1
        GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLeftButtonDownEvent, 1
        GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLeftButtonUpEvent, 1
        GlobalGridImage.EnableImageWindowEvent ds.eImageWindowRightButtonUpEvent, 1
        
        DoEvents
    End If
End Sub

Private Sub CommandButtonRemove_Click()
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Do You Want to Remove Selected Reference Points?"
    Style = VbYesNo + VbQuestion + VbDefaultButton2   ' Define buttons.
    Title = "Remove Reference Points"  ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then    ' User chose Yes.
        GlobalReferencePoints = 0
        If Not GlobalGridImage Is Nothing Then
            If GlobalGridImage.IsValid Then
                GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
            End If
        End If
    End If
End Sub

Private Sub CreateLocationsButton_Click()
Dim XPos As Double
    Dim YPos As Double
    Dim ZPos As Double
    
    Dim idx As Long
    Dim idy As Long
    Dim idt As Long
    Dim x As Double
    Dim y As Double
    Dim det As Double
    Dim A11 As Double
    Dim A12 As Double
    Dim A21 As Double
    Dim A22 As Double
    Dim x1 As Double
    Dim Y1 As Double
    
    Dim result As Long
    Dim ProgressString As String
    Dim Color As Long
    Dim ZeroChanged As Boolean
    Dim SetZeroMarked As Boolean
    Dim idpos As Long
    Dim idold As Long
    Dim OverWriteZ As Boolean
    
    
    
  If GridToggle Then
    
      OverWriteZ = PubFuncOverWriteZ
   
    
    ProgressString = "Please Wait..."
    Color = RGB(&HC0, 0, 0)
    DisplayProgress ProgressString, Color
    DoEvents

    If TextBoxYGrid.Value * TextBoxXGrid.Value <= GlobalMaximumPositions Then
        flgUserChange = False
        GlobalXGrid = TextBoxXGrid.Value
        GlobalYGrid = TextBoxYGrid.Value
        GlobalXStep = TextBoxXStep.Value
        GlobalYStep = TextBoxYStep.Value
        
        
'        Stage.MarkClearAll
        ZPos = GlobalGridStageZ1
        If GlobalReferencePoints >= 1 Then
            XPos = GlobalGridStageX1 - (GlobalGridX1 - 1) * GlobalXStep
            YPos = GlobalGridStageY1 - (GlobalGridY1 - 1) * GlobalYStep
'        Else does not wordk properly
'            x = GetStagePositionX
'            y = GetStagePositionY
'            ConvertToStagePositionXY x, y, XPos, YPos
        End If
        GlobalPositionsRecalled = GlobalPositionsStage
        GlobalPositionsStage = GlobalXGrid * GlobalYGrid
        ReDim GlobalLocationsName(GlobalPositionsStage)
        GlobalCurrentPosition = 1
        If GlobalPositionsStage <= 1 Then
            GlobalPositionsStage = 1
            GlobalCurrentPosition = 1
            StartStopForm.ComboBoxLocation.AddItem "Present"
        Else
            det = -1 * 1 - 0 * 0
          '  det = X11 * X22 - X21 * X12
            If det = 0 Then
                GlobalPositionsStage = 1
                GlobalCurrentPosition = 1
               ' StartStopForm.ComboBoxLocation.AddItem "Present"
            
            Else
                SetZeroMarked = True
                GetGlobalZZero SetZeroMarked, ZeroChanged
                x1 = XPos
                Y1 = YPos
                ReDim GlobalXpos(GlobalXGrid * GlobalYGrid)
                ReDim GlobalYpos(GlobalXGrid * GlobalYGrid)
                If OverWriteZ Then ReDim GlobalZpos(GlobalXGrid * GlobalYGrid)
            
                
                If GlobalMeander Then
                    idpos = 0
                    For idy = 1 To GlobalYGrid Step 2
                        For idx = 1 To GlobalXGrid
                            idt = (idy - 1) * GlobalXGrid + idx
                            If Not GlobalDeActivatedLocations(idx, idy) Then
                                idpos = idpos + 1
                                
                                GlobalLocationsOrder(idx, idy) = idpos
                                GlobalLocationsName(idpos) = "Column" & idx & "_Row" & idy
                                GlobalXpos(idpos) = x1 + (idx - 1) * GlobalXStep
                                GlobalYpos(idpos) = Y1 + (idy - 1) * GlobalYStep
                                If OverWriteZ Then
                                    GlobalZpos(idpos) = ZPos
                                Else
                                    idold = GlobalLocationsOrderOld(idx, idy)
                                    If idold = -1 Then
                                        GlobalZpos(idpos) = ZPos
                                    Else
                                        GlobalZpos(idpos) = GlobalZposOld(idold)
                                    End If
                                End If
                            Else
                                GlobalLocationsOrder(idx, idy) = -1
                            End If
                        Next idx
                        If idt >= GlobalXGrid * GlobalYGrid Then Exit For
                        For idx = 1 To GlobalXGrid
                            idt = idy * GlobalXGrid + GlobalXGrid - idx + 1
                            If Not GlobalDeActivatedLocations(GlobalXGrid - idx + 1, idy + 1) Then
                                idpos = idpos + 1
                                
                                GlobalLocationsOrder(GlobalXGrid - idx + 1, idy + 1) = idpos
                                GlobalLocationsName(idpos) = "Column" & GlobalXGrid - idx + 1 & "_Row" & idy + 1
                                GlobalXpos(idpos) = x1 + (GlobalXGrid - idx) * GlobalXStep
                                GlobalYpos(idpos) = Y1 + idy * GlobalYStep
                                If OverWriteZ Then
                                    GlobalZpos(idpos) = ZPos
                                Else
                                idold = GlobalLocationsOrderOld(GlobalXGrid - idx + 1, idy + 1)
                                    If idold = -1 Then
                                        GlobalZpos(idpos) = ZPos
                                    Else
                                        GlobalZpos(idpos) = GlobalZposOld(idold)
                                    End If
                                End If
                            Else
                                GlobalLocationsOrder(GlobalXGrid - idx + 1, idy + 1) = -1
                            End If
                        Next idx
                    Next idy
                Else
                    idpos = 0
                    For idy = 1 To GlobalYGrid
                        For idx = 1 To GlobalXGrid
                            idt = (idy - 1) * GlobalXGrid + idx
                            If Not GlobalDeActivatedLocations(idx, idy) Then
                                idpos = idpos + 1
                                GlobalLocationsName(idpos) = "Column" & idx & "_Row" & idy
                                GlobalLocationsOrder(idx, idy) = idpos
                                GlobalXpos(idpos) = x1 + (idx - 1) * GlobalXStep
                                GlobalYpos(idpos) = Y1 + (idy - 1) * GlobalYStep
                                If OverWriteZ Then
                                    GlobalZpos(idpos) = ZPos
                                Else
                                    idold = GlobalLocationsOrderOld(idx, idy)
                                    If idold = -1 Then
                                        GlobalZpos(idpos) = ZPos
                                    Else
                                        GlobalZpos(idpos) = GlobalZposOld(idold)
                                    End If
                                End If
                            Else
                                GlobalLocationsOrder(idx, idy) = -1
                            End If
                        Next idx
                    Next idy
                End If
            End If
        End If
        GlobalOrderChanged = False
        GlobalPositionsStage = idpos
        GLobalPositions = GlobalPositionsStage
        flgUserChange = False
        flgUserChange = True
        GlobalIsTile = False
    Else
        MsgBox "Number of locations in the grid cannot exceed " + CStr(GlobalMaximumPositions)
    End If
    Tile
    DisplayProgress GlobalProgressString, GlobalColor
    DoEvents
     AutofocusForm.SetMarkedLocations GlobalPositionsStage
     PubSentStageGrid = True
     
     ElseIf MultipleLocationToggle = True Then
  
    GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
     PutStagePositionsInArray
     Tile
     AutofocusForm.SetMarkedLocations GlobalPositionsStage
     End If
End Sub

Public Sub DisplayGridSelection(x As Long, y As Long, xIndx As Long, yIndx As Long)
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim Start As Long
Dim iX As Long
Dim iY As Long
Dim Xmin As Long
Dim Xmax As Long
Dim Ymin As Long
Dim Ymax As Long
Dim xImage As Long
Dim StartX As Long
Dim yImage As Long
Dim StartY As Long
Dim Found As Boolean
Dim EndX As Long
Dim EndY As Long
Dim XPosition As Long
Dim YPosition As Long
Dim xMod As Long
Dim yMod As Long
Dim xString As String
Dim yString As String
Dim MyString As String

    If GlobalGridImage Is Nothing And GlobalZGridImage Is Nothing Then Exit Sub
    
    XPixels = GlobalGridImage.GetDimensionX
    YPixels = GlobalGridImage.GetDimensionY
    
    XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
    YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
    XPosition = Int(x / XGroup) + 1
    YPosition = Int(y / YGroup) + 1
    xMod = XPosition Mod 3
    yMod = YPosition Mod 3
    xIndx = XPosition / 3
    yIndx = YPosition / 3
TileX = AutofocusForm.TextBoxTileX.Value
TileY = AutofocusForm.TextBoxTileY.Value
    If xMod = 1 Or yMod = 1 Then
        xString = ""
        yString = ""
    Else
        xString = CStr(xIndx)
        yString = CStr(yIndx)
    End If
  '  LabelGrid.ForeColor = RGB(0, 0, 0)
    MyString = "Column= " + xString + vbCrLf + "Row= " + yString
    If Not GlobalOrderChanged Then
        If GlobalLocationsOrder(xIndx, yIndx) > 0 Then
            MyString = MyString + vbCrLf + "Mark= " + CStr((GlobalLocationsOrder(xIndx, yIndx) - 1) * TileX * TileY + 1) + " to " + CStr((GlobalLocationsOrder(xIndx, yIndx)) * TileX * TileY)
        End If
    End If
    If GlobalZmapAquired Then
         If GlobalLocationsOrder(xIndx, yIndx) > -1 Then
            idpos = GlobalLocationsOrder(xIndx, yIndx)
            MyString = MyString + vbCrLf + "z-Value= " + CStr(Round(GlobalZpos(idpos), 3))
            End If
         End If
    DisplayProgress MyString, RGB(0, &HC0, 0)
End Sub

Public Sub DrawCrossGrid(xIndx As Long, yIndx As Long)
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim iX As Long
Dim iY As Long
Dim x1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long

    If (GlobalGridImage Is Nothing) Then Exit Sub
    XPixels = GlobalGridImage.GetDimensionX
    YPixels = GlobalGridImage.GetDimensionY
    XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
    YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
    x1 = XGroup + 3 * XGroup * (xIndx - 1)
    Y1 = YGroup + 3 * YGroup * (yIndx - 1)
    X2 = 3 * XGroup + 3 * XGroup * (xIndx - 1)
    Y2 = 3 * YGroup + 3 * YGroup * (yIndx - 1)
      
    GlobalGridImage.VectorOverlay.Color = RGB(255, 255, 0)
    GlobalGridImage.VectorOverlay.LineWidth = 1
    GlobalGridImage.VectorOverlay.AddSimpleDrawingElement Lsm5Vba.eDrawingModeLine, x1, Y1, X2, Y2
    GlobalGridImage.VectorOverlay.AddSimpleDrawingElement Lsm5Vba.eDrawingModeLine, x1, Y2, X2, Y1
'    dsDoc.VectorOverlay.AddSimpleDrawingElement Lsm5Vba.eDrawingModeCircle, xCross, yCross, xCross, yCross + 30
    GlobalGridImage.RedrawImage
    
End Sub

Public Sub DrawRectangleGrid(x1 As Long, Y1 As Long, X2 As Long, Y2 As Long)

    If (GlobalGridImage Is Nothing) Then Exit Sub
      
    GlobalGridImage.VectorOverlay.Color = RGB(0, 255, 0)
    GlobalGridImage.VectorOverlay.LineWidth = 1
    GlobalGridImage.VectorOverlay.AddSimpleDrawingElement Lsm5Vba.eDrawingModeRectangle, x1, Y1, X2, Y2
    GlobalGridImage.RedrawImage
    
End Sub


Public Sub SetMarkedLocations(Positions As Long)
    Dim idx As Long
    Dim ZeroChanged As Boolean
    Dim SetZeroMarked As Boolean
    SetZeroMarked = False
 '   If GlobalIsStage And GlobalMultiLocation Then
        If Positions > 1 Then
            GetGlobalZZero SetZeroMarked, ZeroChanged
'            If ZeroChanged Then
'                FillLocationList
'            End If
            Lsm5.Hardware.CpStages.MarkClearAll
            For idx = 1 To Positions
                Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idx), GlobalYpos(idx), GlobalZpos(idx)
            Next idx
        End If
 '   End If
End Sub


Private Sub CheckBoxMeander_Click()
 If flgUserChange Then
        GlobalMeander = CheckBoxMeander.Value
    End If

End Sub

'Private Sub CheckBoxKeepSteps_Click()
' If flgUserChange Then
'        GlobalKeepSteps = CheckBoxKeepSteps.Value
'    End If
'End Sub


Private Sub ZMapButton_Click()
Dim text As String
Dim x As Double
Dim y As Double
Dim z As Double
Dim Zbefore As Double
Dim BitsPerSample As Long
Dim bpp As Long
Dim ImgName As String
Dim SpareArray() As Single
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim xIndx As Long
Dim yIndx As Long
Dim Start As Long
Dim iX As Long
Dim iY As Long
Dim PlaneSize As Long
Dim idold As Long

ZValues.Show
InitializeStageProperties
SetStageSpeed 8, True

If PubSentStageGrid = False And (Grid Or StripeScanToggle.Value) Then
    MsgBox "Please send the grid information to stage first!", VbExclamation
    Exit Sub
End If

AutofocusForm.GetBlockValues 'Updates the parameters value for BlockZRange, BlockZStep..
'DisplayProgress "Aquiring Reference", RGB(0, &HC0, 0)
Lsm5.StopScan
Lsm5Vba.Application.ThrowEvent eRootReuse, 0
DoEvents
StoreAquisitionParameters

'got to Refcor1


'Lsm5.Hardware.CpStages.PositionX = GlobalGridStageX1
'Lsm5.Hardware.CpStages.PositionY = GlobalGridStageY1
'While Lsm5.Hardware.CpStages.IsBusy
'    DoEvents
'Wend
'Lsm5.Hardware.CpFocus.Position = GlobalGridStageZ1
'While Lsm5.Hardware.CpFocus.IsBusy
'    DoEvents
'Wend
'Sleep (20)
''Aquire Z-Stack,Caluclate shift
'
'
' Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
' If PubAbort Then GoTo Abort
''Caluclate new z Position, Store Z in Array
'GlobalGridStageZ1 = GlobalGridStageZ1 + ZShift
 GettingZmap = True
GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
    If MultipleLocation Then
    PutStagePositionsInArray
    End If
    

If Grid Then
    If GlobalMeander Then
         For idpos = 1 To GlobalPositionsStage
            If ScanStop Then
                DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                GoTo Abort
             End If
           
          
        If GlobalStageControlZValues = False Then
           If idpos = 1 Then
               Lsm5.Hardware.CpFocus.Position = GlobalGridStageZ1
               Zbefore = GlobalGridStageZ1
               While Lsm5.Hardware.CpFocus.IsBusy
                  DoEvents
               Wend
           Else
               Zbefore = GlobalZpos(idpos - 1)
               Lsm5.Hardware.CpFocus.Position = GlobalZpos(idpos - 1)
               While Lsm5.Hardware.CpFocus.IsBusy
                  DoEvents
               Wend
           End If
          
        Else
            Lsm5.Hardware.CpFocus.Position = GlobalZpos(idpos)
            While Lsm5.Hardware.CpFocus.IsBusy
                  DoEvents
               Wend
        End If
         Lsm5.Hardware.CpStages.PositionX = GlobalXpos(idpos)
        Lsm5.Hardware.CpStages.PositionY = GlobalYpos(idpos)
           Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
                    If ScanStop Then
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        GoTo Abort
                    End If
                    DoEvents
                    Sleep (5)
                Loop
           DisplayProgress "Aquiring idpos " & idpos & ", z= " & GlobalZpos(idpos), RGB(0, &HC0, 0)
            Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
                    If ScanStop Then
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        GoTo Abort
                    End If
             'check if Z shift makes sense
          CheckRefControl BlockZRange
           'Calculate new Z and Store in Array
           GlobalZpos(idpos) = Lsm5.Hardware.CpFocus.Position + BlockZOffset + ZShift
           
           Next idpos
    Else
       For Row = 1 To GlobalYGrid
           For column = 1 To GlobalXGrid
              If Not GlobalDeActivatedLocations(column, Row) Then
                  idpos = GlobalLocationsOrder(column, Row)
                  Lsm5.Hardware.CpStages.PositionX = GlobalXpos(idpos)
                  Lsm5.Hardware.CpStages.PositionY = GlobalYpos(idpos)
                  If Row = 1 And column = 1 Then
                      Lsm5.Hardware.CpFocus.Position = GlobalGridStageZ1
                      Zbefore = GlobalGridStageZ1
                  ElseIf Row <> 1 And column = 1 Then
                  idold = GlobalLocationsOrder(column, Row - 1)
                    If idold = -1 Then
                        Zbefore = GlobalZpos(idpos - 1)
                    Else
                            Zbefore = GlobalZpos(idold)
                    End If
                      Lsm5.Hardware.CpFocus.Position = Zbefore
                Else
                      Zbefore = GlobalZpos(idpos - 1)
                      Lsm5.Hardware.CpFocus.Position = GlobalZpos(idpos - 1)
                  End If
                Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
                    If ScanStop Then
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        GoTo Abort
                    End If
                    DoEvents
                    Sleep (5)
                Loop
                  DisplayProgress "Aquiring idpos " & idpos & ", z= " & GlobalZpos(idpos), RGB(0, &HC0, 0)
                  'Aquire Z-Stack,Caluclate shift
                   Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
                    If ScanStop Then
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        GoTo Abort
                    End If
             'check if Z shift makes sense
          CheckRefControl BlockZRange
           'Calculate new Z and Store in Array
                  GlobalZpos(idpos) = Lsm5.Hardware.CpFocus.Position + BlockZOffset + ZShift
               End If
          Next column
       Next Row
    End If
    Lsm5.Hardware.CpStages.MarkClearAll
        For idpos = 1 To GlobalPositionsStage
     '    GlobalZpos(idpos) = GlobalZpos(idpos)  + ZShift '?????? suchmarke
           Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
        Next idpos
Else ' Zmap in multilocation modus
GlobalPositionsStage = Lsm5.Hardware.CpStages.MarkCount
    For idpos = 1 To GlobalPositionsStage
                    If ScanStop Then
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                       GoTo Abort
                    End If
                    Lsm5.Hardware.CpStages.PositionX = GlobalXpos(idpos)
                  Lsm5.Hardware.CpStages.PositionY = GlobalYpos(idpos)
                  Lsm5.Hardware.CpFocus.Position = GlobalZpos(idpos)
 '                 Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
'            Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToMarkZ (0) 'Moves to the first location marked in the stage control
              Do While Lsm5.Hardware.CpStages.IsBusy Or Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy ' Wait that the movement is done
                    If ScanStop Then
                       DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                      GoTo Abort
                    End If
                    DoEvents
                    Sleep (5)
                Loop
              DisplayProgress "Aquiring idpos " & idpos, RGB(0, &HC0, 0)
            Autofocus_StackShift BlockZRange, BlockZStep, BlockHighSpeed, BlockZOffset
                    If ScanStop Then
                        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
                        GoTo Abort
                    End If
             'check if Z shift makes sense
          CheckRefControl BlockZRange
           'Calculate new Z and Store in Array
           GlobalZpos(idpos) = Lsm5.Hardware.CpFocus.Position + BlockZOffset + ZShift
           
        
'           success = Lsm5.Hardware.CpStages.MarkGet(0, x, y)
 '          success = Lsm5.Hardware.CpStages.MarkClear(0)
'
'            Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ x, y, z
          Next idpos
          Lsm5.Hardware.CpStages.MarkClearAll
          For idpos = 1 To GlobalPositionsStage
          Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
          Next idpos
End If
'Lsm5.Hardware.CpStages.MarkClearAll
'        For idpos = 1 To GlobalPositionsStage
'         GlobalZpos(idpos) = GlobalZpos(idpos) + ZShift
'           Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).lAddMarkZ GlobalXpos(idpos), GlobalYpos(idpos), GlobalZpos(idpos)
'        Next idpos
GlobalZmapAquired = True
DisplayProgress "Zmap ready", RGB(0, &HC0, 0)
If Grid Then
    If Not GlobalZGridImage Is Nothing Then
            If Not GlobalZGridImage.IsValid Then
                Set GlobalZGridImage = Nothing
            Else
                 GlobalZGridImage.CloseAllWindows
                 Set GlobalZGridImage = Nothing
                
            End If
        End If
        If GlobalZGridImage Is Nothing Then
            BitsPerSample = 12
            bpp = 2
            XPixels = 1024
            YPixels = 1024
            XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
            YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
            If XGroup <= YGroup Then
                YGroup = XGroup
            Else
                XGroup = YGroup
            End If
            XPixels = XGroup * (3 * GlobalXGrid + 1)
            YPixels = YGroup * (3 * GlobalYGrid + 1)
            MakeBlankImage GlobalZGridImage, BitsPerSample, bpp, True, ImgName, False, 1, XPixels, YPixels, 3
            GlobalZGridImage.SetTitle "Z Values"
            GlobalZGridImage.NeverAgainScanToTheImage
            
            GlobalZGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowNoButtonMouseMoveEvent, 1
            GlobalZGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLButtonMouseMoveEvent, 1
            GlobalZGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLeftButtonDownEvent, 1
            GlobalZGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLeftButtonUpEvent, 1
            GlobalZGridImage.EnableImageWindowEvent ds.eImageWindowRightButtonUpEvent, 1
        End If
    RedrawZGrid GlobalZGridImage
End If
'RestoreAquisitionParameters
'GettingZmap = False
CheckBoxZMap.Value = True
Abort:
RestoreAquisitionParameters
GettingZmap = False
PubAbort = False
If ScanStop = True Then
        DisplayProgress "Stopped", RGB(&HC0, 0, 0)
        ScanStop = False
End If
End Sub
