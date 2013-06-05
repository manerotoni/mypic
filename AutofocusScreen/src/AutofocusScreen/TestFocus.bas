Attribute VB_Name = "TestFocus"
''''''
'' AFTest1_Click()
'' Perform repeatealy Autofocus with FastZline and acquisition with stage only.
'' Uses No Z-track and Z-track
'''''
'Private Sub AFTest1_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest1Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest1Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'    Dim FilePath As String
'    Dim MaxTestRepeats As Integer
'    Dim TestNr As Integer
'    Dim pixelDwell As Double
'    Dim i As Integer
'    Log = True
'    Dim Zold As Double
'    Zold = posTempZ
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AcquisitionTrack1.Value = AutofocusTrack1.Value
'    AcquisitionTrack2.Value = AutofocusTrack2.Value
'    AcquisitionTrack3.Value = AutofocusTrack3.Value
'    AcquisitionTrack4.Value = AutofocusTrack4.Value
'
'
'
'    '''''''
'    ' No Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = False
'    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'    GlobalBackupRecording.SpecialScanMode = "FocusStep"
'    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest1_FastZLine_Stage_NoTrackZ") Then
'        Exit Function
'    End If
'
'    '''''''
'    ' Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = True
'    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'    GlobalBackupRecording.SpecialScanMode = "FocusStep"
'    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest1_FastZLine_Stage_TrackZ") Then
'        Exit Function
'    End If
'
'    AFTest1Run = True
'End Function
'
'
''''''
'' AFTest2_Click()
'' Perform repeatealy Autofocus with piezo and acquisition with piezo
'' Uses No Z-track and Z-track
'''''
'Private Sub AFTest2_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest2Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest2Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'    Log = True
'    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.Name) Then
'        MsgBox ("No piezo availabe! Cannot start tests.")
'        Exit Function
'    End If
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AcquisitionTrack1.Value = AutofocusTrack1.Value
'    AcquisitionTrack2.Value = AutofocusTrack2.Value
'    AcquisitionTrack3.Value = AutofocusTrack3.Value
'    AcquisitionTrack4.Value = AutofocusTrack4.Value
'    AutofocusMaxSpeed.Value = True
'    AutofocusFastZline = False
'    AutofocusHRZ.Value = True
'
'    '''''''
'    ' No Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = False
'    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
'    GlobalAcquisitionRecording.SpecialScanMode = "ZScanner"
'    GlobalBackupRecording.SpecialScanMode = "ZScanner"
'
'    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest2_Piezo_Piezo_NoTrackZ") Then
'        Exit Function
'    End If
'
'    '''''''
'    ' Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = True
'    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
'    GlobalAcquisitionRecording.SpecialScanMode = "ZScanner"
'    GlobalBackupRecording.SpecialScanMode = "ZScanner"
'
'    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest2_Piezo_Piezo_TrackZ") Then
'        Exit Function
'    End If
'    AFTest2Run = True
'End Function
'
'
''''''
'' AFTest3_Click()
'' Perform repeatealy Autofocus with stage and acquisition with stage
'' Uses No Z-track and Z-track
'''''
'Private Sub AFTest3_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest3Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest3Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'    Log = True
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AcquisitionTrack1.Value = AutofocusTrack1.Value
'    AcquisitionTrack2.Value = AutofocusTrack2.Value
'    AcquisitionTrack3.Value = AutofocusTrack3.Value
'    AcquisitionTrack4.Value = AutofocusTrack4.Value
'    AutofocusMaxSpeed.Value = True
'    AutofocusFastZline = False
'    AutofocusHRZ.Value = False
'
'    '''''''
'    ' No Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = False
'    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
'    GlobalBackupRecording.SpecialScanMode = "FocusStep"
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest3_Stage_Stage_NoTrackZ") Then
'        Exit Function
'    End If
'
'    '''''''
'    ' Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = True
'    GlobalBackupRecording.SpecialScanMode = "FocusStep"
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest3_Stage_Stage_TrackZ") Then
'        Exit Function
'    End If
'    AFTest3Run = True
'End Function
'
''''''
'' AFTest4_Click()
'' Perform repeatealy Autofocus with piezo and acquisition with stage
'' Uses No Z-track and Z-track
'''''
'Private Sub AFTest4_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest4Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest4Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'    Log = True
'    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.Name) Then
'        MsgBox ("No piezo availabe! Cannot start tests.")
'        Exit Function
'    End If
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AcquisitionTrack1.Value = AutofocusTrack1.Value
'    AcquisitionTrack2.Value = AutofocusTrack2.Value
'    AcquisitionTrack3.Value = AutofocusTrack3.Value
'    AcquisitionTrack4.Value = AutofocusTrack4.Value
'    AutofocusMaxSpeed.Value = True
'    AutofocusFastZline = False
'    AutofocusHRZ.Value = True
'
'    '''''''
'    ' No Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = False
'    ActivateTrack GlobalAcquisitionRecording, "Autofocus"
'    GlobalBackupRecording.SpecialScanMode = "FocusStep"
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'
'    If Not RunTestAutofocusButton(RecordingDoc, True, AFTest_Repetitions.Value, "AFTest4_Piezo_Stage_NoTrackZ") Then
'        Exit Function
'    End If
'
'    '''''''
'    ' Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = True
'    GlobalBackupRecording.SpecialScanMode = "FocusStep"
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'    If Not RunTestAutofocusButton(RecordingDoc, False, AFTest_Repetitions.Value, "AFTest4_Piezo_Stage_TrackZ") Then
'        Exit Function
'    End If
'    AFTest4Run = True
'End Function
'
'
''''''
'' AFTest5_Click()
'' Acquire reeatedly images with Fast-Z-Line
'''''
'Private Sub AFTest5_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest5Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest5Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AutofocusTrackZ.Value = False
'    AcquisitionTrack1.Value = False
'    AcquisitionTrack2.Value = False
'    AcquisitionTrack3.Value = False
'    AcquisitionTrack4.Value = False
'    AutofocusMaxSpeed.Value = True
'    AutofocusHRZ.Value = False
'    AutofocusFastZline.Value = True
'    AutofocusLineSize.Value = 256
'    If Not RunTestFastZline(RecordingDoc, 1, AFTest_Repetitions.Value, 1, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'    AutofocusLineSize.Value = 128
'    If Not RunTestFastZline(RecordingDoc, 2, AFTest_Repetitions.Value, 1, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'    AutofocusLineSize.Value = 64
'    If Not RunTestFastZline(RecordingDoc, 3, AFTest_Repetitions.Value, 1, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'    AutofocusLineSize.Value = 256
'    If Not RunTestFastZline(RecordingDoc, 4, AFTest_Repetitions.Value, 2, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'
'
'    AutofocusLineSize.Value = 128
'    If Not RunTestFastZline(RecordingDoc, 5, AFTest_Repetitions.Value, 2, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'    AutofocusLineSize.Value = 256
'    If Not RunTestFastZline(RecordingDoc, 6, AFTest_Repetitions.Value, 3, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'
'
'    AutofocusLineSize.Value = 128
'    If Not RunTestFastZline(RecordingDoc, 7, AFTest_Repetitions.Value, 3, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'
'    AutofocusLineSize.Value = 256
'    If Not RunTestFastZline(RecordingDoc, 8, AFTest_Repetitions.Value, 4, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'
'    AutofocusLineSize.Value = 128
'    If Not RunTestFastZline(RecordingDoc, 9, AFTest_Repetitions.Value, 4, "AFTest5_FastZlineTest", 5000) Then
'        Exit Function
'    End If
'    AFTest5Run = True
'End Function
'
'
'
''''''
'' AFTest6_Click()
'' Perform repeatealy Autofocus with piezo and frame acquisition with piezo at multiposition
'' Uses No Z-track and Z-track
'''''
'Private Sub AFTest6_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest6Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest6Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'    Log = True
'    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.Name) Then
'        MsgBox ("No piezo availabe! Cannot start tests.")
'        Exit Function
'    End If
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AcquisitionTrack1.Value = AutofocusTrack1.Value
'    AcquisitionTrack2.Value = AutofocusTrack2.Value
'    AcquisitionTrack3.Value = AutofocusTrack3.Value
'    AcquisitionTrack4.Value = AutofocusTrack4.Value
'    AutofocusMaxSpeed.Value = True
'    AutofocusFastZline = False
'    AutofocusHRZ.Value = True
'
'
'    '''''''
'    ' Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = True
'
'    MultipleLocationToggle.Value = True
'    GlobalRepetitionNumber = AFTest_Repetitions.Value
'    GlobalRepetitionTime.Value = 0
'    If Not StartSetting() Then
'        Exit Function
'    End If
'    GlobalAcquisitionRecording.SpecialScanMode = "ZScanner"
'
'    GlobalAcquisitionRecording.ScanMode = "Stack"                       'This is defining to acquire a Z stack of Z-Y images
'    GlobalAcquisitionRecording.SamplesPerLine = 32  'If doing frame autofocussing it uses the userdefined frame size
'    GlobalAcquisitionRecording.LinesPerFrame = 32
'    If AutofocusZStep.Value > 0 Then
'        GlobalAcquisitionRecording.FramesPerStack = Round(10 / AutofocusZStep.Value)
'        GlobalAcquisitionRecording.FrameSpacing = AutofocusZStep.Value
'    Else
'        GlobalAcquisitionRecording.FramesPerStack = 10
'        GlobalAcquisitionRecording.FrameSpacing = 10
'    End If
'    TextBoxFileName.Value = "Piezo"
'    'Set counters back to 1
'    RepetitionNumber = 1 ' first time point
'    StartAcquisition BleachingActivated 'This is the main function of the macro
'    AFTest6Run = True
'End Function
'
'
''''''
'' AFTest6_Click()
'' Perform repeatealy Autofocus with piezo and frame acquisition with piezo at multiposition
'' Uses No Z-track and Z-track
'''''
'Private Sub AFTest7_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    AFTest7Run
'    StopAcquisition
'End Sub
'
'Private Function AFTest7Run() As Boolean
'    Running = True
'    Dim RecordingDoc As DsRecordingDoc
'    Log = True
'    If Not Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.Name) Then
'        MsgBox ("No piezo availabe! Cannot start tests.")
'        Exit Function
'    End If
'    If GlobalDataBaseName = "" Then
'        MsgBox ("No outputfolder selected ! Cannot start tests.")
'        Exit Function
'    End If
'
'    'Setup a single recording doc
'    If RecordingDoc Is Nothing Then
'        Set RecordingDoc = Lsm5.NewScanWindow
'        While RecordingDoc.IsBusy
'            Sleep (100)
'            DoEvents
'        Wend
'    End If
'
'    If Not CheckDir(GlobalDataBaseName) Then
'        Exit Function
'    End If
'
'    AcquisitionTrack1.Value = AutofocusTrack1.Value
'    AcquisitionTrack2.Value = AutofocusTrack2.Value
'    AcquisitionTrack3.Value = AutofocusTrack3.Value
'    AcquisitionTrack4.Value = AutofocusTrack4.Value
'    AutofocusMaxSpeed.Value = True
'    AutofocusFastZline = True
'    AutofocusHRZ.Value = False
'
'
'    '''''''
'    ' Z-Tracking, Acquistion after Autofocus
'    '''''''
'    AutofocusTrackZ.Value = True
'
'    MultipleLocationToggle.Value = True
'    GlobalRepetitionNumber = AFTest_Repetitions.Value
'    GlobalRepetitionTime.Value = 0
'    If Not StartSetting() Then
'        Exit Function
'    End If
'    GlobalAcquisitionRecording.SpecialScanMode = "FocusStep"
'
'    GlobalAcquisitionRecording.ScanMode = "Stack"                       'This is defining to acquire a Z stack of Z-Y images
'    GlobalAcquisitionRecording.SamplesPerLine = 8  'If doing frame autofocussing it uses the userdefined frame size
'    GlobalAcquisitionRecording.LinesPerFrame = 8
'    If AutofocusZStep.Value > 0 Then
'        GlobalAcquisitionRecording.FramesPerStack = Round(20 / AutofocusZStep.Value)
'        GlobalAcquisitionRecording.FrameSpacing = AutofocusZStep.Value
'    Else
'        GlobalAcquisitionRecording.FramesPerStack = 10
'        GlobalAcquisitionRecording.FrameSpacing = 10
'    End If
'    TextBoxFileName.Value = "FastZline"
'    'Set counters back to 1
'    RepetitionNumber = 1 ' first time point
'    StartAcquisition BleachingActivated 'This is the main function of the macro
'    AFTest7Run = True
'End Function
'
'
'Private Sub AFTestAll_Click()
'    posTempZ = Lsm5.Hardware.CpFocus.Position
'    Running = True
'    If Not AFTest1Run Then
'        GoTo ScanStop
'    End If
'    If Not AFTest3Run Then
'        GoTo ScanStop
'    End If
'
'    If Not AFTest5Run Then
'        GoTo ScanStop
'    End If
'
'    If Lsm5.Hardware.CpHrz.Exist(Lsm5.Hardware.CpHrz.Name) Then
'        If Not AFTest2Run Then
'            GoTo ScanStop
'        End If
'        If Not AFTest4Run Then
'            GoTo ScanStop
'        End If
'        If Not AFTest6Run Then
'            GoTo ScanStop
'        End If
'        If Not AFTest7Run Then
'            GoTo ScanStop
'        End If
'    End If
'ScanStop:
'    ScanStop = True
'    StopAcquisition
'End Sub
'
'
'''''
''   RunTestAutofocusButton(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer) As Boolean
''   Using the actual setting for autofocusing function runs AutofocusButton. Save images and logfile on the GlobalDataBaseName directory
''       [RecordingDoc] - A recording where images are overwritten
''       [TestNr]       - Number of the test, this sets the name of the image files and logfiles.
''       [MaxTestRepeats] - Maximal number of tests for each repeat
'''''
'Private Function RunTestAutofocusButton(RecordingDoc As DsRecordingDoc, ResetPos As Boolean, MaxTestRepeats As Integer, Optional FileName As String = "AutofocusTest", Optional Pause As Integer = 1000) As Boolean
'
'    Dim FilePath As String
'    Dim TestRepeats As Integer
'    Dim Zold As Double
'    Dim pos As Double
'    TestRepeats = 1
'    LogFileName = GlobalDataBaseName & "\" & FileName & "_Log" & ".txt"
'
'    If Log Then
'        SafeOpenTextFile LogFileName, LogFile, FileSystem
'        LogFile.WriteLine "% Autofocus Test. Repeated AutofocusButton executions. "
'        LogFile.WriteLine "% MaxSpeed " & AutofocusMaxSpeed.Value & ", Zoom " & AutofocusZoom.Value & ", Piezo " & AutofocusHRZ.Value & ", AFTrackZ " & AutofocusTrackZ.Value & _
'        ", AFTrackXY " & AutofocusTrackXY.Value & ", FastZLine" & AutofocusFastZline.Value
'    End If
'    Zold = posTempZ
'    While TestRepeats < MaxTestRepeats + 1
'        DisplayProgress "Running Test " & FileName & ". Repeat " & TestRepeats & "/" & MaxTestRepeats & ".......", RGB(0, &HC0, 0)
'
'        FilePath = GlobalDataBaseName & "\" & FileName & "_" & TestRepeats
'        If Log Then
'            SafeOpenTextFile LogFileName, LogFile, FileSystem
'            LogFile.WriteLine " "
'
'            LogFile.WriteLine "% Save image in file " & FilePath & ".lsm"
'            LogFile.Close
'        End If
'        DoEvents
'        Sleep (Pause)
'        DoEvents
'
'        If ResetPos Then
'            posTempZ = Round(Zold + (1 - 2 * Rnd) * 10, PrecZ)
'        End If
'        Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
'
'        DisplayProgress "Autofocus SetupScanWindow", RGB(0, &HC0, 0)
'        If RecordingDoc Is Nothing Then
'            Set RecordingDoc = Lsm5.NewScanWindow
'            While RecordingDoc.IsBusy
'                Sleep (100)
'                DoEvents
'            Wend
'        End If
'        If Not AutofocusButtonRun(RecordingDoc, GlobalDataBaseName & "\AFimg_" & FileName & "_" & TestRepeats & ".lsm") Then
'            Exit Function
'        End If
'        'save file
'        If ActivateTrack(GlobalAcquisitionRecording, "Acquisition") Then
'            SaveDsRecordingDoc RecordingDoc, FilePath & ".lsm"
'        End If
'        TestRepeats = TestRepeats + 1
'        If ScanStop Then
'            Exit Function
'        End If
'    Wend
'    If Log Then
'        LogFile.Close
'    End If
'    RunTestAutofocusButton = True
'End Function
'
'''''
''   RunTestFastZline(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer, pixelDwell As Double, FrameSize As Integer, pause As Integer) As Boolean
''   Using the actual setting for autofocusing function runs AutofocusButton. Save images and logfile on the GlobalDataBaseName directory
''       [RecordingDoc] - A recording where images are overwritten
''       [TestNr]       - Number of the test, this sets the name of the image files and logfiles.
''       [MaxTestRepeats] - Maximal number of tests for each repeat
'''''
'Private Function RunTestFastZline(RecordingDoc As DsRecordingDoc, TestNr As Integer, MaxTestRepeats As Integer, Optional pixelDwellfactor As Double = 1, Optional FileName As String = "AutofocusTest", Optional Pause As Integer = 5000) As Boolean
'
'    Dim FilePath As String
'    Dim TestRepeats As Integer
'    Dim SuccessRecenter As Boolean
'    Dim time As Double
'    Dim pos As Double ' position temp variable
'    TestRepeats = 1
'    LogFileName = GlobalDataBaseName & "\" & FileName & TestNr & ".txt"
'
'    If Log Then
'        SafeOpenTextFile LogFileName, LogFile, FileSystem
'        LogFile.WriteLine "% FastZlineTest " & TestNr & ". Repeated fast Zline executions. PixelDwellfactor: " & pixelDwellfactor & ", LineSize: " & AutofocusLineSize.Value & ", pause: " & Pause
'        LogFile.WriteLine "% MaxSpeed " & AutofocusMaxSpeed.Value & ", Zoom " & AutofocusZoom.Value & ", Piezo " & AutofocusHRZ.Value & ", AFTrackZ " & AutofocusTrackZ.Value & _
'        ", AFTrackXY " & AutofocusTrackXY.Value
'    End If
'
'    While TestRepeats < MaxTestRepeats + 1
'        DisplayProgress "Running Test " & TestNr & ". Repeat " & TestRepeats & "/" & MaxTestRepeats & ".......", RGB(0, &HC0, 0)
'        FilePath = GlobalDataBaseName & "\" & FileName & TestNr & "_" & TestRepeats
'        If Log Then
'            SafeOpenTextFile LogFileName, LogFile, FileSystem
'            LogFile.WriteLine " "
'            LogFile.WriteLine "% Save image in file " & FilePath & ".lsm"
'            LogFile.Close
'        End If
'        DoEvents
'        Sleep (Pause)
'        DoEvents
'        If Not AutofocusForm.ActivateTrack(GlobalAutoFocusRecording, "Autofocus") Then
'            MsgBox "No track selected for Autofocus! Cannot Autofocus!"
'            Exit Function
'        End If
'        time = Timer
'        Recenter_pre posTempZ, SuccessRecenter, ZENv
'
'        Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime = Lsm5.DsRecording.TrackObjectByMultiplexOrder(0, 1).SampleObservationTime * pixelDwellfactor
'
'        Sleep (Pause)
'        DoEvents
'        If Log Then
'            SafeOpenTextFile LogFileName, LogFile, FileSystem
'            time = Timer - time
'            'pos = Lsm5.Hardware.CpFocus.Position
'            LogFile.WriteLine ("% AutofocusButton: center and wait 1st  Z = " & posTempZ & ", Time required " & time & ", success Recenter " & SuccessRecenter)
''            Sleep (100)
''            If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
''                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ  " & posTempZ & "; obtained Central slide " & pos & "; position " & pos)
''            Else
''                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ  " & posTempZ & "; obtained Central slide " & _
''                Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + pos & "; position " & pos)
''            End If
'            LogFile.Close
'        End If
'
'        If Not ScanToImage(RecordingDoc) Then
'            Exit Function
'        End If
'        time = Timer
'        Recenter_post posTempZ, SuccessRecenter, ZENv
'        DoEvents
'        If Log Then
'            SafeOpenTextFile LogFileName, LogFile, FileSystem
'            time = Timer - time
'            pos = Lsm5.Hardware.CpFocus.Position
'            LogFile.WriteLine ("% AutofocusButton: recenter 1st  Z = " & posTempZ & ", Time required " & time & ", waiting repeats (max 9) " & Round(time / 0.4))
'            If (Lsm5.DsRecording.ScanMode <> "Stack" And Lsm5.DsRecording.ScanMode <> "ZScan") Or AutofocusHRZ Then
'                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ (after img) " & posTempZ & "; obtained Central slide " & Lsm5.Hardware.CpFocus.Position & "; position " & Lsm5.Hardware.CpFocus.Position)
'            Else
'                LogFile.WriteLine ("% AutofocusButton: Target Central slide AQ (after img) " & posTempZ & "; obtained Central slide " & _
'                Lsm5.DsRecording.FrameSpacing * (Lsm5.DsRecording.FramesPerStack - 1) / 2 - Lsm5.DsRecording.Sample0Z + Lsm5.Hardware.CpFocus.Position & "; position " & Lsm5.Hardware.CpFocus.Position)
'            End If
'            LogFile.Close
'        End If
'        SaveDsRecordingDoc RecordingDoc, FilePath & ".lsm"
'        TestRepeats = TestRepeats + 1
'        If ScanStop Then
'            Exit Function
'        End If
'    Wend
'    If Log Then
'        LogFile.Close
'    End If
'    RunTestFastZline = True
'End Function
'
'
