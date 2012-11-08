Attribute VB_Name = "unusedCode"
'''''''''''''''''''''''''''''''''''''''''ExcelXYZstoring'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Excel.Application.Visible = True                               'The Excel stuff is to store the XYZ position of the cells at each time point
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
'''''''''''''''''''''''''''''''''''''''''End ExcelXYZstoring'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''ExcelXYZstoring II'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
'''''''''''''''''''''''''''''''''''''''''End ExcelXYZstoring II'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





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


' Autofocusroutines that worked for Meta
'Public Sub Autofocus_MoveAquisition(Zoffset As Double)
'Dim NoZStack As Boolean
'Const ZBacklash = -50
'Dim ZFocus As Double
'Dim Zbefore As Double
'Dim x As Double
'Dim y As Double
'
'
'
'    RestoreAquisitionParameters
'
'  '  Set GlobalBackupRecording = Nothing
'    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
'    DoEvents
'    AutofocusForm.ActivateAcquisitionTrack
'    If Lsm5.DsRecording.ScanMode = "ZScan" Or Lsm5.DsRecording.ScanMode = "Stack" Then  'Looks if a Z-Stack is going to be acquired
'        NoZStack = False
'    Else
'        NoZStack = True
'    End If
'
'    'Moving to the correct position in Z
'                                          'If using HRZ for autofocusing and there is no Zstack for image acquisition
'     '   ZFocus = Lsm5.Hardware.CpHrz.Position + ZShift - Zoffset
'
'     'Defines the new focus position as the actual position plus the shift and goes back to the object position (that's why you need the offset)
'
'    If HRZ Then
'     ZFocus = Lsm5.Hardware.CpFocus.Position - Zoffset - ZShift
'       Lsm5.Hardware.CpFocus.Position = ZFocus + ZBacklash     'Moves down -50uM (ZBacklash) with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'        Lsm5.Hardware.CpFocus.Position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'    Else
'    ZFocus = Lsm5.Hardware.CpFocus.Position - Zoffset + ZShift
'       Lsm5.Hardware.CpFocus.Position = ZFocus + ZBacklash     'Moves down -50uM (ZBacklash) with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'        Lsm5.Hardware.CpFocus.Position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'    End If
'''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the HRZ can acces the position, and also to wait it is done...
'        Sleep (100)
'        DoEvents
'
'
'
'    'Moving to the correct position in X and Y
'
'    If FrameAutofocussing Then
'        x = Lsm5.Hardware.CpStages.PositionX - XShift  'the fact that it is "-" in this line and "+" in the next line  probably has to do with where the XY of the origin is set (top right corner and not botom left, I think)
'        y = Lsm5.Hardware.CpStages.PositionY - YShift
'        success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
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
'
'Private Sub MovetoCorrectZPosition(Zoffset As Double)
'Const ZBacklash = -50
'Dim ZFocus As Double
'Dim Zbefore As Double
'Dim x As Double
'Dim y As Double
'     ZFocus = Lsm5.Hardware.CpFocus.Position + Zoffset + ZShift
'       Lsm5.Hardware.CpFocus.Position = ZFocus + ZBacklash    'Moves down -50uM (ZBacklash) with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'        Lsm5.Hardware.CpFocus.Position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the HRZ can acces the position, and also to wait it is done...
'        Sleep (100)
'        DoEvents
'End Sub
'
'
'
'
'Public Sub Autofocus_MoveAquisition_HRZ(Zoffset As Double)
'Dim NoZStack As Boolean
'Const ZBacklash = -50
'Dim ZFocus As Double
'Dim Zbefore As Double
'Dim x As Double
'Dim y As Double
'
'
'
'  RestoreAquisitionParameters
'
'    Set GlobalBackupRecording = Nothing
'    Lsm5Vba.Application.ThrowEvent eRootReuse, 0
'    DoEvents
'    AutofocusForm.ActivateAcquisitionTrack
'    If Lsm5.DsRecording.ScanMode = "ZScan" Or Lsm5.DsRecording.ScanMode = "Stack" Then  'Looks if a Z-Stack is going to be acquired
'        NoZStack = False
'    Else
'        NoZStack = True
'    End If
'
'    'Moving to the correct position in Z
'    If HRZ And NoZStack Then                                            'If using HRZ for autofocusing and there is no Zstack for image acquisition
'     Lsm5.Hardware.CpHrz.Stepsize = 0.2
'
'      Lsm5Vba.Application.ThrowEvent eRootReuse, 0
'    DoEvents
'     '   ZFocus = Lsm5.Hardware.CpHrz.Position + ZShift - Zoffset
'
'     'Defines the new focus position as the actual position plus the shift and goes back to the object position (that's why you need the offset)
'
'     ZFocus = Lsm5.Hardware.CpHrz.Position + Zoffset + ZShift
'
'        Lsm5.Hardware.CpHrz.Position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'''''' If I want to do it properly, I should add a lot of controls here, to wait to be sure the HRZ can acces the position, and also to wait it is done...
'
'        DoEvents
'
'    Else                                        'either there is a Z stack for image acquisition or we're using the focuswheel for autofocussing
'        If HRZ Then                             ' Now I'm not sure with the signs and... I some point I just tried random combinations...
'            ZFocus = Lsm5.Hardware.CpHrz.Position - Zoffset - ZShift '         'ZBefore corresponds to the position where the focuswheel was before doing anything. Zshift is the calculated shift
'        Else                                    'If the HRZ is not calibrated the Z shift might be wrong
'            ZFocus = Zbefore + ZShift
'        End If
'
'        Lsm5.Hardware.CpHrz.Position = ZFocus                     'Moves up to the focus position with the focus wheel
'        Do While Lsm5.ExternalCpObject.pHardwareObjects.pFocus.pItem(0).bIsBusy
'            Sleep (20)
'            DoEvents
'        Loop
'    End If
'
'    'Moving to the correct position in X and Y
'
'    If FrameAutofocussing Then
'        x = Lsm5.Hardware.CpStages.PositionX - XShift  'the fact that it is "-" in this line and "+" in the next line  probably has to do with where the XY of the origin is set (top right corner and not botom left, I think)
'        y = Lsm5.Hardware.CpStages.PositionY - YShift
'        success = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).MoveToPosition(x, y)
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
