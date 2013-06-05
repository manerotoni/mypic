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


Public Function ComputeCenterAndAxis(dX As Double, dY As Double)

    Dim i, j, iFrame, channel, ni, bitDepth As Long
    Dim nj As Long
    
    Dim ic, jc, di, dj, PixelSize As Double
    Dim tot As Double
    
    Dim th As Double
    th = 20
    
    
    'Dim ColMax As Integer
    'Dim iRow As Integer
    'Dim nRow As Integer
    'Dim iFrame As Integer
    'Dim gvRow As Variant  ' gv = gray value
    'Dim iCol As Long
    'Dim nCol As Long
    'Dim bitDepth As Long
    'Dim channel As Integer
    'Dim gvMax As Double
    'Dim gvMaxBitRange As Double
    'Dim nSaturatedPixels As Long
    'Dim maxGV_nSat(2) As Double
    
    
    'DisplayProgress "Measuring Exposure...", RGB(0, &HC0, 0)
  
    'ColMax = Lsm5.DsRecordingActiveDocObject.Recording.RtRegionWidth '/ Lsm5.DsRecordingActiveDocObject.Recording.RtBinning
    
    'nRow = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
    'MsgBox "nRow = " + CStr(nRow)
    
'        ElseIf SystemName = "LSM" Then
'            ColMax = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
'            LineMax = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
'        Else
'            MsgBox "The System is not LIVE or LSM! SystemName: " + SystemName
''            Exit Sub
 '       End If
 '   End If
    
    'Initiallize tables to store projected (integrated) pixels values in the 3 dimensions
    'ReDim Intline(nLines) As Long
    
    'iFrame = 0
    'gvMax = -1
        
    'iRow = 0
    'channel = 0
    'bitDepth = 0 ' leaves the internal bit depth
    'gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
    
    
    
    ni = Lsm5.DsRecordingActiveDocObject.Recording.LinesPerFrame
    'nCol = 0
    nj = Lsm5.DsRecordingActiveDocObject.Recording.SamplesPerLine
    
    'Dim image(,) As Variant
    
    'Dim replyCounts(,,) As Short = New Short(2, 1, 2) {}
    
    Dim srcline As Variant
    
    Dim image() As Long
    ReDim image(ni, nj)
    
    
    'Dim x(1 To ni, 1 To 4) As Variant

    'MsgBox "ni = " + CStr(ni) + " nj = " + CStr(nj)
    
   ' image = GetSubRegion(channel, xs, ys, zs, ts
    
    
    'Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
        
    PixelSize = Lsm5.DsRecordingActiveDocObject.Recording.SampleSpacing * 1000000
        
        
    ' get the image  (put into a subprocedure)
    iFrame = 0
    channel = 0
    bitDepth = 0 ' leaves the internal bit depth
    For i = 0 To ni - 1
        srcline = Lsm5.DsRecordingActiveDocObject.ScanLine(channel, 0, iFrame, i, nj, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
        For j = 0 To nj - 1
            image(i, j) = srcline(j)
        Next j
    Next i
    'MsgBox "im = " + CStr(image(100, 100))
        
    ' computer center of mass
    ic = 0
    jc = 0
    tot = 0
    For i = 0 To ni - 1
        For j = 0 To nj - 1
            If (image(i, j) > th) Then
                ic = ic + image(i, j) * i
                jc = jc + image(i, j) * j
                tot = tot + image(i, j)
            End If
        Next j
    Next i
    
    ic = ic / tot
    jc = jc / tot
    'MsgBox "ic = " + CStr(ic) + " jc = " + CStr(jc) + " tot = " + CStr(tot)
    
    dX = (ic - ni / 2) * PixelSize
    dY = (jc - nj / 2) * PixelSize
    
    ' compute displacement vector
    di = 0
    dj = 0
    
    For i = 0 To ni - 1
        For j = 0 To nj - 1
            If (image(i, j) > th) Then
                di = di + image(i, j) * (i - ic) * Sgn(i - ic)
                dj = dj + image(i, j) * (j - jc) * Sgn(i - ic)
            End If
        Next j
    Next i
    
    di = di / tot
    dj = dj / tot
    'MsgBox "di = " + CStr(di) + " dj = " + CStr(dj) + " tot = " + CStr(tot)
        
        
    'PixelSize
        
        
        
    '    For iCol = 0 To nCol - 1            'Now I'm scanning all the pixels in the line
            
     '       If (gvRow(iCol) > gvMax) Then
      '          gvMax = gvRow(iCol)
       '     End If

    
    
    'iFrame = 0
    'gvMax = -1
    'iRow = 0
    'Channel = 0
    'bitDepth = 0 ' leaves the internal bit depth
    'gvRow = Lsm5.DsRecordingActiveDocObject.ScanLine(Channel, 0, iFrame, iRow, nCol, bitDepth) 'this is the lsm function how to read pixel values. It basically reads all the values in one X line. scrline is a variant but acts as an array with all those values stored
    'MsgBox "nCol = " + CStr(nCol)
    'MsgBox "bytes per pixel = " + CStr(bitDepth)

    ' todo: is there another function to find this out??
    'If (bitDepth = 1) Then
    '    gvMaxBitRange = 255
    'ElseIf (bitDepth = 2) Then
    '    gvMaxBitRange = 65536
    'End If
    
    'nSaturatedPixels = 0
 
End Function
