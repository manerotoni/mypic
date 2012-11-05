Attribute VB_Name = "Stage_Grid"

Public Const PubGridPathData = "c:\AIM\macros\datafiles\"
Public Const PubGridPath = "c:\AIM\macros\"
Public Const GlobalMaximumPositions = 800
Public GlobalXStep As Double
Public GlobalYStep As Double
Public GlobalXGrid As Long
Public GlobalYGrid As Long
Public GlobalOrderChanged As Boolean        'checks if values are changed without creating new locations_
                                            '    then  the lost order is not showed, because it not actualized
Public GlobalGridImage As DsRecordingDoc
Public GlobalZGridImage As DsRecordingDoc
Public GlobalGridX1 As Long
Public GlobalGridY1 As Long
Public GlobalGridStageX1 As Double
Public GlobalGridStageY1 As Double
Public GlobalGridStageZ1 As Double
Public GlobalGridX2 As Long
Public GlobalGridY2 As Long
Public GlobalGridStageX2 As Double
Public GlobalGridStageY2 As Double
Public GlobalGridStageZ2 As Double
Public GlobalReferencePoints As Long
'Public GlobalKeepSteps As Boolean
Public GlobalMeander As Boolean
Public GlobalPositionsRecalled As Long
Public GlobalPositionsStage As Long
Public GlobalCurrentPosition As Long
Public GlobalXpos() As Double
Public GlobalYpos() As Double
Public GlobalZpos() As Double
Public GlobalLocationsName() As String
Public GlobalLocationsNameOld() As String
Public GlobalZposOld() As Double
Public GlobalXposOld() As Double
Public GlobalYposOld() As Double
Public GlobalRelativeZ() As Double
Public dsDoc As DsRecordingDoc
Public Stage As CpStages
Public GlobalProgressString As String
Public GlobalColor As Long
Public GettingZmap As Boolean
Public GlobalZmapAquired As Boolean
Public idpos As Long
Public PubAbort As Boolean
Public Grid As Boolean

Public XR As Long
Public YR As Long
Public x As Long
Public y As Long
    



Public GlobalDeActivatedLocations() As Boolean
Public GlobalLocationsOrder() As Long
Public GlobalLocationsOrderOld() As Long



Public Const ZBacklash = -50


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


Public Sub RedrawGrid(GridImage As DsRecordingDoc)
Dim BitsPerSample As Long
Dim bpp As Long
Dim ImgName As String
Dim LsmMath As New LsmVectorMath
Dim SpareArrayRed() As Single
Dim SpareArrayBlue() As Single
Dim SpareArrayGreen() As Single
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim xIndx As Long
Dim yIndx As Long
Dim Start As Long
Dim ix As Long
Dim iy As Long
Dim PlaneSize As Long

    If GridImage Is Nothing Then Exit Sub
    XPixels = GridImage.GetDimensionX
    YPixels = GridImage.GetDimensionY
    
    XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
    YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
    
    PlaneSize = LsmMath.ImagePlaneSizeXY(GridImage)
    ReDim SpareArrayRed(PlaneSize)
    ReDim SpareArrayBlue(PlaneSize)
    ReDim SpareArrayGreen(PlaneSize)
    For yIndx = 1 To GlobalYGrid
        For xIndx = 1 To GlobalXGrid
            Start = YGroup * XPixels + (3 * YGroup * XPixels) * (yIndx - 1) + XGroup + _
            3 * XGroup * (xIndx - 1) - 1
            For iy = 1 To 2 * YGroup
                For ix = 1 To 2 * XGroup
                    If Not GlobalDeActivatedLocations(xIndx, yIndx) Then
                        SpareArrayGreen(Start + ix + (iy - 1) * XPixels) = 0
                        SpareArrayBlue(Start + ix + (iy - 1) * XPixels) = 4000
                        SpareArrayRed(Start + ix + (iy - 1) * XPixels) = 0
                    Else
                        SpareArrayRed(Start + ix + (iy - 1) * XPixels) = 1000
                        SpareArrayGreen(Start + ix + (iy - 1) * XPixels) = 0
                        SpareArrayBlue(Start + ix + (iy - 1) * XPixels) = 0
                    End If
                Next ix
            Next iy
        Next xIndx
    Next yIndx
    LsmMath.WriteImagePlaneXY GridImage, 0, 0, 0, PlaneSize, SpareArrayRed(0)
    LsmMath.WriteImagePlaneXY GridImage, 1, 0, 0, PlaneSize, SpareArrayGreen(0)
    LsmMath.WriteImagePlaneXY GridImage, 2, 0, 0, PlaneSize, SpareArrayBlue(0)
    DoEvents

End Sub
Public Sub RedrawZGrid(GridImage As DsRecordingDoc)
Dim BitsPerSample As Long
Dim bpp As Long
Dim ImgName As String
Dim LsmMath As New LsmVectorMath
Dim SpareArrayBlue() As Single
Dim SpareArrayRed() As Single
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim xIndx As Long
Dim yIndx As Long
Dim Start As Long
Dim ix As Long
Dim iy As Long
Dim PlaneSize As Long
Dim MinZValue As Double
Dim MaxZValue As Double
Dim ColorStep As Double
Dim ZGridRange As Double
Dim ZGridImage As DsRecording

    MinZValue = 10000           'I choose any high number that canot be reached
    
    MaxZValue = -10000
    For idpos = 1 To GlobalPositionsStage
        If GlobalZpos(idpos) < MinZValue Then
            MinZValue = GlobalZpos(idpos)
        End If
        If GlobalZpos(idpos) > MaxZValue Then
            MaxZValue = GlobalZpos(idpos)
        End If
    Next idpos
ZGridRange = MaxZValue - MinZValue

If ZGridRange <= 0 Then
MsgBox "ZGridRange <= 0!" + vbCrLf + "MaxValue =" + CStr(MaxZValue) + vbCrLf + ";MinValue = " + CStr(MinZValue)
Exit Sub
End If

ColorStep = 4000 / ZGridRange



 '   If ZGridImage Is Nothing Then Exit Sub
    XPixels = GridImage.GetDimensionX
    YPixels = GridImage.GetDimensionY
    
    XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
    YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
    
    PlaneSize = LsmMath.ImagePlaneSizeXY(GridImage)
    ReDim SpareArrayBlue(PlaneSize)
    ReDim SpareArrayRed(PlaneSize)
    For yIndx = 1 To GlobalYGrid
        For xIndx = 1 To GlobalXGrid
         idpos = GlobalLocationsOrder(xIndx, yIndx)
            Start = YGroup * XPixels + (3 * YGroup * XPixels) * (yIndx - 1) + XGroup + _
            3 * XGroup * (xIndx - 1) - 1
            
                For iy = 1 To 2 * YGroup
                    For ix = 1 To 2 * XGroup
                        If Not GlobalDeActivatedLocations(xIndx, yIndx) Then
                          SpareArrayBlue(Start + ix + (iy - 1) * XPixels) = (GlobalZpos(idpos) - MinZValue) * ColorStep + 96
                          SpareArrayRed(Start + ix + (iy - 1) * XPixels) = 0
                        Else
                            SpareArrayRed(Start + ix + (iy - 1) * XPixels) = 500
                            SpareArrayBlue(Start + ix + (iy - 1) * XPixels) = 0
                        End If
                    Next ix
                Next iy
            
         Next xIndx
    Next yIndx
    LsmMath.WriteImagePlaneXY GridImage, 2, 0, 0, PlaneSize, SpareArrayBlue(0)
    LsmMath.WriteImagePlaneXY GridImage, 0, 0, 0, PlaneSize, SpareArrayRed(0)
    DoEvents
    DisplayProgress "highest Z: " + CStr(MaxZValue) + vbCrLf + "lowest Z: " + CStr(MinZValue), RGB(0, &HC0, 0)
End Sub
Public Sub DrawCrossGrid(xIndx As Long, yIndx As Long)
Dim XPixels As Long
Dim YPixels As Long
Dim XGroup As Long
Dim YGroup As Long
Dim ix As Long
Dim iy As Long
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

Public Sub ConvertToStagePositionXY(XP As Double, YP As Double, Xnew As Double, Ynew As Double)

    Dim bExchangeXY As Boolean
    Dim bMirrorX As Boolean
    Dim bMirrorY As Boolean
    Dim dExchange As Double
    Dim x As Double
    Dim y As Double
    
    x = XP
    y = YP
    On Error GoTo oldversion
    
    CoordinateConversion bExchangeXY, bMirrorX, bMirrorY
    
    If bMirrorX Then
        x = -x
    End If
'    If Not bMirrorY Then
    If bMirrorY Then
        y = -y
    End If
        
    If bExchangeXY Then
        Ynew = x
        Xnew = y
    Else
        Xnew = x
        Ynew = y
    End If
            
    Exit Sub
oldversion:

    Xnew = X11 * x + X21 * y
    Ynew = X12 * x + X22 * y
    
nostage:

End Sub

Public Sub DoMouseEventsMulti(ByVal EventNr As Long, ByVal Param As Variant)
    Dim x1 As Double
    Dim Y1 As Double
    Dim X2 As Double
    Dim Y2 As Double
    Dim X3 As Double
    Dim Y3 As Double
    
    Dim Xtemp As Double
    Dim Ytemp As Double
    Dim xtemp1 As Double
    Dim ytemp1 As Double
'
'    Dim x As Long
'    Dim y As Long
    Dim z As Long
    Dim t As Long
    Dim c As Long
    
'    Dim XR As Long
'    Dim YR As Long

    Dim eps As Long
    eps = 5
    Dim cond1 As Boolean
    Dim cond2 As Boolean
    Dim cond3 As Boolean
    Dim Angle As Double
    
    Dim Positions As Long
    Dim XPos() As Double
    Dim YPos() As Double
    Dim ZPos() As Double
    Dim ZeroChanged As Boolean
    Dim SetZeroMarked As Boolean
    Dim Row As Long
    Dim XPixel As Long
    Dim YPixel As Long
    Dim MyString As String
    Dim Count As Long
    Dim Style, Title, Response
    Dim xIndx As Long
    Dim yIndx As Long
    Dim DiffX As Double
    Dim DiffY As Double
    Dim DiffXcorr As Double
    Dim DiffYcorr As Double
    Dim Xcorr As Double
    Dim Ycorr As Double
    
    Dim DiffXGrid As Long
    Dim DiffYGrid As Long
  
'    Set Stage = Lsm5.Hardware.CpStages
On Error GoTo marke1

Select Case EventNr
    Case eEventDsScanStopping
        DisplayProgress "Stopping", RGB(&HC0, &HC0, 0)
        flgEvent = 8
    Case eEventDsScanStopped
        flgEvent = 7
    Case Else
        If (EventNr = Lsm5Vba.eImageWindowNoButtonMouseMoveEvent) Then
            If (dsDoc Is Nothing) Then
                
                    If Not GlobalGridImage Is Nothing Then
                     
                    
                        If GlobalGridImage.GetCurrentMousePosition(c, t, z, y, x) <> 0 Then
                        AutofocusForm.DisplayGridSelection x, y, xIndx, yIndx
                        End If
                    End If
              
                If Not GlobalZGridImage Is Nothing Then
                    
                        If GlobalZGridImage.GetCurrentMousePosition(c, t, z, y, x) <> 0 Then
                        AutofocusForm.DisplayGridSelection x, y, xIndx, yIndx
                        End If
                    End If
            Else
'                If dsDoc.GetCurrentMousePosition(C, T, Z, Y, X) <> 0 Then
'                    CenterForm.Label2 = "dx=" + _
'                    Strings.Format(CDbl(X - xCross) * dsDoc.VoxelSizeX * 10 ^ 6, "0.00") + _
'                    " " + Strings.Chr(181) + "m" + Strings.Chr(10) + _
'                    "dy=" + Strings.Format(CDbl(Y - yCross) * dsDoc.VoxelSizeY * 10 ^ 6, "0.00") + _
'                    " " + Strings.Chr(181) + "m"
'                End If
            End If
        ElseIf (EventNr = DS45.eImageWindowRightButtonUpEvent) Then
            If Not GlobalGridImage Is Nothing Then
                If GlobalGridImage.GetCurrentMousePosition(c, t, z, y, x) <> 0 Then
                    AutofocusForm.DisplayGridSelection x, y, xIndx, yIndx
                  
                
                    Style = vbOKOnly + VbQuestion + VbDefaultButton2 ' Define buttons.
                    Title = "Selecting Reference Locations"  ' Define title.

                    Style = VbYesNo + VbQuestion + VbDefaultButton2 ' Define buttons.

                    If GlobalReferencePoints = 2 Then
                        MyString = "Is Stage Positioned at the Grid Location:" + vbCrLf + "Column=" + CStr(xIndx) + _
                        vbCrLf + "Row=" + CStr(yIndx) + "?" + vbCrLf + "Do You Want to Use it as a Reference?" + vbCrLf + _
                        "Previous Reference Points will be Deleted!"
                    ElseIf GlobalReferencePoints = 1 Then
                        MyString = "Is Stage Positioned at the Grid Location:" + vbCrLf + "Column=" + CStr(xIndx) + _
                        vbCrLf + "Row=" + CStr(yIndx) + "?" + vbCrLf + "Do You Want to Use it as a Second Reference?"
                    ElseIf GlobalReferencePoints = 0 Then
                        MyString = "Is Stage Positioned at the Grid Location:" + vbCrLf + "Column=" + CStr(xIndx) + _
                        vbCrLf + "Row=" + CStr(yIndx) + "?" + vbCrLf + "Do You Want to Use it as a First Reference?"
                    End If
                    Response = MsgBox(MyString, Style, Title)
                    If Response = vbYes Then    ' User chose Yes.
                        If GlobalReferencePoints = 2 Then
                            GlobalReferencePoints = 1
                            GlobalGridX1 = xIndx
                            GlobalGridY1 = yIndx
                            ReadLoc GlobalGridStageX1, GlobalGridStageY1
                            GlobalGridStageZ1 = Lsm5.Hardware.CpFocus.Position
                            GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
                            AutofocusForm.DrawCrossGrid GlobalGridX1, GlobalGridY1
                        ElseIf GlobalReferencePoints = 1 Then
                            GlobalGridX2 = xIndx
                            GlobalGridY2 = yIndx
                            ReadLoc GlobalGridStageX2, GlobalGridStageY2
                            GlobalGridStageZ2 = Lsm5.Hardware.CpFocus.Position
                            DiffX = GlobalGridStageX2 - GlobalGridStageX1
                            DiffY = GlobalGridStageY2 - GlobalGridStageY1
'                            If Abs(DiffX) < 50 Or Abs(DiffY) < 50 Then
'                                MsgBox "Selected Reference Locations are too Close!" + vbCrLf + _
'                                "Please Select New Reference Point!"
'                            Else
                                DiffXGrid = GlobalGridX2 - GlobalGridX1
                                DiffYGrid = GlobalGridY2 - GlobalGridY1
                                If Abs(DiffXGrid) > 0 And Abs(DiffYGrid) > 0 Then
'                                    If GlobalKeepSteps Then
'                                        DiffXcorr = GlobalXStep * DiffXGrid
'                                        DiffYcorr = GlobalYStep * DiffYGrid
'                                        If (Abs(DiffXcorr - DiffX) >= Abs(GlobalXStep)) Or (Abs(DiffYcorr - DiffY) >= Abs(GlobalYStep)) Then
'                                            MsgBox "The Difference Between Marked and Corrected Reference Points is Greater then the Grid Step!" + vbCrLf + _
'                                            "Please Check if the Grid Steps are Correct!"
'                                        Else
'                                            Xcorr = (DiffXcorr - DiffX) / 2
'                                            Ycorr = (DiffYcorr - DiffY) / 2
'                                            GlobalGridStageX1 = GlobalGridStageX1 - Xcorr
'                                            GlobalGridStageX2 = GlobalGridStageX2 + Xcorr
'                                            GlobalGridStageY1 = GlobalGridStageY1 - Ycorr
'                                            GlobalGridStageY2 = GlobalGridStageY2 + Ycorr
'                                            GlobalReferencePoints = 2
'                                            GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
'                                            AutofocusForm.DrawCrossGrid GlobalGridX1, GlobalGridY1
'                                            AutofocusForm.DrawCrossGrid GlobalGridX2, GlobalGridY2
'
'                                        End If
'                                    Else
                                        GlobalXStep = DiffX / DiffXGrid
                                        GlobalYStep = DiffY / DiffYGrid
                                        AutofocusForm.TextBoxXStep.Value = GlobalXStep
                                        AutofocusForm.TextBoxYStep.Value = GlobalYStep
                                        GlobalReferencePoints = 2
                                        GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
                                        AutofocusForm.DrawCrossGrid GlobalGridX1, GlobalGridY1
                                        AutofocusForm.DrawCrossGrid GlobalGridX2, GlobalGridY2

'                                    End If
'                                End If
                            End If

                        ElseIf GlobalReferencePoints = 0 Then
                            GlobalReferencePoints = 1
                            GlobalGridX1 = xIndx
                            GlobalGridY1 = yIndx
                            ReadLoc GlobalGridStageX1, GlobalGridStageY1
                            GlobalGridStageZ1 = Lsm5.Hardware.CpFocus.Position
                            GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
                            AutofocusForm.DrawCrossGrid GlobalGridX1, GlobalGridY1

                        End If
                    End If
                
                End If
            End If
        ElseIf (EventNr = Lsm5Vba.eImageWindowLeftButtonDownEvent) Then
            If Not GlobalGridImage Is Nothing Then
                If GlobalGridImage.GetCurrentMousePosition(c, t, z, y, x) <> 0 Then
 '                   GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
'                    GlobalGridImage.VectorOverlay.AddSimpleDrawingElement Lsm5Vba.eDrawingModeRectangle, X, Y, X, Y
'                    GlobalGridImage.RedrawImage
                    AutofocusForm.DrawRectangleGrid x, y, x, y
                End If
            End If
       
        ElseIf (EventNr = Lsm5Vba.eImageWindowLButtonMouseMoveEvent) Then
               If Not GlobalGridImage Is Nothing Then
                If GlobalGridImage.GetCurrentMousePosition(c, t, z, y, x) <> 0 Then
                    Count = GlobalGridImage.VectorOverlay.GetNumberDrawingElements
                    If GlobalGridImage.VectorOverlay.GetKnot(Count - 1, 0, XR, YR) Then
                        GlobalGridImage.VectorOverlay.RemoveDrawingElement Count - 1
                        AutofocusForm.DrawRectangleGrid XR, YR, x, y
                        AutofocusForm.DisplayGridSelection x, y, xIndx, yIndx
                    End If
                End If
            End If
       
        ElseIf (EventNr = Lsm5Vba.eImageWindowLeftButtonUpEvent) Then
        If (XR = x And YR = y) Then
        GridOnOff x, y, XR, YR
        GoTo Continue3
        Else
        GoTo Continue1
        End If
 
Continue1:
            If Not GlobalGridImage Is Nothing Then
                If GlobalGridImage.GetCurrentMousePosition(c, t, z, y, x) <> 0 Then
                    Count = GlobalGridImage.VectorOverlay.GetNumberDrawingElements
                    If GlobalGridImage.VectorOverlay.GetKnot(Count - 1, 0, x, y) Then
                        If GlobalGridImage.VectorOverlay.GetKnot(Count - 1, 1, XR, YR) Then
                            If (XR <> x) Or (YR <> y) Then
                            SelectLocs.Show
'                                Style = vbOKOnly + VbQuestion + VbDefaultButton2 ' Define buttons.
'                                Title = "Select/Deselect Locations"  ' Define title.
'
'                '                GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLButtonMouseMoveEvent, 0
'                '                GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLeftButtonDownEvent, 0
'                '                GlobalGridImage.EnableImageWindowEvent Lsm5Vba.eImageWindowLeftButtonUpEvent, 0
'
'                                Style = VbYesNo + VbQuestion + VbDefaultButton2 ' Define buttons.
'                                MyString = "Select Locations - Click YES;" + vbCrLf + "Deselect Locations - Click NO"
'                                Response = MsgBox(MyString, Style, Title)
'                                If Response = vbYes Then    ' User chose Yes.
'                                    GridSelection x, y, XR, YR, True
'                                Else
'                                    GridSelection x, y, XR, YR, False
'                                End If
Continue3:
                     If Not GlobalGridImage Is Nothing Then
                                GlobalGridImage.VectorOverlay.RemoveAllDrawingElements
                                If GlobalReferencePoints = 2 Then
                                    AutofocusForm.DrawCrossGrid GlobalGridX1, GlobalGridY1
                                    AutofocusForm.DrawCrossGrid GlobalGridX2, GlobalGridY2
                                ElseIf GlobalReferencePoints = 1 Then
                                    AutofocusForm.DrawCrossGrid GlobalGridX1, GlobalGridY1
                                End If
                                RedrawGrid GlobalGridImage
                               
                            End If
                           End If
                        End If
                    End If
                End If
            End If
        End If
    End Select
'Set Stage = Nothing
marke1:
End Sub

Public Sub GridSelection(x As Long, y As Long, XR As Long, YR As Long, Activate As Boolean)
    Dim XPixels As Long
    Dim YPixels As Long
    Dim XGroup As Long
    Dim YGroup As Long
    Dim xIndx As Long
    Dim yIndx As Long
    Dim Start As Long
    Dim ix As Long
    Dim iy As Long
    Dim Xmin As Long
    Dim Xmax As Long
    Dim Ymin As Long
    Dim Ymax As Long
    Dim xImage As Long
    Dim StartX As Long
    Dim yImage As Long
    Dim StartY As Long
    Dim Found As Boolean

    If GlobalGridImage Is Nothing Then Exit Sub
    If x >= XR Then
        Xmin = XR
        Xmax = x
    Else
        Xmin = x
        Xmax = XR
    End If
    If y >= YR Then
        Ymin = YR
        Ymax = y
    Else
        Ymin = y
        Ymax = YR
    End If
    
    XPixels = GlobalGridImage.GetDimensionX
    YPixels = GlobalGridImage.GetDimensionY
    
    XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
    YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
    For yIndx = 1 To GlobalYGrid
        For xIndx = 1 To GlobalXGrid
            StartX = XGroup + 3 * XGroup * (xIndx - 1)
            StartY = YGroup + 3 * YGroup * (yIndx - 1)
            Found = False
            For iy = 1 To 2 * YGroup
                For ix = 1 To 2 * XGroup
                    xImage = StartX + ix
                    yImage = StartY + iy
                    If xImage >= Xmin And xImage <= Xmax And yImage >= Ymin And yImage <= Ymax Then
                        GlobalDeActivatedLocations(xIndx, yIndx) = Not Activate
                        Found = True
                        Exit For
                    End If
                    If Found Then Exit For
                Next ix
            Next iy
        Next xIndx
    Next yIndx
    GlobalOrderChanged = True
End Sub

Public Sub GridOnOff(x As Long, y As Long, XR As Long, YR As Long)
    Dim XPixels As Long
    Dim YPixels As Long
    Dim XGroup As Long
    Dim YGroup As Long
    Dim xIndx As Long
    Dim yIndx As Long
    Dim Start As Long
    Dim ix As Long
    Dim iy As Long
    Dim Xmin As Long
    Dim Xmax As Long
    Dim Ymin As Long
    Dim Ymax As Long
    Dim xImage As Long
    Dim StartX As Long
    Dim yImage As Long
    Dim StartY As Long
    Dim Found As Boolean

    If GlobalGridImage Is Nothing Then Exit Sub
    If x >= XR Then
        Xmin = XR
        Xmax = x
    Else
        Xmin = x
        Xmax = XR
    End If
    If y >= YR Then
        Ymin = YR
        Ymax = y
    Else
        Ymin = y
        Ymax = YR
    End If
    
    XPixels = GlobalGridImage.GetDimensionX
    YPixels = GlobalGridImage.GetDimensionY
    
    XGroup = Int(XPixels / (3 * GlobalXGrid + 1))
    YGroup = Int(YPixels / (3 * GlobalYGrid + 1))
    For yIndx = 1 To GlobalYGrid
        For xIndx = 1 To GlobalXGrid
            StartX = XGroup + 3 * XGroup * (xIndx - 1)
            StartY = YGroup + 3 * YGroup * (yIndx - 1)
            Found = False
            For iy = 1 To 2 * YGroup
                For ix = 1 To 2 * XGroup
                    xImage = StartX + ix
                    yImage = StartY + iy
                    If xImage >= Xmin And xImage <= Xmax And yImage >= Ymin And yImage <= Ymax Then
                        GlobalDeActivatedLocations(xIndx, yIndx) = Not GlobalDeActivatedLocations(xIndx, yIndx)
                        Found = True
                        Exit For
                    End If
                    If Found Then Exit For
                Next ix
            Next iy
        Next xIndx
    Next yIndx
    GlobalOrderChanged = True
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
       ' If TrackParameters.IsTrackUsed(lTrack) Then
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
       ' End If
    Next lTrack
End Sub
