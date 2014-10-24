VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Roi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Roi Class
'Allows to store a Roi check for its consitency and create it
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public roiType As String   'circle, rectangle, polyline, ellipse
Public roiAim As String    'acquisition, bleach, analysis
Private roiX() As Double   'set of X coordinates in pixels
Private roiY() As Double   'set of Y coordinates in pixels
Private roiAimDic As Dictionary 'map name to number
Private roiTypeDic As Dictionary ' map name to number


Private Sub populateDic()
    Set roiAimDic = New Dictionary
    Set roiTypeDic = New Dictionary
    roiAimDic.Add "", ""
    roiAimDic.Add "nothing", 0
    roiAimDic.Add "bleach", AimVectorOverlay40.eVectorOverlayAcquisitionFlagsBleach

    roiAimDic.Add "bleachAnalyse", AimVectorOverlay40.eVectorOverlayAcquisitionFlagsBleach Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
    roiAimDic.Add "acquisition", AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAcquisition
    roiAimDic.Add "acquisitionBleach", AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAcquisition Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis _
            Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsBleach
    roiAimDic.Add "acquisitionAnalyse", AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAcquisition Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
    roiAimDic.Add "analyse", AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
    roiTypeDic.Add "circle", eImageVectorOverlayElementCircle
    roiTypeDic.Add "rectangle", eImageVectorOverlayElementRectangle
    roiTypeDic.Add "polyline", eImageVectorOverlayElementClosedPolyLine
    roiTypeDic.Add "ellipse", eImageVectorOverlayElementEllipse
    
End Sub

Private Function createRoiTxt(roiTypeI As Integer, roiAimI As Integer) As Boolean
    Dim i As Integer
    Dim Keys() As Variant
    Dim Items() As Variant
    
    If (roiAimDic Is Nothing) Or (roiTypeDic Is Nothing) Then
        populateDic
    End If
    Keys = roiAimDic.Keys
    Items = roiAimDic.Items
    Me.roiAim = ""
    Me.roiType = ""
    For i = 0 To UBound(Items)
        If Items(i) = roiAimI Then
            Me.roiAim = Keys(i)
            Exit For
        End If
    Next i
    
    If Me.roiAim = "" Then
        MsgBox "roi with code roiAim : " & roiAimI & " is not implemented"
        Exit Function
    End If
    
    Keys = roiTypeDic.Keys
    Items = roiTypeDic.Items
    For i = 0 To UBound(Items)
        If Items(i) = roiTypeI Then
            Me.roiType = Keys(i)
            Exit For
        End If
    Next i
    
    If Me.roiType = "" Then
        MsgBox "roi with code roiType : " & roiTypeI & " is not implemented"
        Exit Function
    End If
    createRoiTxt = True
    
End Function


Public Sub setX(X() As Double)
    roiX = X
End Sub


Public Sub setY(Y() As Double)
    roiY = Y
End Sub


Public Function getX() As Double()
    getX = roiX
End Function


Public Function getY() As Double()
    getY = roiY
End Function

Public Sub setRoiFomZen(roiNr As Integer)
    On Error GoTo ErrorHandle
    Dim vo As AimImageVectorOverlay
    Dim i As Long
    Dim X() As Double
    Dim Y() As Double
    Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions
    'Debug.Print "nr Elements "; vo.GetNumberElements
    If roiNr <= vo.GetNumberElements Then
         If createRoiTxt(vo.ElementType(roiNr), vo.ElementAcquisitionFlags(roiNr)) Then
            ReDim X(0 To vo.GetElementNumberKnots(roiNr) - 1)
            ReDim Y(0 To vo.GetElementNumberKnots(roiNr) - 1)
            For i = 0 To vo.GetElementNumberKnots(roiNr) - 1
                vo.GetElementKnot roiNr, i, X(i), Y(i), 0, 0
            Next i
            Me.setX X
            Me.setY Y
        Else
             MsgBox "Roi.setRoiFromZen: problems in generating Roi"
        End If
    Else
        MsgBox "Roi.setRoiFromZen no roi nr " & roiNr + 1
    End If
Exit Sub
ErrorHandle:
   MsgBox "Error in Roi.setRoiFromZen " & Err.Description
End Sub

Public Sub setRoi(roiType As String, roiAim As String, X() As Double, Y() As Double)
    Me.roiType = roiType
    Me.roiAim = roiAim
    roiX = X
    roiY = Y
    If Not roiConsistency Then
        MsgBox "Wrong dimensions/type/aim for  roi " + roiType + " " + roiAim
    End If
End Sub


'''''
' remove all vector elements
'''''
Public Function ClearVectorElements() As Boolean
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions
    vo.Cleanup
End Function

Public Function roiConsistency() As Boolean
    Dim i As Integer
    If (roiAimDic Is Nothing) Or (roiTypeDic Is Nothing) Then
        populateDic
    End If
    
    If roiType <> "" And roiAim <> "" Then
        If UBound(roiX) <> UBound(roiY) Then
            MsgBox "For Roi: roiX and roiY coordinates need to be the same"
            Exit Function
        End If
    Else
        MsgBox "For Roi: a type and aim need to be defined"
        Exit Function
    End If
     
    If Not roiTypeDic.Exists(roiType) Then
        MsgBox "Roi.roiConsistency: Does not understand the type of Roi" + roiType + ". Types are circle, rectangle, polyline, ellipse"
        Exit Function
    End If
    
    Select Case roiType
        Case "circle":
            If UBound(roiX) <> 1 Or UBound(roiY) <> 1 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define 2 points (in px)" + vbCrLf _
                & "roiX=center_X, point_on_circle_X" & vbCrLf & "roiY=center_Y, point_on_circle_Y"
                Exit Function
            End If
        Case "rectangle":
            If UBound(roiX) <> 1 Or UBound(roiY) <> 1 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define 2 points (in px)" + vbCrLf _
                & "roiX=upper_left_X, lower_right_X" & vbCrLf & "roiY=upper_left_Y, lower_right_Y"
                Exit Function
            End If
        Case "polyline":
            If UBound(roiX) <> UBound(roiY) Or UBound(roiY) < 2 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define at least 3 points (in px)" + vbCrLf _
                & "roiX=point1_X, point2_X ..." & vbCrLf & "roiY=point1_Y, point2_Y..."
                Exit Function
            End If
        Case "ellipse":
            If UBound(roiX) <> 2 Or UBound(roiX) <> 2 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define 3 points (in px). Center point_on_axis1 point_on_axis2" + vbCrLf _
                & "roiX=center_X, axis1_X, axis2_X" & vbCrLf & "roiY= center_Y, axis1_Y, axis2_Y"
                Exit Function
            End If
        Case Else:
            MsgBox "roiConsistency: Does not understand the type of Roi" + roiType + ". Types are circle, rectangle, polyline, ellipse"
            Exit Function
    End Select
    
    If Not roiAimDic.Exists(Me.roiAim) Then
        MsgBox "Roi.roiConsistency: Does not understand the aim of Roi, roiAim = " & roiAim & vbCrLf & _
        ". Aims are acquisition, bleach, bleachAnalysis, acquisitionBleach, acquisitionAnalysis, acquisitionBleachAnalysis"
        Exit Function
    End If
    
    roiConsistency = True
End Function

''''
' Make a Vectorelement (a ROI) to be used for bleaching or imaging
''''
Public Function MakeVectorElement() As Boolean
    If Not roiConsistency Then
        Exit Function
    End If
    
    If (roiAimDic Is Nothing) Or (roiTypeDic Is Nothing) Then
        populateDic
    End If
    
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.ScanController.AcquisitionRegions

    Dim i As Integer
    Dim ElementNumber As Long
    
    vo.AddElement roiTypeDic.item(roiType)
    ElementNumber = vo.GetNumberElements - 1
    Sleep 50 ' this pause is require to finish setting the elements
    vo.ElementAcquisitionFlags(ElementNumber) = roiAimDic.item(roiAim)
    Sleep 50 ' this pause is require to finish setting the elements
    For i = 0 To UBound(roiX)
        vo.AppendElementKnot ElementNumber, roiX(i), roiY(i), 0, 0
    Next i

    MakeVectorElement = True
End Function

