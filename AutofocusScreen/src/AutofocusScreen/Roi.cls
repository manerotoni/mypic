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

Public roiType As String   'circle, rectangle, polyline, ellipse
Public roiAim As String    'acquisition, bleach, analysis
Private roiX() As Double   'set of X coordinates in pixels
Private roiY() As Double   'set of Y coordinates in pixels

Public Sub setX(X() As Double)
    roiX = X
End Sub


Public Sub setY(Y() As Double)
    roiY = Y
End Sub


Public Function getX() As Double
    getX = roiX
End Function


Public Function getY() As Double
    getY = roiY
End Function

Public Sub setRoi(roiTypeI As String, roiAimI As String, X() As Double, Y() As Double)
    roiType = roiTypeI
    roiAim = roiAimI
    roiX = X
    roiY = Y
    If Not roiConsistency Then
        MsgBox "Wrong dimensions for  roi " + roiType + " " + roiAim
    End If
End Sub


'''''
' remove all vector elements
'''''
Public Function ClearVectorElements() As Boolean
    Dim vo As AimImageVectorOverlay
    Set vo = AcquisitionController.AcquisitionRegions
    vo.Cleanup
End Function

Public Function roiConsistency() As Boolean
    Dim i As Integer
    If roiType <> "" And roiAim <> "" Then
        If UBound(roiX) <> UBound(roiY) Then
            MsgBox "For Roi: roiX and roiY coordinates need to be the same"
            Exit Function
        End If
    Else
        MsgBox "For Roi: a type and aim need to be defined"
        Exit Function
    End If
    
    Select Case roiType
        Case "circle", "Circle":
            If UBound(roiX) <> 1 Or UBound(roiY) <> 1 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define 2 points (in px)" + vbCrLf _
                & "roiX=center_X, point_on_circle_X" & vbCrLf & "roiY=center_Y, point_on_circle_Y"
                Exit Function
            End If
        Case "rectangle", "Rectangle":
            If UBound(roiX) <> 1 Or UBound(roiY) <> 1 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define 2 points (in px)" + vbCrLf _
                & "roiX=upper_left_X, lower_right_X" & vbCrLf & "roiY=upper_left_Y, lower_right_Y"
                Exit Function
            End If
        Case "polyline", "Polyline":
            If UBound(roiX) <> UBound(roiY) Or UBound(roiY) < 2 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define at least 3 points (in px)" + vbCrLf _
                & "roiX=point1_X, point2_X ..." & vbCrLf & "roiY=point1_Y, point2_Y..."
                Exit Function
            End If
        Case "ellipse", "Ellipse": ''Check this
            If UBound(roiX) <> 2 Or UBound(roiX) <> 2 Then
                MsgBox "roiConsistency: For a roiType=" + roiType + "ROI you have to define 3 points (in px). Center point_on_axis1 point_on_axis2" + vbCrLf _
                & "roiX=center_X, axis1_X, axis2_X" & vbCrLf & "roiY= center_Y, axis1_Y, axis2_Y"
                Exit Function
            End If
        Case Else:
            MsgBox "roiConsistency: Does not understand the type of Roi" + roiType + ". Types are circle, rectangle, polyline, ellipse"
            Exit Function
    End Select
    
    Select Case roiAim:
        Case "acquisition", "Acquisition":
        Case "bleach", "Bleach":
        Case "analysis", "Analysis"
        Case Else:
            Exit Function
    End Select
    roiConsistency = True
    
End Function

''''
' Make a Vectorelement (a ROI) to be used for bleaching or imaging
''''
Public Function MakeVectorElement() As Boolean
    If Not roiConsitency Then
        Exit Function
    End If
    Dim vo As AimImageVectorOverlay
    Set vo = AcquisitionController.AcquisitionRegions

    Dim i As Integer
    Dim ElementNumber As Long
    Select Case roiType
        Case "circle", "Circle":
            vo.AddElement eImageVectorOverlayElementCircle
        Case "rectangle", "Rectangle":
            vo.AddElement eImageVectorOverlayElementRectangle
        Case "polyline", "Polyline":
            vo.AddElement eImageVectorOverlayElementClosedPolyLine
        Case "ellipse", "Ellipse":
            vo.AddElement eImageVectorOverlayElementEllipse
    End Select
    
    ElementNumber = vo.GetNumberElements - 1
    For i = 0 To UBound(roiX)
        vo.AppendElementKnot ElementNumber, X(i), Y(i), 0, 0
    Next i
    Sleep 50 ' this pause is require to finish setting the elements
    Select Case roiAim
        Case "Acquisition", "acquisition":
            vo.ElementAcquisitionFlags(ElementNumber) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAcquisition
            vo.ElementColor(ElementNumber) = "&H0000C000" 'this is green
        Case "Bleach", "bleach":
            vo.ElementAcquisitionFlags(ElementNumber) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsBleach Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
            vo.ElementColor(ElementNumber) = "&H000000FF" ' this is red
            
        Case "Analysis", "analysis":
            vo.ElementAcquisitionFlags(ElementNumber) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
            vo.ElementColor(ElementNumber) = "&H00FFFF00" ' this is cyan
        Case Else:
            MsgBox "MakeVectorElement: Does not understand the type of task. Use acquisition, bleach, or analysis"
            Exit Function
    End Select
End Function

