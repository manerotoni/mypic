Attribute VB_Name = "RoiControl"
'''
' Module for Roi control. Due to roi class this module is close to become obsolete
''''

Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




'''''
' remove all vector elements
'''''
Public Function ClearVectorElements() As Boolean
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.Scancontroller.AcquisitionRegions
    vo.Cleanup
End Function

''''
' Make a Vectorelement (a ROI) to be used for bleaching or imaging
'   TypeVectorOverlay (In) - speifies type of ROI. "circle", "reactangle", "polyline", "ellipse"
'   X, Y              (In) - X and Y coordinates in pixel!! Upper left corner of image is 0, 0
'   Aim               (In) - Either "acquisition", "bleaching" (also includes analysis) or "analysis"
''''
Public Function MakeVectorElement(ByVal TypeVectorOverlay As String, X() As Double, Y() As Double, ByVal aim As String) As Boolean
'    Dim AcquisitionParameter As AimAcquisitionController40.AimAcquisitionParameters
'    Set AcquisitionParameter = Lsm5.ExternalDsObject.Scancontroller
    ' Get the Acquisition/Bleach ROIs
    Dim vo As AimImageVectorOverlay
    Set vo = Lsm5.ExternalDsObject.Scancontroller.AcquisitionRegions

    Dim i As Integer
    Dim ElementNumber As Long
    Select Case TypeVectorOverlay
        Case "circle", "Circle":
            If UBound(X) <> 1 Or UBound(Y) <> 1 Then
                MsgBox "MakeVectorElement: For a circle you need to define 2 points (in px). center point_on_circle"
                Exit Function
            End If
            vo.AddElement eImageVectorOverlayElementCircle
        Case "rectangle", "Rectangle":
            If UBound(X) <> 1 Or UBound(Y) <> 1 Then
                MsgBox "MakeVectorElement: For a square you need to define 2 points (in px). upper_left_corner and lower_right_corner"
                Exit Function
            End If
            'add a rectangle
            vo.AddElement eImageVectorOverlayElementRectangle
        Case "polyline", "Polyline":
            If UBound(X) <> UBound(Y) Or UBound(Y) < 2 Then
                MsgBox "MakeVectorElement: For a polyline you need to define at least 3 points (in px)"
                Exit Function
            End If
            vo.AddElement eImageVectorOverlayElementClosedPolyLine
        Case "ellipse", "Ellipse":
            If UBound(X) <> UBound(Y) Or UBound(X) < 2 Then
                MsgBox "MakeVectorElement: For an ellipse you need to define at least 3 points (in px). Center line_axis_1 line_axis_2"
                Exit Function
            End If
            vo.AddElement eImageVectorOverlayElementEllipse
        Case Else:
            MsgBox "MakeVectorElement: Does not understand the type of Roi. Types are circle, rectangle, polyline, ellipse"
            Exit Function
    End Select
    
    ElementNumber = vo.GetNumberElements - 1
    For i = 0 To UBound(X)
        vo.AppendElementKnot ElementNumber, X(i), Y(i), 0, 0
    Next i
    Sleep 50 ' this pause is require to finish setting the elements
    Select Case aim
        Case "Acquisition", "acquisition":
            vo.ElementAcquisitionFlags(ElementNumber) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAcquisition
            vo.ElementColor(ElementNumber) = "&H0000C000" 'this is green
        Case "Bleach", "bleach":
            vo.ElementAcquisitionFlags(ElementNumber) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsBleach Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
        Case "Analysis", "analysis":
            vo.ElementAcquisitionFlags(ElementNumber) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
        Case Else:
            MsgBox "MakeVectorElement: Does not understand the type of task. Use acquisition, bleach, or analysis"
            Exit Function
    End Select
    
End Function

Private Sub TestMakeVectorElement()
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.Scancontroller
    Dim X() As Double
    Dim Y() As Double

        ' Get the Acquisition/Bleach ROIs
    Dim vo As AimImageVectorOverlay
    Set vo = AcquisitionController.AcquisitionRegions
    vo.Cleanup
    ' add a circle
    ReDim X(0 To 1)
    ReDim Y(0 To 1)
    X(0) = 256
    Y(0) = 256
    X(1) = 256
    Y(1) = 200
    MakeVectorElement "circle", X, Y, "acquisition"
    
    X(0) = 256
    Y(0) = 256
    X(1) = 200
    Y(1) = 200
    MakeVectorElement "rectangle", X, Y, "bleach"
    
    ReDim X(2)
    ReDim Y(2)
    X(0) = 256
    Y(0) = 256
    X(1) = 200
    Y(1) = 200
    Y(2) = 100
    X(2) = 150
    MakeVectorElement "polyline", X, Y, "acquisition"
End Sub

