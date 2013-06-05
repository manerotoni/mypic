Attribute VB_Name = "Roi_Control"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Macro1()
    '**************************************
    'Recorded: 11/21/2012
    'Description:
    '**************************************
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
     
    ZEN.GUI.Acquisition.Snap.Execute

End Sub


Sub GetVectorElements(Rois As AimImageVectorOverlay)
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
    Set Rois = AcquisitionController.AcquisitionRegions
    
End Sub

Sub TranslateVectorElements(Rois As AimImageVectorOverlay, X As Double, Y As Double, Z As Double)
    Dim i As Long
    Rois.Copy Rois, 1, 0, 0, 0, X, 0, 1, 0, 0, Y, 0, 0, 1, 0, Z, 0, 0, 0, 0, 0
End Sub

Sub GetCenterVectorElements(Rois As AimImageVectorOverlay, Element As Long, XCenter As Double, YCenter As Double)
    Dim Knot As Long
    Dim T As Double
    Dim Z As Double
    Dim X As Double
    Dim Y As Double
    XCenter = 0
    YCenter = 0
    If Element < Rois.GetNumberElements And Element > -1 Then
        Select Case Rois.ElementType(Element)
            Case eImageVectorOverlayElementRectangle
                For Knot = 0 To Rois.ElementKnotSize(Element) - 2
                   Rois.GetElementKnot Element, Knot, X, Y, Z, T
                   XCenter = XCenter + X
                   YCenter = YCenter + Y
                Next Knot
                XCenter = XCenter / (Rois.ElementKnotSize(Element) - 1)
                YCenter = YCenter / (Rois.ElementKnotSize(Element) - 1)
            Case eImageVectorOverlayElementCircle
                Rois.GetElementKnot Element, 0, XCenter, YCenter, Z, T
            Case eImageVectorOverlayElementClosedPolyLine
                For Knot = 0 To Rois.ElementKnotSize(Element) - 2
                   Rois.GetElementKnot Element, Knot, X, Y, Z, T
                   XCenter = XCenter + X
                   YCenter = YCenter + Y
                Next Knot
                XCenter = XCenter / (Rois.ElementKnotSize(Element) - 1)
                YCenter = YCenter / (Rois.ElementKnotSize(Element) - 1)
        End Select
    End If
End Sub

Sub MakeVectorElements()
   
    Dim ImageVectorOverlay As AimImage40.AimImageVectorOverlay
    Dim AcquisitionController As AimAcquisitionController40.AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
   
    ' Get the Acquisition/Bleach ROIs
    Dim vo As AimImageVectorOverlay
    Set vo = AcquisitionController.AcquisitionRegions
    ' remove all elements

    '
    Dim Knot As Long
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim T As Double

    vo.Cleanup
    ' add a circle
    vo.AddElement eImageVectorOverlayElementCircle
    vo.AppendElementKnot 0, 100, 100, 0, 0
    vo.AppendElementKnot 0, 50, 50, 0, 0
    
    'add a rectangle
    vo.AddElement eImageVectorOverlayElementRectangle
    vo.AppendElementKnot 1, 100, 100, 0, 0
    vo.AppendElementKnot 1, 50, 50, 0, 0
    
    ' add a polygon
    vo.AddElement eImageVectorOverlayElementClosedPolyLine
    vo.AppendElementKnot 2, 232, 200, 0, 0
    vo.AppendElementKnot 2, 268, 154, 0, 0
    vo.AppendElementKnot 2, 359, 297, 0, 0
    
    vo.AddElement eImageVectorOverlayElementEllipse
    vo.AppendElementKnot 3, 232, 200, 0, 0
    vo.AppendElementKnot 3, 268, 154, 0, 0
    vo.AppendElementKnot 3, 359, 297, 0, 0
    vo.AppendElementKnot 3, 600, 297, 0, 0
'    Dim X As Double
'    Dim Y As Double
'    Dim Z As Double
'    vo.GetElementVectorEnd 1, X, Y, Z, 0
'    vo.GetElementVectorStart 1, X, Y, Z, 0
   ' vo.Copy vo, 0, 0, 0, 0, -10, 0, 0, 0, 0, -10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    'vo.Copy vo, 1, 0, 0, 0, -10, 0, 1, 0, 0, -10, 1, 1, 1, 1, -10, 1, 1, 1, 1, 0
    ' wait a short while, otherwise the main programm is sometimes not finished with setting the elements
    Sleep 100
    
    ' Set flags (Use "Or" to concatenate multiple flags)
    vo.ElementAcquisitionFlags(0) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAcquisition
    vo.ElementAcquisitionFlags(1) = AimVectorOverlay40.eVectorOverlayAcquisitionFlagsBleach Or AimVectorOverlay40.eVectorOverlayAcquisitionFlagsAnalysis
    vo.ElementColor(1) = "&H0000C000" ' this is green
    'this activate or not the rois
    vo.ElementValid(0) = True
    vo.ElementValid(1) = False
    'this activates or not ROIS and show them ?? need to be tested
    vo.ElementDisabled(0) = False
    vo.ElementDisabled(1) = False
    Dim XCenter As Double
    Dim YCenter As Double
    GetCenterVectorElements vo, 2, XCenter, YCenter
    
End Sub

Sub Macro2()
    '**************************************
    'Recorded: 02/04/2013
    'Description:
    '**************************************
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
     
    ZEN.GUI.Acquisition.Regions.RegionList.ByIndex = 0

    ZEN.SetDouble "AcqisitionRegion.CenterPixelsX", 50
End Sub
Sub Macro3()
    '**************************************
    'Recorded: 02/04/2013
    'Description:
    '**************************************
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
     
    ZEN.GUI.Acquisition.Snap.Execute

    ZEN.GUI.Acquisition.EnableRegions.Value = False

    ZEN.GUI.Acquisition.Snap.Execute

End Sub
