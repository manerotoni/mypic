VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public Recording As DsRecording
Private pAcquireTrack() As Boolean 'must be private for no obvious reasons of VBA
Public pixelDwell As Double      'pixelDwell of acquisition
Public TrackNumber As Integer
Public TimeBetweenStacks As Double
Public UseRoi As Boolean
Public CorrX  As Double ' corretion for double directional scanning
Public CorrY As Double  ' corretion for double directional scanning

Private Rois() As Roi
Private pRoiMaxNr As Integer

Public Property Get AcquireTrack() As Boolean()
    AcquireTrack = pAcquireTrack
End Property

Public Property Let AcquireTrack(Value() As Boolean)
    Dim i As Integer
    ReDim pAcquireTrack(UBound(Value))
    For i = 0 To UBound(Value)
        pAcquireTrack(i) = Value(i)
    Next i
End Property


Public Sub SetAcquireTrack(iTrack As Integer, Value As Boolean)
    If UBound(pAcquireTrack) >= iTrack Then
        pAcquireTrack(iTrack) = Value
    End If
End Sub

Public Function GetAcquireTrack(iTrack As Integer) As Boolean
    If UBound(pAcquireTrack) >= iTrack Then
        GetAcquireTrack = pAcquireTrack(iTrack)
    End If
End Function

Public Sub ReDimAcquireTrack(Value As Integer)
    ReDim Preserve pAcquireTrack(Value)
End Sub


Public Function GetRecording() As DsRecording
    Set GetRecording = Recording
End Function

Public Sub SetRecording(Value As DsRecording)
   Set Recording = Lsm5.CreateBackupRecording
   Recording.Copy Value
   TrackNumber = Recording.GetNormalTrackCount
End Sub

Public Sub setRoiMaxNr(roiMaxNr As Integer)
    pRoiMaxNr = roiMaxNr
    If pRoiMaxNr = 0 Then
        Erase Rois
    Else
        ReDim Rois(0 To pRoiMaxNr - 1)
    End If
End Sub

Public Function getRoiMaxNr() As Integer
    getRoiNr = pRoiMaxNr
End Function

Public Sub clearRois()
    ClearVectorElements
    pRoiMaxNr = 0
    Erase Rois
End Sub

Public Sub setRoi(ARoi As Roi, roiNr As Integer)
    If UBound(Rois) <= roiNr Then
        Rois(roiNr).setRoi ARoi.roiType, ARoi.roiAim, ARoi.getX, ARoi.getY
    Else
        MsgBox (" No Roi of nr" + roiNr + " initialize Rois first ")
    End If
End Sub

Public Function getRoi(roiNr As Integer) As Roi
    If UBound(Rois) <= roiNr Then
        Set getRoi = Rois(roiNr)
    Else
        MsgBox (" No Roi of nr" + roiNr + " initialize Rois first ")
    End If
End Function



Public Function createRois() As Boolean
    Dim ARoi As Roi
    Dim i As Integer
    Rois(i).ClearVectorElements
    For i = 0 To pRoiMaxNr - 1
            If Not Rois(i).MakeVectorElement Then
                Exit Function
            End If
        Next i
    End If
    createRois = True
End Function

'Public Property Get TrackNumber(Value As Integer)
'    pTrackNumber =
'    pRecording = Lsm5.CreateBackupRecording
'    pRecording.Copy DsRecording
'End Property
