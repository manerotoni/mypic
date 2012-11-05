Attribute VB_Name = "Module1"
Sub Macro1()
    '**************************************
    'Recorded: 06/13/2012
    'Descrption:
    '**************************************
    Dim RecordingDoc As DsRecordingDoc
    Dim Recording As DsRecording
    Dim Track As DsTrack
    Dim Laser As DsLaser
    Dim DetectionChannel As DsDetectionChannel
    Dim IlluminationChannel As DsIlluminationChannel
    Dim DataChannel As DsDataChannel
    Dim BeamSplitter As DsBeamSplitter
    Dim Timers As DsTimers
    Dim Markers As DsMarkers
    Dim Success As Integer
    Set Recording = Lsm5.DsRecording
     
    Recording.Sample0Z = 1#
     
    '************* End ********************
End Sub
Sub Macro2()
    '**************************************
    'Recorded: 06/13/2012
    'Descrption:
    '**************************************
    Dim RecordingDoc As DsRecordingDoc
    Dim Recording As DsRecording
    Dim Track As DsTrack
    Dim Laser As DsLaser
    Dim DetectionChannel As DsDetectionChannel
    Dim IlluminationChannel As DsIlluminationChannel
    Dim DataChannel As DsDataChannel
    Dim BeamSplitter As DsBeamSplitter
    Dim Timers As DsTimers
    Dim Markers As DsMarkers
    Dim Success As Integer
    Set Recording = Lsm5.DsRecording
     
    Lsm5.DsRecording.Sample0Z = 1#
    MsgBox ":asdsds"
    
     
    '************* End ********************
End Sub
