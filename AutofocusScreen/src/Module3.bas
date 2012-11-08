Attribute VB_Name = "Module3"
Sub Macro1()
    '**************************************
    'Recorded: 06/15/2012
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
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 0#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = -1#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = -1#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 0#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 0#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 1#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 1#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 2#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 2#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 3#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 3#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 4#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 4#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 5#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 5#
    Recording.FramesPerStack = 3
    Recording.Sample0Z = 1#
     
    '************* End ********************
End Sub
