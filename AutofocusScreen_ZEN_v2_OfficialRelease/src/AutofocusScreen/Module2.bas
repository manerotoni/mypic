Attribute VB_Name = "Module2"
Sub Macro1()
    '**************************************
    'Recorded: 06/14/2012
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
     
    Recording.Sample0Z = 335.71
    Recording.SpecialScanMode = "FocusStep"
    Recording.scanMode = "Stack"
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Set DataChannel = Track.DataChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Set DataChannel = Track.DataChannelObjectByIndex(1, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Track.HdrImagingMode = 0
    Track.HdrNumFrames = 1
    Set DataChannel = Track.DataChannelObjectByIndex(1, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Track.HdrIntensity = 10
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Track.HdrImagingMode = 0
    Set Track = Recording.TrackObjectByIndex(0, Success)
    Recording.FrameSpacing = 0.666667
    Recording.FramesPerStack = 4
    Recording.FrameSpacing = 0.666667
    Recording.FramesPerStack = 4
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set DataChannel = Track.DataChannelObjectByName("ChS1", Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(2, Success)
    Set DataChannel = Track.DataChannelObjectByName("Ch2", Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(11, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Recording.FrameSpacing = 0.5
    Recording.FramesPerStack = 5
    Recording.FrameSpacing = 0.5
    Recording.FramesPerStack = 5
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set DataChannel = Track.DataChannelObjectByName("ChS1", Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(2, Success)
    Set DataChannel = Track.DataChannelObjectByName("Ch2", Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(11, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Recording.Sample0Z = 1#
     
    '************* End ********************
End Sub
Sub Macro2()
    '**************************************
    'Recorded: 06/14/2012
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
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Recording.Sample0Z = 1.25
    Recording.FramesPerStack = 6
    Recording.Sample0Z = 1.25
    Recording.FramesPerStack = 6
    Recording.Sample0Z = 1.5
    Recording.FramesPerStack = 7
    Recording.Sample0Z = 1.5
    Recording.FramesPerStack = 7
    Recording.Sample0Z = 1.75
    Recording.FramesPerStack = 8
    Recording.Sample0Z = 1.75
    Recording.FramesPerStack = 8
    Recording.SpecialScanMode = "ZScanner"
    Recording.Sample0Z = 1.75
    Recording.FramesPerStack = 8
    Recording.FrameSpacing = 0.5
    Recording.Sample0Z = 1.75
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Recording.FrameSpacing = 1.5
    Recording.Sample0Z = 5.25
    Recording.FrameSpacing = 1.5
    Recording.Sample0Z = 5.25
    Set Track = Recording.TrackObjectByMultiplexOrder(0, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set DataChannel = Track.DataChannelObjectByName("ChS1", Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(2, Success)
    Set DataChannel = Track.DataChannelObjectByName("Ch2", Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(11, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
    Set DetectionChannel = Track.DetectionChannelObjectByIndex(0, Success)
    Set Track = Recording.TrackObjectByMultiplexOrder(1, Success)
     
    '************* End ********************
End Sub
