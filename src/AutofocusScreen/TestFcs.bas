Attribute VB_Name = "TestFcs"
'Public FcsJob As AFcsJob
'Public Type FcsJobType
'    Name As String
'    LaserActive() As Boolean
'    BleachActive() As Boolean
'    LaserTransmission() As Double
'    BleachTransmission() As Double
'End Type
'
'
'Public Sub TestFcsClass()
'    Set ZEN = Lsm5.CreateObject("Zeiss.Micro.AIM.ApplicationInterface.ApplicationInterface")
'    Dim FcsControl As AimFcsController
'    Set FcsControl = Fcs
'    Debug.Print "Channels " & FcsControl.AcquisitionParameters.Channels '.ChannelDetectorA(0)
'    Debug.Print "Active " & FcsControl.AcquisitionParameters.ChannelEnabled(16) '.ChannelDetectorA(0)
'    Debug.Print "Category " & FcsControl.AcquisitionParameters.ChannelDetectorA(4)
'    If FcsJob Is Nothing Then
'        Set FcsJob = New AFcsJob
'        FcsJob.setJobNoAi "current"
'    End If
'    FcsJob.setJobNoAi "current"
'    If Not FcsJob.putJobNoAi Then
'        FcsJob.setJobNoAi "current"
'    End If
'    Debug.Print FcsJob.jobDescriptor
'    'FcsJob.setLightPath
'
'End Sub
'Public Sub loadAFcsJob()
'
''ZEN2011 or up
'    Dim FcsData As AimFcsData
'    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
'    Set ZEN = Application.ApplicationInterface
'    ZEN.gui.Fcs.LightPath.Lasers.ByIndex = 1 'set 458
'    ZEN.gui.Fcs.LightPath.Lasers.On.Value = True
'    'ZEN.gui.Fcs.LightPath.BleachLasers.IsEnabled = True
'    'ZEN.gui.Fcs.SaveMethod.Save.Execute
'    NewFcsRecord GlobalFcsRecordingDoc, FcsData, "Test"
'    Debug.Print ZEN.gui.Fcs.LightPath.Config.CurrentItem
'    'ZEN.CommandExecute "Fcs.BeamPath.Save"
'    'ZEN.CommandExecute "SimpleInput.Ok"
'    Dim FcsControl As AimFcsController
'    Set FcsControl = Fcs
'    FcsControl.BeamPathParameters.AttenuatorOn(1) = False
'    FcsData.DataSet(0).AcquisitionParameters.Copy FcsControl.AcquisitionParameters
'    FcsData.DataSet(0).AcquisitionParameters.MeasurementTime = 10
'    'FcsControl.AcquisitionParameters.MeasurementTime = 2
'    Dim AqPar As AimFcsAcquisitionParameters
'
'    AqPar.MeasurementTime = 0.1
'
'
'    'ZEN.gui.Fcs.method.Load "fcs1"
'    'ZEN.gui.Fcs.BeamPath.Save
'
''    starts with 1 this is the acquisition power
''    Power = FcsControl.BeamPathParameters.AttenuatorPower(2)
''    Power = FcsControl.BeamPathParameters.BleachAttenuatorPower(1)
''    FcsControl.BeamPathParameters.AttenuatorOn(1) = False
''    Dim FcsControl As AimFcsController
''    Set FcsControl = Fcs
''    ZEN.GUI.Fcs.LightPath.BleachLasers.ByIndex = 1
''    ZEN.GUI.Fcs.LightPath.BleachLasers.Transmission.Value = 0.1
'
'End Sub
''    starts with 1 this is the acquisition power
''    Power = FcsControl.BeamPathParameters.AttenuatorPower(2)
''    Power = FcsControl.BeamPathParameters.BleachAttenuatorPower(1)
''    FcsControl.BeamPathParameters.AttenuatorOn(1) = False
''    Dim FcsControl As AimFcsController
''    Set FcsControl = Fcs
''    ZEN.GUI.Fcs.LightPath.BleachLasers.ByIndex = 1
''    ZEN.GUI.Fcs.LightPath.BleachLasers.Transmission.Value = 0.1
'
