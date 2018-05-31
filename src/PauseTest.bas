Attribute VB_Name = "PauseTest"
' Read every 100ms the registry
' Pause commands do not wait for Z-stack to finish. Image analysis needs to take care of it. For instance by sending first a wait command, then checking if file saving is finished
' When Pause is resumed imaging continues straight away


Sub Pause_Zeiss_micro_aim()
    ' Workflow continue after Pause.Execute
    '**************************************
    'Recorded: 05/31/2018
    'Description:
    '**************************************
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
     
    ZEN.GUI.Acquisition.StartExperiment.AsyncMode = True
    ZEN.GUI.Acquisition.StartExperiment.Execute
    SleepWithEvents 200
    'ZEN.GUI.Acquisition.TimeSeries.Pause.Execute
    
    'SleepWithEvents 200
    'ZEN.GUI.Acquisition.TimeSeries.Pause.Execute
    
End Sub


Sub Pause_LSM5()
    ' Workflow continue after PauseTimeSeries
    '**************************************
    'Recorded: 05/31/2018
    'Description:
    '**************************************
    Lsm5.StartAcquisition
    SleepWithEvents 200
    
    'Lsm5.PauseTimeSeries
    
    'SleepWithEvents 200
    'Lsm5.ResumeTimeSeries 1
    
End Sub



Sub Pause_AcquisitionController()

        

    ' Workflow continue after TimeSeriesPause
    '**************************************
    'Recorded: 05/31/2018
    'Description:
    '**************************************
    Dim AcquisitionController As AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
    Dim RecordingDoc As DsRecordingDoc
    
    Lsm5.StartAcquisition
    'AcquisitionController.
    'While Lsm5.DsRecordingActiveDocObject.IsBusy
    '    AcquisitionController.TimeSeriesPause True
    
    '    SleepWithEvents 200
    '   AcquisitionController.TimeSeriesPause False
    'Wend
End Sub



Sub AutoSave_Aim()
    ' Workflow continue after Pause.Execute
    '**************************************
    'Recorded: 05/31/2018
    'Description:
    '**************************************
    'Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    'Set ZEN = Application.ApplicationInterface
    'Lsm5.Options.UseAutosaveName
    'Lsm5.Options.UseAutosaveName = True
    Dim AcquisitionControllerOptions As AimScanControllerOptions
    Dim AcquisitionController As AimScanController
    Set AcquisitionController = Lsm5.ExternalDsObject.ScanController
    Set AcquisitionControllerOptions = AcquisitionController.ScanOptions
    AcquisitionControllerOptions.AutoSave = True
    AcquisitionControllerOptions.AutoSaveName = "AutoName"
    AcquisitionControllerOptions.AutoSaveDirectory = "D:\Antonio\Test\tmp"
    
    AcquisitionControllerOptions.AutoSaveSeparateFiles(3) = True
    
    
    
    '
    'Debug.Print AcquisitionController.ScanOptions.AutoSave
    AcquisitionController.StartGrab eGrabModeSingle
    
    'Lsm5.Options.AutoSaveBaseName = "DE_1"
    'Debug.Print Lsm5.Options.AutoSaveBaseName
    
    'ZEN.GUI.Acquisition.TimeSeries.Pause.Execute
    
    'SleepWithEvents 200
    'ZEN.GUI.Acquisition.TimeSeries.Pause.Execute
    
End Sub
