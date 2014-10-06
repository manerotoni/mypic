VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AFcsJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''
' Class for storing FCS job parameters
''''

Option Explicit
Public Name As String
Public LightPathConfig As String
Private LaserActive() As Boolean
Private LaserName() As String
Private BleachActive() As Boolean
Private LaserTransmission() As Double
Private BleachTransmission() As Double
Private MeasurementTime As Double
Private Repetitions As Integer
Private BleachTime As Double
Private PositionsEnabled As Boolean
Private TimeSeriesEnabled As Boolean
Private ptimeToAcquire As Double 'time to acquire a job in seconds

'Private Filters() As Long
'Private Channels() As Long
'
'Private BeamSplitterVis As Long
'Private BeamSplitterInVis As Long
'Private IRBlockFilter As Long
'Private BeamSplitter2 As Long
'Private EmissionFilterAPD1 As Long
'Private EmissionFilterAPD2 As Long
'Private SelectCh1 As Boolean
'Private SelectCh2 As Boolean
'Private SelectCh1Ch2 As Boolean
'Private SelectCh2Ch1 As Boolean

Public Property Get timeToAcquire() As Double
    timeToAcquire = ptimeToAcquire
End Property
    

Public Property Let timeToAcquire(value As Double)
    ptimeToAcquire = value
End Property

'''
' setJob(ZEN As Object, ZENv As Integer)
'   loads FCSJob from ZEN into Macro. Calls respective function
'''
Public Function SetJob(ZEN As Object, ZenV As Integer) As Boolean
    If ZenV > 2010 Then
       SetJob = setJobAi(ZEN)
    Else
       SetJob = setJobNoAi("current")
    End If
End Function

''''
'   reads FCS settings from ZEN into Macro. This command is used for ZENv < 2011
'   In the absence of ZEN_Micro_AIM_ApplicationInterface
'   Only stores which laser is active, its transmission and the measurement time and repetitions
'''
Public Function setJobNoAi(ConfigName As String) As Boolean
    Dim FcsControl As AimFcsController
    Set FcsControl = Fcs
    Dim i As Integer
    Dim NrLasers As Long
    
    LightPathConfig = ConfigName
    NrLasers = FcsControl.HardwareInformation.GetNumberAttenuators
   
    ReDim BleachActive(0 To NrLasers - 1)
    ReDim BleachTransmission(0 To NrLasers - 1)
    ReDim LaserActive(0 To NrLasers - 1)
    ReDim LaserTransmission(0 To NrLasers - 1)
    ReDim LaserName(0 To NrLasers - 1)
    For i = 0 To NrLasers - 1
          LaserName(i) = CStr(1000000000 * FcsControl.HardwareInformation.GetAttenuatorWavelength(i))
          LaserActive(i) = FcsControl.BeamPathParameters.AttenuatorOn(i)
          BleachActive(i) = FcsControl.BeamPathParameters.BleachAttenuatorOn(i)
          BleachTransmission(i) = FcsControl.BeamPathParameters.BleachAttenuatorPower(i)
          LaserTransmission(i) = FcsControl.BeamPathParameters.AttenuatorPower(i)
    Next i
    MeasurementTime = FcsControl.AcquisitionParameters.MeasurementTime
    Repetitions = FcsControl.AcquisitionParameters.MeasurementRepeat
    BleachTime = FcsControl.AcquisitionParameters.BleachTime
    setJobNoAi = True
End Function


''''
'   loads FCSJob from ZEN into Macro. Uses the GUI (only works for ZENv > 2010)
'   recquires ZEN_Micro_AIM_ApplicationInterface
'   force to save the lightpath config. This config can then be reloaded when needed
'   Save also settings for lasers. It is compatible with putJobNoGui
'''
Public Function setJobAi(ZEN As Object) As Boolean
    On Error GoTo ErrorHandle
    Dim FcsControl As AimFcsController
    Set FcsControl = Fcs
    Dim i As Integer
    Dim NrLasers As Integer
    If Not ZEN Is Nothing Then
        NrLasers = ZEN.gui.Fcs.LightPath.Lasers.ItemCount
        ZEN.CommandExecute "Fcs.BeamPath.Save"
        'get the current saved item. Loading this item is in priciple enough for the experiment
        LightPathConfig = ZEN.gui.Fcs.LightPathConfig.CurrentItem
        
        ReDim LaserActive(0 To NrLasers - 1)
        ReDim LaserTransmission(0 To NrLasers - 1)
        ReDim LaserName(0 To NrLasers - 1)
        
        For i = 0 To NrLasers - 1
            ZEN.gui.Fcs.LightPath.Lasers.ByIndex = i
            LaserActive(i) = ZEN.gui.Fcs.LightPath.Lasers.On.value
            LaserTransmission(i) = ZEN.gui.Fcs.LightPath.Lasers.Transmission.value / 100
            LaserName(i) = ZEN.gui.Fcs.LightPath.Lasers.ByName
        Next i
        
        ReDim BleachActive(0 To NrLasers - 1)
        ReDim BleachTransmission(0 To NrLasers - 1)
        
        For i = 0 To NrLasers - 1
            ZEN.gui.Fcs.LightPath.BleachLasers.ByIndex = i
            BleachActive(i) = ZEN.gui.Fcs.LightPath.BleachLasers.On.value
            BleachTransmission(i) = ZEN.gui.Fcs.LightPath.BleachLasers.Transmission.value / 100
        Next i
        MeasurementTime = FcsControl.AcquisitionParameters.MeasurementTime
        Repetitions = FcsControl.AcquisitionParameters.MeasurementRepeat
        BleachTime = FcsControl.AcquisitionParameters.BleachTime
        PositionsEnabled = ZEN.gui.Fcs.EnablePositions.value
        TimeSeriesEnabled = ZEN.gui.Fcs.EnableTimeSeries.value
        Me.timeToAcquire = 0
        setJobAi = True
        
    Else
        Exit Function
    End If
    Exit Function
ErrorHandle:
    MsgBox "Error in AFcsJob.setJobAi " + Err.Description
    LogManager.UpdateErrorLog "Error in AFcsJob.setJobGui " + Err.Description
End Function


'''
' putJob(ZEN As Object, ZENv As Integer)
'   put stored settings from Macro back into ZEN. Calls respective function depending on ZEN version
'''
Public Function PutJob(ZEN As Object, ZenV As Integer) As Boolean
    If ZenV > 2010 Then
        PutJob = putJobAi(ZEN)
    Else
        PutJob = putJobNoAi
    End If
End Function


'''
'     Transfers settings from Macro into ZEN using the GUI (ZENv > 2011)
'     recquires ZEN_Micro_AIM_ApplicationInterface
'     Complete setting with name LightPathConfig for light path can be uploaded
'     also the Measurement time etc are stored
'''
Public Function putJobAi(ZEN As Object) As Boolean

    On Error GoTo ErrorHandle
    If Not ZEN Is Nothing Then
        'get the current saved item. Loading this item is in priciple enough for the experiment
        If ZEN.gui.Fcs.LightPathConfig.isValidItem(LightPathConfig) Then
            ZEN.gui.Fcs.LightPathConfig.Load LightPathConfig
        Else
            If LightPathConfig <> "" Then
                MsgBox "AFcsJob.putJobGui. For Fcs Job " & Name & ". No LightPathconfig with the name " & LightPathConfig
            Else
                MsgBox "AFcsJob.putJobGui. Fcs Job " & Name & " has not been initiated use ZEN->Macro button"
            End If
            Exit Function
        End If
        ZEN.gui.Fcs.EnablePositions.value = PositionsEnabled
        ZEN.gui.Fcs.EnableTimeSeries.value = TimeSeriesEnabled
        ZEN.SetDouble "Fcs.Acquisition.MeasureTime", MeasurementTime
        ZEN.SetDouble "Fcs.Acquisition.Repetitions", Repetitions
        ZEN.SetDouble "Fcs.Acquisition.BleachTime", BleachTime
        putJobAi = True
    
    
        Exit Function
    End If
    
    Exit Function
ErrorHandle:
    MsgBox "Error in AFcsJob.setJobGui " + Err.Description
    LogManager.UpdateErrorLog "Error in AFcsJob.setJobGui " + Err.Description
End Function

     
     
'''
' putJobNoGui()
'   Transfer settings from Macro into ZEN without using the GUI
'   In the absence of ZEN_Micro_AIM_ApplicationInterface (ZENv <2011) we only save which lasers are on, laser power, and time lapse.
'   All other settings are not saved and use the
'   current status
'''
Public Function putJobNoAi() As Boolean
    Dim FcsControl As AimFcsController
    Set FcsControl = Fcs
    Dim i As Integer
    Dim NrLasers As Long
    NrLasers = FcsControl.HardwareInformation.GetNumberAttenuators
    For i = 0 To NrLasers - 1
          FcsControl.BeamPathParameters.AttenuatorOn(i) = LaserActive(i)
          FcsControl.BeamPathParameters.BleachAttenuatorOn(i) = BleachActive(i)
          FcsControl.BeamPathParameters.BleachAttenuatorPower(i) = BleachTransmission(i)
          FcsControl.BeamPathParameters.AttenuatorPower(i) = LaserTransmission(i)
    Next i
    FcsControl.AcquisitionParameters.MeasurementTime = MeasurementTime
    FcsControl.AcquisitionParameters.MeasurementRepeat = Repetitions
    FcsControl.AcquisitionParameters.BleachTime = BleachTime
End Function

''''
' Create a string with the description of a Job
''''
Public Function jobDescriptor() As String
    Dim i As Integer
    Dim Description As String
    Description = "FcsJob: " & Name & vbCrLf
    If LightPathConfig <> "" Then
        Description = Description & "Configuration: " & LightPathConfig & vbCrLf
        Description = Description & "FcsLasers: " & vbCrLf
        For i = 0 To UBound(LaserName)
            If LaserActive(i) Then
                Description = Description & "  " & LaserName(i) & " at " & CStr(Round(LaserTransmission(i) * 100, 3)) & "%" & vbCrLf
            End If
        Next i
        Description = Description & "BleachLasers: " & vbCrLf
        For i = 0 To UBound(LaserName)
            If BleachActive(i) Then
                Description = Description & "  " & LaserName(i) & " at  " & CStr(Round(BleachTransmission(i) * 100, 3)) & "%" & vbCrLf
            End If
        Next i
        Description = Description & "MeasurementTime = " & MeasurementTime & " s" & vbCrLf
        Description = Description & "Repetitions = " & Repetitions & vbCrLf
        Description = Description & "Bleach Time = " & BleachTime & " s" & vbCrLf
 
    Else
        Description = Description & "Job has not been set"
    End If
    jobDescriptor = Description
End Function

'''
' split a String into two different Strigs depending on their maximal length
''''
Public Function splittedJobDescriptor(MaxLines As Integer, jobDescription As String) As String()
    Dim i As Integer
    Dim Label1() As String
    Dim Label2() As String
    Dim Output() As String
    ReDim Output(1)
    Label1 = Split(jobDescription, vbCrLf)
    If UBound(Label1) > MaxLines - 1 Then
        ReDim Label2(0 To UBound(Label1) - MaxLines)
        For i = 0 To UBound(Label1) - MaxLines
            Label2(i) = Label1(MaxLines + i)
        Next i
        ReDim Preserve Label1(0 To MaxLines - 1)
    Else
        ReDim Label2(0)
    End If
    Output(0) = Join(Label1, vbCrLf)
    Output(1) = Join(Label2, vbCrLf)
    splittedJobDescriptor = Output
End Function

''
' create a String to save of the "complete" job desriptin
''
Public Function jobDescriptorSettings() As String
    Dim Descriptor As String
    Descriptor = "JobFcsName " & Me.Name & vbCrLf
    Descriptor = Descriptor & "LightPathConfig " & LightPathConfig & vbCrLf
    Descriptor = Descriptor & "PositionsEnabled " & PositionsEnabled & vbCrLf
    Descriptor = Descriptor & "TimeSeriesEnabled " & TimeSeriesEnabled & vbCrLf
    Descriptor = Descriptor & "MeasurementTime " & MeasurementTime & vbCrLf
    Descriptor = Descriptor & "Repetitions " & Repetitions & vbCrLf
    Descriptor = Descriptor & "BleachTime " & BleachTime & vbCrLf
    Descriptor = Descriptor & "EndJobFcsDef " & vbCrLf
    jobDescriptorSettings = Descriptor
End Function

Public Sub testDescriptor()
    consistencyDescriptor (jobDescriptorSettings)
End Sub

Private Function consistencyDescriptor(Descriptor As String)
    Dim i As Integer
    Dim DerivedDescr() As String
    Dim KeyItem() As String
    DerivedDescr = Split(Descriptor, vbCrLf)
    For i = 0 To UBound(DerivedDescr)
        'Debug.Print DerivedDescr(i)
        Sleep (100)
        KeyItem = Split(DerivedDescr(i), " ")
        If UBound(KeyItem) = 1 Then
            If Not changeJobFromDescriptor(KeyItem(0), KeyItem(1)) Then
                Exit Function
            End If
        End If
    Next i
    consistencyDescriptor = True
End Function


''''
' change a job from a descriptor
'''
Public Function changeJobFromDescriptor(Descriptor As String, value As Variant)
    
    Select Case Descriptor
        Case "JobFcsName":
            Name = CStr(value)
        Case "LightPathConfig":
            LightPathConfig = CStr(value)
        Case "PositionsEnabled":
            PositionsEnabled = CBool(value)
        Case "TimeSeriesEnabled":
            TimeSeriesEnabled = CBool(value)
        Case "MeasurementTime":
            MeasurementTime = CDbl(value)
        Case "Repetitions":
            Repetitions = CInt(value)
        Case "BleachTime":
            BleachTime = CDbl(value)
        Case "EndJobFcsDef":
            
        Case Else:
            MsgBox " No Item " & Descriptor
            Exit Function
    End Select
    changeJobFromDescriptor = True
End Function



' some commands one can use
'    Dim Power As Double
'    starts with 1 this is the acquisition power
'    Power = FcsControl.BeamPathParameters.AttenuatorPower(2)
'    Power = FcsControl.BeamPathParameters.BleachAttenuatorPower(1)
'    FcsControl.BeamPathParameters.AttenuatorOn(1) = False
'    Dim FcsControl As AimFcsController
'    Set FcsControl = Fcs
'    ZEN.GUI.Fcs.LightPath.BleachLasers.ByIndex = 1
'    ZEN.GUI.Fcs.LightPath.BleachLasers.Transmission.Value = 0.1
        'load Filters we only save the laser power
        '        BeamSplitterVis = ZEN.gui.Fcs.LightPath.BeamSplitterVis.ByIndex
        '        BeamSplitterInVis = ZEN.gui.Fcs.LightPath.BeamSplitterInVis.ByIndex
        '        IRBlockFilter = ZEN.gui.Fcs.LightPath.IRBlockFilter.ByIndex
        '        BeamSplitter2 = ZEN.gui.Fcs.LightPath.BeamSplitter2.ByIndex
        '        EmissionFilterAPD1 = ZEN.gui.Fcs.LightPath.EmissionFilterAPD1.ByIndex
        '        EmissionFilterAPD2 = ZEN.gui.Fcs.LightPath.EmissionFilterAPD2.ByIndex
        '        SelectCh1 = ZEN.gui.Fcs.LightPath.SelectCh1.Value
        '        SelectCh2 = ZEN.gui.Fcs.LightPath.SelectCh2.Value
        '        SelectCh1Ch2 = ZEN.gui.Fcs.LightPath.SelectCh1Ch2.Value
        '        SelectCh2Ch1 = ZEN.gui.Fcs.LightPath.SelectCh2Ch1.Value
            
    'there are too many things one shall record we only set the laser power
    'NrFilters = FcsControl.HardwareInformation.GetNumberFilterSets
    'For i = 0 To NrFilters - 1
    '    Filters(i) = FcsControl.BeamPathParameters.FilterSetPosition(i)
    'Next i
    ' NrChannels = FcsControl.AcquisitionParameters.Channels
'    For i = 0 To NrChannels - 1
'        Channels(i) = FcsControl.AcquisitionParameters.ChannelEnabled(i)
'    Next i
    'SelectCh1 = FcsControl.AcquisitionParameters.ChannelDetectorA
    'dumper = FcsControl.BeamPathParameters.Dump
'    Debug.Print FcsControl.AcquisitionParameters.ChannelDetectorA
'    Debug.Print "Position Filter " & FcsControl.BeamPathParameters.FilterSetPosition(0)
