Attribute VB_Name = "A_Main_ModifySeries"
Option Explicit

Public flgBreak As Boolean
Public GlobalTimeStart() As Double
Public GlobalImageIndex() As Long
Public GlobalNodes() As AimExperimentTreeNode
Public TimeStampChange As Boolean
Public ZStepChange As Boolean
Public TimeNumberChange As Boolean
Public GlobalSetWlChange As Boolean
Public GlobalStartWl As Double
Public GlobalStepWl As Double
Public GlobalStartWlTmp As Double
Public GlobalStepWlTmp As Double
Public GlobalFileOptions As Long
Public GlobalNumberOfStacks As Long
Public GlobalTimeIntv As Double
Public GlobalTimeStampDate As Date
Public GlobalTimeStampTime As Date
Public GlobalZStep As Long
Public GlobalFileSource As Long
Public GlobalDirName As String
Public GlobalFiles() As String
Public GlobalNumberFiles As Long
Public GlobalUseBrowser As Boolean
Public GlobalUseChannelColor As Boolean
Public GlobalSingleImage As DsRecordingDoc
Public GlobalImageDocument As AimExperimentTreeNode
Public GlobalImage As AimImage
Public GlobalNumberOfChannels As Long


Declare Sub Sleep Lib "kernel32" (ByVal Time As Long)

Public Sub Main()
    Dim SystemVersion As String
    GlobalUseBrowser = False
    GlobalProjectName = "ModifySeriesZEN.lvb"
    GlobalHelpName = "ModifySeries.rtf"
    GlobalHelpNamePDF = "ModifySeries.pdf"
    
    GlobalMacroKey = "AutoModify"
    GlobalAutoStoreKey = "AutoModifyStore"
    GlobalUseChannelColor = False
    
    GetPathAndVersion GlobalPath, GlobalSystemVersion, GlobalMacrosPath
    
    If GlobalSystemVersion >= 45 Then
        DatabaseDialog.Show 0
    Else
        MsgBox "Program Requires LSM Release ZEN 2007 or Later"
    End If
End Sub

Public Sub SetDefaultWl(NoOfChannels As Long, StartWl As Double, StepWl As Double)

    StartWl = 400
    If NoOfChannels <= 5 Then
        StepWl = 10
    Else
        StepWl = 10
    End If

End Sub

Public Sub AutoStoreModify(Optional opt As Boolean)
    Dim key As String
    Dim myKey As String
    
    key = "UI\" + GlobalMacroKey
    myKey = key + "\" + GlobalAutoStoreKey
    If Lsm5.tools.RegExistKey(myKey) Then
        Lsm5.tools.RegDeleteKey (myKey)
    End If
    DatabaseDialog.GetFormControls
    Lsm5.tools.RegCreateKey (myKey)
    
    StoreRegistryModify key, myKey
End Sub

Public Sub ReadRegistryModify(key As String, myKey As String)
    GlobalNumberOfStacks = Lsm5.tools.RegLongValue(myKey, "GlobalNumberOfStacks")
    GlobalTimeIntv = Lsm5.tools.RegDoubleValue(myKey, "GlobalTimeIntv")
    GlobalDirName = ""
    'Lsm5.tools.RegStringValue(myKey, "Directory")
    GlobalFileSource = Lsm5.tools.RegLongValue(myKey, "GlobalFileSource")
       
End Sub

Public Sub StoreRegistryModify(key As String, myKey As String)
    Lsm5.tools.RegLongValue(myKey, "GlobalNumberOfStacks") = GlobalNumberOfStacks
    Lsm5.tools.RegDoubleValue(myKey, "GlobalTimeIntv") = GlobalTimeIntv
    Lsm5.tools.RegStringValue(myKey, "Directory") = GlobalDirName
    Lsm5.tools.RegLongValue(myKey, "GlobalFileSource") = GlobalFileSource

End Sub

Public Sub AutoRecallModify(Optional opt As Boolean)
    Dim key As String
    Dim myKey As String
    
    User_flg = False
    key = "UI\" + GlobalMacroKey
    myKey = key + "\" + GlobalAutoStoreKey
    
    If Lsm5.tools.RegExistKey(myKey) Then
        ReadRegistryModify key, myKey
    Else
        Lsm5.tools.RegCreateKey (myKey)
        StoreRegistryModify key, myKey
    End If
    
    DatabaseDialog.SetFormControls
    User_flg = True
End Sub

Public Function MakeDestinationDS(DestinationImageDocument As RecordingDocument, _
                                 DestinationImage As AimImage, _
                                 SizeX As Long, _
                                 SizeY As Long, _
                                 SizeZ As Long, _
                                 SizeT As Long, _
                                 SizeC As Long, _
                                 DataType As Long) As Boolean

    Dim DS As Document
'    Set DS = Lsm5.ExternalDsObject
    MakeDestinationDS = False

    Set DestinationImageDocument = Lsm5.ExternalDsObject.MakeNewImageDocument(SizeX, _
                                                           SizeY, _
                                                           SizeZ, _
                                                           SizeT, _
                                                           SizeC, _
                                                           DataType, _
                                                           1)

    If Not (DestinationImageDocument Is Nothing) Then Set DestinationImage = DestinationImageDocument.Image(0, False)
    If (DestinationImage Is Nothing) Then
        MsgBox "Cannot Create New Image!"
        Exit Function
    Else
        MakeDestinationDS = True
    End If
End Function

