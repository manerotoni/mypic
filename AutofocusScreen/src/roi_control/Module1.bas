Attribute VB_Name = "Module1"
Sub Macro1()
    '**************************************
    'Recorded: 11/21/2012
    'Description:
    '**************************************
    Dim ZEN As Zeiss_Micro_AIM_ApplicationInterface.ApplicationInterface
    Set ZEN = Application.ApplicationInterface
     
    ZEN.GUI.Acquisition.Snap.Execute

End Sub
