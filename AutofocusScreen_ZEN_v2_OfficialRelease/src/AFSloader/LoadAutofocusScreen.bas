Attribute VB_Name = "LoadAutofocusScreen"
Option Explicit

Public gKey As Long
Public Const LoaderMacroFileName As String = "AFSloader_AutofocusScreen_ZEN_version2.lvb"
Public Const MainMacroFileName As String = "AutofocusScreen_ZEN_version2.lvb"
Public Const StartupFunction = "newMacros.A_Setup"
'


Public Sub Start()
    Dim fs As New FileSystemObject, libPath As String, p As Lsm5VbaProject
    libPath = fs.GetParentFolderName(GetProjectObject(LoaderMacroFileName, 1).ProjectPath) & "\" & MainMacroFileName
    If Not fs.FileExists(libPath) Then
        MsgBox "AutoFocusScreen macro not found."
        ProjectUnLoad LoaderMacroFileName
    End If
    ShowVBAEditor 0
    gKey = Timer + 13957532
    Set p = ProjectLoad(libPath, 0)
    If p Is Nothing Then ProjectUnLoad LoaderMacroFileName Else p.ExecuteLine StartupFunction
End Sub

