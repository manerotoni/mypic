Attribute VB_Name = "CommonDialogAPI"
'CommonDialog allows for openSave file windows to be used
' From http://www.activevb.de/tipps/vb6tipps/tipp0368.html
' http://msdn.microsoft.com/en-us/library/ms645524%28VS.85%29.aspx
'--------- Anfang Modul "Module1" alias Module1.bas ---------
Option Explicit

Private Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
  
Private Declare Function GetWindowLongA Lib "user32" _
(ByVal hWnd As Long, _
ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLongA Lib "user32" _
(ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long
        
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
        Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const OFN_ALLOWMULTISELECT As Long = &H200&
Public Const OFN_CREATEPROMPT As Long = &H2000&
Public Const OFN_ENABLEHOOK As Long = &H20&
Public Const OFN_ENABLETEMPLATE As Long = &H40&
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80&
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400&
Public Const OFN_FILEMUSTEXIST As Long = &H1000&
Public Const OFN_HIDEREADONLY As Long = &H4&
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8&
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000&
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100&
Public Const OFN_OVERWRITEPROMPT As Long = &H2&
Public Const OFN_PATHMUSTEXIST As Long = &H800&
Public Const OFN_READONLY As Long = &H1&
Public Const OFN_SHAREAWARE As Long = &H4000&
Public Const OFN_SHAREFALLTHROUGH As Long = 2&
Public Const OFN_SHARENOWARN As Long = 1&
Public Const OFN_SHAREWARN As Long = 0&
Public Const OFN_SHOWHELP As Long = &H10&



Public Function ShowOpen(Filter As String, Flags As Long, Optional fileName As String = "", Optional initDir As String = "", Optional DialogTitle As String = "Open") As String
    Dim Buffer As String
    Dim Result As Long
    Dim ComDlgOpenFileName As OPENFILENAME
    
    Buffer = fileName & String$(128 - Len(fileName), 0)
    
    With ComDlgOpenFileName
        .lStructSize = Len(ComDlgOpenFileName)
        .Flags = Flags
        .nFilterIndex = 1&
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter
        .lpstrInitialDir = initDir
        .lpstrTitle = DialogTitle
    End With
    
    Result = GetOpenFileName(ComDlgOpenFileName)
    
    If Result <> 0 Then
        ShowOpen = Left$(ComDlgOpenFileName.lpstrFile, _
                   InStr(ComDlgOpenFileName.lpstrFile, _
                   Chr$(0)) - 1)
    End If
End Function

Public Function ShowSave(Filter As String, Flags As Long, fileName As String, Optional initDir As String = "", Optional DialogTitle As String = "Save As") As String
                           
    Dim Buffer As String
    Dim Result As Long
    Dim ComDlgOpenFileName As OPENFILENAME
    
    Buffer = fileName & String$(128 - Len(fileName), 0)
    
    With ComDlgOpenFileName
        .lStructSize = Len(ComDlgOpenFileName)
        .Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
        .nFilterIndex = 1&
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter
        .lpstrInitialDir = initDir
        .lpstrTitle = DialogTitle
    End With
    
    Result = GetSaveFileName(ComDlgOpenFileName)
    
    If Result <> 0 Then
        ShowSave = Left$(ComDlgOpenFileName.lpstrFile, _
                   InStr(ComDlgOpenFileName.lpstrFile, _
                   Chr$(0)) - 1)
    End If
End Function
'---------- Ende Modul "Module1" alias Module1.bas ----------
'-------------- Ende Projektdatei Project1.vbp --------------


'''''''''
'Minimize button for Macro window
''''''


Sub FormatUserForm(UserFormCaption As String)
     
    Dim hWnd            As Long
    Dim exLong          As Long
     
    hWnd = FindWindowA(vbNullString, UserFormCaption)
    exLong = GetWindowLongA(hWnd, -16)
    If (exLong And &H20000) = 0 Then
        SetWindowLongA hWnd, -16, exLong Or &H20000
    Else
    End If
     
End Sub
''''''''

