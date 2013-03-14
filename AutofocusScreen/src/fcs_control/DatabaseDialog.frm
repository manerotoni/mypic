VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatabaseDialog 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "DatabaseDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatabaseDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub BrowseButton_Click()
    Dim Name As String
    Name = FileNameTextBox
    UseCommonDialog Name
    FileNameTextBox = Name
    Lsm5.tools.RegStringValue("UI\" + GlobalMacroKey, "Directory") = FileNameTextBox
End Sub

Public Sub SetControlFlags()
    With CommonDialog
       .FLAGS = cdlOFNPathMustExist
       .FLAGS = .FLAGS Or cdlOFNHideReadOnly
       .FLAGS = .FLAGS Or cdlOFNNoChangeDir
       .FLAGS = .FLAGS Or cdlOFNExplorer
       .FLAGS = .FLAGS Or cdlOFNNoValidate
       .FileName = "*.*"
    End With

End Sub

Private Sub UseCommonDialog(MyPath As String)

    Dim flgUserChangeSaved As Boolean
    Dim lFreeSpace As Double
    Dim lSpace As Long
    Dim lngth As Long
    Dim Name1 As String
    Dim Start As Long
    Dim pos As Long
    Dim bslash As String
    Dim idx As Long
    Dim NumOfPositions As Long
    Dim tmpString As String
    Dim MyFile As String
    Dim driveString As String
    Dim fsTemp As New Scripting.FileSystemObject
    
    flgUserChangeSaved = flgUserChange
    flgUserChange = False
    
    'Initialize Common Dialog control
    If GlobalSystemVersion >= 30 Then
        SetControlFlags
    End If
                                
    If Not fsTemp.FolderExists(MyPath) Then
        MyPath = "C:\"
    End If
    lngth = Len(MyPath)
    If lngth >= 3 Then
        tmpString = Strings.Right(MyPath, 1)
        If tmpString <> "\" Then
            MyPath = MyPath + "\"
            lngth = lngth + 1
        End If
        tmpString = Strings.Left(MyPath, lngth - 1)
        MyFile = Dir(tmpString, vbDirectory)
        If MyFile <> "" Then
        Else
            tmpString = "C:\"
            MyPath = tmpString
            ChDir tmpString
        End If
        driveString = Strings.Left(MyPath, 1)
        ChDrive driveString
        ChDir tmpString
        
    End If
    CommonDialog.FileName = MyPath + "*.*"
    CommonDialog.Filter = "Temporary Files Folder ( *.* ) |*.*"
    
    CommonDialog.ShowOpen
    CommonDialog.FLAGS = 0
    tmpString = CommonDialog.FileTitle
    lngth = Len(CommonDialog.FileName)
    If lngth > 0 Then
        Name1 = Strings.Left(CommonDialog.FileName, lngth)
        Start = 1
        bslash = "\"
        pos = Start
        Do While pos > 0
            pos = InStr(Start, Name1, bslash)
            If pos > 0 Then
                Start = pos + 1
            End If
        Loop
        tmpString = Strings.Left(Name1, Start - 1)
        If Len(tmpString) >= 3 Then
            MyPath = tmpString
            pos = InStr(MyPath, ":")
            If pos <> 0 Then
                Name1 = Strings.Left(MyPath, pos) + "\"
            Else
                Name1 = MyPath
            End If
            CheckDiskSpace Name1, lFreeSpace, lSpace
'            If lFreeSpace < 10 ^ 8 Then
'                MsgBox "Warning! Drive contains only " + Strings.Format(lFreeSpace / 10 ^ 6, "0.00") + " MB of free space or do not exists! Please check the destination!"
'            Else
'                MsgBox "Information! Drive contains  " + Strings.Format(lFreeSpace / 10 ^ 6, "0.00") + " MB of free space"
'            End If
        End If
    End If
    
    flgUserChange = flgUserChangeSaved
End Sub
