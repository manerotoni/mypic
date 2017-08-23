VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatabaseDialog 
   Caption         =   "Concat"
   ClientHeight    =   11625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   OleObjectBlob   =   "DatabaseDialog.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatabaseDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''Version Description''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Concat v2.0.6
'''''''''''''''''''''End: Version Description'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const Version = " v2.0.6"
Option Explicit 'force to declare all variables
Private Pattern() As String

Dim DatabaseDialogLoaded As Boolean




Private Sub Recursive_Click()
    If Recursive Then
        ConcatenatetimeMarkedLocation.Enabled = False
    Else
        ConcatenatetimeMarkedLocation.Enabled = True
    End If
End Sub

Private Sub StopButton_Click()
    flgBreak = True
End Sub

Private Sub UserForm_Activate()
    If Not DatabaseDialogLoaded Then
        LoadWindowPosition
    End If
    DatabaseDialogLoaded = True
    ReDim Pattern(4)
    Pattern(0) = "(\w+\d+)_+T(\d+).(lsm|czi)" 'new standard formt
    Pattern(1) = "(\d+\-*\d+\-*\d+\-*\(\d+,\d+\)\-*\w+\-*\w+)\-*(\d+).lsm" 'Cellbase Format
    Pattern(2) = "(\w+)_R(\d+).(lsm|czi)" 'old obsolete format
    Pattern(3) = "(\w+\-*W\w+\-*P\w+)\-*T(\d+).lsm" 'old intermediate format obsolete
    ConcatenatetimeMarkedLocation.Enabled = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    AutoStoreModify
    SaveWindowPosition
End Sub

Private Sub OptionAllFiles_Click()
    If User_flg Then
        GlobalFileOptions = 3
        OptionAllFiles.Value = True
        OptionZStack.Value = False
        OptionTimeStack.Value = False
        Frame1.Caption = "All images"
        FillListBox
        Image1.Visible = False
        SetButtons
        DisplayProgress "Select image...", RGB(&HC0, &HC0, 0)
    End If
End Sub

Private Sub OptionTimeStack_Click()
    If User_flg Then
        GlobalFileOptions = 1
        OptionTimeStack.Value = True
        OptionZStack.Value = False
        OptionAllFiles.Value = False
        Frame1.Caption = "Time series"
        FillListBox
        Image1.Visible = False
        SetButtons
        DisplayProgress "Select image...", RGB(&HC0, &HC0, 0)
    End If
End Sub

Private Sub OptionZStack_Click()
    If User_flg Then
        GlobalFileOptions = 2
        OptionZStack.Value = True
        OptionTimeStack.Value = False
        OptionAllFiles.Value = False
        Frame1.Caption = "Z-Stacks"
        FillListBox
        Image1.Visible = False
        SetButtons
        DisplayProgress "Select image...", RGB(&HC0, &HC0, 0)
    End If
End Sub

''''''
' UserForm_Initialize()
'   Function called from e.g. AutoFocusForm.Show
'   Load and initialize form
''
Private Sub UserForm_Initialize()
    'Setting of some global variables
    
    Me.Caption = Me.Caption + Version + " for ZEN "
    FormatUserForm (Me.Caption) ' make minimizing button available
    flgBreak = False
    SetDefaultWl 4, GlobalStartWl, GlobalStepWl
    AutoRecallModify
    FillListBox
    SetFormControls
    SetButtons
    SetControls
    User_flg = True
    OptionAllFiles_Click
    ZENv = getVersionNr
    
End Sub


Public Sub FillListBox()
    Dim ViewerGuiServer As AimViewerGuiServer
    Dim ExperimentTree As AimExperimentTreeNode
    Dim NumberOfSelected As Long
    Dim Image As AimImage
    Dim Node As Long
    Dim Files() As String
    Dim PathNames() As String
    Dim NumberFiles As Long
    Dim index As Long
    Dim DimensionT As Long
    Dim DimensionZ As Long
    If GlobalFileSource = 0 Then
        Set ViewerGuiServer = Lsm5.ViewerGuiServer
        Set ExperimentTree = ViewerGuiServer.ExperimentTree
        
        If OptionZStack.Value Then
            ImagesListBox.MultiSelect = fmMultiSelectExtended
        ElseIf OptionTimeStack.Value Then
            ImagesListBox.MultiSelect = fmMultiSelectExtended
        Else
            ImagesListBox.MultiSelect = fmMultiSelectExtended
        End If
    
        ImagesListBox.Clear
        NumberOfSelected = 0
        
        For Node = 0 To ExperimentTree.NumberChildren - 1
            If ExperimentTree.Child(Node).Type = eExperimentTeeeNodeTypeLsm Then
                Set Image = ExperimentTree.Child(Node).Image(0)
                If Not Image Is Nothing Then
                
                    If (OptionTimeStack.Value And Image.ImageMemory.GetDimensionT > 1) _
                       Or (OptionZStack.Value And Image.ImageMemory.GetDimensionZ > 1) _
                       Or ((Not OptionTimeStack.Value) And (Not OptionZStack.Value)) Then
            
                        ImagesListBox.AddItem ExperimentTree.Child(Node).Name
                        ReDim Preserve GlobalNodes(NumberOfSelected + 1)
                        Set GlobalNodes(NumberOfSelected + 1) = ExperimentTree.Child(Node)
                        NumberOfSelected = NumberOfSelected + 1
                    End If
                End If
            End If
        Next Node
    Else
        GlobalNumberFiles = 0
        ImagesListBox.Clear
    
        If ListFiles(Files, PathNames, NumberFiles, FileNameTextBox, "*.lsm") Then
            Image1.Visible = False
            For index = 0 To NumberFiles - 1
                If ReadImageInformation(Image, DimensionT, DimensionZ, PathNames(index)) Then
                    If (OptionTimeStack.Value And DimensionT > 1) _
                       Or (OptionZStack.Value And DimensionZ > 1) _
                       Or ((Not OptionTimeStack.Value) And (Not OptionZStack.Value)) Then
                
                        ImagesListBox.AddItem Files(index)
                        ReDim Preserve GlobalFiles(GlobalNumberFiles + 1)
                        GlobalFiles(GlobalNumberFiles + 1) = PathNames(index)
                        GlobalNumberFiles = GlobalNumberFiles + 1
                    End If
                End If
            Next index
            SetButtons
        End If
        If ListFiles(Files, PathNames, NumberFiles, FileNameTextBox, "*.czi") Then
            Image1.Visible = False
            For index = 0 To NumberFiles - 1
                If ReadImageInformation(Image, DimensionT, DimensionZ, PathNames(index)) Then
                    If (OptionTimeStack.Value And DimensionT > 1) _
                       Or (OptionZStack.Value And DimensionZ > 1) _
                       Or ((Not OptionTimeStack.Value) And (Not OptionZStack.Value)) Then
                
                        ImagesListBox.AddItem Files(index)
                        ReDim Preserve GlobalFiles(GlobalNumberFiles + 1)
                        GlobalFiles(GlobalNumberFiles + 1) = PathNames(index)
                        GlobalNumberFiles = GlobalNumberFiles + 1
                    End If
                End If
            Next index
            SetButtons
        End If
    End If
End Sub

Private Function ReadImageInformation(Image As AimImage, DimensionT As Long, DimensionZ As Long, filename As String) As Boolean

    ReadImageInformation = False
    
On Error GoTo ErrorExit

    Dim Import As AimImageImport
'    Set SourceImage = Lsm5.CreateObject("AimImage.Image")
'    Set SourceImageDocument = Lsm5.CreateObject("AimExperiment.TreeNode")
    Set Import = Lsm5.CreateObject("AimImageImportExport.Import")
'    Set ImageCopy = Lsm5.CreateObject("AimImageProcessing.Copy")
    
    Import.filename = filename
    Import.ReadFullSizeFileInformation Image
    DimensionT = Import.FileInfoSize(eAimImportExportCoordinateT)
    DimensionZ = Import.FileInfoSize(eAimImportExportCoordinateZ)
    ReadImageInformation = True
    Exit Function
    
ErrorExit:
End Function

Private Sub BrowseButton_Click()
    Dim Name As String
    Name = FileNameTextBox
    If GlobalUseBrowser Then
        If BrowseForFolder(Name, "Input directory") Then
            FileNameTextBox = Name
            Lsm5.tools.RegStringValue("UI\" + GlobalMacroKey, "Directory") = FileNameTextBox
        End If
    Else
        UseCommonDialog Name
        FileNameTextBox = Name
        Lsm5.tools.RegStringValue("UI\" + GlobalMacroKey, "Directory") = FileNameTextBox
    End If
End Sub

Private Sub UseCommonDialog(MyPath As String)
    Dim lpReOpenBuff As OFSTRUCT
    Dim wStyle As Long
    Dim hFile As Long
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
    CommonDialog.filename = MyPath + "*.*"
    CommonDialog.Filter = "Temporary Files Folder ( *.* ) |*.*"
    CommonDialog.ShowOpen
    CommonDialog.FLAGS = 0
    tmpString = CommonDialog.FileTitle
    lngth = Len(CommonDialog.filename)
    If lngth > 0 Then
        Name1 = Strings.Left(CommonDialog.filename, lngth)
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

Public Sub SetControlFlags()
    With CommonDialog
       .FLAGS = cdlOFNPathMustExist
       .FLAGS = .FLAGS Or cdlOFNHideReadOnly
       .FLAGS = .FLAGS Or cdlOFNNoChangeDir
       .FLAGS = .FLAGS Or cdlOFNExplorer
       .FLAGS = .FLAGS Or cdlOFNNoValidate
       .filename = "*.*"
    End With

End Sub






Private Sub SetControls()
    User_flg = False
    If GlobalFileSource = 0 Then
        FileNameTextBox.Visible = False
        BrowseButton.Visible = False
        Label1.Visible = False
        OptionButton1.Value = True
        OptionButton2.Value = False
        
    Else
        FileNameTextBox.Visible = True
        BrowseButton.Visible = True
        Label1.Visible = True
        OptionButton1.Value = False
        OptionButton2.Value = True
        
    End If
    User_flg = True
End Sub

Private Sub FileNameTextBox_Change()
    If User_flg Then
        FillListBox
    End If
End Sub

Private Sub OptionButton1_Click()
    If User_flg Then
        GlobalFileSource = 0
        OptionButton1.Value = True
        OptionButton2.Value = False
        FillListBox
        Image1.Visible = False
        SetControls
        SetButtons
        DisplayProgress "Select image...", RGB(&HC0, &HC0, 0)
    End If

End Sub

Private Sub OptionButton2_Click()
    If User_flg Then
        GlobalFileSource = 1
        OptionButton1.Value = False
        OptionButton2.Value = True
        FillListBox
        Image1.Visible = False
        SetControls
        SetButtons
        DisplayProgress "Select image...", RGB(&HC0, &HC0, 0)
    End If

End Sub




Private Sub CommandButton1_Click()
    flgBreak = True
End Sub




Private Sub ExitButton_Click()
    AutoStoreModify
    SaveWindowPosition

    End
End Sub

Private Sub HelpButton_Click()
    DisplayHelp GlobalHelpNamePDF, GlobalHelpName

'    Dim dblTask As Double
'    Dim MacroPath As String
'    Dim MyPath As String
'    Dim bslash As String
'    Dim success As Integer
'    Dim pos As Integer
'    Dim Start As Integer
'    Dim count As Long
'    Dim ProjName As String
'    Dim indx As Integer
'
'    count = ProjectCount()
'    For indx = 0 To count - 1
'        MacroPath = ProjectPath(indx, success)
'        ProjName = ProjectTitle(indx, success)
'        If StrComp(ProjName, GlobalProjectName, vbTextCompare) = 0 Then
'            Start = 1
'            bslash = "\"
'            pos = Start
'            Do While pos > 0
'                pos = InStr(Start, MacroPath, bslash)
'                If pos > 0 Then
'                    Start = pos + 1
'                End If
'            Loop
'            MyPath = Left(MacroPath, Start - 1)
'            MyPath = MyPath + GlobalHelpName
'            dblTask = Shell("C:\Program Files\Windows NT\Accessories\wordpad.exe " + MyPath, vbNormalFocus)
'            Exit For
'        End If
'    Next indx
End Sub

Private Sub ImagesListBox_change()

If User_flg Then
    SetButtons
End If
End Sub

Private Sub SetButtons()
    Dim result As Boolean
    Dim SourceImageDocument As AimExperimentTreeNode
    Dim SourceImage As AimImage
    Dim Import As AimImageImport
    Dim Thumbnail As AimImageBitmap
On Error GoTo NoImage
    Set SourceImage = Lsm5.CreateObject("AimImage.Image")
    Set SourceImageDocument = Lsm5.CreateObject("AimExperiment.TreeNode")
    Set Import = Lsm5.CreateObject("AimImageImportExport.Import")
    Set Thumbnail = Lsm5.CreateObject("AimImageBitmap.Bitmap")
    
    User_flg = False
    If (ImagesListBox.ListIndex <> -1) Then
        If OptionZStack.Value Then
            OptionTimeStack.Value = False
            OptionAllFiles.Value = False
        ElseIf OptionTimeStack.Value Then
            OptionTimeStack.Value = True
            OptionZStack.Value = False
            OptionAllFiles.Value = False
        Else
            OptionTimeStack.Value = False
            OptionZStack.Value = False
            OptionAllFiles.Value = True
        End If
        If GlobalFileSource = 0 Then
            If LoadSourceImage(SourceImageDocument, SourceImage, ImagesListBox.ListIndex) Then
                'Set Thumbnail = SourceImageDocument.Thumbnail(0, 200, 200, SourceImage.ImageMemory.GetDimensionZ() / 2, SourceImage.ImageMemory.GetDimensionT() / 2, Nothing)
                Set Thumbnail = SourceImageDocument.Thumbnail
                If Not Thumbnail Is Nothing Then
                    'Image1.Picture = TransferPicture(Thumbnail).Picture
                    'Image1.Visible = True
                    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
                Else
                    Image1.Visible = False
                    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
                End If
            Else
                Image1.Visible = True
                DisplayProgress "Select image...", RGB(&HC0, &HC0, 0)
            End If
        Else
            Import.filename = GlobalFiles(ImagesListBox.ListIndex + 1)
            'Import.ReadFullSizeFileInformation SourceImage
            'Import.ReadThumbnail Thumbnail, Import.FileInfoSize(eAimImportExportCoordinateT) / 2, _
            'Import.FileInfoSize(eAimImportExportCoordinateZ) / 2, 128, 128
            'Image1.Visible = True
            'Image1.Picture = TransferPicture(Thumbnail).Picture
            User_flg = True
            Exit Sub
        
        End If
    Else
        Image1.Visible = False
        DisplayProgress "Select File...", RGB(&HC0, &HC0, 0)
    End If
NoImage:
    User_flg = True
End Sub




Public Sub GetFormControls()
    User_flg = False
    GlobalDirName = FileNameTextBox
    If GlobalFileOptions = 1 Then
        OptionZStack.Value = True
        OptionTimeStack.Value = False
        OptionAllFiles.Value = False
    ElseIf GlobalFileOptions = 2 Then
        OptionZStack.Value = False
        OptionTimeStack.Value = True
        OptionAllFiles.Value = False
    ElseIf GlobalFileOptions = 3 Then
        OptionZStack.Value = False
        OptionTimeStack.Value = False
        OptionAllFiles.Value = True
    End If
    User_flg = True
End Sub

Public Sub SetFormControls()
    User_flg = False
    FileNameTextBox = GlobalDirName
    If GlobalFileOptions = 1 Then
        OptionZStack.Value = True
        OptionTimeStack.Value = False
        OptionAllFiles.Value = False
    ElseIf GlobalFileOptions = 2 Then
        OptionZStack.Value = False
        OptionTimeStack.Value = True
        OptionAllFiles.Value = False
    ElseIf GlobalFileOptions = 3 Then
        OptionZStack.Value = False
        OptionTimeStack.Value = False
        OptionAllFiles.Value = True
    End If
    SetControls
    User_flg = True
End Sub

Sub Heapsort(arr() As Double, hcount As Long, art() As Long)
Dim i As Long
Dim L As Long
Dim Ir As Long
Dim Rra As Double
Dim Tra As Double
Dim J As Long

If hcount > 1 Then
  For i = 1 To hcount
    art(i) = i
  Next i
  L = CInt(hcount / 2) + 1
  Ir = hcount
Cont:
  If L > 1 Then
    L = L - 1
    Rra = arr(art(L))
    Tra = art(L)
  Else
    Rra = arr(art(Ir))
    Tra = art(Ir)
    art(Ir) = art(1)
    Ir = Ir - 1
    If Ir = 1 Then
      art(1) = Tra
      GoTo Done
    End If
  End If
  i = L
  J = L + L
back:
  If J <= Ir Then
    If J < Ir Then
      If arr(art(J)) < arr(art(J + 1)) Then
        J = J + 1
      End If
    End If
    If Rra < arr(art(J)) Then
      art(i) = art(J)
      i = J
      J = J + J
    Else
      J = Ir + 1
    End If
    GoTo back
  End If
  art(i) = Tra
  GoTo Cont
Done:
End If
End Sub

Public Function LoadSourceImage(Document As AimExperimentTreeNode, _
                                 Image As AimImage, index As Long) As Boolean
    Dim ViewerGuiServer As AimViewerGuiServer
    Dim ViewerContext As AimImageViewerContext
    Dim MyProgress As AimProgress
    Dim ProgressFiFo As IAimProgressFifo

    LoadSourceImage = False
    If GlobalFileSource = 0 Then
        If GlobalNodes(index + 1) Is Nothing Then Exit Function
        Set Document = GlobalNodes(index + 1)
        If Document Is Nothing Then Exit Function
        Set Image = Document.Image(0)
        If Image Is Nothing Then Exit Function
        LoadSourceImage = True
    Else
        If (index = -1) Then Exit Function
        Set ViewerGuiServer = Lsm5.ViewerGuiServer
        ViewerGuiServer.LoadFile GlobalFiles(index + 1), True
        Set ViewerContext = ViewerGuiServer.CurrentViewer
        If ViewerContext Is Nothing Then Exit Function
        Set Document = ViewerContext.ExperimentTreeNode
    
        Set Image = Document.Image(0)
        If Image Is Nothing Then Exit Function
        Set ProgressFiFo = Image
        ProgressFiFo.Get MyProgress
        If Not MyProgress Is Nothing Then
            While Not MyProgress.Ready
                Sleep 100
                DoEvents
            Wend
        End If
        If Image Is Nothing Then Exit Function
        LoadSourceImage = True
        
        
    End If
End Function

Private Function MakeDestination(Document As AimExperimentTreeNode, _
                                 Image As AimImage, _
                                 SizeX As Long, _
                                 SizeY As Long, _
                                 SizeZ As Long, _
                                 SizeT As Long, _
                                 SizeC As Long, _
                                 DataType As enumAimImageDataType) As Boolean
                                 
    Dim ImageMemory As AimImageMemory

    MakeDestination = False
    Set Document = Lsm5.NewDocument
    If Document Is Nothing Then Exit Function
    Set Image = Document.Image(0)
    If Image Is Nothing Then Exit Function
    Set ImageMemory = Image
    If ImageMemory Is Nothing Then Exit Function
    
On Error GoTo ErrorExit
    ImageMemory.Create SizeC, SizeT, SizeZ, SizeY, SizeX, DataType
    
    MakeDestination = True
    Exit Function
ErrorExit:
    MsgBox "Cannot Create New Image!"
End Function

Private Function WaitProgress(Progress As AimProgress)

    WaitProgress = True
    While Not Progress.Ready
        DoEvents
        Sleep (100)
        If flgBreak Then
            User_flg = True
            WaitProgress = False
            Exit Function
        End If
    Wend

End Function

Private Function DoConcatenate_Time(Name2 As String) As Boolean

    Dim SourceImageNodeDocument As AimExperimentTreeNode
    Dim DestinationImageDocument As AimExperimentTreeNode
    
'    Dim DestinationImageDocument As RecordingDocument
    Dim SourceImage As AimImage
    Dim DestinationImage As AimImage
    Dim ImageCopy As AimImageCopy
    Dim index As Long
    Dim SizeT As Long
    Dim SizeZ As Long
    Dim SizeY As Long
    Dim SizeX As Long
    Dim SizeC As Long
    
    Dim Td As Long
    Dim Ts As Long
    Dim Cs As Long
    Dim Cd As Long
    Dim TimeSystemStart As Double
    Dim TimeSystemStartLast As Double

    Dim TimeDifference As Double
    Dim TimeDifferenceLast As Double
    Dim NumberOfSelected As Long
    Dim PreviousTime As Double
    Dim TimeStart() As Double
    Dim TimeStamp() As Double
    Dim TimeSort() As Double
    Dim TimeStampMax As Double
    Dim TimeStartMin As Double
    Dim Nodes() As AimExperimentTreeNode
    Dim IndexArray() As Long
    Dim NumberOfImages As Long
    Dim SlctFileName() As String
    Dim ViewerGuiServer As AimViewerGuiServer
    Dim ViewerContext As AimImageViewerContext
    Dim MyProgress As AimProgress
    Dim ProgressFiFo As IAimProgressFifo
    
    Dim Import As AimImageImport
    Dim TmpImage As AimImage
    Dim DataType As Long
    Dim TmpTimeStamp As Double
    Dim Count As Long
    Dim EventTimeStamp As Double
    Dim EventType As Long
    Dim EventDescription As String
    Dim Es As Long

On Error GoTo Finish

    flgBreak = False
    User_flg = False
    Set TmpImage = Lsm5.CreateObject("AimImage.Image")
'    Set SourceImageDocument = Lsm5.CreateObject("AimExperiment.TreeNode")
    Set Import = Lsm5.CreateObject("AimImageImportExport.Import")
    Set ImageCopy = Lsm5.CreateObject("AimImageProcessing.Copy")

    If (ImagesListBox.ListIndex <> -1) Then
        Set ViewerGuiServer = Lsm5.ViewerGuiServer
        DisplayProgress "Working...", RGB(0, &HC0, 0)
        
        SizeT = 0
        
        NumberOfSelected = 0
        TimeStampMax = 0
        TimeStartMin = 10 ^ 10
        NumberOfImages = ImagesListBox.ListCount
        For index = 1 To NumberOfImages
            If ImagesListBox.Selected(index - 1) Then
                ReDim Preserve TimeStart(NumberOfSelected + 1)
                ReDim Preserve Nodes(NumberOfSelected + 1)
                ReDim Preserve TimeStamp(NumberOfSelected + 1)
                ReDim Preserve TimeSort(NumberOfSelected + 1)
                ReDim Preserve SlctFileName(NumberOfSelected + 1)
                
                If GlobalFileSource = 0 Then
                    Set SourceImageNodeDocument = GlobalNodes(index)
                    Set SourceImage = SourceImageNodeDocument.Image(0)
                    Set Nodes(NumberOfSelected + 1) = SourceImageNodeDocument
                Else
                    Set SourceImage = Lsm5.CreateObject("AimImage.Image")
'                    Set SourceImage = New AimImage
                    Import.filename = GlobalFiles(index)
                    Import.ReadFullSizeFileInformation SourceImage
                
                End If
                If GlobalFileSource = 0 Then
                    SlctFileName(NumberOfSelected + 1) = SourceImageNodeDocument.Name
                Else
                    SlctFileName(NumberOfSelected + 1) = GlobalFiles(index)
                End If
                TimeStart(NumberOfSelected + 1) = CDbl(SourceImage.Characteristics.AcquisitionDateAndTime)
                TimeStamp(NumberOfSelected + 1) = SourceImage.ImageMemory.TimeStamp(0)
                If TimeStamp(NumberOfSelected + 1) > TimeStampMax Then TimeStampMax = TimeStamp(NumberOfSelected + 1)
                If TimeStart(NumberOfSelected + 1) < TimeStartMin Then TimeStartMin = TimeStart(NumberOfSelected + 1)
                If GlobalFileSource = 0 Then
                    SizeT = SizeT + SourceImage.ImageMemory.GetDimensionT
                    If SourceImage.ImageMemory.GetDimensionZ >= SizeZ Then SizeZ = SourceImage.ImageMemory.GetDimensionZ
                    If SourceImage.ImageMemory.GetDimensionY >= SizeY Then SizeY = SourceImage.ImageMemory.GetDimensionY
                    If SourceImage.ImageMemory.GetDimensionX >= SizeX Then SizeX = SourceImage.ImageMemory.GetDimensionX
                    If SourceImage.ImageMemory.GetDimensionC >= SizeC Then SizeC = SourceImage.ImageMemory.GetDimensionC
                    
                Else
                    SizeT = SizeT + Import.FileInfoSize(eAimImportExportCoordinateT)
                    If Import.FileInfoSize(eAimImportExportCoordinateZ) >= SizeZ _
                    Then SizeZ = Import.FileInfoSize(eAimImportExportCoordinateZ)
                    If Import.FileInfoSize(eAimImportExportCoordinateY) >= SizeY _
                    Then SizeY = Import.FileInfoSize(eAimImportExportCoordinateY)
                    If Import.FileInfoSize(eAimImportExportCoordinateX) >= SizeX _
                    Then SizeX = Import.FileInfoSize(eAimImportExportCoordinateX)
                    If Import.FileInfoSize(eAimImportExportCoordinateC) >= SizeC _
                    Then SizeC = Import.FileInfoSize(eAimImportExportCoordinateC)
                    
                End If
                NumberOfSelected = NumberOfSelected + 1
            End If
            DisplayProgress "Reading File Info..." + Strings.Format(100 * index / NumberOfImages, "0") + "%", RGB(0, &HC0, 0)
            DoEvents
        Next index
                
        If NumberOfSelected < 2 Then
            MsgBox "Select two or more time series Images!"
            DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
            GoTo Finish
        End If
'        GlobalUseChannelColor = True
        If TimeStampMax = 0 Then
            TimeStampMax = 1
        End If
        For index = 1 To NumberOfSelected
            TimeSort(index) = (TimeStart(index) - TimeStartMin) * 24 * 3600 + TimeStamp(index) / TimeStampMax / 10 ^ 3
        Next index
        
        ReDim IndexArray(NumberOfSelected)
        Heapsort TimeSort, NumberOfSelected, IndexArray
        
        Td = 0
        PreviousTime = 0
        DisplayProgress "Copying Files...", RGB(0, &HC0, 0)
        
        For index = 1 To NumberOfSelected
            If GlobalFileSource = 0 Then
                Set SourceImageNodeDocument = Nodes(IndexArray(index))
                Set SourceImage = SourceImageNodeDocument.Image(0)
            Else
            
'                Set SourceImage = New AimImage
                Set SourceImage = Lsm5.CreateObject("AimImage.Image")
                Import.filename = SlctFileName(IndexArray(index))
                Import.ReadFullSizeFileInformation SourceImage
                If Import.FileInfoChannelDataType(0) = eAimImageDataTypeU8 Then
                    DataType = 1
                Else
                    DataType = 2
                End If
                SourceImage.ImageMemory.Create Import.FileInfoSize(eAimImportExportCoordinateC), _
                                                    Import.FileInfoSize(eAimImportExportCoordinateT), _
                                                    Import.FileInfoSize(eAimImportExportCoordinateZ), _
                                                    Import.FileInfoSize(eAimImportExportCoordinateY), _
                                                    Import.FileInfoSize(eAimImportExportCoordinateX), _
                                                    Import.FileInfoChannelDataType(0)
                                                    
                Import.Import SourceImage
            
          
            End If
            
            If index = 1 Then
            
                                       
                 If Not MakeDestination(DestinationImageDocument, _
                                       DestinationImage, _
                                       SizeX, _
                                       SizeY, _
                                       SizeZ, _
                                       SizeT, _
                                       SizeC, _
                                       SourceImage.ImageMemory.GetDataType(0)) Then GoTo Finish
                                       
            
                ImageCopy.SourceImage = SourceImage
                ImageCopy.DestinationImage = DestinationImage
                ImageCopy.ImageParameterCopyFlags = eAimImageParameterCopyAll
            
                ImageCopy.CreateDestinationMemory eAimImageDataTypeInvalid, _
                                                  True
                                                  
                DestinationImage.ImageMemory.Resize SizeT, _
                                                    SizeZ, _
                                                    SizeY, _
                                                    SizeX, _
                                                    eAimImageResizeTypePreserve
                DestinationImageDocument.Name = Name2
                ImageCopy.Start
                If Not WaitProgress(ImageCopy) Then GoTo Finish
            Else
                ImageCopy.SourceImage = SourceImage
                ImageCopy.DestinationImage = DestinationImage
                ImageCopy.DestinationStart(eAimImageOperationCoordinateT) = Td

                ImageCopy.ImageParameterCopyFlags = eAimImageParameterCopyTimeStamps ' Or eAimImageParameterCopyEventList
                If GlobalUseChannelColor And GlobalSystemVersion >= 50 Then
                    For Cd = 0 To DestinationImage.ImageMemory.GetDimensionC - 1
                        For Cs = 0 To SourceImage.ImageMemory.GetDimensionC - 1
                            If SourceImage.DisplayParameters.ChannelInformation.ChannelColor(Cs) = _
                                       DestinationImage.DisplayParameters.ChannelInformation.ChannelColor(Cd) _
                                       Then
                                ImageCopy.DestinationStart(eAimImageOperationCoordinateC) = Cd
                                ImageCopy.SourceStart(eAimImageOperationCoordinateC) = Cs
                                ImageCopy.Size(eAimImageOperationCoordinateC) = Cd + 1
                                ImageCopy.Start
                                If Not WaitProgress(ImageCopy) Then GoTo Finish
                            End If
                        Next Cs
                    Next Cd
                Else
                    For Cd = 0 To DestinationImage.ImageMemory.GetDimensionC - 1
                        For Cs = 0 To SourceImage.ImageMemory.GetDimensionC - 1
                            If StrComp(SourceImage.DisplayParameters.ChannelInformation.ChannelName(Cs), _
                                       DestinationImage.DisplayParameters.ChannelInformation.ChannelName(Cd), _
                                       vbTextCompare) = 0 Then
                                ImageCopy.DestinationStart(eAimImageOperationCoordinateC) = Cd
                                ImageCopy.SourceStart(eAimImageOperationCoordinateC) = Cs
                                ImageCopy.Size(eAimImageOperationCoordinateC) = Cd + 1
                                ImageCopy.Start
                                If Not WaitProgress(ImageCopy) Then GoTo Finish
                            End If
                        Next Cs
                    Next Cd
                
                End If
            End If
            
            If index > 1 Then
                TimeDifference = CDbl((TimeStart(IndexArray(index)) - TimeStart(IndexArray(1))) * 24 * 3600)
                TimeSystemStart = TimeDifference - SourceImage.ImageMemory.TimeStamp(0)
                If TimeSystemStart > TimeSystemStartLast + 2 Then TimeSystemStartLast = TimeSystemStart
                
            Else
                TimeDifference = 0
                PreviousTime = SourceImage.ImageMemory.TimeStamp(0)
                TimeSystemStart = TimeDifference - SourceImage.ImageMemory.TimeStamp(0)
                TimeSystemStartLast = TimeDifference - SourceImage.ImageMemory.TimeStamp(0)
                
            End If
            
            For Ts = 0 To SourceImage.ImageMemory.GetDimensionT - 1
                TmpTimeStamp = SourceImage.ImageMemory.TimeStamp(Ts)
                DestinationImage.ImageMemory.TimeStamp(Td) = TmpTimeStamp _
                                                           + TimeSystemStartLast
                Td = Td + 1
            Next Ts
            
            Count = SourceImage.EventList.Count
            For Es = 0 To Count - 1
                EventTimeStamp = SourceImage.EventList.Time(Es)
                EventType = SourceImage.EventList.Type(Es)
                EventDescription = SourceImage.EventList.Description(Es)
                DestinationImage.EventList.Append EventTimeStamp + TimeSystemStartLast, EventType, EventDescription
            Next Es
        Next index
        DestinationImageDocument.Name = Name2
        
        Lsm5Vba.Application.ThrowEvent eRootReuse, 0
        DoEvents
    End If
    DoConcatenate_Time = True
Finish:
    flgBreak = False
    User_flg = True
End Function


Public Function SingleImage() As DsRecordingDoc
    Dim ChNum As Long
    Dim bpp As Long
    Dim Tags As AimImage40.AimImageApplicationTags
    Dim OtherTags As AimImage40.AimImageApplicationTags
    Dim ViewerGuiServer As AimViewerGuiServer40.AimViewerGuiServer
    Dim Tree As AimExperiment40.AimExperimentTreeNode
    Dim Node As AimExperiment40.AimExperimentTreeNode
    Dim GlobalSingleImageValid As Boolean
    Dim index As Long
    
    GlobalSingleImageValid = False
    
    If Not GlobalSingleImage Is Nothing Then
        If GlobalSystemVersion < 50 Then
            GlobalSingleImageValid = GlobalSingleImage.IsValid
        Else
            If Not EngelImageToHechtImage(GlobalSingleImage) Is Nothing Then
                Set Tags = EngelImageToHechtImage(GlobalSingleImage).Image(0, False)
                If Not Tags Is Nothing Then
                    Set ViewerGuiServer = Lsm5.ViewerGuiServer
                    If Not ViewerGuiServer Is Nothing Then
                        Set Tree = ViewerGuiServer.ExperimentTree
                        If Not Tree Is Nothing Then
                            For index = 0 To Tree.NumberChildren - 1
                                Set Node = Tree.Child(index)
                                If Not Node Is Nothing Then
                                    If Node.Type = eExperimentTeeeNodeTypeLsm Then
                                        Set OtherTags = Node.Image(0)
                                        If Not OtherTags Is Nothing Then
                                            Tags.SetBooleanValue "MultiTimeImageTest", False
                                            If Not OtherTags.GetBooleanValue("MultiTimeImageTest", True) Then
                                                Tags.SetBooleanValue "MultiTimeImageTest", True
                                                If OtherTags.GetBooleanValue("MultiTimeImageTest", False) Then
                                                    GlobalSingleImageValid = True
                                                    Tags.Remove "MultiTimeImageTest"
                                                    Exit For
                                                End If
                                            End If
                                            Tags.Remove "MultiTimeImageTest"
                                        End If
                                    End If
                                End If
                            Next index
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Not GlobalSingleImageValid Then
        ChNum = Lsm5.DsRecording.NumberOfChannels
        Set GlobalSingleImage = Lsm5.MakeNewImageDocument(512, 512, 1, 1, ChNum, bpp, True)
    End If
    Set SingleImage = GlobalSingleImage
End Function










Public Sub DisplayProgress(state As String, Color As Long)
    If (Color & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Or ((Color / 256) & &HFF) > 128 Then
        ProgressLabel.ForeColor = 0
    Else
        ProgressLabel.ForeColor = &HFFFFFF
    End If
    ProgressLabel.BackColor = Color
    ProgressLabel.Caption = state
    DoEvents
End Sub

Public Sub GetPureName(Path As String, Name As String)

    Dim pos As Integer
    Dim Start As Integer
    Dim bslash As String
    Dim lngth As String
    Dim dot As String
    Dim tmpName As String
    
        Start = 1
        bslash = "\"
        dot = "."
        pos = Start
        Do While pos > 0
            pos = InStr(Start, Path, bslash)
            If pos > 0 Then
                Start = pos + 1
            End If
        Loop
        lngth = Strings.Len(Path)
        tmpName = Strings.Right(Path, lngth - Start + 1)
        Start = 1
        pos = Start
        Do While pos > 0
            pos = InStr(Start, tmpName, dot)
            If pos > 0 Then
                Start = pos + 1
            End If
        Loop
        Name = Strings.Left(tmpName, Start - 2)
        
    End Sub

Public Function LoadWindowPosition()
    Dim PosKey As String
    
    PosKey = Lsm5.tools.GetWindowPositionKey() + "\" + Caption
    Left = Lsm5.tools.RegLongValue(PosKey, "Left")
    Top = Lsm5.tools.RegLongValue(PosKey, "Top")
    If Left < 1 Then Left = 0
    If Top < 1 Then Top = 0
    
    If Left = 0 And Top = 0 Then
                'Center frm
                Left = 300
                Top = 300
'    SaveWindowPosition
                Exit Function
    End If
End Function



Public Sub SaveWindowPosition()
    Dim PosKey As String
    
    PosKey = Lsm5.tools.GetWindowPositionKey() + "\" + Caption
    Lsm5.tools.RegLongValue(PosKey, "Left") = CInt(Left)
    Lsm5.tools.RegLongValue(PosKey, "Top") = CInt(Top)
End Sub





' m1tle mitosys modification
Private Sub SelectLocationTextBox_Change()
    EnsureOnlyNumbers
End Sub

' m1tle mitosys modification
Private Sub EnsureOnlyNumbers()
    If TypeName(Me.ActiveControl) = "TextBox" Then
        With Me.ActiveControl
            If Not IsNumeric(.Value) And .Value <> vbNullString Then
                MsgBox "Sorry, only numbers allowed"
                .Value = vbNullString
            End If
        End With
    End If
End Sub




''''
'   CreateListToConcatenate()
'   Reads file from list and concatenate them according to their end index
''''
Private Function CreateListToConcatenate() As ImageName()

    Dim Images() As ImageName  'this array contains the base-name (first entry) and all files belonging to it
    ReDim Images(0)
    Dim iPattern As Integer
    Images(0).BaseName = ""
    Dim Match As MatchCollection
    Dim index As Integer
    Dim i As Integer
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp") ' an object to do regular expression operations
    Dim List
    List = ImagesListBox.List
    'check through the full list of image files

    For index = 0 To ImagesListBox.ListCount - 1
        'We can use different pattern to match the files
        ' try to match Pattern for file naming used by AutofocusScreen
        RegEx.Pattern = Pattern(0)
        For iPattern = 0 To UBound(Pattern, 1) - 1
            RegEx.Pattern = Pattern(iPattern)
            Set Match = RegEx.Execute(ImagesListBox.List(index, 0))
            If Match.Count > 0 Then
                Exit For
            Else
                GoTo NextPattern
            End If
NextPattern:
        Next iPattern
        If Match.Count = 0 Then
            GoTo NextIndex
        Else
                        
            If Images(0).BaseName = "" Then 'initialize the Images
                ReDim Images(0)
                Images(0).BaseName = Match.Item(0).SubMatches.Item(0)
                Images(0).Suffix = Match.Item(0).SubMatches.Item(2)
                ReDim Images(0).ListOfNames(0)
                Images(0).ListOfNames(0) = ImagesListBox.List(index, 0)
            Else
                RegEx.Pattern = "^" & Match.Item(0).SubMatches.Item(0)
                For i = 0 To UBound(Images)
                    If RegEx.Test(Images(i).BaseName) And Right(ImagesListBox.List(index, 0), 3) = Images(i).Suffix Then
                        ReDim Preserve Images(i).ListOfNames(UBound(Images(i).ListOfNames) + 1)
                        Images(i).ListOfNames(UBound(Images(i).ListOfNames)) = ImagesListBox.List(index, 0)
                        GoTo NextIndex
                    End If
                Next i
                ReDim Preserve Images(UBound(Images) + 1)
                ' if I am here no matches was found a create a new BaseName
                Images(UBound(Images)).BaseName = Match.Item(0).SubMatches.Item(0)
                Images(UBound(Images)).Suffix = Match.Item(0).SubMatches.Item(2)
                ReDim Images(i).ListOfNames(0)
                Images(UBound(Images)).ListOfNames(0) = ImagesListBox.List(index, 0)
            End If
        End If
NextIndex:
    Next index
    CreateListToConcatenate = Images
End Function

Private Sub SelectLocationNew(loc As Integer, Images() As ImageName)
    Dim index As Integer
    Dim indexI As Integer
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp") ' an object to do regular expression operations
    For index = 0 To ImagesListBox.ListCount - 1
        ImagesListBox.Selected(index) = False
        RegEx.Pattern = "^" & Images(loc).BaseName
        User_flg = False
        If RegEx.Test(ImagesListBox.List(index, 0)) Then
            For indexI = 0 To UBound(Images(loc).ListOfNames)
                RegEx.Pattern = Images(loc).ListOfNames(indexI)
                If RegEx.Test(ImagesListBox.List(index, 0)) Then
                    ImagesListBox.Selected(index) = True
                End If
            Next indexI
        End If
    Next index
    User_flg = True
End Sub

Private Sub SelectLocation(loc As Integer)
    Dim indexL As Integer
    Dim indexL2 As Integer
    Dim index As Integer
    For index = 0 To ImagesListBox.ListCount - 1
        indexL = InStr(1, ImagesListBox.List(index, 0), "_L")
        indexL2 = InStr(indexL + 1, ImagesListBox.List(index, 0), "_")
        ImagesListBox.Selected(index) = False
        If Not ((indexL = 0) Or (indexL2 = 0)) Then
            If loc = Val(Mid(ImagesListBox.List(index, 0), indexL + 2, indexL2 - indexL - 2)) Then
                ImagesListBox.Selected(index) = True
            End If
        End If
    Next index
End Sub

'' m1tle mitosys modification
Private Sub ConcatenateTimePerLocationButton_Click()
    Set colFolders = New Collection
    Dim Folder As Variant
    mainDir = FileNameTextBox
    colFolders.Add mainDir
    If Recursive Then
        ListDirectories colFolders, FileNameTextBox
    End If
    
    For Each Folder In colFolders
        FileNameTextBox = Folder
        FillListBox
        If Not ConcatenateTimePerLocation Then
            GoTo Abort
        End If
    Next Folder
Abort:
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
End Sub


' m1tle mitosys modification
Private Function ConcatenateTimePerLocation() As Boolean
    Dim Images() As ImageName
    Dim fileformat As enumAimExportFormat
    Images = CreateListToConcatenate
    ' find minimum and maximum location number
    Dim outputfile As String
    Dim index As Integer
    If UBound(Images) = 0 And Images(0).BaseName = "" Then
        ConcatenateTimePerLocation = True
        Exit Function
    End If
    For index = 0 To UBound(Images)
        fileformat = eAimExportFormatLsm5
        outputfile = mainDir & Images(index).BaseName & ".lsm"
        If ZENv > 2010 Then
            If Images(index).Suffix = "czi" Then
                fileformat = 42
                outputfile = mainDir & Images(index).BaseName & "." & Images(index).Suffix
            End If
        End If
            
        If UBound(Images(index).ListOfNames) > 0 Then
            DoEvents
            SelectLocationNew index, Images
            If Not DoConcatenate_Time(Images(index).BaseName) Then
                Exit Function
            End If
            DisplayProgress "Saving File ...", RGB(0, &HC0, 0)
                
            SaveDsRecordingDoc Lsm5.DsRecordingActiveDocObject, outputfile, fileformat
            Lsm5.DsRecordingActiveDocObject.CloseAllWindows
            
        End If
    Next index
    ConcatenateTimePerLocation = True
    Exit Function
End Function


Private Sub ConcatenatetimeMarkedLocation_Click()
    Dim Images() As ImageName
    Dim fileformat As enumAimExportFormat
    Images = CreateListToConcatenate
    ' find minimum and maximum location number
    Dim outputfile As String
    Dim index As Integer
    Dim indexI As Integer
    Dim indexList() As Boolean
    ReDim indexList(0 To UBound(Images, 1))
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp") ' an object to do regular expression operations
    'find index that correspond to marked location
    For indexI = 0 To ImagesListBox.ListCount - 1
        If ImagesListBox.Selected(indexI) Then
            For index = 0 To UBound(Images, 1)
                RegEx.Pattern = "^" & Images(index).BaseName
                If RegEx.Test(ImagesListBox.List(indexI, 0)) Then
                    indexList(index) = True
                    Exit For
                End If
            Next index
        End If
    Next indexI
            
    For index = 0 To UBound(Images, 1)
        If indexList(index) Then
            fileformat = eAimExportFormatLsm5
            outputfile = mainDir & Images(index).BaseName & ".lsm"
            If ZENv > 2010 Then
                If Images(index).Suffix = "czi" Then
                    fileformat = 42
                    outputfile = mainDir & Images(index).BaseName & "." & Images(index).Suffix
                End If
            End If
            If UBound(Images(index).ListOfNames) > 0 Then
                SelectLocationNew index, Images
                If Not DoConcatenate_Time(Images(index).BaseName) Then
                    GoTo Abort
                End If
                SaveDsRecordingDoc Lsm5.DsRecordingActiveDocObject, outputfile, fileformat
            End If
        End If
    Next index
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
    Exit Sub
Abort:
    DisplayProgress "Ready", RGB(&HC0, &HC0, 0)
End Sub

' Copied and adapted from MultiTimeSeries macro
Public Function SaveDsRecordingDoc(Document As DsRecordingDoc, filename As String, fileformat As enumAimExportFormat) As Boolean
    Dim Export As AimImageExport
    Dim Image As AimImageMemory
    Dim Error As AimError
    Dim Planes As Long
    Dim Plane As Long
    Dim Horizontal As enumAimImportExportCoordinate
    Dim Vertical As enumAimImportExportCoordinate

    On Error GoTo Done

    'Set Image = EngelImageToHechtImage(Document).Image(0, True)
    If Not Document Is Nothing Then
        Set Image = Document.RecordingDocument.Image(0, True)
    End If
    
    Set Export = Lsm5.CreateObject("AimImageImportExport.Export.4.5")
'    Set Export = New AimImageExport
    Export.filename = filename
    Export.Format = fileformat
    Export.StartExport Image, Image
    Set Error = Export
    Error.LastErrorMessage
    
    Planes = 1
    Export.GetPlaneDimensions Horizontal, Vertical
    
    Select Case Vertical
        Case eAimImportExportCoordinateY:
             Planes = Image.GetDimensionZ * Image.GetDimensionT
        Case eAimImportExportCoordinateZ:
            Planes = Image.GetDimensionT
    End Select
    
    For Plane = 0 To Planes - 1
        DoEvents
        Export.ExportPlane Nothing
    Next Plane
    Export.FinishExport
    SaveDsRecordingDoc = True
    Exit Function
Done:
    MsgBox "Check Temporary Files Folder! Cannot Save Temporary File(s)!"
    SaveDsRecordingDoc = False
    Export.FinishExport
    
End Function
