VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StoreApplyForm 
   Caption         =   "Store/Apply Grid"
   ClientHeight    =   1350
   ClientLeft      =   645
   ClientTop       =   930
   ClientWidth     =   5550
   OleObjectBlob   =   "StoreApplyForm.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "StoreApplyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 
Private Sub CommandButtonDelete_Click()
Dim x As Long

filnam = PubGridPathData + CStr(ComboBoxGrid.Value) + ".txt"
Msg = "Do You want to delete grid '" + ComboBoxGrid.Value + "' ?"
        Style = VbYesNo + VbCritical + VbDefaultButton2 ' Define buttons.
        Title = "Grid Delete"  ' Define title.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            Kill filnam
            x = ComboBoxGrid.ListIndex
            ComboBoxGrid.RemoveItem (x)
            ComboBoxGrid.Cut
            ComboBoxGrid.Value = ""
            FillRecipeList
        End If
End Sub

Private Sub UserForm_Initialize()
FilerQuery
FillRecipeList

End Sub


Private Sub CommandButtonApply_Click()
Dim filenam As String
Dim hFile As Long
Dim fs
Dim idx As Long
Dim idy As Long
Dim Value As String
filnam = PubGridPathData + CStr(ComboBoxGrid.Value) + ".txt"
hFile = FreeFile

Open filnam For Input As hFile
Input #hFile, GlobalXGrid
Input #hFile, GlobalYGrid
Input #hFile, GlobalXStep
Input #hFile, GlobalYStep


AutofocusForm.TextBoxXGrid.Value = GlobalXGrid
AutofocusForm.TextBoxYGrid.Value = GlobalYGrid
AutofocusForm.TextBoxXStep.Value = GlobalXStep
AutofocusForm.TextBoxYStep.Value = GlobalYStep

ReDim GlobalDeActivatedLocations(GlobalXGrid, GlobalYGrid)
For idy = 1 To GlobalYGrid
     For idx = 1 To GlobalXGrid
        Input #hFile, Value
         GlobalDeActivatedLocations(idx, idy) = CBool(Value)
     Next idx
Next idy
On Error Resume Next ' in older verios the meander varable wasnot saved to avoid, that you cannot use this files
Input #hFile, Value
GlobalMeander = CBool(Value)

AutofocusForm.CheckBoxMeander.Value = GlobalMeander
Close #hFile
AutofocusForm.ShowGrid
End Sub

Private Sub CommandButtonClose_Click()
   Unload StoreApplyForm
End Sub

Private Sub CommandButtonStore_Click()
Dim filenam As String
Dim filenew As String

Dim fso As New FileSystemObject
'Dim fold As Folder
'Set fold = fso.GetFolder(PubGridPath)


filenam = PubGridPathData + CStr(ComboBoxGrid.Value) + ".txt"
filenew = CStr(ComboBoxGrid.Value)

If fso.FileExists(filenam) = False Then
    CreateNewFolder filenam
Else
        Msg = "File '" + ComboBoxGrid.Value + "' already exists." + vbCrLf + "Do You want to overwrite '" + ComboBoxGrid.Value + "' ?"
        Style = VbYesNo + VbCritical + VbDefaultButton2 ' Define buttons.
        Title = "Grid Delete"  ' Define title.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            CreateNewFolder filenam
        Else
            Exit Sub
        End If
End If


FillRecipeList
MsgBox "Grid " + filenew + " is stored successfully!"
End Sub

Sub FillRecipeList()
    Dim Count As Long
    Dim idx As Long
    Dim filenam As String
    Dim path As String
    Dim Success As Boolean
    Dim result As Long
    Dim key As String
    ComboBoxGrid.Clear
    Dim fso As New FileSystemObject
    Dim fold As Folder
    Dim f As File
    Dim name As String
    Dim lenfilnam As Integer
    Dim i As Integer
    
    
    Set fold = fso.GetFolder(PubGridPathData)
     For Each f In fold.Files
        filenam = fso.GetFileName(f)
     lenfilnam = Len(filenam)
     name = Left(filenam, lenfilnam - 4)
     
      ComboBoxGrid.AddItem name
      Next
     i = ComboBoxGrid.ListCount
     If i = 0 Then Exit Sub
      ComboBoxGrid.Value = ComboBoxGrid.List(0)

End Sub

Private Sub FilerQuery()
Dim fso As New FileSystemObject
Dim fs
Dim fold As Folder
Set fold = fso.GetFolder(PubGridPath)
If fso.FolderExists(PubGridPathData) = False Then
    fso.CreateFolder (PubGridPathData)
End If

End Sub

Private Sub CreateNewFolder(filnam As String)
Dim hFile As Long
Dim fs
Dim idx As Long
Dim idy As Long
hFile = FreeFile
Set fs = CreateObject("Scripting.FileSystemObject")
        fs.CreateTextFile filnam          'Create a file
    Open filnam For Output As hFile         'open file
    Print #hFile, GlobalXGrid
    Print #hFile, GlobalYGrid
    Print #hFile, AutofocusForm.TextBoxXStep.Value
    Print #hFile, AutofocusForm.TextBoxYStep.Value
    For idy = 1 To GlobalYGrid
     For idx = 1 To GlobalXGrid
        Print #hFile, GlobalDeActivatedLocations(idx, idy)
     Next idx
Next idy
    Print #hFile, AutofocusForm.CheckBoxMeander.Value
Close #hFile
End Sub
