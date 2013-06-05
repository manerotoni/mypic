VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StoreApply 
   Caption         =   "Store/Apply Grid"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   OleObjectBlob   =   "StoreApply.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StoreApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Form was kept to keep the possibility of saving the macro setting into the registry.
' It is not used in current state of the macro





Private Sub CommandButtonClose_Click()
    Unload StoreApplyForm
End Sub


Sub FillRecipeList()
    Dim Count As Long
    Dim idx As Long
    Dim MyMacroKey As String
    Dim path As String
    Dim success As Boolean
    Dim result As Long
    Dim key As String
    ComboBoxRecipe.Clear
    key = "UI\" + GlobalMacroKey
    success = tools.RegExistKey(key)
    If success Then
        Count = tools.RegCountSubKeys(key)
    Else
        success = tools.RegCreateKey(key)
    End If
    If Count > 0 Then
        For idx = 0 To Count - 1
            ComboBoxRecipe.AddItem tools.RegSubkeyName(idx, key)
        Next idx
        If (Count > 0) Then
            ComboBoxRecipe.ListIndex = 0
        End If
        ComboBoxRecipe.ListIndex = 0
    End If
End Sub


Private Sub CommandButtonDelete_Click()
    Dim myKey As String
    Dim success As Boolean
    Dim deleteOK As Boolean
    Dim idx As Long
    Dim lockNo As Long
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    AutofocusForm.GetBlockValues
    deleteOK = True
    myKey = "UI\" + GlobalMacroKey + "\" + ComboBoxRecipe.Value
    success = tools.RegExistKey(myKey)
    If success Then
        Msg = "Do You want to delete recipe " + ComboBoxRecipe.Value + "?"
        Style = VbYesNo + VbCritical + VbDefaultButton2 ' Define buttons.
        Title = "Recipe Delete"  ' Define title.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
        Else    ' User chose No.
            deleteOK = False
        End If
    Else
    End If
    If deleteOK Then
        success = tools.RegDeleteKey(myKey)
    End If
    Unload StoreApply
End Sub


Private Sub CommandButtonStore_Click()
Dim filenam As String
Dim path As String
Dim hFile As Long
Dim fs
Dim idx As Long
Dim idy As Long


path = "c:\AIM\macros\datafiles\"
filnam = path + CStr(ComboBoxGrid.Value) + ".txt"
hFile = FreeFile

Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile filnam          'Create a file
Open filnam For Output As hFile         'open file
Print #hFile, TextBoxXGrid.Value
Print #hFile, TextBoxYGrid.Value
Print #hFile, TextBoxXStep.Value
Print #hFile, TextBoxYGrid.Value

For idy = 1 To GlobalYGrid
     For idx = 1 To GlobalXGrid
                       
        Print #hFile, GlobalDeActivatedLocations(idx, idy)
                                
     Next idx
Next idy
Print #hFile, AutofocusForm.CheckBoxMeander.Value
   
End Sub


Private Sub UserForm_Initialize()
    FillRecipeList
End Sub
