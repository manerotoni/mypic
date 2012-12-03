Attribute VB_Name = "ExportMacro"
'Sub ExportVBAFiles()
'  Dim pVBAProject As Lsm5.L
'  Dim vbComp As VBComponent  'VBA module, form, etc...
'  Dim strDocPath As String   'Current document path
'  Dim strSavePath As String  'Path to save the exported files to
'
'  ' strSavePath will be the pathname of the document with a _VBACode suffix
'  ' If you want to export the code for Normal instead, change the following
'  ' line to:
'  ' strDocPath = Application.Templates.Item(0)
'  strDocPath = Application.Templates.Item(Application.Templates.Count - 1)
'
'  strSavePath = Left(strDocPath, Len(strDocPath) - 4)
'  strSavePath = strSavePath & "_VBACode"
'
'  ' If this folder doesn't exist, create it
'  If Dir(strSavePath, vbDirectory) = "" Then
'    MkDir strSavePath
'  End If
'
'  ' Get the VBA project
'  ' If you want to export code for Normal instead, paste this macro into
'  ' ThisDocument in the Normal VBA project and change the following line to:
'  ' Set pVBAProject = ThisDocument.VBProject
'  Set pVBAProject = Application.Document.VBProject
'
'  ' Loop through all the components (modules, forms, etc) in the VBA project
'  For Each vbComp In pVBAProject.VBComponents
'    Select Case vbComp.Type
'    Case vbext_ct_StdModule
'      vbComp.Export strSavePath & "\" & vbComp.name & ".bas"
'    Case vbext_ct_Document, vbext_ct_ClassModule
'      ' ThisDocument and class modules
'      vbComp.Export strSavePath & "\" & vbComp.name & ".cls"
'    Case vbext_ct_MSForm
'      vbComp.Export strSavePath & "\" & vbComp.name & ".frm"
'    Case Else
'      vbComp.Export strSavePath & "\" & vbComp.name
'    End Select
'  Next
'    MsgBox "VBA files have been exported to: " & strSavePath
'End Sub
'
