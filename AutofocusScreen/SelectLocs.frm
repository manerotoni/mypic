VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectLocs 
   Caption         =   "Select/Deselct Locations"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   OleObjectBlob   =   "SelectLocs.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SelectLocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub DeselctButton_Click()
GridSelection x, y, XR, YR, False
SelectLocs.Hide
End Sub

Private Sub SelectButton_Click()
GridSelection x, y, XR, YR, True
SelectLocs.Hide
End Sub
