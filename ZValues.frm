VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZValues 
   Caption         =   "ZValues"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   OleObjectBlob   =   "ZValues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ZValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
GlobalStageControlZValues = True
ZValues.Hide
End Sub

Private Sub CommandButton2_Click()
GlobalStageControlZValues = False
ZValues.Hide
End Sub
