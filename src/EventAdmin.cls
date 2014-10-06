VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Ready()
Public Event Busy(Task As Integer)
Private isReady As Boolean

Public Sub setReady(Optional Force As Boolean)
    If Not isReady Or Force Then
        isReady = True
        RaiseEvent Ready
    End If
End Sub

Public Sub setBusy(Task As Integer, Optional Force As Boolean)
    If isReady Or Force Then
        RaiseEvent Busy(Task)
        isReady = False
    End If
End Sub

Public Sub initialize()
    isReady = True
End Sub
