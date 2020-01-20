VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Value As Variant
Public nItem As StackItem

Private Sub Class_Initialize()
    Set nItem = Nothing
End Sub

Private Sub Class_Terminate()
    Set nItem = Nothing
End Sub


