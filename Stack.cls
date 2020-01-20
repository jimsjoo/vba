VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' From "VBA Developer's Handbook"
' by Ken Getz and Mike Gilbert
' Copyright 1997; Sybex, Inc. All rights reserved.

' Stack class.

Dim Top As StackItem

Public Function Pop() As Variant
    If Not IsEmpty Then
        ' Get the value from the current top stack element.
        ' Then, get a reference to the new stack top.
        Pop = Top.Value
        Set Top = Top.nItem
    End If
End Function

Public Sub Push(ByVal varText As Variant)
    ' Add a new item to the top of the stack.
    Dim NewTop As New StackItem
        
    NewTop.Value = varText
    Set NewTop.nItem = Top
    Set Top = NewTop
End Sub

Property Get GetTop() As Variant
    'GetTop = IIf(IsEmpty, vbNull, Top.Value)
    If IsEmpty Then
        GetTop = Null
    Else
        GetTop = Top.Value
    End If
End Property

Property Get IsEmpty() As Boolean
    ' Is the stack empty?  It can
    ' only be empty if Top is Nothing.
    'On Error Resume Next
    IsEmpty = (Top Is Nothing)
End Property

Private Sub Class_Initialize()
    Set Top = Nothing
End Sub

Private Sub Class_Terminate()
    Set Top = Nothing
End Sub

