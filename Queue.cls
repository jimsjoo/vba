VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim qFront As QueueItem
Dim qRear As QueueItem

Public Sub Enqueue(newItem As Variant)
    Dim qNew As New QueueItem
    
    qNew.Value = newItem
    
    If IsEmpty Then
        Set qFront = qNew
        Set qRear = qNew
    Else
        Set qRear.nItem = qNew
        Set qRear = qNew
    End If
End Sub

Public Function Dequeue() As Variant
    If IsEmpty Then
        Dequeue = Null
    Else
        Dequeue = qFront.Value

        If qFront Is qRear Then
            Set qFront = Nothing
            Set qRear = Nothing
        Else
            Set qFront = qFront.nItem
        End If
    End If
End Function

Property Get IsEmpty() As Boolean
    IsEmpty = ((qFront Is Nothing) And (qRear Is Nothing))
End Property

Property Get Peek() As Variant
    If Not IsEmpty Then Peek = qFront.Value
End Property

Private Sub Class_Initialize()
    Set qFront = Nothing
    Set qRear = Nothing
End Sub

Private Sub Class_Terminate()
    Set qFront = Nothing
    Set qRear = Nothing
End Sub
