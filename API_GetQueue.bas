Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function GetQueueStatus _
    Lib "user32" (ByVal fuFlags As Long) As Long
Private Const QS_KEY = &H1
Private Const QS_MOUSEMOVE = &H2
Private Const QS_MOUSEBUTTON = &H4
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)

Private Sub demoDoEventsWithAPI()
    Dim i           As Long
    Dim BeginTime   As Double
    Dim EndTime     As Double
    
    BeginTime = Timer
    
    For i = 1 To 1000000
        If GetQueueStatus(QS_INPUT) <> 0 Then DoEvents
    Next
    
    EndTime = Timer
    
    Debug.Print EndTime - BeginTime
End Sub


