Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Const QS_KEY As Integer = 1
Const QS_MOUSEMOVE As Integer = 2
Const QS_MOUSEBUTTON As Integer = 4
Const QS_POSTMESSAGE As Integer = 8
Const QS_TIMER As Integer = 16
Const QS_PAinteger As Integer = 32
Const QS_SENDMESSAGE As Integer = 64
Const QS_HOTKEY As Integer = 128
Const QS_ALLPOSTMESSAGE As Integer = 256
Const QS_MOUSE As Integer = (QS_MOUSEMOVE _
    Or QS_MOUSEBUTTON)
Const QS_INPUT As Integer = (QS_MOUSE Or QS_KEY)
Const QS_ALLEVENTS As Integer = (QS_INPUT _
    Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAinteger Or QS_HOTKEY)
Const QS_ALLINPUT As Integer = (QS_SENDMESSAGE _
    Or QS_PAinteger Or QS_TIMER Or QS_POSTMESSAGE _
    Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
    
Sub DoEventsTester()
    Dim i           As Long
    Dim start       As Long
    Dim take1       As Single
    Dim take2       As Single
    
    start = Timer
    For i = 1 To 100000
        DoEvents
    Next
    take1 = Timer - start
    Debug.Print "With DoEvents by itself: "; take1; " seconds"
    
    start = Timer
    For i = 1 To 100000
        '// 'Message in the queue will be flushed
        If GetQueueStatus(QS_ALLINPUT) <> 0 Then DoEvents
    Next
    take2 = Timer - start
    Debug.Print "With GetQueueStatus check: "; take2; " seconds"
End Sub

