Attribute VB_Name = "Module3"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)

Private Sub demoSleepAPI()
    
    Debug.Print ">Now time is ", Now
    Sleep 3000  '// Will pause for 3 seconds
'   Or you can use 'WaitSeconds 3'
    Debug.Print ">After using Sleep API function ", Now
End Sub

Sub WaitSeconds(waitTime As Integer)
    Application.Wait (Now + TimeValue("00:00:" _
        & Format(waitTime, "00")))
End Sub

