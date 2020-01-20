Attribute VB_Name = "Module4"
Option Explicit

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Function ExcelInstances() As Long
    Dim hWndDesk As Long
    Dim hWndXL As Long

    'Get a handle to the desktop
    hWndDesk = GetDesktopWindow
    Do
        'Get the next Excel window
        hWndXL = FindWindowEx(GetDesktopWindow, hWndXL, _
          "XLMAIN", vbNullString)

        'If we got one, increment the count
        If hWndXL > 0 Then
            ExcelInstances = ExcelInstances + 1
        End If

        'Loop until we've found them all
    Loop Until hWndXL = 0
End Function

Sub demoCountExcelINstance()
    Debug.Print ExcelInstances()
End Sub
