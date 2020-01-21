Attribute VB_Name = "Module4"
Option Explicit

Private Declare Function FindWindow _
    Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow _
    Lib "user32" (ByVal hwnd As Long, _
        ByVal nCmdShow As Long) As Long
    
Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9

Sub demoSizeExcelWindow()
    Dim hInst As Long
    hInst = FindWindow(vbNullString, Application.Caption)
    If Not IsNull(hInst) Then
        ShowWindow hInst, SW_MINIMIZE
    End If
End Sub


