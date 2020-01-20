Attribute VB_Name = "Module3"
Option Explicit

#If VBA7 Then   'or: #If Win64 Then  'Win64=true, Win32=true, Win16= false
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Public Sub DetectExcel()    'This procedure detects a running Excel app and registers it
    Const WM_USER = 1024
    Dim hwnd As Long

    hwnd = FindWindow("XLMAIN", 0)  'If Excel is running this API call returns its handle
    If hwnd = 0 Then Exit Sub       '0 means Excel not running

    'Else Excel is running so use the SendMessage API function
    'to enter it in the Running Object Table

    SendMessage hwnd, WM_USER + 18, 0, 0
End Sub

Public Sub GetDataFromExternalXLInstanceAPI()
    Dim xlApp As Object
    Dim xlNotRunning As Boolean 'Flag for final reference release

    On Error Resume Next        'Check if Excel is already running; defer error trapping
        Set xlApp = GetObject(, "Excel.Application")    'If it's not running an error occurs
        xlNotRunning = (Err.Number <> 0)
        Err.Clear               'Clear Err object in case of error
    On Error GoTo 0             'Reset error trapping

    DetectExcel                 'If Excel is running enter it into the Running Object table
    Set xlApp = GetObject("C:\Users\JIMSJOO\Documents\ipo_sample.xlsx")      'Set object reference to the file

    'Show Excel through its Application property
    xlApp.Application.Visible = True
    'Show the actual window of the file using the Windows collection of the xlApp object ref
    xlApp.Parent.Windows(1).Visible = True

    '... Process file

    'If Excel was not running when this started, close it using the App's Quit method
    If xlNotRunning = True Then xlApp.Application.Quit
    Set xlApp = Nothing    'Release reference to the application and spreadsheet
End Sub
