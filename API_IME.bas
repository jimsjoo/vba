Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function ImmGetContext _
    Lib "imm32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ImmSetConversionStatus _
    Lib "imm32.dll" (ByVal himc As Long, _
        ByVal dw1 As Long, ByVal dw2 As Long) As Long
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

Private Const IME_CMODE_NATIVE = &H1
Private Const IME_CMODE_HANGEUL = IME_CMODE_NATIVE
Private Const IME_CMODE_ALPHANUMERIC = &H0
Private Const IME_SMODE_NONE = &H0

Private Function GetAppHandle() As Long
    Dim dVersion    As Double
    
    dVersion = Application.Version
    
    If dVersion < 10 Then
        GetAppHandle = FindWindow("XLMAIN", _
            Application.Caption)
    Else
        GetAppHandle = Application.hwnd
    End If
End Function

Private Sub demoAPI()
    Dim hWndIME     As Long
    Dim hWndApp     As Long

    '// get Excel's handle
    hWndApp = GetAppHandle
    
    '// get IME's handle
    hWndIME = ImmGetContext(hWndApp)
    
    '// change it into English IME
    Call ImmSetConversionStatus(hWndIME, _
        IME_CMODE_ALPHANUMERIC, IME_SMODE_NONE)
    InputBox "English IME"
    
    '// change it into Korean IME
    Call ImmSetConversionStatus(hWndIME, _
        IME_CMODE_HANGEUL, IME_SMODE_NONE)
    InputBox "Korean IME"
    
End Sub

