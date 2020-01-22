VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// How to use this class
'Dim MyXLApp As New CApp
'
'Sub WrappingApp()
'    Set MyXLApp.App = Application
'End Sub
'
'Sub UnWrappingApp()
'    If Not (MyXLApp Is Nothing) Then Set MyXLApp = Nothing
'End Sub
'
'Sub demo_AddWorkbook()
'    MyXLApp.AddNewWorkbook
'End Sub

Dim WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Sub AddNewWorkbook()
    App.Workbooks.Add
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Excel.Workbook)
    Dim Wnd     As Window
    
    For Each Wnd In Application.Windows
        Wnd.WindowState = xlNormal
    Next
    Application.Windows.Arrange ArrangeStyle:=xlArrangeStyleTiled
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    MsgBox "A workbook Opened", vbInformation, "Workbook Opened"
End Sub

Private Sub Class_Initialize()
    Set App = Application
    App.Visible = True
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub
