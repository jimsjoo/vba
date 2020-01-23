Attribute VB_Name = "FileBrowser"
Option Explicit

Type thOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As String
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare Function th_apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As thOPENFILENAME) As Boolean
Declare Function th_apiGetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As thOPENFILENAME) As Boolean
Declare Function CommDlgExtendetError Lib "commdlg32.dll" () As Long

Private Const thOFN_READONLY = &H1
Private Const thOFN_OVERWRITEPROMPT = &H2
Private Const thOFN_HIDEREADONLY = &H4
Private Const thOFN_NOCHANGEDIR = &H8
Private Const thOFN_SHOWHELP = &H10
Private Const thOFN_NOVALIDATE = &H100
Private Const thOFN_ALLOWMULTISELECT = &H200
Private Const thOFN_EXTENSIONDIFFERENT = &H400
Private Const thOFN_PATHMUSTEXIST = &H800
Private Const thOFN_FILEMUSTEXIST = &H1000
Private Const thOFN_CREATEPROMPT = &H2000
Private Const thOFN_SHAREWARE = &H4000
Private Const thOFN_NOREADONLYRETURN = &H8000
Private Const thOFN_NOTESTFILECREATE = &H10000
Private Const thOFN_NONETWORKBUTTON = &H20000
Private Const thOFN_NOLONGGAMES = &H40000
Private Const thOFN_EXPLORER = &H80000
Private Const thOFN_NODEREFERENCELINKS = &H100000
Private Const thOFN_LONGNAMES = &H200000

Function StartIt()
    Dim strFilter As String
    Dim lngFlags As Long
    strFilter = thAddFilterItem(strFilter, "Excel Files (*.xls)", "*.XLS")
    strFilter = thAddFilterItem(strFilter, "Text Files(*.txt)", "*.TXT")
    strFilter = thAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    Startform.filenameinput.Value = thCommonFileOpenSave(InitialDir:="C:\Windows", Filter:=strFilter, FilterIndex:=3, Flags:=lngFlags, DialogTitle:="File Browser")
    Debug.Print Hex(lngFlags)
End Function

Function GetOpenFile(Optional varDirectory As Variant, Optional varTitleForDialog As Variant) As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    Dim varFileName As Variant
    lngFlags = thOFN_FILEMUSTEXIST Or thOFN_HIDEREADONLY Or thOFN_NOCHANGEDIR
    
    If IsMissing(varDirectory) Then varDirectory = ""
    End If
    
    If IsMissing(varTitleForDialog) Then varTitleForDialog = ""
    End If
    
    strFilter = thAddFilterItem(strFilter, "Excel (*.xls)", "*.XLS")
    varFileName = thCommonFileOpenSave(OpenFile:=True, InitialDir:=varDirectory, Filter:=strFilter, Flags:=lngFlags, DialogTitle:=varTitleForDialog)
    
    If Not IsNull(varFileName) Then varFileName = TrimNull(varFileName)
    End If
    
    GetOpenFile = varFileName
    
End Function

Function thCommonFileOpenSave(Optional ByRef Flags As Variant, Optional ByVal InitialDir As Variant, Optional ByVal Filter As Variant, _
                               Optional ByVal FilterIndex As Variant, Optional ByVal DefaultEx As Variant, Optional ByVal fileName As Variant, _
                               Optional ByVal DialogTitle As Variant, Optional ByVal hwnd As Variant, Optional ByVal OpenFile As Variant) As Variant
                               
    Dim OFN As thOPENFILENAME
    Dim strFileName As String
    Dim FileTitle As String
    Dim fResult As Boolean
            
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultEx) Then DefaultEx = ""
    If IsMissing(fileName) Then fileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(hwnd) Then hwnd = 0
    If IsMissing(OpenFile) Then OpenFile = True
    
    strFileName = Left(fileName & String(256, 0), 256)
    FileTitle = String(256, 0)
    
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = FileTitle
        .nMaxFileTitle = Len(FileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultEx
        .strInitialDir = InitialDir
        .hInstance = 0
        .lpfnHook = 0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
        
    If OpenFile Then fResult = th_apiGetOpenFileName(OFN) Else fResult = th_apiGetSaveFileName(OFN)
    
        
    If fResult Then
        If Not IsMissing(Flags) Then Flags = OFN.Flags
        thCommonFileOpenSave = TrimNull(OFN.strFile)
        Else
        thCommonFileOpenSave = vbNullString
    End If
        
End Function

Function thAddFilterItem(strFilter As String, strDescription As String, Optional varItem As Variant) As String

    If IsMissing(varItem) Then varItem = "*.*"
    thAddFilterItem = strFilter & strDescription & vbNullChar & varItem & vbNullChar
    
End Function

Private Function TrimNull(ByVal strItem As String) As String
    Dim intPos As Integer
    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = Left(strItem, intPos - 1)
        Else
        TrimNull = strItem
    End If
        
End Function
