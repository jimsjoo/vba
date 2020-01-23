VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Startform 
   Caption         =   "Filebrowser Demo"
   ClientHeight    =   1896
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6204
   OleObjectBlob   =   "FileBrowser.frx":0000
   StartUpPosition =   2  '화면 가운데
End
Attribute VB_Name = "Startform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmainaction_Click()
        Workbooks.Open (filenameinput)
        Startform.Hide
End Sub

Private Sub cmdBrowseButton_Click()
    Call StartIt
End Sub

Private Sub UserForm_Click()

End Sub
