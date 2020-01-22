VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataCombo 
   Caption         =   "Data Combo Box"
   ClientHeight    =   1890
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4260
   OleObjectBlob   =   "frmDataCombo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// declare DataComboBox with its event
Private WithEvents mdcbCombo As DataComboBox
Attribute mdcbCombo.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
    Dim strValue As String
    
    '// save data of combobox to strValue
    strValue = Me.cboData.Text
    
    '// call DataComboBox's AddDataItem method
    mdcbCombo.AddDataItem strValue
End Sub

Private Sub UserForm_Initialize()
    '// when this form begins,  make DataComboBox into instance
    Set mdcbCombo = New DataComboBox
    
    '// connect object  to frmDataCombo's combobox
    Set mdcbCombo.ComboBox = Me.cboData
End Sub

Private Sub mdcbCombo_ItemAdded(strValue As String, blnCancel As Boolean)
    Dim lngCount    As Long
    Dim lngLoop     As Long
    Dim strItem     As String
    Dim varItem     As Variant

    '// count of list
    lngCount = mdcbCombo.ComboBox.ListCount

    '// while the loop, detemine if new Item already added.
    '// if blnCancel is True, new item already exists
    For lngLoop = 1 To lngCount
        strItem = mdcbCombo.ComboBox.List(lngLoop - 1)
        If strItem = strValue Then
            blnCancel = True
            Exit For
        End If
    Next
End Sub

