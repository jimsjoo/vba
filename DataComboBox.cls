VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// define DataComboBox's an event, 'ItemAdded'
Public Event ItemAdded(strValue As String, blnCancel As Boolean)

'// p_cboCombo is a variable that refers to Userform's combobox
Private p_cboCombo As ComboBox

Public Property Set ComboBox(cboCombo As ComboBox)
'// reference to Userform's combobox
    If p_cboCombo Is Nothing Then
        Set p_cboCombo = cboCombo
    End If
End Property

Public Property Get ComboBox() As ComboBox
'// return combox
    Set ComboBox = p_cboCombo
End Property

Public Function AddDataItem(strValue As String)
'// blnCancel is a variable that determine
'// whether strValue can be added

    Dim blnCancel As Boolean
    
    blnCancel = False
    
'// Trigger User-defined Event, ItemAdded
    RaiseEvent ItemAdded(strValue, blnCancel)

'// if blnCancel is True, strValue not added
    If blnCancel = False Then
        Me.ComboBox.AddItem strValue
    End If
End Function
