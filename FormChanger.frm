VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTest 
   Caption         =   "Sizeable UserForm"
   ClientHeight    =   4125
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3960
   OleObjectBlob   =   "FormChanger.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************************
'*
'* MODULE NAME:     SIZEABLE USERFORM
'* AUTHOR:          STEPHEN BULLEN, Business Modelling Solutions Ltd.
'*                  TIM CLEM
'*
'* CONTACT:         Stephen@BMSLtd.co.uk
'* WEB SITE:        http://www.BMSLtd.co.uk
'*
'* DESCRIPTION:     Makes a userform resizeable and reacts to the form being resized
'*
'* THIS MODULE:     Handles the userform's resizing
'*
'* PROCEDURES:
'*  UserForm_Activate    Instantiates a class module to make the form sizeable
'*  UserForm_Resize      Sizes and positions the controls on the form when resized
'*  btnOK_Click          Closes the form when the OK button is clicked
'*
'***************************************************************************

Option Explicit
Option Compare Text

'Declare a new instance of our form changer
Dim oFormChanger As New CFormChanger

Private Sub btnChangeIcon_Click()
Dim vFile As Variant

vFile = Application.GetOpenFilename("Icon files (*.ico;*.exe;*.dll),*.ico;*.exe;*.dll", 0, "Open Icon File", "Open", False)

'Showing dialog sets the form modeless, so check it
oFormChanger.Modal = cbModal

If vFile = False Then Exit Sub

oFormChanger.IconPath = vFile

End Sub

Private Sub UserForm_Activate()

    'Initialise everything to show
    cbModal.Value = True
    cbCaption.Value = True
    cbCloseBtn.Value = True
    cbIcon.Value = True
    cbMaximize.Value = True
    cbMinimize.Value = True
    cbSizeable.Value = True
    cbSysmenu.Value = True
    cbTaskBar.Value = False
    cbSmallCaption.Value = False

    'Set the form changer to change this userform
    Set oFormChanger.Form = Me

    'Make sure everything is in the right place to start with
    UserForm_Resize

End Sub

Private Sub UserForm_Resize()

    Dim dFrameCols As Double, dFrameRows As Double, dFrameHeight As Double
    Dim i As Integer, j As Integer

    'Standard control gap of 6pts
    Const dGAP As Integer = 6

    'Exit the sub if we've been minimized
    If Me.InsideWidth = 0 Then Exit Sub
    
    'Set controls that don't move/size
    With lblMessage     'The position of the "Message" label
        .Top = dGAP
        .Left = dGAP
    End With

    With tbMessage      'The position of the message box (the size changes, not the position)
        .Top = dGAP + lblMessage.Height + dGAP
        .Left = dGAP
    End With

    fraStyle.Left = dGAP

    'Don't let the form get less than a certain height - must have at least the message and button
    If Me.InsideHeight < lblMessage.Height + btnOK.Height + fraStyle.Height + dGAP * 5 Then

        'Reset the height, allowing for the form's border (Height - InsideHeight)
        Me.Height = lblMessage.Height + btnOK.Height + fraStyle.Height + dGAP * 5 + Me.Height - Me.InsideHeight
    End If

    'Don't let the form get less than a certain width - must be as wide as the biggest check box, plus the standard gap
    If Me.InsideWidth < cbMaximize.Width + fraStyle.Width - fraStyle.InsideWidth + dGAP * 4 Then

        'Reset the width, allowing for the form's border (Width - InsideWidth)
        Me.Width = cbMaximize.Width + fraStyle.Width - fraStyle.InsideWidth + dGAP * 4
    End If

    'Work out the new dimensions of the frame (as the check boxes move within the frame)
    With fraStyle
        dFrameCols = Application.Max(1, (Me.InsideWidth - dGAP * 3 - (.Width - .InsideWidth)) \ (cbMaximize.Width + dGAP))
        dFrameRows = .Controls.Count / dFrameCols

        If dFrameRows <> Int(dFrameRows) Then dFrameRows = Int(dFrameRows) + 1
        dFrameHeight = dFrameRows * cbMaximize.Height + dGAP + .Height - .InsideHeight
    End With

    'Don't allow the form width to decrease so that there's no room for the checkboxes
    'i.e. decreasing the width causes the check boxes to require an extra row, which doesn't fit.
    If Me.InsideHeight <= btnOK.Height + lblMessage.Height + dFrameHeight + dGAP * 5 Then

        'Reset the width, allowing for the form's border (Width - InsideWidth)
        Me.Width = fraStyle.Width + dGAP * 2 + Me.Width - Me.InsideWidth

        'Recalculate the frame's dimensions with the changed form's width
        With fraStyle
            dFrameCols = Application.Max(1, (Me.InsideWidth - dGAP * 3 - (.Width - .InsideWidth)) \ (cbMaximize.Width + dGAP))
            dFrameRows = .Controls.Count / dFrameCols

            If dFrameRows <> Int(dFrameRows) Then dFrameRows = Int(dFrameRows) + 1
            dFrameHeight = dFrameRows * cbMaximize.Height + dGAP + .Height - .InsideHeight
        End With

    End If

    'Set the OK button to be in the middle at the bottom
    With btnOK
        .Left = (Me.InsideWidth - btnOK.Width) / 2
        .Top = Me.InsideHeight - btnOK.Height - dGAP
    End With

    'Set the frame to be as wide as the box and move the check boxes in it to fit
    With fraStyle
        .Width = Me.InsideWidth - dGAP * 2
        .Height = dFrameHeight

        'Reposition the controls in the frame, according to their tab order
        For i = 0 To .Controls.Count - 1
            For j = 0 To .Controls.Count - 1
                With .Controls(j)
                    If .TabIndex = i Then
                        .Left = (i Mod dFrameCols) * (cbMaximize.Width + dGAP) + dGAP
                        .Top = Int(i / dFrameCols) * cbMaximize.Height + dGAP
                    End If
                End With
            Next
        Next

        .Top = btnOK.Top - dGAP - .Height
    End With

    'Userform is big enough, so set the message box's height and width to fill it
    With tbMessage
        .Width = Me.InsideWidth - dGAP * 2

        'Don't allow the height to go negative
        .Height = Application.Max(0, fraStyle.Top - .Top - dGAP)
    End With

End Sub

Private Sub cbModal_Change()
    oFormChanger.Modal = cbModal.Value
    CheckEnabled
End Sub

Private Sub cbSizeable_Change()
    oFormChanger.Sizeable = cbSizeable.Value

    CheckBorderStyle
End Sub

Private Sub cbCaption_Change()
    oFormChanger.ShowCaption = cbCaption.Value

    CheckBorderStyle
    CheckEnabled
End Sub

Private Sub cbSysmenu_Change()
    oFormChanger.ShowSysMenu = cbSysmenu.Value
    CheckEnabled
End Sub

Private Sub cbTaskBar_Change()
    oFormChanger.ShowTaskBarIcon = cbTaskBar.Value
    CheckEnabled
End Sub

Private Sub cbSmallCaption_Change()
    oFormChanger.SmallCaption = cbSmallCaption.Value
    CheckEnabled
End Sub

Private Sub cbIcon_Change()
    oFormChanger.ShowIcon = cbIcon.Value
    CheckEnabled
End Sub

Private Sub cbCloseBtn_Change()
    oFormChanger.ShowCloseBtn = cbCloseBtn.Value
    CheckEnabled
End Sub

Private Sub cbMaximize_Change()
    oFormChanger.ShowMaximizeBtn = cbMaximize.Value
    CheckEnabled
End Sub

Private Sub cbMinimize_Change()
    oFormChanger.ShowMinimizeBtn = cbMinimize.Value
    CheckEnabled
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub CheckBorderStyle()

    'If the useform is not sizeable and doesn't have a caption,
    'Windows draws it without a border, and we need to apply our
    'own 3D effect.
    If Not (cbSizeable Or cbCaption) Then
        Me.SpecialEffect = fmSpecialEffectRaised
    Else
        Me.SpecialEffect = fmSpecialEffectFlat
    End If

End Sub

Private Sub CheckEnabled()
    
    'Without a system menu, we can't have the close, max or min buttons
    cbSysmenu.Enabled = cbCaption
    cbCloseBtn.Enabled = cbSysmenu And cbCaption
    cbIcon.Enabled = cbSysmenu And cbCaption And Not cbSmallCaption
    cbMaximize.Enabled = cbSysmenu And cbCaption
    cbMinimize.Enabled = cbSysmenu And cbCaption

    btnChangeIcon.Enabled = cbIcon

End Sub
