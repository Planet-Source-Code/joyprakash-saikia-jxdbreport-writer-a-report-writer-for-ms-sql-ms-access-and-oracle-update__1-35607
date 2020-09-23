VERSION 5.00
Begin VB.UserControl ctlLabel 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   405
   ScaleWidth      =   1200
   Begin VB.Label lblText 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "ctlLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ctl As Label
Attribute ctl.VB_VarHelpID = -1
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Private Sub ctl_Click()
        RaiseEvent Click
End Sub

Private Sub ctl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub ctl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
        Set ctl = lblText
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set ctl = lblText
End Sub
Public Sub Adjust()
 UserControl.Width = lblText.Width
 UserControl.Height = lblText.Height

End Sub

Private Sub UserControl_Resize()
lblText.Width = UserControl.Width
lblText.Height = UserControl.Height
End Sub

Private Sub UserControl_Show()
UserControl.BackStyle = 1
'Set UserControl.Parent = Parent.picDoc
UserControl.BackColor = vbWhite 'Parent.BackColor
End Sub
