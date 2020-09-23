VERSION 5.00
Object = "{A97B8938-0414-11D5-83E3-008048D61E92}#2.0#0"; "SOFTBTTN.OCX"
Begin VB.Form frmPreview 
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8700
   ControlBox      =   0   'False
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCont 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5295
      ScaleWidth      =   7755
      TabIndex        =   3
      Top             =   480
      Width           =   7755
      Begin VB.PictureBox picCmd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         ScaleHeight     =   225
         ScaleWidth      =   3405
         TabIndex        =   8
         Top             =   5040
         Width           =   3435
         Begin VB.Label lblPageSize 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   11
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label lblRecord 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmdDummy 
         Caption         =   "//"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7380
         MaskColor       =   &H8000000C&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5040
         Width           =   255
      End
      Begin VB.VScrollBar vs 
         Height          =   1215
         Left            =   7380
         TabIndex        =   6
         Top             =   3720
         Width           =   255
      End
      Begin VB.HScrollBar hs 
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   4980
         Width           =   2175
      End
      Begin VB.PictureBox picDoc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4545
         ScaleWidth      =   3405
         TabIndex        =   4
         Top             =   120
         Width           =   3435
         Begin JxDBRpt.ctlLabel ctlLabel1 
            Height          =   225
            Index           =   0
            Left            =   2430
            TabIndex        =   12
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
         End
      End
   End
   Begin VB.PictureBox picControl 
      BackColor       =   &H00B4B4B4&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtPage 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4590
         TabIndex        =   9
         Text            =   "page/page"
         Top             =   60
         Width           =   1035
      End
      Begin VB.ComboBox cboZoom 
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Text            =   "100"
         Top             =   60
         Width           =   1695
      End
      Begin SoftButton.SoftBttn SoftBttn1 
         Height          =   345
         Left            =   30
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         Picture         =   "frmPreview.frx":000C
         PictureAlignment=   4
         BackColor       =   -2147483634
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   -1  'True
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   7
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn cmdPrint 
         Height          =   435
         Left            =   660
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   767
         Picture         =   "frmPreview.frx":045E
         PictureAlignment=   3
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn cmdClose 
         Height          =   315
         Left            =   6600
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         PictureAlignment=   3
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   0
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   "Close"
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   4
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Arial"
         FontSize        =   9
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn cmdNavigate 
         Height          =   315
         Index           =   0
         Left            =   3870
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmPreview.frx":2C10
         PictureAlignment=   0
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn cmdNavigate 
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmPreview.frx":306A
         PictureAlignment=   0
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn cmdNavigate 
         Height          =   315
         Index           =   2
         Left            =   5670
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmPreview.frx":34C4
         PictureAlignment=   0
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin SoftButton.SoftBttn cmdNavigate 
         Height          =   315
         Index           =   3
         Left            =   6030
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmPreview.frx":391E
         PictureAlignment=   0
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         MaskColor       =   16777215
         UseMaskColor    =   -1  'True
         PlaySounds      =   0   'False
         Object.ToolTipText     =   ""
         Caption         =   ""
         ShrinkIcon      =   0   'False
         IconHeight      =   16
         IconWidth       =   16
         TextAlignment   =   3
         ForeColor       =   0
         PopUpButton     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   0
      End
      Begin VB.Label lblView 
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   90
         Width           =   495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781CD0208"
'**********************************************************************
'
'           Module Name: frmPreview
'           Purpose    : To Show The Report in the Window
'           Author     : Joyprakash Saikia

'**********************************************************************
Option Explicit
'This Constant is Used for Debugging Purpose
'This is not used on this submission

Private Const MOD_NAME = "frmPreview"

'The  Following Functions are used for Graphical Manipulation of
'Print Preview
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
    
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' To Show Hand Cursor on the Anchor Field
'  Not Implemented Yet
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_ARROW = 32512&
Private Const IDC_HAND = 32649&


Private Const STRETCH_HALFTONE  As Long = &H4&

Private Const SW_RESTORE        As Long = &H9&
Private Const LeftMargin = 120
Private Const TopMargin = 120
Dim oRpt As JxDBReport
Dim ZoomRatio As Single
Private lPage As Long 'hold current page
Private lPageMax As Long 'hold total report page(s)
Dim bmZoomChanged As Boolean

'Private hHelp As New HTMLHelp  'to Display Popup Help


Private Sub cboZoom_Click()
        '************************************************************
        '          Description:
        '                 This Routine is to Zoom the PictureBox
        '          OutPut:
        '                   View Aspect Ratio is Changed
        '
        '************************************************************

    If Val(cboZoom.Text) Then
        ZoomRatio = CSng(cboZoom.Text) / 100
    Else
        Select Case cboZoom.Text
            Case "Fit Width"
                ZoomRatio = GetRatioFitWidth
            Case "Fit Page"
                ZoomRatio = GetRatioFitPage
            Case "Fit Height"
                ZoomRatio = GetRatioFitHeight
        End Select
        
    End If
    
    SetZoom
End Sub

Private Sub cboZoom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim fZoom As Single
    If KeyCode = vbKeyReturn Then 'if Enter or Return Key is Pressed then Change it
         On Error GoTo ErrHandler
         fZoom = CSng(cboZoom.Text)
         
        If fZoom < 10 Then
            fZoom = 10
        End If
        If fZoom > 200 Then
            fZoom = 200
        End If
        ZoomRatio = fZoom / 100
        cboZoom.Text = fZoom
        SetZoom
    End If
ErrHandler:
    
End Sub

Private Sub cboZoom_KeyPress(KeyAscii As Integer)
    If Not (IsNumeric(Chr$(KeyAscii)) Or KeyAscii = 8) Then KeyAscii = 0
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdClose_MouseEnter()
    cmdClose.ForeColor = vbRed
    cmdClose.FontBold = True
End Sub

Private Sub cmdClose_MouseExit()
    cmdClose.ForeColor = vbBlack
End Sub


Private Sub cmdNavigate_Click(Index As Integer)
    Select Case Index
    
        Case 0 'first page
            lPage = 1
            
        Case 1 ' previous page
            lPage = lPage - 1
        Case 2 ' next page
            lPage = lPage + 1
            If lPage > lPageMax Then
                cmdNavigate(2).Enabled = False
            End If
        Case 3 ' last page
            lPage = lPageMax
    End Select
    UpdateNavigationCmd
        
    PreviewPage lPage, lPage
    
    
End Sub

Private Sub UpdateNavigationCmd()
    If lPage = lPageMax Then
        cmdNavigate(3).Enabled = False
        cmdNavigate(2).Enabled = False
    Else
        cmdNavigate(3).Enabled = True
        cmdNavigate(2).Enabled = True
    End If
    If lPage = 1 Then
        cmdNavigate(0).Enabled = False
        cmdNavigate(1).Enabled = False
    Else
        cmdNavigate(0).Enabled = True
        cmdNavigate(1).Enabled = True
    End If
End Sub

Private Sub cmdNavigate_MouseEnter(Index As Integer)
    cmdNavigate(Index).BackColor = &HB4B4B4
End Sub

Private Sub cmdNavigate_MouseExit(Index As Integer)
    cmdNavigate(Index).BackColor = -2147483633
End Sub

'##ModelId=3B3781D003DE
Private Sub cmdPrint_Click()
    oRpt.ShowPrinterDialog lPage, Me.hwnd
End Sub

'##ModelId=3B3781D10064
Private Sub cmdSave_Click()
    ShowSave Me.hwnd, "*.JxDB", "Save Report As"
End Sub

'Private Sub ctlLabel1_Click(inDex As Integer)
'
'Dim strPopup As String
'  With hHelp
'    .HHPopupType = HH_TEXT_POPUP
'    .HHPopupCustomColors = True
'    .HHPopupCustomBackColor = &HC6FEFF
''    strPopup = "This is a Pop Display Of Description " & _
''        Chr(10) & "Like Addresses, Phone No or SSN " & Chr(10) & Chr(10) & _
''        " This is a way to provide eye catching infotmation to Users."
'
'    .HHPopupText = ctlLabel1(inDex).Tag
'    .HHPopupTextColor = vbBlack
'    .HHPopupTextFont = "Arial"
'    .HHPopupTextSize = "8"
'    .HHDisplayPopup Me.hwnd
'
'  End With
'End Sub

Private Sub ctlLabel1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

'Private Sub ctlLabel1_MouseMove(inDex As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'
'If x >= 0 And y >= 0 And x < ctlLabel1(inDex).Width And y < ctlLabel1(inDex).Height Then
'
'    SetCursor LoadCursor(0, IDC_HAND)
'Else
' If Not hHelp Is Nothing Then
'        hHelp.HHClose
' End If
'End If
'
'End Sub


Private Sub Form_Activate()
    Me.Caption = oRpt.ReportTitle
    ctlLabel1(0).ctl.Caption = "Name"
    ctlLabel1(0).ctl.FontBold = True
    ctlLabel1(0).ctl.AutoSize = True
    ctlLabel1(0).Adjust
    ctlLabel1(0).Visible = False
    If bmZoomChanged Then
        SetZoom
    End If
    
    bmZoomChanged = False
   
End Sub

Private Sub Form_Load()
  Dim a As Long

    cboZoom.AddItem "200"
    cboZoom.AddItem "150"
    cboZoom.AddItem "125"
    cboZoom.AddItem "100"
    cboZoom.AddItem "75"
    cboZoom.AddItem "50"
    cboZoom.AddItem "25"
    cboZoom.AddItem "Fit Width"
    cboZoom.AddItem "Fit Height"
    cboZoom.AddItem "Fit Page"
    
    
    ZoomRatio = 1
    lPage = 1
    bmZoomChanged = True
    UpdateNavigationCmd
    UpdateToolBar
    DisplayReportInfo

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Close any open popups
  Me.SetFocus
  
  ' Clean up the HH class
'  hHelp.HHClose
'  Set hHelp = Nothing
End Sub

Private Sub Form_Resize()

    If Me.ScaleWidth < 4000 Then
        Me.Width = 4000
    End If
    If Me.ScaleHeight < 4000 Then
        Me.Height = 4000
    End If
   
    If Me.WindowState = 1 Then Exit Sub
    picCont.Move 0, picControl.Height, Me.ScaleWidth, Me.ScaleHeight - picControl.Height
    vs.Move picCont.ScaleWidth - vs.Width, 0, vs.Width, picCont.Height - hs.Height  '- TopMargin
    hs.Move picCmd.ScaleWidth, picCont.ScaleHeight - hs.Height, picCont.ScaleWidth - vs.Width - picCmd.ScaleWidth, hs.Height
    cmdDummy.Move hs.Width + picCmd.ScaleWidth, vs.Height
    picCmd.Top = hs.Top
    picCmd.Left = 0
    SetScroll

 End Sub

Public Property Set Document(vDocument As JxDBReport)
    Set oRpt = vDocument
    lPageMax = oRpt.TotalPages
   
End Property

Private Sub SetZoom()
    
    picDoc.Visible = False
    picDoc.ScaleMode = vbTwips
    With picDoc
        .Height = oRpt.ReportHeight * ZoomRatio
        .Width = oRpt.ReportWidth * ZoomRatio
        .Refresh
    End With
    SetScroll
    
    
    PreviewPage lPage, lPage

    picDoc.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRpt = Nothing
End Sub

Private Sub hs_Change()

    picDoc.Left = -hs.Value * 10
    If hs.Value = 0 Then picDoc.Left = LeftMargin
    If hs.Value = hs.Max Then
        picDoc.Left = picDoc.Left - LeftMargin - vs.Width
    End If
    
End Sub

 Friend Sub PreviewPage(ByVal StartPage As Integer, ByVal EndPage As Integer)

    Set oRpt.Target = picDoc
 
    oRpt.PreviewRatio = ZoomRatio
    oRpt.PreviewIt StartPage, EndPage
    UpdatePageStatus
End Sub


Private Sub hs_Scroll()
    hs_Change
    
End Sub
Private Sub mnuFileExit_Click()
    cmdClose_Click
End Sub

Private Sub picDoc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu mnuFile, vbPopupMenuRightButton
End Sub


Private Sub SoftBttn1_Click()
Unload Me
End Sub

Private Sub txtPage_GotFocus()
    txtPage.SelStart = 0
    txtPage.SelLength = Len(txtPage.Text)
End Sub


Private Sub txtPage_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lPageToView As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    lPageToView = Val(txtPage.Text) ' To take the First Numeric Value
    If (lPageToView = 0) Or (lPageToView > lPageMax) Then
        MsgBox LoadResString(MSG_PAGE_NOTEXIST), vbInformation, APP_NAME
        UpdatePageStatus
        Exit Sub
    End If
    'we should get a valid page to view here
    lPage = lPageToView
    UpdateNavigationCmd
    PreviewPage lPageToView, lPageToView
    
End Sub



Private Sub vs_Change()
    
    Dim iY As Long

    iY = -vs.Value     'to avoid overflow
    iY = iY * 10
    picDoc.Top = iY
    If vs.Value = 0 Then picDoc.Top = TopMargin 'Min Value
    If vs.Value = vs.Max Then
        picDoc.Top = picDoc.Top - TopMargin - hs.Height
    End If
End Sub

Private Sub UpdatePageStatus()
    Dim sText As String
    sText = CStr(lPage) & " / " & lPageMax
    txtPage.Text = sText
End Sub

Private Sub vs_Scroll()
    vs_Change
End Sub

Private Function GetRatioFitWidth() As Single
    Dim fWidth As Single
    Dim fXratio As Single
    picDoc.ScaleMode = vbTwips
    fWidth = picCont.Width - vs.Width - LeftMargin - LeftMargin
    fXratio = fWidth / oRpt.ReportWidth
    
    GetRatioFitWidth = fXratio
    'SetZoom
End Function

Private Function GetRatioFitPage() As Single
    Dim fHeight As Single
    Dim fWidth As Single
    Dim fYratio As Single, fXratio As Single
    picDoc.ScaleMode = vbTwips
    'use height first
    fYratio = GetRatioFitHeight
    'try using width
    fXratio = GetRatioFitWidth
    GetRatioFitPage = fXratio
    
    If fYratio < fXratio Then
        GetRatioFitPage = fYratio
    End If
    
End Function
Private Function GetRatioFitHeight() As Single
    Dim fHeight As Single
    
    Dim fYratio As Single
    picDoc.ScaleMode = vbTwips
    'use height first
    fHeight = picCont.Height - hs.Height - TopMargin - TopMargin
    fYratio = fHeight / oRpt.ReportHeight
    GetRatioFitHeight = fYratio
End Function
Private Sub SetScroll()
    If picDoc.Width > picCont.Width - vs.Width Then
        picDoc.Left = LeftMargin
        hs.Enabled = True
        hs.Min = 0
        hs.SmallChange = 10
        hs.LargeChange = picCont.Width / 10
        hs.Max = (picDoc.Width - picCont.Width) / 10
    Else
        'center
        picDoc.Left = (picCont.Width - picDoc.Width - vs.Width) / 2
        'picDoc.Left = LeftMargin '  (picCont.Width - picDoc.Width) / 2
        hs.Enabled = False
    End If
    If picDoc.Height > picCont.Height Then
        vs.Enabled = True
        picDoc.Top = TopMargin
        vs.Min = 0
        vs.Max = (picDoc.Height - picCont.Height) / 10
        vs.SmallChange = 10
        vs.LargeChange = picCont.Width / 10
    Else
        picDoc.Top = (picCont.Height - picDoc.Height - hs.Height) / 2
        vs.Enabled = False
    End If
    vs.Value = 0
    hs.Value = 0
End Sub
Private Sub DisplayReportInfo()
    Dim fHeight As Single, fWidth As Single
    'show total record processed
    lblRecord.Caption = "Record:" & oRpt.RecordProcessed
    'show report paper dimension
    fHeight = ConvertFromTwip(JxDBRptScaleCm, oRpt.ReportHeight)
    fWidth = ConvertFromTwip(JxDBRptScaleCm, oRpt.ReportWidth)
    lblPageSize = Format(fWidth, "##0.00") & "cm x " & Format(fHeight, "##0.00") & "cm"
End Sub
    
Sub UpdateToolBar()
Dim Count As Long
picControl.BackColor = Me.BackColor
For Count = 0 To 3
     cmdNavigate(Count).BackColor = Me.BackColor
     
Next
cmdClose.BackColor = Me.BackColor
End Sub
