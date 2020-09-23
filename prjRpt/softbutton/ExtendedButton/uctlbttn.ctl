VERSION 5.00
Begin VB.UserControl SoftBttn 
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "uctlbttn.ctx":0000
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   ToolboxBitmap   =   "uctlbttn.ctx":0047
End
Attribute VB_Name = "SoftBttn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Soft Button Control"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'************************************************************************
'Author            :   Vijay Phulwadhawa     Date    : 23/02/2001 12:59:20 PM
'Project Name      :   Insert_Project_Name
'Form/Class Name   :   SoftBttn (Code)
'Version           :   6.00
'Description       :   <Purpose>
'Links             :   <Links With Any Other Form Modules>
'Change History    :
'Date      Author      Description Of Changes          Reason Of Change
'************************************************************************


'-------------------------------------------------------------------------
'This user control mimics the behavior of a button in the VB5 toolbar.
'When no mouse is over the button, it appears flat, like an image with a
'transparent background.  When a mouse is over the button, a soft 3D edge
'is drawn.  When the mouse is pressed a sunken 3D edge is drawn.  This control
'creates its own Tooltips because the VB intrinsic tooltips are disabled
'by the controls use of SetCapture.
'Needed files:
'   modTlTip.bas    module providing WinProc needed for ToolTips
'   clsDraw.cls     Class module provides transparent and disabled drawing
'                   procedures
'   modSftBn.bas    Declarations of needed Types, API Functions, and Constants
'-------------------------------------------------------------------------

Public Enum PictureAlignment
TopLeft
TopCenter
TopRight
LeftCenter
Center
RightCenter
BottomLeft
BottomCenter
BottomRight
End Enum


'Property Name constants
Private Const msURL_PICTURE_NAME = "URLPicture"
Private Const msBACK_COLOR_NAME = "BackColor"
Private Const msPICTURE_NAME = "Picture"
Private Const msENABLED_NAME = "Enabled"
Private Const msMASK_COLOR_NAME = "MaskColor"
Private Const msUSE_MASK_COLOR_NAME = "UseMaskColor"
Private Const msPLAY_SOUNDS_NAME = "PlaySounds"
Private Const msTOOL_TIP_TEXT_NAME = "ToolTipText"
Private Const msPICTURE_ALIGNMENT = "PictureAlignment"
Private Const msTEXT_ALIGNMENT = "TextAlignment"
Private Const msCAPTION = "Caption"
Private Const msSHRINKICON = "ShrinkIcon"
Private Const msICONHEIGHT = "IconHeight"
Private Const msICONWIDTH = "IconWidth"
Private Const msFORECOLOR = "ForeColor"
Private Const msPOPUPBUTTON = "PopUpButton"

'Property Values
Private m_bEnabled As Boolean
Private m_clrMaskColor As OLE_COLOR
Private m_bUseMaskColor As Boolean
Private m_bPlaySounds As Boolean
Private m_picPictured As Picture
Private m_sToolTipText As String
Private m_sURLPicture As String
Private m_PicAlign As PictureAlignment
Private m_TextAlign As PictureAlignment
Private m_Caption As String
Private m_ShrinkIcon As Boolean
Private m_IconHeight As Integer
Private m_IconWidth As Integer
Private m_ForeColor As Long
Private m_PopUpButton As Boolean
'Class level variables
Private msToolTipBuffer As String         'Tool tip text; This string must have
                                          'module or global level scope, because
                                          'a pointer to it is copied into a
                                          'ToolTipText structure
Private mbClearURLOnly As Boolean
Private mbClearPictureOnly As Boolean
Private mbToolTipNotInExtender As Boolean
Private moDrawTool As clsDrawPictures
Private mbGotFocus As Boolean
Private mbMouseOver As Boolean
Private miCurrentState As Integer
Private mWndProcNext As Long            'The address entry point for the subclassed window
Private mHWndSubClassed As Long         'hWnd of the subclassed window
Private mbLeftMouseDown As Boolean
Private mbLeftWasDown As Boolean
Private mudtButtonRect As RECT
Private mudtPictureRect As RECT
Private mudtPicturePoint As POINTAPI
Private mbPropertiesLoaded As Boolean
Private mbEnterOnce As Boolean
Private mbMouseDownFired As Boolean
Private mlhHalftonePal As Long

Private mudText As POINTAPI

#If DEBUGSUBCLASS Then                      'Tool that checks if in break mode and then
    Private moProcHook As Object            'Sends messages to mWndPRocNext instead of
#End If                                     'Address of my function

Public Event Click()
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event PopUp()
Public Event MouseEnter()
Public Event MouseExit()


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
'debug.Print "UserControl_AccessKeyPress"
MakeClick
End Sub

'****************************
'UserControl event procedures
'****************************
Private Sub UserControl_Click()
'debug.Print "UserControl_Click"
    mbMouseOver = False     'Set mbMouseOver = False so that SetCapture
                            'will be called again.  For some reason, capture
                            'is released during MouseClick
    MouseOver
End Sub

Private Sub UserControl_EnterFocus()
'debug.Print "UserControl_EnterFocus"
    '-------------------------------------------------------------------------
    'Purpose:   If tabstop property is true, show button raised so that user
    '           can see that button received focus.
    '-------------------------------------------------------------------------
    On Error GoTo UserControl_EnterFocusError
    'Error may occur if TabStop property is not available
    If UserControl.Extender.TabStop Then
        On Error Resume Next
        mbGotFocus = True
        If Not miCurrentState = giRAISED Then DrawButtonState giRAISED
    End If
    UpdateCaption
UserControl_EnterFocusError:
    Exit Sub
End Sub

Private Sub UserControl_ExitFocus()
'debug.Print "UserControl_ExitFocus"
    '-------------------------------------------------------------------------
    'Purpose:   Flatten button if the mouse is not over it
    '-------------------------------------------------------------------------
    If Not mbMouseOver And m_PopUpButton Then DrawButtonState giFLATTENED
    mbGotFocus = False
    UpdateCaption
End Sub

Private Sub UserControl_Initialize()
    'debug.Print "UserControl_Initialize"
    mbPropertiesLoaded = False
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
    Set moDrawTool = New clsDrawPictures
    mlhHalftonePal = CreateHalftonePalette(UserControl.hdc)
    
End Sub

Private Sub UserControl_InitProperties()
'debug.Print "UserControl_InitProperties"
    '-------------------------------------------------------------------------
    'Purpose:   Set the default properties to be displayed the first time this
    '           control is placed on a container
    '-------------------------------------------------------------------------
    On Error Resume Next
    'Error may occur if TabStop property is not available
    BackColor = UserControl.BackColor
    Enabled = True
    UseMaskColor = False
    MaskColor = vbWhite
    UserControl.Extender.TabStop = False
    UserControl.ScaleMode = vbPixels
    ToolTipText = ""
    Caption = "Soft Command 1"
    PictureAlignment = Center
    ShrinkIcon = False
    IconHeight = 16
    IconWidth = 16
    TextAlignment = BottomCenter
    m_ForeColor = vbBlack
    mbPropertiesLoaded = True
    m_PopUpButton = False
    Set UserControl.Font = Ambient.Font
    Refresh
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'debug.Print "UserControl_KeyDown"
If KeyCode = vbKeySpace Then
DrawButtonState giSUNKEN
UpdateCaption
End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
'debug.Print "UserControl_KeyPress"
If KeyAscii = vbKeyReturn Then
MakeClick
UpdateCaption
End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'debug.Print "UserControl_KeyUp"
If KeyCode = vbKeySpace Then
DrawButtonState giRAISED
MakeClick
UpdateCaption
End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'debug.Print "UserControl_MouseDown"
    '-------------------------------------------------------------------------
    'Purpose:   If the mouse is over the button and the left button is down
    '           show that the button is sunken and set a flag that the button
    '           is down
    '-------------------------------------------------------------------------
    With UserControl
        If Button = vbLeftButton And (x >= 0 And x <= .ScaleWidth) And (y >= 0 And y <= .ScaleHeight) Then
            mbLeftMouseDown = True
            DrawButtonState giSUNKEN
            UpdateCaption
        End If
    End With
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'debug.Print "UserControl_MouseMove"
    '-------------------------------------------------------------------------
    'Purpose:   If the mouse is over show the button raised.  If the mouse is
    '           over and the left mouse is down and was down when the mouse left
    '           the button show button sunken.  If mouse is off button, flatten,
    '           unless left button is down show the mouse raised until the button
    '           is released.
    '-------------------------------------------------------------------------
    With UserControl
        If (x <= .ScaleWidth And x >= 0) And (y <= .ScaleHeight And y >= 0) Then
            If mbLeftWasDown Then
                mbLeftMouseDown = True
                mbLeftWasDown = False
                DrawButtonState giSUNKEN
                UpdateCaption
            Else
                If Button <> vbLeftButton Then MouseOver
            End If
            RaiseEvent MouseMove(Button, Shift, x, y)
        Else
            If mbLeftMouseDown Then
                mbLeftWasDown = True
                mbLeftMouseDown = False
                DrawButtonState giRAISED
                UpdateCaption
                RaiseEvent MouseMove(Button, Shift, x, y)
            ElseIf Not mbLeftWasDown Then
                Flatten
            Else
                RaiseEvent MouseMove(Button, Shift, x, y)
            End If
        End If
    End With

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'debug.Print "UserControl_MouseUp"
    '-------------------------------------------------------------------------
    'Purpose:   If the left mouse was down and the left button is up and the mouse
    '           is over the button raise a Click event.  If left button was down
    '           and mouse is off button flatten button.
    '-------------------------------------------------------------------------
    With UserControl
        If (x >= 0 And x <= .ScaleWidth) And (y >= 0 And y <= .ScaleHeight) Then
            If (mbLeftMouseDown Or mbLeftWasDown) And Button = vbLeftButton Then
                mbLeftMouseDown = False
                DrawButtonState giRAISED
                MakeClick
            End If
        ElseIf mbLeftWasDown And Button = vbLeftButton Then
            mbLeftWasDown = False
            Flatten
        Else
            mbMouseOver = False     'Set mbMouseOver = False so that SetCapture
                                    'will be called again.  For some reason, capture
                                    'is released during MouseUp
        End If
    End With
    UpdateCaption
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
'debug.Print "UserControl_Paint"
DrawButtonState miCurrentState
UpdateCaption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'debug.Print "UserControl_ReadProperties"
    Dim sUrl As String
    Dim picMine As Picture
    On Error Resume Next
    ' Read in the properties that have been saved into the PropertyBag...
    With PropBag
        sUrl = .ReadProperty(msURL_PICTURE_NAME, "")        ' Read URLPicture property value
        If Len(sUrl) <> 0 Then                         ' If a URL has been entered...
            URLPicture = sUrl                        ' Attempt to download it now, URL may be unavailable at this time
        Else
            Set picMine = PropBag.ReadProperty(msPICTURE_NAME, UserControl.Picture) ' Read Picture property value
            If Not (picMine Is Nothing) Then            ' URL is not available
                Set Picture = picMine                   ' Use existing picture (This is used only if URL is empty)
            End If
        End If
    End With
    BackColor = PropBag.ReadProperty(msBACK_COLOR_NAME, UserControl.BackColor)
    Enabled = PropBag.ReadProperty(msENABLED_NAME, True)
    MaskColor = PropBag.ReadProperty(msMASK_COLOR_NAME, vbWhite)
    UseMaskColor = PropBag.ReadProperty(msUSE_MASK_COLOR_NAME, False)
    PlaySounds = PropBag.ReadProperty(msPLAY_SOUNDS_NAME, False)
    ToolTipText = PropBag.ReadProperty(msTOOL_TIP_TEXT_NAME, "")
    PictureAlignment = PropBag.ReadProperty(msPICTURE_ALIGNMENT, Center)
    TextAlignment = PropBag.ReadProperty(msTEXT_ALIGNMENT, BottomCenter)
    Caption = PropBag.ReadProperty(msCAPTION, "Soft Command 1")
    ShrinkIcon = PropBag.ReadProperty(msSHRINKICON, False)
    IconHeight = PropBag.ReadProperty(msICONHEIGHT, 16)
    IconWidth = PropBag.ReadProperty(msICONWIDTH, 16)
    ForeColor = PropBag.ReadProperty(msFORECOLOR, vbBlack)
    PopUpButton = PropBag.ReadProperty(msPOPUPBUTTON, False)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontName = PropBag.ReadProperty("FontName", "")
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 0)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.AccessKeys = PropBag.ReadProperty("HotKey", "")
    InstanciateToolTipsWindow
    mbPropertiesLoaded = True
End Sub

Private Sub UserControl_Terminate()
'debug.Print "UserControl_Terminate"
    Set moDrawTool = Nothing
    DeleteObject mlhHalftonePal
    glToolsCount = glToolsCount - 1
    UnSubClass
    If gbToolTipsInstanciated And glToolsCount = 0 Then
        DestroyWindow gHWndToolTip
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'debug.Print "UserControl_WriteProperties"
    If Len(m_sURLPicture) <> 0 Then
        PropBag.WriteProperty msURL_PICTURE_NAME, m_sURLPicture
    Else
        PropBag.WriteProperty msPICTURE_NAME, m_picPictured
    End If
    
    If Len(m_PicAlign) <> 0 Then
        PropBag.WriteProperty msPICTURE_ALIGNMENT, m_PicAlign
    Else
        PropBag.WriteProperty msPICTURE_ALIGNMENT, Center
    End If
    
    PropBag.WriteProperty msBACK_COLOR_NAME, UserControl.BackColor
    PropBag.WriteProperty msENABLED_NAME, m_bEnabled
    PropBag.WriteProperty msMASK_COLOR_NAME, m_clrMaskColor
    PropBag.WriteProperty msUSE_MASK_COLOR_NAME, m_bUseMaskColor
    PropBag.WriteProperty msPLAY_SOUNDS_NAME, m_bPlaySounds
    PropBag.WriteProperty msTOOL_TIP_TEXT_NAME, m_sToolTipText
    PropBag.WriteProperty msCAPTION, m_Caption
    PropBag.WriteProperty msSHRINKICON, m_ShrinkIcon
    PropBag.WriteProperty msICONHEIGHT, m_IconHeight
    PropBag.WriteProperty msICONWIDTH, m_IconWidth
    PropBag.WriteProperty msTEXT_ALIGNMENT, m_TextAlign
    PropBag.WriteProperty msFORECOLOR, m_ForeColor
    PropBag.WriteProperty msPOPUPBUTTON, m_PopUpButton
    
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("HotKey", UserControl.AccessKeys, "")
End Sub

Private Sub UserControl_Resize()
'debug.Print "UserControl_Resize"
    'Reevaluate coordinates
    'Repaint control
    PositionChanged
    DrawButtonState miCurrentState
    PositionChanged
    UpdateCaption
End Sub

'**********************
'Public Properties
'**********************
Public Property Let ToolTipText(ByVal sToolTip As String)
'debug.Print "Let ToolTipText"
    m_sToolTipText = sToolTip
    'If this property gets called with more than an empty
    'string, I know for sure that there is not a ToolTipText
    'extender property
    If Len(sToolTip) <> 0 Then mbToolTipNotInExtender = True
    PropertyChanged (msTOOL_TIP_TEXT_NAME)
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get ToolTipText"
    ToolTipText = m_sToolTipText
End Property



Public Property Let PlaySounds(ByVal bPlaySounds As Boolean)
'debug.Print "Let PlaySounds"
    'The following line of code ensures that the integer
    'value of the boolean parameter is either
    '0 or -1.  It is known that Access 97 will
    'set the boolean's value to 255 for true.
    'In this case a P-Code compiled VB5 built
    'OCX will return True for the expression
    '(Not [boolean variable that ='s 255]).  This
    'line ensures the reliability of boolean operations
    If CBool(bPlaySounds) Then bPlaySounds = True Else bPlaySounds = False
    m_bPlaySounds = bPlaySounds
    PropertyChanged (msPLAY_SOUNDS_NAME)
End Property

Public Property Get PlaySounds() As Boolean
Attribute PlaySounds.VB_Description = "Returns or sets whether or not system sounds are played for user generated events."
Attribute PlaySounds.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get PlaySounds"
    PlaySounds = m_bPlaySounds
End Property

Public Property Let MaskColor(ByVal clrMaskColor As OLE_COLOR)
'debug.Print "Let MaskColor"
    'If there is a valid picture, repaint control
    m_clrMaskColor = clrMaskColor
    If m_bUseMaskColor And Not m_picPictured Is Nothing Then DrawButtonState miCurrentState
    PropertyChanged (msMASK_COLOR_NAME)
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns or sets a color in a button's picture to be a 'mask' (that is, transparent)."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
'debug.Print "Get MaskColor"
    MaskColor = m_clrMaskColor
End Property

Public Property Let UseMaskColor(ByVal bUseMaskColor As Boolean)
'debug.Print "Let UseMaskColor"
    'If true, use the mask color.  Mask color only applies
    'to bitmaps not icons.
    'Repaint control
    'Validate whether correct picture type is provided
    
    'The following line of code ensures that the integer
    'value of the boolean parameter is either
    '0 or -1.  It is known that Access 97 will
    'set the boolean's value to 255 for true.
    'In this case a P-Code compiled VB5 built
    'OCX will return True for the expression
    '(Not [boolean variable that ='s 255]).  This
    'line ensures the reliability of boolean operations
    If CBool(bUseMaskColor) Then bUseMaskColor = True Else bUseMaskColor = False
    m_bUseMaskColor = bUseMaskColor
    If Not m_picPictured Is Nothing Then
        If m_picPictured.Type = vbPicTypeBitmap Then DrawButtonState miCurrentState
    End If
    PropertyChanged (msUSE_MASK_COLOR_NAME)
End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns or sets a value that determines whether the color assigned in the MaskColor property is used as a 'mask'. (That is, used to create transparent regions.)"
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get UseMaskColor"
    UseMaskColor = m_bUseMaskColor
End Property

Public Property Let BackColor(ByVal clrBackColor As OLE_COLOR)
'debug.Print "Let BackColor"
    'Control will be repainted because VB will
    'fire paint event
    UserControl.BackColor = clrBackColor
    DrawButtonState miCurrentState
    PropertyChanged (msBACK_COLOR_NAME)
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute BackColor.VB_UserMemId = -501
'debug.Print "Get BackColor"
    BackColor = UserControl.BackColor
End Property

Public Property Set Picture(ByVal picPicture As Picture)
Attribute Picture.VB_Description = "Returns or sets a graphic to be displayed  on the control."
Attribute Picture.VB_ProcData.VB_Invoke_PropertyPutRef = "StandardPicture;Appearance"
'debug.Print "Set Picture"
    'Validate what kind of picture is passed
    'Only allow bitmaps and icons
    'If not in runtime display message that UseMaskColor can't be
    'used with icons, if picture is icon.
    'If picture is icon, make sure UseMaskColor is false
    'Paint Control
    If Not picPicture Is Nothing Then
        With picPicture
            If (.Type <> vbPicTypeBitmap) And (.Type <> vbPicTypeNone) And (.Type <> vbPicTypeIcon) Then
                If Not UserControl.Ambient.UserMode Then
                    MsgBox LoadResString(giINVALID_PIC_TYPE), vbOKOnly, UserControl.Name
                End If
                Exit Property
            End If
        End With
    End If
    If Not mbClearPictureOnly Then
        mbClearURLOnly = True       ' If Picture property is not being set by the URLPicture
        URLPicture = ""             ' property then clear the URLPicture value...
        mbClearURLOnly = False
    End If
    
    If (Not picPicture Is Nothing) Then
        If (picPicture.Handle = 0) Then Set picPicture = Nothing
    End If
    Set m_picPictured = picPicture
    PositionChanged
    DrawButtonState miCurrentState
    PropertyChanged (msPICTURE_NAME)
End Property

Public Property Get Picture() As Picture
'debug.Print "Get Picture"
    Set Picture = m_picPictured
End Property

Public Property Let URLPicture(ByVal Url As String)
'debug.Print "Let URLPicture"
    If (m_sURLPicture <> Url) Then                   ' Do only if value has changed...
        mbClearPictureOnly = Not mbClearURLOnly      ' If Picture property is not being set by the URLPicture
                                                     ' property then clear the URLPicture value...
        m_sURLPicture = Url                          ' Save URL string value to global variable
        PropertyChanged msURL_PICTURE_NAME           ' Notify property bag of property change

        If Not mbClearURLOnly Then
            On Error GoTo ErrorHandler               ' Handle Error if URL is unavailable or Invalid...
            If (Url <> "") Then
                UserControl.AsyncRead Url, vbAsyncTypePicture, msPICTURE_NAME ' Begin async download of picture file...
            Else
                Set Picture = Nothing
            End If
        End If
    End If
ErrorHandler:
    mbClearPictureOnly = False
End Property

Public Property Get URLPicture() As String
Attribute URLPicture.VB_Description = "Returns or sets the URL address of a picture to be used instead of the Picture property."
Attribute URLPicture.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get URLPicture"
    URLPicture = m_sURLPicture                         ' Return URL string value
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
'debug.Print "Let Enabled"
    'If button is raised, flatten it
    'Draw disabled appearance of picture
    Dim lresult As Long
    'The following line of code ensures that the integer
    'value of the boolean parameter is either
    '0 or -1.  It is known that Access 97 will
    'set the boolean's value to 255 for true.
    'In this case a P-Code compiled VB5 built
    'OCX will return True for the expression
    '(Not [boolean variable that ='s 255]).  This
    'line ensures the reliability of boolean operations
    If CBool(bEnabled) Then bEnabled = True Else bEnabled = False
    UserControl.Enabled = bEnabled
    m_bEnabled = bEnabled
    If bEnabled Then
        'If m_PopUpButton = True Then
        DrawButtonState miCurrentState
        'Else
        'DrawButtonState giRAISED
        'End If
    Else
        If miCurrentState = giRAISED Then
            'Call flatten as if button does not have focus
            If mbGotFocus Then
                'Get rid of focus
                mbGotFocus = False
            End If
            Flatten
        End If
        DrawButtonState giDISABLED
    End If
    PropertyChanged (msENABLED_NAME)
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
Attribute Enabled.VB_UserMemId = -514
'debug.Print "Get Enabled"
    Enabled = UserControl.Enabled
End Property

'*************************
'Private Procedures
'*************************

Private Sub MakeClick()
'debug.Print "MakeClick"
    '-------------------------------------------------------------------------
    'Purpose:   Raise a Click event to container, play sound
    '-------------------------------------------------------------------------
    '-----------------------------------------
    '- Added for sound support
    '-----------------------------------------
    If m_bPlaySounds Then PlaySound EVENT_MENU_COMMAND, 0, SND_SYNC
    '-----------------------------------------
    UpdateCaption
    RaiseEvent Click
End Sub

Private Sub MouseOver()
'debug.Print "MouseOver"
    '-------------------------------------------------------------------------
    'Purpose:   Call whenever the mouse is over the button and
    '           button needs raised appearance and capture set
    '-------------------------------------------------------------------------
    If miCurrentState <> giRAISED Then
    DrawButtonState giRAISED
    UpdateCaption
    End If
    If Not mbMouseOver Then
        Capture True
        mbMouseOver = True
        '-----------------------------------------
        '- Added for sound support
        '-----------------------------------------
        If Not mbEnterOnce Then
            RaiseEvent MouseEnter
            RaiseEvent PopUp
            If m_bPlaySounds Then PlaySound EVENT_MENU_POPUP, 0, SND_SYNC
            mbEnterOnce = True
        End If
        '-----------------------------------------
    End If
End Sub

Private Sub Flatten()
'debug.Print "Flatten"
    '-------------------------------------------------------------------------
    'Purpose:   Call whenever the mouse is off the control
    '           and capture needs released and button needs
    '           flattened appearance
    '-------------------------------------------------------------------------
    If mbMouseOver Then Capture False
    mbMouseOver = False
    If (Not mbGotFocus) And miCurrentState <> giFLATTENED Then
    'If m_PopUpButton = True Then
    RaiseEvent MouseExit
    DrawButtonState giFLATTENED
    'End If
    UpdateCaption
    End If
    '-----------------------------------------
    '- Added for sound support
    '-----------------------------------------
    '   PlaySound EVENT_MENU_POPUP, 0, SND_SYNC
    mbEnterOnce = False
    '-----------------------------------------
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
'debug.Print "UserControl_AsyncReadComplete"
    On Error GoTo ErrorHandler
    
    If (AsyncProp.PropertyName = msPICTURE_NAME) Then ' Picture download is complete
        mbClearPictureOnly = True
        Set Picture = AsyncProp.Value           ' Store picture data to property...
    End If
ErrorHandler:
    mbClearPictureOnly = False
End Sub

Private Sub AddTool(hWnd As Long)
'debug.Print "AddTool"
    '-------------------------------------------------------------------------
    'Purpose:   Add a tool to the ToolTips object
    'In:        [hWnd]
    '               hWnd of Tool being added
    '-------------------------------------------------------------------------
                   
    Dim ti As TOOLINFO
    
    With ti
        .cbSize = Len(ti)
        .uId = hWnd
        .hWnd = hWnd
        .hinst = App.hInstance
        .uFlags = TTF_IDISHWND
        .lpszText = LPSTR_TEXTCALLBACK
    End With
    
    SendMessage gHWndToolTip, TTM_ADDTOOL, 0, ti
    SendMessage gHWndToolTip, TTM_ACTIVATE, 1, ByVal hWnd
    Exit Sub
End Sub

Private Sub InstanciateToolTipsWindow()
'debug.Print "InstanciateToolTipsWindow"
    '-------------------------------------------------------------------------
    'Purpose:   Instanciate needed collections.
    '           Create ToolTips Class window
    '-------------------------------------------------------------------------
    glToolsCount = glToolsCount + 1
    If UserControl.Ambient.UserMode Then
        If Not gbToolTipsInstanciated Then
            gbToolTipsInstanciated = True
            InitCommonControls
            gHWndToolTip = CreateWindowEX(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, 0, _
                      CW_USEDEFAULT, CW_USEDEFAULT, _
                      CW_USEDEFAULT, CW_USEDEFAULT, _
                      0, 0, _
                      App.hInstance, _
                      ByVal 0)
            SendMessage gHWndToolTip, TTM_ACTIVATE, 1, ByVal 0
            
            #If DEBUGSUBCLASS Then
                If goWindowProcHookCreator Is Nothing Then Set goWindowProcHookCreator = CreateObject("DbgWindowProc.WindowProcHookCreator")
            #End If
        End If
        'Sub class this code module to receive
        'window messages for the Usercontrol
        SubClass UserControl.hWnd
        'Add Register Usercontrol with ToolTip window
        AddTool UserControl.hWnd
    End If
End Sub

Private Sub SubClass(hWnd)
'debug.Print "SubClass"
    '-------------------------------------------------------------------------
    'Purpose:   Subclass control so that the ToolTip Need text message can be
    '           handled.  Store handle of class as UserData of control window
    '-------------------------------------------------------------------------
    Dim lresult As Long
    
    UnSubClass
    
    #If DEBUGSUBCLASS Then
        'If in debug, SubClass window using address of WindowProcHook
        'Let WindowProcHook CallWindowProc at address of my function
        'if in run mode but call the previous address if in break mode
        'this prevents crashes in break mode
        Set moProcHook = goWindowProcHookCreator.CreateWindowProcHook
        With moProcHook
            .SetMainProc AddressOf SubWndProc
            mWndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, CLng(.ProcAddress))
            .SetDebugProc mWndProcNext
        End With
    #Else
        mWndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)
    #End If
    
    If mWndProcNext Then
        mHWndSubClassed = hWnd
        lresult = SetWindowLong(hWnd, GWL_USERDATA, ObjPtr(Me))
    End If
End Sub

Private Sub UnSubClass()
'debug.Print "UnSubClass"
    '-------------------------------------------------------------------------
    'Purpose:   Unsubclass control
    '-------------------------------------------------------------------------
    If mWndProcNext Then
        SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcNext
        mWndProcNext = 0
        
        #If DEBUGSUBCLASS Then
            Set moProcHook = Nothing
        #End If
        
    End If
End Sub

Private Sub Capture(bCapture As Boolean)
'debug.Print "Capture"
    '-------------------------------------------------------------------------
    'Purpose:   Is the only place where setcapture and releasecapture are called
    '           setcapture may be called after mouse clicks because VB seems to
    '           release capture on my behalf.
    '-------------------------------------------------------------------------
    If bCapture Then
        SetCapture UserControl.hWnd
    Else
        ReleaseCapture
    End If
End Sub

Private Sub PositionChanged()
'debug.Print "PositionChanged"
    '-------------------------------------------------------------------------
    'Purpose:   Calculate needed coordinates for painting the control
    '-------------------------------------------------------------------------
    On Error GoTo PositionChangedError
    With mudtButtonRect
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    If Not m_picPictured Is Nothing Then
        With mudtPicturePoint
            .x = CLng(UserControl.ScaleX(m_picPictured.Width, vbHimetric, vbPixels))
            .y = CLng(UserControl.ScaleY(m_picPictured.Height, vbHimetric, vbPixels))
        End With
    
Dim TempPoint As POINTAPI
TempPoint = mudtPicturePoint
If m_picPictured.Type = vbPicTypeIcon Then
    If m_ShrinkIcon = True Then
        TempPoint.x = m_IconWidth
        TempPoint.y = m_IconHeight
    End If
End If
        With mudtPictureRect
            '.Left = CLng((mudtButtonRect.Right - mudtPicturePoint.x) / 2)
            '.Top = CLng((mudtButtonRect.Bottom - mudtPicturePoint.y) / 2)
            
            
            Select Case m_PicAlign
            
            Case TopLeft 'OK
            .Left = 1
            .Top = 1
            
            Case TopCenter 'OK
            .Left = CLng((mudtButtonRect.Right - TempPoint.x) / 2)
            .Top = 1
            
            Case TopRight 'OK
            .Left = CLng((mudtButtonRect.Right - TempPoint.x)) - 1
            .Top = 1
            
            Case LeftCenter 'OK
            .Left = 1
            .Top = CLng((mudtButtonRect.Bottom - TempPoint.y) / 2)
            
            Case Center 'OK
            .Left = CLng((mudtButtonRect.Right - TempPoint.x) / 2)
            .Top = CLng((mudtButtonRect.Bottom - TempPoint.y) / 2)
            
            Case RightCenter 'OK
            .Left = CLng((mudtButtonRect.Right - TempPoint.x)) - 1
            .Top = CLng((mudtButtonRect.Bottom - TempPoint.y) / 2)
            
            Case BottomLeft 'OK
            .Left = 1
            .Top = CLng(mudtButtonRect.Bottom - TempPoint.y) - 1
            
            Case BottomCenter 'OK
            .Left = CLng((mudtButtonRect.Right - TempPoint.x) / 2)
            .Top = CLng(mudtButtonRect.Bottom - TempPoint.y) - 1
            
            Case BottomRight 'OK
            .Left = CLng((mudtButtonRect.Right - TempPoint.x)) - 1
            .Top = CLng(mudtButtonRect.Bottom - TempPoint.y) - 1

             End Select
             
            .Right = .Left + TempPoint.x
            .Bottom = .Top + TempPoint.y
            
            'UserControl_Resize
            'Label1.Left = .Left
            'Label1.Top = .Top
        End With
        
        '''''''''''''''''''''
'        GetTextExtentPoint32 UserControl.hdc, m_Caption, Len(m_Caption), mudText
'
'        With mudText
'        Select Case m_TextAlign
'
'            Case TopLeft 'OK
'            .x = 1
'            .y = 1
'
'            Case TopCenter 'OK
'            .x = CLng((mudtButtonRect.Right - .x) / 2)
'            .y = 1
'
'            Case TopRight 'OK
'            .x = CLng((mudtButtonRect.Right - .x)) - 1
'            .y = 1
'
'            Case LeftCenter 'OK
'            .x = 1
'            .y = CLng((mudtButtonRect.Bottom - .y) / 2)
'
'            Case Center 'OK
'            .x = CLng((mudtButtonRect.Right - .x) / 2)
'            .y = CLng((mudtButtonRect.Bottom - .y) / 2)
'
'            Case RightCenter 'OK
'            .x = CLng((mudtButtonRect.Right - .x)) - 1
'            .y = CLng((mudtButtonRect.Bottom - .y) / 2)
'
'            Case BottomLeft 'OK
'            .x = 1
'            .y = CLng(mudtButtonRect.Bottom - .y) - 1
'
'            Case BottomCenter 'OK
'            .x = CLng((mudtButtonRect.Right - .x) / 2)
'            .y = CLng(mudtButtonRect.Bottom - .y) - 1
'
'            Case BottomRight 'OK
'            .x = CLng((mudtButtonRect.Right - .x)) - 1
'            .y = CLng(mudtButtonRect.Bottom - .y) - 1
'
'             End Select
'
'        UpdateCaption
'        'TextOut lhdcMem, .x, .y, m_Caption, Len(m_Caption)
'        End With
        '''''''''''''''''''''
UpdateTextCoordinates
        
        
    End If
    Exit Sub
PositionChangedError:
    Exit Sub
End Sub

Private Sub DrawButtonState(iState As Integer)
'debug.Print "DrawButtonState"
    '-------------------------------------------------------------------------
    'Purpose:   Draw the button in whatever state it needs to be in
    '-------------------------------------------------------------------------
    Dim lhbmMemory As Long
    Dim lhbmMemoryOld As Long
    Dim lhdcMem As Long 'HDC
    Dim lBackColor As Long
    Dim udtPictureRect As RECT
    Dim bUseMask As Boolean
    Dim lhPal As Long
    Dim lhPalOld As Long
    Dim lhbrBack As Long
    Dim bHaveAmbientPalette As Boolean
    
    On Error GoTo DrawButtonState_Error
    If mbPropertiesLoaded Then
        miCurrentState = iState
        udtPictureRect = mudtPictureRect
        On Error Resume Next
        'Error will occur if the Ambient.Palette is not supported
        bHaveAmbientPalette = (Not UserControl.Ambient.Palette Is Nothing)
        If Err.Number <> 0 Then bHaveAmbientPalette = False
        Err.Clear
        If bHaveAmbientPalette Then
            'If the Palette or hPal property fails
            'resume next and use the halftone palette
            lhPal = UserControl.Ambient.Palette.hPal
            If lhPal = 0 Then lhPal = mlhHalftonePal
            Err.Clear
        Else
            lhPal = mlhHalftonePal    'If there is no specified palette
                                      'use the halftone palette.
        End If
        On Error GoTo DrawButtonState_Error
        'If button is sunken offset the picture coordinates
        'so that the picture looks like it is in sunken
        'perspective
        If iState = giSUNKEN Then
            With udtPictureRect
                .Right = .Right + glSUNKEN_OFFSET
                .Left = .Left + glSUNKEN_OFFSET
                .Top = .Top + glSUNKEN_OFFSET
                .Bottom = .Bottom + glSUNKEN_OFFSET
            End With
        End If
        
        'Create memory DC and bitmap to do all of the painting work
        lhdcMem = CreateCompatibleDC(UserControl.hdc)
        lhbmMemory = CreateCompatibleBitmap(UserControl.hdc, mudtButtonRect.Right, mudtButtonRect.Bottom)
        lhbmMemoryOld = SelectObject(lhdcMem, lhbmMemory)
        lhPalOld = SelectPalette(lhdcMem, lhPal, True)
        RealizePalette lhdcMem
        
        'fill the memory DC with the background color of the screen dc
        OleTranslateColor UserControl.BackColor, 0, lBackColor
        SetBkColor lhdcMem, lBackColor
        lhbrBack = CreateSolidBrush(lBackColor)
        FillRect lhdcMem, mudtButtonRect, lhbrBack
        If Not m_picPictured Is Nothing Then
            If m_picPictured.Type = vbPicTypeBitmap Then
                If m_bUseMaskColor Then bUseMask = True
            End If
            If Not m_bEnabled Then
                'If button is disabled draw disabled picture on memory dc
                moDrawTool.DrawDisabledPicture lhdcMem, m_picPictured, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lBackColor, bUseMask, m_clrMaskColor, lhPal
            ElseIf bUseMask Then
                'if using mask color draw transparent bitmap on memory dc
                moDrawTool.DrawTransparentBitmap lhdcMem, m_picPictured, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, m_clrMaskColor, lhPal
                'moDrawTool.DrawTransparentBitmap lhdcMem, m_picPictured, udtPictureRect.Left, udtPictureRect.Top, 16, 16, m_clrMaskColor, lhPal
            Else
                'otherwise draw picture with no effects on button
                If m_picPictured.Type = vbPicTypeBitmap Then
                    moDrawTool.DrawBitmapToHDC lhdcMem, m_picPictured, udtPictureRect.Left, udtPictureRect.Top, mudtPicturePoint.x, mudtPicturePoint.y, lhPal
                ElseIf m_picPictured.Type = vbPicTypeIcon Then
                    'DrawIcon lhdcMem,  udtPictureRect.Left, udtPictureRect.Top, m_picPictured.Handle
                    If m_ShrinkIcon = True Then
                    DrawIconEx lhdcMem, udtPictureRect.Left, udtPictureRect.Top, m_picPictured.Handle, m_IconWidth, m_IconHeight, 0, lhbrBack, DI_NORMAL
                    Else
                    DrawIconEx lhdcMem, udtPictureRect.Left, udtPictureRect.Top, m_picPictured.Handle, mudtPicturePoint.x, mudtPicturePoint.y, 0, lhbrBack, DI_NORMAL
                    End If
                End If
            End If
        End If
        
DrawButtonState_DrawFrame:
        'Draw Frame Needed
        If iState = 0 And m_PopUpButton = False Then iState = giRAISED
        Select Case iState
            Case giFLATTENED, giDISABLED
            'debug.Print "m_PopUpButton : "; m_PopUpButton
                If m_PopUpButton Then
                If Not UserControl.Ambient.UserMode Then
                    DrawEdge lhdcMem, mudtButtonRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
                End If
                End If
            Case giRAISED
                DrawEdge lhdcMem, mudtButtonRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
            Case giSUNKEN
                DrawEdge lhdcMem, mudtButtonRect, BDR_SUNKENOUTER, BF_RECT Or BF_SOFT
        End Select
        
        'Copy the destination to the screen
        BitBlt UserControl.hdc, 0, 0, mudtButtonRect.Right, mudtButtonRect.Bottom, lhdcMem, 0, 0, vbSrcCopy

DrawButtonStateCleanUp:
        DeleteObject lhbrBack
        SelectPalette lhdcMem, lhPalOld, True
        RealizePalette (lhdcMem)
        DeleteObject SelectObject(lhdcMem, lhbmMemoryOld)
        DeleteDC lhdcMem
    End If
    Exit Sub
DrawButtonState_Error:
    Select Case Err.Number
        Case giOBJECT_VARIABLE_NOT_SET
            Resume DrawButtonState_DrawFrame
        Case giINVALID_PICTURE
            Resume DrawButtonState_DrawFrame
        Case Else
            Resume DrawButtonStateCleanUp
    End Select
End Sub

'*************************
'Friend Methods
'*************************
Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    '-------------------------------------------------------------------------
    'Purpose:   Handles window messages specific to subclassed window associated
    '           with this object.  Is called by SubWndProc in standard module.
    '           Relays all mouse messages to ToolTip window, and returns a value
    '           for ToolTip NeedText message.
    '-------------------------------------------------------------------------
    Dim msgStruct As MSG
    Dim hdr As NMHDR
    Dim ttt As ToolTipText
    On Error GoTo WindowProc_Error
    Select Case uMsg
        Case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP
            With msgStruct
                .lParam = lParam
                .wParam = wParam
                .message = uMsg
                .hWnd = hWnd
            End With
            SendMessage gHWndToolTip, TTM_RELAYEVENT, 0, msgStruct
        Case WM_NOTIFY
            CopyMemory hdr, ByVal lParam, Len(hdr)
            If hdr.code = TTN_NEEDTEXT And hdr.hwndFrom = gHWndToolTip Then
                'Get the tooltip text from the UserControl class object
                'If the host for this control provides a ToolTipText property
                'on the extender object (as in VB5).  The ToolTipText property
                'declares will not be hit.  Therefore, the user's ToolTipText
                'may be found either in the Extender.ToolTipText property or
                'in my own member variable m_sToolTipText
                'Error may occur if ToolTipText property is not available
                On Error Resume Next
                If mbToolTipNotInExtender Then
                    msToolTipBuffer = StrConv(m_sToolTipText, vbFromUnicode)
                Else
                    msToolTipBuffer = StrConv(UserControl.Extender.ToolTipText, vbFromUnicode)
                End If
                If Err.Number = 0 Then
                    CopyMemory ttt, ByVal lParam, Len(ttt)
                    ttt.lpszText = StrPtr(msToolTipBuffer)
                    CopyMemory ByVal lParam, ttt, Len(ttt)
                End If
            End If
        Case WM_CANCELMODE
            'A window has been put over this one
            'flatten the button
            Flatten
            mbGotFocus = False
            mbLeftMouseDown = False
            mbLeftWasDown = False
            mbMouseDownFired = False
    End Select
WindowProc_Resume:
    WindowProc = CallWindowProc(mWndProcNext, hWnd, uMsg, wParam, ByVal lParam)
    Exit Function
WindowProc_Error:
    Resume WindowProc_Resume
End Function


Public Property Get PictureAlignment() As PictureAlignment
'debug.Print "Get PictureAlignment"
PictureAlignment = m_PicAlign
End Property

Public Property Let PictureAlignment(ByVal NewValue As PictureAlignment)
'debug.Print "Let PictureAlignment"
m_PicAlign = NewValue
PropertyChanged "PictureAlignment"
Refresh
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get Caption"
Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
'debug.Print "Let Caption"
m_Caption = NewValue
PropertyChanged "Caption"
UpdateCaption
Refresh
End Property

Public Sub UpdateCaption()
'debug.Print "UpdateCaption"
If Len(m_Caption) <> 0 Then
    'debug.Print mudText.x; mudText.y; miCurrentState
    If m_picPictured Is Nothing Then
        UpdateTextCoordinates
        TextOut UserControl.hdc, mudText.x, mudText.y, m_Caption, Len(m_Caption)
        
        Else
        
        Select Case miCurrentState
            Case giSUNKEN
            TextOut UserControl.hdc, mudText.x, mudText.y + 1, m_Caption, Len(m_Caption)
            Case giRAISED
            TextOut UserControl.hdc, mudText.x, mudText.y, m_Caption, Len(m_Caption)
            Case Else
            TextOut UserControl.hdc, mudText.x, mudText.y, m_Caption, Len(m_Caption)
        End Select
        
    End If
End If
End Sub

Public Sub Refresh()
'debug.Print "Refresh"
UserControl_Resize
End Sub

Public Property Get ShrinkIcon() As Boolean
Attribute ShrinkIcon.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get ShrinkIcon"
ShrinkIcon = m_ShrinkIcon
End Property

Public Property Let ShrinkIcon(ByVal NewValue As Boolean)
'debug.Print "Let ShrinkIcon"
m_ShrinkIcon = NewValue
PropertyChanged "ShrinkIcon"
Refresh
End Property

Public Property Get IconHeight() As Integer
Attribute IconHeight.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get IconHeight"
IconHeight = m_IconHeight
End Property

Public Property Let IconHeight(ByVal NewValue As Integer)
'debug.Print "Let IconHeight"
If NewValue > 0 Then
m_IconHeight = NewValue
Refresh
Else
MsgBox "Invalid Value", vbCritical, "Soft Button"
End If
End Property

Public Property Get IconWidth() As Integer
Attribute IconWidth.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get IconWidth"
IconWidth = m_IconWidth
End Property

Public Property Let IconWidth(ByVal NewValue As Integer)
'debug.Print "Let IconWidth"
If NewValue > 0 Then
m_IconWidth = NewValue
Refresh
Else
MsgBox "Invalid Value", vbCritical, "Soft Button"
End If

End Property
Public Property Get TextAlignment() As PictureAlignment
'debug.Print "Get TextAlignment"
TextAlignment = m_TextAlign
End Property

Public Property Let TextAlignment(ByVal NewValue As PictureAlignment)
'debug.Print "Let TextAlignment"
m_TextAlign = NewValue
PropertyChanged "TextAlignment"
Refresh
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
'debug.Print "Get Font"
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
'debug.Print "Set Font"
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = UserControl.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    UserControl.FontTransparent() = New_FontTransparent
    PropertyChanged "FontTransparent"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AccessKeys
Public Property Get HotKey() As String
Attribute HotKey.VB_Description = "Returns or sets a string that contains the keys that will act as the access keys (or hot keys) for the control."
Attribute HotKey.VB_ProcData.VB_Invoke_Property = "General"
    HotKey = UserControl.AccessKeys
End Property

Public Property Let HotKey(ByVal New_HotKey As String)
    UserControl.AccessKeys() = New_HotKey
    PropertyChanged "HotKey"
End Property


Public Property Get PopUpButton() As Boolean
Attribute PopUpButton.VB_ProcData.VB_Invoke_Property = "General"
'debug.Print "Get PopUpButton"
PopUpButton = m_PopUpButton
End Property

Public Property Let PopUpButton(ByVal NewValue As Boolean)
'debug.Print "Let PopUpButton"
m_PopUpButton = NewValue
PropertyChanged "PopUpButton"
Refresh
End Property

Public Sub UpdateTextCoordinates()
GetTextExtentPoint32 UserControl.hdc, m_Caption, Len(m_Caption), mudText
        
With mudText
Select Case m_TextAlign

    Case TopLeft 'OK
    .x = 1
    .y = 1
    
    Case TopCenter 'OK
    .x = CLng((mudtButtonRect.Right - .x) / 2)
    .y = 1
    
    Case TopRight 'OK
    .x = CLng((mudtButtonRect.Right - .x)) - 1
    .y = 1
    
    Case LeftCenter 'OK
    .x = 1
    .y = CLng((mudtButtonRect.Bottom - .y) / 2)
    
    Case Center 'OK
    .x = CLng((mudtButtonRect.Right - .x) / 2)
    .y = CLng((mudtButtonRect.Bottom - .y) / 2)
    
    Case RightCenter 'OK
    .x = CLng((mudtButtonRect.Right - .x)) - 1
    .y = CLng((mudtButtonRect.Bottom - .y) / 2)
    
    Case BottomLeft 'OK
    .x = 1
    .y = CLng(mudtButtonRect.Bottom - .y) - 1
    
    Case BottomCenter 'OK
    .x = CLng((mudtButtonRect.Right - .x) / 2)
    .y = CLng(mudtButtonRect.Bottom - .y) - 1
    
    Case BottomRight 'OK
    .x = CLng((mudtButtonRect.Right - .x)) - 1
    .y = CLng(mudtButtonRect.Bottom - .y) - 1

End Select
'UpdateCaption

End With
End Sub

