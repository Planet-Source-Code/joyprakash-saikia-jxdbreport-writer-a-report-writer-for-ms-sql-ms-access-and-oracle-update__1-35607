Attribute VB_Name = "modSoftButton"
Option Explicit
'************************************************************************
'Author            :   Vijay Phulwadhawa     Date    : 23/02/2001 12:59:27 PM
'Project Name      :   Insert_Project_Name
'Form/Class Name   :   modSoftButton (Code)
'Version           :   6.00
'Description       :   <Purpose>
'Links             :   <Links With Any Other Form Modules>
'Change History    :
'Date      Author      Description Of Changes          Reason Of Change
'************************************************************************


'-------------------------------------------------------------------------
'This module provides all needed Type, API, and Constant declarations
'-------------------------------------------------------------------------

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    rct As RECT
    hinst As Long
    lpszText As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Type ToolTipText
    hdr As NMHDR
    lpszText As Long
    szText As String * 80
    hinst As Long
    uFlags As Long
End Type

'vijay
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Const DI_IMAGE = &H2
Public Const DI_MASK = &H1
Public Const DI_NORMAL = &H3
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public DestSize As POINTAPI
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Const MOD_ALT = &H1
#If UNICODE Then
    Public Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
    Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#Else
    Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
    Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

'Misc Constants
Public Const H_MAX As Long = &HFFFFFFFF + 1
Public Const TOOLTIPS_CLASS As String = "tooltips_class32"
Public Const WS_EX_TOPMOST = &H8&
Public Const CW_USEDEFAULT  As Long = &H80000000
Public Const glSUNKEN_OFFSET = 1
Public Const GDI_ERROR = &HFFFFFFFF

'Windows Messages
Public Const WM_CANCELMODE = &H1F

'Resource String Indexes
Public Const giINVALID_PIC_TYPE As Integer = 10

'Get Windows Long Constants
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)

'Draw State constants
'Image type
Public Const DST_ICON = &H3&
Public Const DST_BITMAP = &H4&
'State type
Public Const DSS_DISABLED = &H20&

'Raster Operation Codes
Public Const PSDPxax = &HB8074A
Public Const DSna = &H220326 '0x00220326

'System colors
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16

'Messages to relay to ToolTip
Public Const WM_USER = &H400
Public Const WM_NOTIFY = &H4E
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208

'ToolTip style
Public Const TTF_IDISHWND = &H1

'Tool Tip messages
Public Const TTM_ACTIVATE = (WM_USER + 1)
#If UNICODE Then
    Public Const TTM_ADDTOOLW = (WM_USER + 50)
    Public Const TTM_ADDTOOL = TTM_ADDTOOLW
#Else
    Public Const TTM_ADDTOOLA = (WM_USER + 4)
    Public Const TTM_ADDTOOL = TTM_ADDTOOLA
#End If
Public Const TTM_RELAYEVENT = (WM_USER + 7)

'ToolTip Notification
Public Const TTN_FIRST = (H_MAX - 520&)
#If UNICODE Then
    Public Const TTN_NEEDTEXTW = (TTN_FIRST - 10&)
    Public Const TTN_NEEDTEXT = TTN_NEEDTEXTW
#Else
    Public Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
    Public Const TTN_NEEDTEXT = TTN_NEEDTEXTA
#End If

'Misc ToolTip
Public Const LPSTR_TEXTCALLBACK As Long = -1

'DrawEdge constants
Public Const BDR_RAISEDOUTER As Long = &H1
Public Const BDR_SUNKENOUTER As Long = &H2

' Border flags
Public Const BF_LEFT As Long = &H1
Public Const BF_TOP As Long = &H2
Public Const BF_RIGHT As Long = &H4
Public Const BF_BOTTOM As Long = &H8
Public Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_SOFT As Long = &H1000      ' For softer buttons

Public Const SND_SYNC = 0
Public Const EVENT_MENU_COMMAND = "MenuCommand"
Public Const EVENT_MENU_POPUP = "MenuPopup"


'Button States
Public Const giFLATTENED As Integer = 0
Public Const giRAISED As Integer = 1
Public Const giSUNKEN As Integer = 3
Public Const giDISABLED As Integer = 4

'VB Errors
Public Const giOBJECT_VARIABLE_NOT_SET As Integer = 91
Public Const giINVALID_PICTURE As Integer = 481
Public Const giDLL_FUNCTION_NOT_FOUND As Integer = 453

'Windows Errors
Public Const ERROR_CALL_NOT_IMPLEMENTED As Long = 120

