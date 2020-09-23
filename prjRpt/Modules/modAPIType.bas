Attribute VB_Name = "modAPIType"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781EE00D2"
Option Explicit
'**********************************************************************
'
'           Module Name: modApiType.bas
'           Purpose    : Declare User Defined Type used by WINDOWS API
'           Author     : Joyprakash Saikia
'**********************************************************************
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type size
    x As Long
    y As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type
'This  Structure is Used for Drawing or manipulating A window region
' Here It is used for Printing A rectangular Region to Window or Printer
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'The  Following Structure is used by Commom dialog Control to Gather the Information
' from Open , Save , Saveas Etc.
Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'This Structure is needed by the "PrintDialogA" Function of Common Dialog Control Directly

Public Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Public Type PAGESETUPDLG
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
'Used by "PrintDialogA" function for retriving or Setting for Printer Object
Public Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPixel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    ' The following only appear in Windows 95, 98, 2000
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    ' The following only appear in Windows 2000
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

' For Device Name ( printers)
Public Type DEVNAMES
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 100
End Type
