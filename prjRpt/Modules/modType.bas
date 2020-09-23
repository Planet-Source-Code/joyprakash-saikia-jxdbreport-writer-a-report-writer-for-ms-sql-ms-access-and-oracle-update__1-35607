Attribute VB_Name = "modType"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781EE01E0"
Option Explicit
'**********************************************************************
'           Module Name: modAPIConst.bas
'
'           Purpose    : Declare User Defined Type (UDT) used internally by the application
'           Created On :
'
'**********************************************************************
'for error handling
Public Type UDTErrSave
    ErrNumber As Long 'error number
    ErrSource As String 'error source
    ErrDesc As String 'err source
    ErrHelpFile As String 'help file
    ErrHelpContext As Long 'help context id
End Type

'report data file signature
Type UDTRptPageHeader
    NewPage As String * 3 ' = "PG#"
    PageNo As Long ' page number starting from one onwards
End Type
'report data file item
Type UDTRptItem
    FontName As String * 20 ' 20 bytes
    XPos As Single '4 bytes
    YPos As Single '4 bytes
    FontSize As Single '4 bytes
    FontUnderline As Byte '1 byte
    FontBold As Byte '1
    FontStrikethru As Byte '1
    FontItalic As Byte '1
    Allign As Integer '
    ItemType As Integer '2
    FontColor As Long '4
    PopUpValue As String 'to Displaying Popup For an Anchor Field
    Value As String * 255 'is it enough?
End Type
Type UDTRptPageIndex
    PageNo As Long 'report page #
    PageSize As Long 'report page size without Page header size
    PageOffset As Long 'offset of starting page in the report data file
End Type

Public ItemInstance_Count As Long
Public GroupInstance_Count As Long
