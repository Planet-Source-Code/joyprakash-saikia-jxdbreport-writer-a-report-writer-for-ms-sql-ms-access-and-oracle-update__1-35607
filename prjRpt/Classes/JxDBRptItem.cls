VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JxDBRptItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781E30000"
'**********************************************************************
'
'           Module Name: JxDBRptItem.cls
'
'           Purpose    : Hold information about report item to be printed
'
'           Author     : Joyprakash saikia
'
'**********************************************************************

Option Explicit
'This Constant is Used for Debugging Purpose
'This is not used on this submission
Private Const MOD_NAME = "JxDBReport.JxDBRptItem"
'The Following Variables are used by the Font and Other Properties

Private mvarFont As StdFont
Private mvarFontName As String
Private mvarFontSize As Integer
Private mvarFontBold As Boolean
Private mvarFontUnderline As Boolean
Private mvarFontStrikeThrough As Boolean
Private mvarFormatString As String
Private mvarPreAdvanceLine As Boolean
Private mvarPostAdvanceLine As Boolean
Private mvarXPosition As Single
Private mvarYPosition As Single
Private mvarItemName As String

'   Report item enum
'   If you want increase Item Type you Can put it Here
Public Enum JxDBRptItemType
    JxDBItemLabel = 0
    JxDBItemCounter = 1
    JxDBItemDataField = 2
    JxDBItemCurrentDate = 3
    JxDBItemCurrentPage = 4
    JxDBItemTotalPage = 5
    JxDBItemFormula = 6     ' For Report Defined Formula
    JxDBItemRunningTotal = 7 'for Accumulated Sum
    JxDBItemPageXofY = 8
    JxDBItemCustom = 256 'custom item(allowed to be customised during runtime)
    JxDBItemLabelAnchor = 9 'For Future Use (might be on next version)
End Enum
'Printing allignment enum
Public Enum JxDBPrintAllign
    JxDBAllignLeft = 0
    JxDBAllignCenter = 1
    JxDBAllignRight = 2
End Enum
' Here is the Report Define formulas
' The possible Other formuli is Standard Deviation

Public Enum JxDBRptFormulaType
    JxDBRptFormulaSum = 0
    JxDBRptFormulaAverage = 1
    JxDBRptFormulaMedian = 2
    JxDBRptFormulaHighest = 3
    JxDBRptFormulaLowest = 4
    JxDBRptFormulaRunningTotal = 5
End Enum
'The Following Variables are used by the Other Properties
Private mvarFormulaType As JxDBRptFormulaType
Private mvarValue As Variant
Private mvarFieldName As String
Private mvarPrintAllign As JxDBPrintAllign
Private mvarItemType As JxDBRptItemType
Private mvarValueExtra As Variant 'Used internally by the application for 1.Average Calculation 2. Highest 3. Lowest
Private mvarHeight As Single 'item height in Twips
Private mvarWidth As Single
Private mvarFontColor As Long

Public Property Let FontColor(ByVal vData As Long)
    '************************************************************
    'Description :
    '          Sets the Fore Color of an Item in BBGGBB
    '************************************************************
    mvarFontColor = vData
End Property


Public Property Get FontColor() As Long
    '************************************************************
    'Description :
    '          Returns the ForeColor of an Item in BBGGBB format
    '************************************************************

    FontColor = mvarFontColor
End Property

Public Property Let Width(ByVal vData As Single)
    '************************************************************
    'Note:
    '        Due to the Name of this property, sometimes VB Shows Conflict with VBA Width Function
    '        If you Get this Problem ,Please Change it as you Like.
    '************************************************************

    mvarWidth = vData
End Property
Public Property Get Width() As Single
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Single)
    '************************************************************
    'Note :
    '        Due to the Name of this property, sometimes VB Shows Conflict with VBA
    '        Height Function
    '        If you Get this Problem ,Please Change it as you Like.
    '************************************************************
    mvarHeight = vData
End Property


Public Property Get Height() As Single
    Height = mvarHeight
End Property

Friend Property Let ValueExtra(ByVal vData As Variant)
    mvarValueExtra = vData
End Property
Friend Property Get ValueExtra() As Variant
    ValueExtra = mvarValueExtra
End Property

Public Property Let ItemType(ByVal vData As JxDBRptItemType)
    '************************************************************
    'Description :
    '          Sets the Type of an Item
    '************************************************************

    mvarItemType = vData
End Property
Public Property Get ItemType() As JxDBRptItemType
     ItemType = mvarItemType
End Property

Public Property Let PrintAllign(ByVal vData As JxDBPrintAllign)
    '************************************************************
    'Description :
    '          Sets the Allignment of an Item
    '************************************************************

    mvarPrintAllign = vData
End Property


Public Property Get PrintAllign() As JxDBPrintAllign
    PrintAllign = mvarPrintAllign
End Property



'##ModelId=
Public Property Let FieldName(ByVal vData As String)
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
    FieldName = mvarFieldName
End Property

Public Property Let Value(ByVal vData As Variant)
    mvarValue = vData
End Property


Public Property Get Value() As Variant
    Value = mvarValue
    
End Property

Public Property Let FormulaType(ByVal vData As JxDBRptFormulaType)
    mvarFormulaType = vData
End Property
Public Property Get FormulaType() As JxDBRptFormulaType
    FormulaType = mvarFormulaType
End Property

Public Property Let ItemName(ByVal vData As String)
    mvarItemName = vData
End Property


Public Property Get ItemName() As String
    ItemName = mvarItemName
End Property

Public Property Let YPosition(ByVal vData As Single)
    mvarYPosition = vData
End Property


Public Property Get YPosition() As Single
    YPosition = mvarYPosition
End Property
Public Property Let XPosition(ByVal vData As Single)
    mvarXPosition = vData
End Property


Public Property Get XPosition() As Single
    XPosition = mvarXPosition
End Property




Public Property Let PostAdvanceLine(ByVal vData As Boolean)
    mvarPostAdvanceLine = vData
End Property


Public Property Get PostAdvanceLine() As Boolean
    PostAdvanceLine = mvarPostAdvanceLine
End Property

Public Property Let PreAdvanceLine(ByVal vData As Boolean)
    mvarPreAdvanceLine = vData
End Property


Public Property Get PreAdvanceLine() As Boolean
    PreAdvanceLine = mvarPreAdvanceLine
End Property

Public Property Let FormatString(ByVal vData As String)
    mvarFormatString = vData
End Property


Public Property Get FormatString() As String
    FormatString = mvarFormatString
End Property


Public Property Let FontStrikeThrough(ByVal vData As Boolean)
    mvarFont.Strikethrough = vData
End Property


Public Property Get FontStrikeThrough() As Boolean
    FontStrikeThrough = mvarFont.Strikethrough
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    mvarFontUnderline = vData
End Property


Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarFontUnderline
End Property
Public Property Let FontBold(ByVal vData As Boolean)
    mvarFont.Bold = vData
End Property


Public Property Get FontBold() As Boolean
    FontBold = mvarFont.Bold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    mvarFont.Italic = vData
End Property


Public Property Get FontItalic() As Boolean
    FontItalic = mvarFont.Italic
End Property

Public Property Let FontSize(ByVal vData As Integer)
    mvarFont.size = vData
End Property


Public Property Get FontSize() As Integer
    FontSize = mvarFont.size
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFont.Name = vData
End Property


Public Property Get FontName() As String
    FontName = mvarFont.Name
End Property

Public Property Set Font(ByVal vData As StdFont)
    Set mvarFont = vData
End Property
Public Property Get Font() As StdFont
    Set Font = mvarFont
End Property

Private Sub Class_Initialize()
        '************************************************************
        '          Description:
        'Initializes Default Font and its properties
        '************************************************************
    Set mvarFont = New StdFont
    Me.Font.Name = "Courier New"
    Me.Font.size = 8
    Me.Height = 200
End Sub

