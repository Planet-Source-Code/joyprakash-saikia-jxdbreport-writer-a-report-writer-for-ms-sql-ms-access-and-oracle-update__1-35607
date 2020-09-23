VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTest 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhelp 
      BackColor       =   &H80000016&
      Caption         =   "See &Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Preview Report with all header"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      MaskColor       =   &H00D38525&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   3345
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000016&
      Caption         =   "Print Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1230
      Width           =   3345
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000016&
      Caption         =   "Preview report without Second Group "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      MaskColor       =   &H00D38525&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   3345
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Page Orientation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   1
      Left            =   5280
      TabIndex        =   17
      Top             =   1620
      Width           =   2145
      Begin VB.OptionButton porientation 
         BackColor       =   &H80000016&
         Caption         =   "Portrait"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton porientation 
         BackColor       =   &H80000016&
         Caption         =   "Landscape"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   6
         Top             =   810
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   150
         Picture         =   "frmTest.frx":0000
         Top             =   210
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   1
         Left            =   120
         Picture         =   "frmTest.frx":0AC2
         Top             =   690
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Page Margins (inch)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   1
      Left            =   3780
      TabIndex        =   12
      Top             =   420
      WhatsThisHelpID =   20001
      Width           =   3675
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   2
         Left            =   2700
         TabIndex        =   3
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox pagemargin 
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   4
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   4
         Mask            =   "#.##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   15
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Top:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   14
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1860
         TabIndex        =   13
         Top             =   540
         Width           =   825
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTest.frx":152C
      Left            =   6180
      List            =   "frmTest.frx":152E
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   1245
   End
   Begin VB.Label lblVote 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Please Vote me If you Like it "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   750
      TabIndex        =   18
      Top             =   3090
      Width           =   5265
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size"
      Height          =   285
      Left            =   4980
      TabIndex        =   11
      Top             =   30
      Width           =   1125
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Dim WithEvents oRpt As JxDBReport
Attribute oRpt.VB_VarHelpID = -1
Private Sub SetupReportNew()
    Dim Xposition As Long
    Set oRpt = New JxDBReport
    Dim oRptItem As JxDBRptItem

    Dim oGroup As JxDBRpt.JxDBRptGroup
    'Dim TRG0
    Dim SSQL As String
    oRpt.PaperSize = Combo1.ListIndex
    oRpt.ReportTitle = "Car Model Master"
    'Page header setup
    Set oRptItem = New JxDBRptItem

    'set Page Header
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Car Model Master"
        .Font.Name = "Arial"
        .FontBold = True
        .Xposition = Xposition
        .Font.Size = 14
        .PrintAllign = JxDBAllignCenter
        .PostAdvanceLine = 1
        .Height = 350
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmRptTitle"

    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemCurrentDate
        .Font.Name = "Arial"
        .FormatString = "hh:mm dd/mm/yyyy"
        .Xposition = Xposition ' set to zero allow page right justified
        .PrintAllign = JxDBAllignRight
        .PostAdvanceLine = 1
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmDate"
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Customer Code"
        .Font.Bold = True
        .Xposition = (Xposition + 10)
        .FontName = "Arial"
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmLblVendor"
    Set oRptItem = New JxDBRptItem
    Xposition = Xposition + 2000
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Customer Name"
        '.Font.Name = "Arial"
        .Font.Bold = True
        .Xposition = 3000 'Xposition
        .PrintAllign = JxDBAllignLeft
        .PostAdvanceLine = 1
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmLblVendorName"
    oRpt.PageHeader.RepeatHeader = True

    'we're done with page header





    'New Terminal Name Group
    Set oGroup = New JxDBRptGroup
    oGroup.GroupName = "trml_nm"

    Set oRptItem = New JxDBRptItem

    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Treminal Name"
        '.Font.Name = "Arial"
        .Font.Bold = True
        .Xposition = 500 'Xposition
        .PrintAllign = JxDBAllignLeft
    End With
    oGroup.AddPrintItem oRptItem
    oGroup.AddGroupField "trml_nm"
    oGroup.CheckForBreak = True
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "trml_nm"
        .Xposition = 2000
        .PostAdvanceLine = 1
        .FontBold = True
        .FontUnderline = True
        .Font.Name = "Arial"
        .Font.Italic = True

        .FontSize = 9
    End With

'    oGroup.AddPrintItem oRptItem
   oGroup.AddBreakItem oRptItem
     oGroup.PrintGroupBreak = True

    'group  break field for vendor group

    oRpt.ReportGroups.Add oGroup

     'GROUP Customer
    Set oGroup = New JxDBRptGroup
    oGroup.GroupName = "Customer Code"
    oGroup.AddGroupFields "Cust_c", "Cust_nm"
    'vendor ID
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "Cust_c"
        .Xposition = 10
        .FontBold = True
         .FontSize = 9
         .PreAdvanceLine = 1
       ' .Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    'vendor Name
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "cust_nm"
        .Xposition = 1000
        .FontSize = 9
        .PostAdvanceLine = 1
        .FontBold = True
        '.Font.Name = "Arial"
        .FontColor = vbBlue
    End With
    oGroup.AddPrintItem oRptItem
    oGroup.PrintGroupBreak = True
    oRpt.ReportGroups.Add oGroup


    'GROUP Detail
    Set oGroup = New JxDBRptGroup
    oGroup.AddGroupFields Trim$("carmdl_c"), Trim$("carmdl_nm"), Trim$("user_id"), "ent_dt", "trml_nm"
    oGroup.CheckForBreak = True
    oGroup.GroupName = "detail"
    'oGroup.PrintGroup = False 'flag for group printing
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "carmdl_c"
        .Xposition = 2000
        .FontSize = 9
        .Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    'amount field
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "Carmdl_nm"
        .Font.Name = "Arial"
        .Xposition = 3000
        '.FormatString = "###,##0.00"
        .PrintAllign = JxDBAllignLeft
        .FontSize = 9
        '.PostAdvanceLine = 1
    End With
    oGroup.AddPrintItem oRptItem

'    Set oRptItem = New JxDBRptItem
'    With oRptItem
'        .ItemType = JxDBItemDataField ' JxDBItemLabel
'        .Value = "user_id"
'        .FieldName = "user_id"
'        .Xposition = 5000
'        '.FormatString = "###,##0.00"
'        .PrintAllign = JxDBAllignLeft
'        .Font.Name = "Arial"
'        .FontSize = 9
'    End With
'    oGroup.AddPrintItem oRptItem
    'update Date
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField ' JxDBItemLabel
        .FieldName = "upd_dt"
        .Xposition = 6000
        .FormatString = "YYYY/MM/DD"
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
        .FontSize = 9
    End With
    oGroup.AddPrintItem oRptItem
    'Entered Date
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField ' JxDBItemLabel

        .FieldName = "ent_dt"
        .Xposition = 7000
        .FormatString = "YYYY/MM/DD"
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"

        .PostAdvanceLine = 1
        .FontSize = 9
    End With
    oGroup.AddPrintItem oRptItem
    oGroup.PrintGroupBreak = True

    oRpt.ReportGroups.Add oGroup, "detail"



  Dim DB As Connection
  Set DB = New Connection
  Dim rst As Recordset
  Dim sDbPath As String

  sDbPath = App.Path & "\test.mdb"
  DB.CursorLocation = adUseClient
  'Db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDbPath 'C:\MY DOCUMENTS\test.mdb;"
  DB.Open "PROVIDER=SQLOLEDB;Data Source=nilgiri;Initial Catalog=acsdb;user id=acsuser1;password=admin"

  Set rst = New Recordset

  'this SQL stmt is generated by MS Access 2000 SQL wizard
  'to get the group break correctly, the ORDER clause must satisfy the report grouping
'  sSQL = "SELECT [invoice].[cocode], [invoice].[vendor], [invoice].[invoice], [invoice].[item], [invoice].[amount], [invoice_sum].[inv_date], [vendor].[vendor_name] " _
'        & " FROM (invoice INNER JOIN invoice_sum ON ([invoice].[invoice]=[invoice_sum].[invoice]) AND ([invoice].[vendor]=[invoice_sum].[vendor]) AND ([invoice].[cocode]=[invoice_sum].[cocode])) INNER JOIN vendor ON [invoice_sum].[vendor]=[vendor].[id] " _
'        & "ORDER BY invoice.vendor,invoice.cocode, invoice.invoice, invoice.item"
    SSQL = "SELECT DISTINCT  carmdl_mst.trml_nm,carmdl_mst.cust_c,customer_mst.cust_nm,  carmdl_mst.carmdl_c, carmdl_mst.carmdl_nm,carmdl_mst.user_id,carmdl_mst.upd_dt,carmdl_mst.ent_dt" & _
           " From customer_mst , carmdl_mst " & _
           " Where  customer_mst.cust_c = carmdl_mst.cust_c " & _
           " Order By  carmdl_mst.trml_nm, carmdl_mst.cust_c,customer_mst.cust_nm,carmdl_mst.carmdl_c,carmdl_mst.carmdl_nm,carmdl_mst.user_id,carmdl_mst.upd_dt,carmdl_mst.ent_dt"

    rst.Open SSQL, DB, adOpenForwardOnly, adLockReadOnly
    Set rst.ActiveConnection = Nothing
    Set oRpt.DataSource = rst

    'set report margin
    oRpt.SetBottomMargin 1.25, JxDBRptScaleInch
    oRpt.SetTopMargin 0.25, JxDBRptScaleInch
    oRpt.SetPageFooterMargin 0.5, JxDBRptScaleInch
    oRpt.SetLeftMargin 1000
    oRpt.SetRightMargin 2, JxDBRptScaleInch

End Sub
Private Sub SetupReport()

  
    Dim oRptItem As JxDBRptItem

    Dim oGroup As JxDBRptGroup
    Dim SSQL As String
    If oRpt Is Nothing Then
        Set oRpt = New JxDBReport
    End If
    oRpt.PaperSize = Combo1.ItemData(Combo1.ListIndex)
    oRpt.ReportTitle = "Invoice Listing"
    If porientation(0).Value = True Then
        oRpt.PageOrientation = JxDBRptPotrait
    Else
        oRpt.PageOrientation = JxDBRptLandscape
    End If
    Set oRptItem = New JxDBRptItem

    'set Page Header
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Invoice Listing"
        .Font.Name = "Arial"
        .FontBold = True
        .Xposition = 0 ' set to zero allow page Left justified
        .Font.Size = 18
        .PrintAllign = JxDBAllignLeft
        .PostAdvanceLine = 1
        .Height = 350
        .FontColor = vbBlue
    End With


    oRpt.PageHeader.AddItem oRptItem, "itmRptTitle"
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Time : "
        .Font.Name = "Arial"
        .Xposition = 8200
        .PrintAllign = JxDBAllignLeft
        .FontBold = True
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmlblTime"

    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemCurrentDate
        .Font.Name = "Arial"
        .FormatString = "hh:mm"
        .Xposition = 9800
        .PrintAllign = JxDBAllignRight
        .PostAdvanceLine = 1
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmTime"
    
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Date  : "
        .Font.Name = "Arial"
        .FontBold = True
        .Xposition = 8200
        .PrintAllign = JxDBAllignLeft
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmlblDate"
    
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemCurrentDate
        .Font.Name = "Arial"
        .FormatString = "dd/mm/yyyy"
        .Xposition = 9800
        .PrintAllign = JxDBAllignRight
        .PostAdvanceLine = 1
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmDate"
       'Put a Line
        Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "_____________________________________________________________________________"
        .Xposition = 0
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
        .FontSize = "12"
        .FontBold = True
        .PostAdvanceLine = 1
        .Height = 300
    End With
        oRpt.PageHeader.AddItem oRptItem, "itmLine1"
    'Print Labels with default Font
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Vendor"
        .Font.Bold = True
        .PrintAllign = JxDBAllignLeft
        .Xposition = 10
        .FontSize = 10
        '.FontName = "Arial"
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmLblVendor"
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Vendor Name"
        .Font.Bold = True
        .Xposition = 1000
        .PrintAllign = JxDBAllignLeft
        .FontSize = 10
    End With
    oRpt.PageHeader.AddItem oRptItem, "itmLblVendorName"
    
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "_____________________________________________________________________________"
        .Xposition = 0
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
        .FontSize = "12"
        .FontBold = True
        .PostAdvanceLine = 1
        .Height = 300
    End With
        oRpt.PageHeader.AddItem oRptItem, "itmLine"

    oRpt.PageHeader.RepeatHeader = True
    
'    oRpt.ReportHeader.AddItem oRptItem
'    oRpt.ReportHeader.RepeatHeader = True
    'we're done with page header

    'SETUP cocode GROUP
    Set oGroup = New JxDBRptGroup
    oGroup.GroupName = "cocode"

  
    oGroup.AddGroupField "cocode"
    oGroup.PrintGroupBreak = True
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Company -->"
        
        .Font.Bold = True
         .FontSize = 10
        .Xposition = 10
        .PrintAllign = JxDBAllignLeft
        .PreAdvanceLine = 1
    End With
    oGroup.AddPrintItem oRptItem
    
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .Font.Bold = True
        .ItemType = JxDBItemDataField
        
        .FieldName = "coCode"
        .FontSize = 10
        .Font.Italic = True
        .PrintAllign = JxDBAllignLeft
        .Xposition = 1500
        .FontColor = vbBlue
    End With
   
    oGroup.AddPrintItem oRptItem
     Set oRptItem = New JxDBRptItem
   
    With oRptItem
        
        .ItemType = JxDBItemDataField
        
        .FieldName = "co_name"
        .Font.Underline = True
        .PrintAllign = JxDBAllignLeft
        .Xposition = 1800
        .FontSize = 10
        .PostAdvanceLine = 1
        .FontColor = vbRed
        .Height = 400
    End With
   
    oGroup.AddPrintItem oRptItem
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .Font.Bold = True
        .ItemType = JxDBItemFormula
        .FontName = "Arial"
        .FieldName = "amount"
        .FormulaType = JxDBRptFormulaSum
        .Font.Underline = True
        .FormatString = "###,##0.00"
        .PrintAllign = JxDBAllignRight
        .Xposition = 4000
        .PreAdvanceLine = 1
        .FontColor = vbRed
    End With
    oGroup.AddBreakItem oRptItem


    'invoice total label
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Company Total"
        .Font.Size = 10
        '.Font.Name = "Arial"
        .Font.Bold = True
        .PostAdvanceLine = 1
        .Xposition = 1000
    End With
    oGroup.AddBreakItem oRptItem
    oRpt.ReportGroups.Add oGroup, "company"
   Set oGroup = Nothing

    'GROUP VENDOR
    Set oGroup = New JxDBRptGroup
    oGroup.GroupName = "Vendor"
    oGroup.AddGroupField "vendor"
 
    'vendor ID
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "vendor"
        .Xposition = 10
       ' .Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    'vendor Name
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "vendor_name"
        .Xposition = 1000
        .PostAdvanceLine = 1
        '.Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    oGroup.PrintGroupBreak = True

    'group  break field for vendor group
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemFormula
        '.Font.Name = "Arial"
        .FormulaType = JxDBRptFormulaSum
        .FieldName = "amount"
        .Xposition = 4000
        .Font.Bold = True
        .PrintAllign = JxDBAllignRight
        .FormatString = "###,##0.00"
        .FontColor = vbBlue 'we can have multiple color!
    End With
    oGroup.AddBreakItem oRptItem
    'just a label for total
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Total Vendor"
        .Xposition = 1000
        .PostAdvanceLine = 1
       ' .Font.Name = "Arial"
    End With
    oGroup.AddBreakItem oRptItem
    Set oRptItem = New JxDBRptItem
    With oRptItem
        '.PreAdvanceLine = 1
        .ItemType = JxDBItemLabel
        .Value = "Highest Invoice"
        .Xposition = 1000
        .PrintAllign = JxDBAllignLeft
       ' .Font.Name = "Arial"
       .FontUnderline = True
    End With
    oGroup.AddBreakItem oRptItem, "lblHigh"
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemFormula
        .FormulaType = JxDBRptFormulaHighest
        .FieldName = "amount"
        .FormatString = "###,##0.00"
        .Xposition = 4000
        .PrintAllign = JxDBAllignRight
        .Font.Name = "Arial"
       ' .PostAdvanceLine = 1
    End With
    oGroup.AddBreakItem oRptItem, "itmAmtHigh"
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Smallest Invoice"
        .Xposition = 4200
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
    End With
    oGroup.AddBreakItem oRptItem, "lblSmallest"
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemFormula
        .FormulaType = JxDBRptFormulaLowest
        .FieldName = "amount"
        .FontColor = vbMagenta
        .FormatString = "###,##0.00"
        .Xposition = 7000
        .PrintAllign = JxDBAllignRight
        .Font.Name = "Arial"
        .PostAdvanceLine = 1
        .Height = 300
    End With
    oGroup.AddBreakItem oRptItem, "itmSmallest"


    'add group
    oRpt.ReportGroups.Add oGroup, "vendor"


    'GROUP INVOICE
    Set oGroup = New JxDBRptGroup
    oGroup.GroupName = "invoice"
    oGroup.AddGroupFields "invoice"

    'add counter field
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemCounter
        .FormatString = "###,##"
        .Xposition = 100
      '  .Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    'invoice# field
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "invoice"
        .Xposition = 700
        .Font.Name = "Arial"
       ' .PostAdvanceLine = 1
    End With
    oGroup.AddPrintItem oRptItem
    'invoice date field
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "inv_date"
        .Xposition = 2000
        .FormatString = "dd/mm/yyyy"
        .PostAdvanceLine = 1
       ' .Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    oGroup.PrintGroupBreak = True

    'group break items for invoice group
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemFormula
        .Font.Italic = True
        '.HasFormula = True
         .PrintAllign = JxDBAllignRight
        .FormulaType = JxDBRptFormulaSum
        .Font.Italic = True
        .FieldName = "amount"
        .Xposition = 4000
        .FormatString = "###,##0.00"
        '.Font.Name = "Arial"

    End With
    oGroup.AddBreakItem oRptItem
    oGroup.PrintGroupBreak = True

    'total field label
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Total Invoice"
        .Xposition = 1000
        .PrintAllign = JxDBAllignLeft
        '.PostAdvanceLine = 1
        '.Font.Name = "Arial"
    End With
    oGroup.AddBreakItem oRptItem
    'setup line
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "___________________________________________________________________________________________"
        .Xposition = 0
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
        .FontSize = "9"
        .FontBold = True
        .PostAdvanceLine = 1
        .Height = 300
    End With
    oGroup.AddBreakItem oRptItem
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Highest Invoice Item"
        .Xposition = 1000
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
    End With
    oGroup.AddBreakItem oRptItem
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemFormula
        .FormulaType = JxDBRptFormulaHighest
        .FieldName = "amount"
        
        .FormatString = "###,##0.00"
        .Xposition = 4000
        .PrintAllign = JxDBAllignRight
        .Font.Name = "Arial"
    End With
    oGroup.AddBreakItem oRptItem
    'smallest invoice item
     Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Smallest Invoice Item"
        .Xposition = 4200
        .PrintAllign = JxDBAllignLeft
        .Font.Name = "Arial"
    End With
    oGroup.AddBreakItem oRptItem
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemFormula
        .FormulaType = JxDBRptFormulaLowest
        .FieldName = "amount"
        .FormatString = "###,##0.00"
        .Xposition = 7000
        .FontColor = vbRed
        .PrintAllign = JxDBAllignRight
        .Font.Name = "Arial"
        .PostAdvanceLine = 1
    End With
    oGroup.AddBreakItem oRptItem

    oRpt.ReportGroups.Add oGroup

    'GROUP Detail
    Set oGroup = New JxDBRptGroup
    oGroup.AddGroupFields "item"
    oGroup.CheckForBreak = False
    oGroup.GroupName = "detail"
    'oGroup.PrintGroup = False 'flag for group printing
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "item"
        .Xposition = 2000
       ' .Font.Name = "Arial"
    End With
    oGroup.AddPrintItem oRptItem
    'amount field
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemDataField
        .FieldName = "amount"
        .Font.Name = "Arial"
        .Xposition = 4000
        .FormatString = "###,##0.00"
        .PrintAllign = JxDBAllignRight

    End With
    oGroup.AddPrintItem oRptItem
    
    'running total on amount field
    Set oRptItem = New JxDBRptItem
    With oRptItem

        .ItemType = JxDBItemRunningTotal
        .FieldName = "amount"
        .Xposition = 6000
        .FormatString = "###,##0.00"
        .PrintAllign = JxDBAllignRight
        .Font.Name = "Arial"

        '.PostAdvanceLine = 1
    End With
    oGroup.AddPrintItem oRptItem
    Set oRptItem = New JxDBRptItem
    With oRptItem

        .ItemType = JxDBItemCustom
        .PrintAllign = JxDBAllignCenter

        .Xposition = 7000
        .PostAdvanceLine = 1
    End With
    oGroup.AddPrintItem oRptItem

    oGroup.PrintGroupBreak = True

    oRpt.ReportGroups.Add oGroup, "detail"

    ''setup report footer
    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemPageXofY
        .Font.Name = "Arial"
    End With
    oRpt.PageFooter.AddItem oRptItem, "itmPageXofY"

    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemLabel
        .Value = "Created by Joyprakash Saikia (footer) "
        .PrintAllign = JxDBAllignCenter
        .Font.Name = "Arial"
        .FontBold = True
    End With
    oRpt.PageFooter.AddItem oRptItem

    Set oRptItem = New JxDBRptItem
    With oRptItem
        .ItemType = JxDBItemTotalPage
        .PrintAllign = JxDBAllignRight
        .Font.Name = "Arial"
    End With
     oRpt.PageFooter.AddItem oRptItem
  oRpt.PageFooter.RepeatHeader = True

  Dim DB As Connection
  Set DB = New Connection
  Dim rst As Recordset
  Dim sDbPath As String

  sDbPath = App.Path & "\test.mdb"
  DB.CursorLocation = adUseClient
  DB.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDbPath

  Set rst = New Recordset

  
  'to get the group break correctly, the ORDER clause must satisfy the report grouping
  SSQL = "SELECT [invoice].[cocode], [invoice].[vendor], [invoice].[invoice], [invoice].[item], [invoice].[amount], [invoice_sum].[inv_date], [vendor].[vendor_name] " _
        & " FROM ((invoice INNER JOIN invoice_sum ON ([invoice].[invoice]=[invoice_sum].[invoice]) AND ([invoice].[vendor]=[invoice_sum].[vendor]) AND ([invoice].[cocode]=[invoice_sum].[cocode])) INNER JOIN vendor ON [invoice_sum].[vendor]=[vendor].[id] " _
        & " INNER JOIN    vendor ON (`invoice_sum`.`vendor` = `vendor`.`id`))" _
        & " INNER JOIN    Co_master ON " _
        & " [invoice].[cocode] = [co_master].[co_code]" _
        & "ORDER BY invoice.cocode,invoice.vendor, invoice.invoice, invoice.item"
  SSQL = "SELECT `invoice`.`cocode`, `co_master`.`Co_name`, " _
       & " `invoice`.`vendor`, `invoice`.`invoice`, `invoice`.`item`, " _
       & " `invoice`.`amount`, `invoice_sum`.`inv_date`, `vendor`.`vendor_name` " _
       & " FROM ((invoice INNER JOIN " _
       & " invoice_sum ON " _
       & " (`invoice`.`invoice` = `invoice_sum`.`invoice`) AND " _
       & " (`invoice`.`vendor` = `invoice_sum`.`vendor`) AND " _
       & " (`invoice`.`cocode` = `invoice_sum`.`cocode`)) INNER JOIN " _
       & " vendor ON (`invoice_sum`.`vendor` = `vendor`.`id`)) " _
       & " INNER JOIN  Co_master ON " _
       & " `invoice`.`cocode` = `co_master`.`co_code` " _
       & " ORDER BY invoice.cocode, invoice.vendor, invoice.invoice,    invoice.Item"
     oRpt.DataSource.Open SSQL, DB, adOpenStatic, adLockReadOnly
    
    'set report margin

    oRpt.SetLeftMargin pagemargin(0).Text, JxDBRptScaleInch
    oRpt.SetRightMargin pagemargin(1).Text, JxDBRptScaleInch
    oRpt.SetTopMargin pagemargin(2).Text, JxDBRptScaleInch
    oRpt.SetBottomMargin pagemargin(3).Text, JxDBRptScaleInch
    
    oRpt.SetPageFooterMargin 0.35, JxDBRptScaleInch 'you can even put margin for Footer
    Set oRptItem = Nothing
End Sub



Private Sub cmdhelp_Click()
    ShellExecute Me.hwnd, vbNullString, App.Path & "\chm\jxdbreport.chm", vbNullString, App.Path & "\chm", 1&
End Sub

Private Sub Command1_Click()
    'To display Report with All Gorups
    Dim oGroup As JxDBRptGroup
    Dim itms As JxDBRptItems
    Command2_Click
    oRpt.PaperSize = Combo1.ListIndex
    SetupReport
    oRpt.SetBottomMargin 0.25, JxDBRptScaleInch
    oRpt.ReportType = JxDBReportPrePrinted
   
    oRpt.Preview
   '
  For Each oGroup In oRpt.ReportGroups
    Set itms = oGroup.PrintItems
    Set itms = Nothing
 Next
  '  oRpt.destroyall
    Set oRpt = Nothing
End Sub

Private Sub Command2_Click()
    If oRpt Is Nothing Then
        Set oRpt = New JxDBReport
    Else
     Set oRpt = Nothing 'Destroy The Previous Object
     Set oRpt = New JxDBReport
    End If
    
End Sub

Private Sub Command3_Click()
Command2_Click
     SetupReport
    oRpt.PrintReport True, 2, 2, Me.hwnd
End Sub

Private Sub Command4_Click()
   
     Dim oGroup As JxDBRptGroup
     Command2_Click
         oRpt.PaperSize = Combo1.ListIndex
        
    SetupReport
   
   oRpt.SetBottomMargin 3.45, JxDBRptScaleInch
   'Disaplya Only the Detail and Suppress the Remain groups
    For Each oGroup In oRpt.ReportGroups
              'Debug.Print oGroup.GroupName; " with level="; oGroup.GroupLevel
      
        If oGroup.GroupName = "invoice" Then
            oGroup.PrintGroup = True
        Else
            oGroup.PrintGroup = False
            
            
        End If
        
        
    Next
    oRpt.Preview
     ' oRpt.ReportWidth = 0
End Sub
Private Sub Form_Load()
Dim tp As Integer
    App.HelpFile = App.Path & "\help\JxDBReport.chm::/popup.txt"

    Combo1.Clear
   Combo1.AddItem "Letter"
   Combo1.ItemData(0) = 1
   Combo1.AddItem "LetterSmall"
   Combo1.ItemData(1) = 2
   Combo1.AddItem "Tabloid"
   Combo1.ItemData(2) = 3
   Combo1.AddItem "Ledger"
   Combo1.ItemData(3) = 4
   Combo1.AddItem "PaperA3"
   Combo1.ItemData(4) = 8
   Combo1.AddItem "PaperA4"
   Combo1.ItemData(5) = 9
   Combo1.AddItem "PaperA5"
   Combo1.ItemData(6) = 11
   Combo1.AddItem "PaperB4"
   Combo1.ItemData(7) = 12
   Combo1.AddItem "Custom"
   Combo1.ItemData(8) = 256
   Combo1.ListIndex = 1
    pagemargin(0).Text = "0.25"
    pagemargin(1).Text = "0.25"
    pagemargin(2).Text = "0.25"
    pagemargin(3).Text = "0.25"
    porientation_Click 0

   If oRpt Is Nothing Then
        Set oRpt = New JxDBReport
    End If
    'oRpt.ShowAbout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRpt = Nothing
End Sub
Private Sub lblVote_Click()
  ShellExecute Me.hwnd, vbNullString, "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=35607&optCodeRatingValue=5&intUserRatingTotal=0&intNumOfUserRatings=0", vbNullString, App.Path & "\chm", 1&
End Sub
Private Sub lblVote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(ByVal 0&, 32649&)
End Sub
Private Sub oRpt_PrintCustomItem(ReportItem As JxDBRpt.JxDBRptItem)
  ' Debug.Print ReportItem.Value
    If oRpt.DataSource.Fields("amount").Value > 10000 Then
        ReportItem.FontUnderline = True
        ReportItem.Value = "Marginal Amount" 'oRpt.DataSource.Fields("amount").Value '"Big amount"
        ReportItem.PrintAllign = JxDBAllignRight
        ReportItem.Xposition = 9500
        ReportItem.FontColor = vbBlue
        ReportItem.FONTSTRIKETHROUGH = False
    Else
        ReportItem.Value = "Small amount"
        ReportItem.FontColor = vbRed
        'ReportItem.FONTSTRIKETHROUGH = True
    End If


End Sub

Private Sub oRpt_Printing(ByVal CurrentPage As Long, Cancel As Boolean)
    '
    If CurrentPage > 1000 Then
        Cancel = True
    End If
End Sub

Private Sub porientation_Click(Index As Integer)
    Image1(Index).Visible = True
    If Index = 0 Then
        Image1(Index + 1).Visible = False
    Else
        Image1(Index - 1).Visible = False
    End If
End Sub
