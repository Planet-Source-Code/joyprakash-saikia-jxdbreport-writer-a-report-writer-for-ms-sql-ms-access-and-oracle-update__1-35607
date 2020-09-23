VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JxDBReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781AC0122"
'**********************************************************************
'
'           Module Name: JxDBReport.cls
'           Purpose    : Main class for JxDB Report application
'           Author     : Joyprakash Saikia
'
'**********************************************************************
Option Explicit
Option Base 0

'This Constant is Used for Debugging Purpose
'This is not used on this submission

Private Const MOD_NAME = "JxDBReport"
'for page report info
Private Type JxDBRptPageInfo
  Width As Single ' Report Width
  Height As Single 'Report Height
  AvailWidth As Single 'Total printable width
  AvailHeight As Single 'Total printable width (not including footer margin)
  LeftM As Single 'Left Margin
  RightM As Single 'Right Margin
  TopM As Single 'Top Margin
  BottomM As Single 'Bottom Margin
  RHeaderM As Single 'Report Header Margin(used once only)
  HeaderM As Single 'page header margin
  FooterM As Single 'footer margin
  RFooterM As Single 'Report footer Margin
  CurrentPage As Integer 'Current page number
  TotalPage As Integer 'Total report page(s)
  CurrentY As Single ' Hold current height of the report
  NewPage As Boolean 'NewPage Flag
  hFile As Long 'handle to report data file
  hIndexFile As Long 'handle to report index file
End Type
'pre defined paper size enum
Public Enum JxDBPaperSize
  JxDBRptPaperLetter = 1
  JxDBRptPaperLetterSmall = 2
  JxDBRptPaperTabloid = 3
  JxDBrptPaperLedger = 4
  JxDBRptPaperA3 = 8
  JxDBRptPaperA4 = 9
  JxDBRptPaperA5 = 11
  JxDBRptPaperB4 = 12
  JxDBRptPaperCustom = 256
End Enum
'pre defined Scale enum
Public Enum JxDBRptScale
  JxDBRptScaleDefault = 0 'default scale
  JxDBRptScaleTwip = 1
  JxDBRptScaleChar = 4
  JxDBRptScaleInch = 5
  JxDBRptScaleMm = 6
  JxDBRptScaleCm = 7
End Enum
Public Enum JxDBRptOrientation
    JxDBRptPotrait = 1
    JxDBRptLandscape = 2
End Enum

'report type enum
'std report will have header & footer
Public Enum JxDBReportType
    JxDBReportStandard = 1
    JxDBReportPrePrinted = 2
End Enum
'report datasource
Dim rst As Recordset
'Report items & collections
' The Report Header is not used on this Submission
Public ReportHeader As JxDBRptHeaderFooter ' Report Header
'Page Header Object
Public PageHeader As JxDBRptHeaderFooter 'page Header
'Page Footer Object
Public PageFooter As JxDBRptHeaderFooter 'page footer
' The Report Footer is not used on this Submission
Public ReportFooter As JxDBRptHeaderFooter 'Report Footer
' The Groups Collection
Public ReportGroups As JxDBRptGroups

'Default values

Private Const m_defReportType = 1
Private Const m_PaperSize = JxDBPaperSize.JxDBRptPaperA4
Private Const m_defReportScale = JxDBRptScale.JxDBRptScaleTwip
' Followings are the Private Variable for Properties
Private mvarReportType As JxDBReportType
Private mvarPaperSize As JxDBPaperSize
Private mvarReportHeight As Single
Private mvarReportWidth As Single
Private mvarReportScale As JxDBRptScale
Private mvarBottomMargin As Single
Private mvarTopMargin As Single
Private mvarLeftMargin As Single
Private mvarRightMargin As Single
Private mvarMaxPage As Integer 'Hold Max Pages
Private mvarTotalPages As Long
Private mvarReportTitle As String
Private mvarPreviewRatio As Single
Private mvarRecordProcessed As Long 'no of record processed
Private bPrinter As Boolean
Private mvarPageOrientation As JxDBRptOrientation

 'flag for newpage while printing/previewing
Private NewPage As Boolean
'To Store Target obj for printing whether picturebox or printer
Private PrintTarget As Object
Private bPrint As Boolean 'flag to indicate whether to print or not useful for generating/printing specific page(s)
Private bInitial As Boolean 'flag to indicate initial printing
'Flag for Special Consideration while Target is Printer
Private bmPrinterFlag As Boolean

Private bmChanged As Boolean
' File is Opened or Created
Private bmFileOpened As Boolean
Private bmFileCreated As Boolean
'flag to Display report without recreating Report file
Private ShowwithPrevData As Boolean
'report page size
Dim ReportInfo As JxDBRptPageInfo

'public events
Public Event Printing(ByVal CurrentPage As Long, Cancel As Boolean)
Public Event PrintGroupItem(ByVal oGroup As JxDBRptGroup) ' to be removed on submission
Public Event PrintCustomItem(ReportItem As JxDBRptItem)

'UDTs for report data and index file
' To store in the File

Private tmItemInfo As UDTRptItem
Private tmPageHeader As UDTRptPageHeader
Private tmPageIndex As UDTRptPageIndex

Dim iCount As Integer


Private Property Let RecordProcessed(ByVal vData As Long)
        '************************************************************
        'Description:
        ' To Display On Preview at the Bottom
        '************************************************************
        mvarRecordProcessed = vData
End Property

Public Property Let RefreshReport(Data As Boolean)
        '************************************************************
        'Description:
        ' True - If File is to be created
        ' False- Donot Create the File
        '************************************************************
        bmFileCreated = Not Data
End Property
Friend Property Get RecordProcessed() As Long
    RecordProcessed = mvarRecordProcessed
End Property

Public Property Let PageOrientation(ByVal vData As JxDBRptOrientation)
        '************************************************************
        'Description:
        ' Portrait or Landscape
        '************************************************************
        mvarPageOrientation = vData
End Property
Public Property Get PageOrientation() As JxDBRptOrientation
    PageOrientation = mvarPageOrientation
End Property

Friend Property Let PreviewRatio(ByVal vData As Single)
    mvarPreviewRatio = vData
End Property

Friend Property Get PreviewRatio() As Single
    PreviewRatio = mvarPreviewRatio
End Property
Public Property Let ReportTitle(ByVal vData As String)
    mvarReportTitle = vData
End Property
Public Property Get ReportTitle() As String
    ReportTitle = "JxDBReport - " & mvarReportTitle
End Property

Public Property Set DataSource(ByVal vData As Recordset)
Attribute DataSource.VB_Description = "' Description:  Data Source is an ADO recordSet Object \r\n             Responsible for providing the Report with Data\r\n\r\n'\r\n\r\n"
    Set rst = vData
End Property

Public Property Get DataSource() As Recordset

    If rst Is Nothing Then Set rst = New ADODB.Recordset: rst.CursorLocation = adUseClient
    Set DataSource = rst
End Property
Public Property Get TotalPages() As Long
    TotalPages = mvarTotalPages
End Property

Public Property Let ReportScale(ByVal vData As JxDBRptScale)
     mvarReportScale = vData
End Property
Public Property Get ReportScale() As JxDBRptScale
    ReportScale = mvarReportScale
End Property

Public Sub SetBottomMargin(ByVal NewMargin As Single, Optional ByVal MarginScale As JxDBRptScale = JxDBRptScaleDefault)
        '************************************************************
        'Description:
        ' The Input Parameter is Converted into appropriate Scale and
        ' Make it as new Bottom Margin
        '************************************************************
    If MarginScale = JxDBRptScaleDefault Then
        MarginScale = Me.ReportScale
    End If
    
    ReportInfo.BottomM = ConvertToTwip(MarginScale, NewMargin)
    
    'check if Margin is Greater than the Report Height
    
    If ReportInfo.BottomM > ReportInfo.Height Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        ReportInfo.BottomM = 0
        Exit Sub
    End If
    ' ReCalculate the Size for the Data
    GetPaperSize
End Sub

Public Sub SetPageFooterMargin(ByVal NewMargin As Single, Optional ByVal MarginScale As JxDBRptScale = JxDBRptScaleDefault)
        '************************************************************
        'Description:
        ' The Input Parameter is Converted into appropriate Scale and
        ' Make it as new Page Footer Margin
        '************************************************************

    If MarginScale = JxDBRptScaleDefault Then
        MarginScale = Me.ReportScale
    End If
    ReportInfo.FooterM = ConvertToTwip(MarginScale, NewMargin)
    'check if Margin is Greater than the Report Height
    If ReportInfo.FooterM > ReportInfo.Height Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        ReportInfo.FooterM = 0
        Exit Sub
    End If
    ' ReCalculate the Size for the Data
    GetPaperSize
End Sub

Public Sub SetPageHeaderMargin(ByVal NewMargin As Single, Optional ByVal MarginScale As JxDBRptScale = JxDBRptScaleDefault)
        '************************************************************
        'Description:
        ' The Input Parameter is Converted into appropriate Scale and
        ' Make it as new Page Footer Margin
        '************************************************************
    If MarginScale = JxDBRptScaleDefault Then
        MarginScale = Me.ReportScale
    End If
    ReportInfo.HeaderM = ConvertToTwip(MarginScale, NewMargin)
        'check if Margin is Greater than the Report Height
    If ReportInfo.HeaderM > ReportInfo.Height Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        ReportInfo.HeaderM = 0
        Exit Sub
    End If
    ' ReCalculate the Size for the Data
    GetPaperSize
End Sub

Public Sub SetTopMargin(ByVal NewMargin As Single, Optional ByVal MarginScale As JxDBRptScale = JxDBRptScaleDefault)
        '************************************************************
        'Description:
        ' The Input Parameter is Converted into appropriate Scale and
        ' Make it as new Page Top Margin
        '************************************************************
    
    If MarginScale = JxDBRptScaleDefault Then
        MarginScale = Me.ReportScale
    End If
    ReportInfo.TopM = ConvertToTwip(MarginScale, NewMargin)
    'check for validity
    If ReportInfo.TopM > ReportInfo.Height Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        ReportInfo.TopM = 0
        Exit Sub
    End If
        ' ReCalculate the Size for the Data
    GetPaperSize
End Sub


Public Sub SetRightMargin(ByVal NewMargin As Single, Optional ByVal MarginScale As JxDBRptScale = JxDBRptScaleDefault)
        '************************************************************
        'Description:
        ' The Input Parameter is Converted into appropriate Scale and
        ' Make it as new Page Right Margin
        '************************************************************

    If MarginScale = JxDBRptScaleDefault Then
        MarginScale = Me.ReportScale
    End If
    ReportInfo.RightM = ConvertToTwip(MarginScale, NewMargin)
    'check for validity
    If ReportInfo.RightM > ReportInfo.Width Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        ReportInfo.LeftM = 0
        Exit Sub
    End If
End Sub

Public Sub SetLeftMargin(ByVal NewMargin As Single, Optional ByVal MarginScale As JxDBRptScale = JxDBRptScaleDefault)
        '************************************************************
        'Description:
        ' The Input Parameter is Converted into appropriate Scale and
        ' Make it as new Page Footer Margin
        '************************************************************

    If MarginScale = JxDBRptScaleDefault Then
        MarginScale = Me.ReportScale
    End If
    ReportInfo.LeftM = ConvertToTwip(MarginScale, NewMargin)
    'check if Margin is Greater than the Report Width
    If ReportInfo.LeftM > ReportInfo.Width Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        ReportInfo.LeftM = 0
        Exit Sub
    End If
End Sub

Public Sub SetCustomPaper(ByVal PaperScale As JxDBRptScale, ByVal NewWidth As Single, ByVal NewHeight As Single)
        '************************************************************
        'Description:
        ' Called by PaperSize Property
        ' Also used when ReportSize is Custom
        '************************************************************
    
    ReportInfo.Width = ConvertToTwip(PaperScale, NewWidth)
    ReportInfo.Height = ConvertToTwip(PaperScale, NewHeight)
    GetPaperSize
End Sub

Public Property Let ReportWidth(ByVal vData As Single)
        '************************************************************
        'Description:
        ' if you Donot provide the Width , How can I calculate
        ' an arbitary One
        '************************************************************
    If vData = 0 Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        Exit Property
    End If
    mvarReportWidth = vData
    
End Property
Public Property Get ReportWidth() As Single
    ReportWidth = ReportInfo.Width
End Property

Public Property Let ReportHeight(ByVal vData As Single)
        '************************************************************
        'Description:
        ' if you Donot provide the Height of a Report , How can I calculate
        ' an arbitary One
        '************************************************************
    If vData <= 0 Then
        Err.Raise vbObjectError + ERR_INVALID_PROP_VALUE, MOD_NAME, LoadResString(ERR_INVALID_PROP_VALUE)
        Exit Property
    End If
    ReportInfo.Height = ConvertToTwip(ReportScale, vData)
End Property
Public Property Get ReportHeight() As Single
    ReportHeight = ConvertFromTwip(ReportScale, ReportInfo.Height)
End Property

Public Property Let PaperSize(ByVal vData As JxDBPaperSize)
        '************************************************************
        'Description:
        ' Property to assign the Paper Size
        '************************************************************

    mvarPaperSize = vData
    If vData = JxDBRptPaperCustom Then Exit Property
    Select Case vData
        Case JxDBPaperSize.JxDBRptPaperA3
            SetCustomPaper JxDBRptScaleMm, 297, 420
        Case JxDBPaperSize.JxDBRptPaperA4
            SetCustomPaper JxDBRptScaleMm, 210, 297
        Case JxDBPaperSize.JxDBRptPaperA5
            SetCustomPaper JxDBRptScaleMm, 148, 210
        Case JxDBPaperSize.JxDBRptPaperB4
            SetCustomPaper JxDBRptScaleMm, 250, 354
        Case JxDBPaperSize.JxDBrptPaperLedger
            SetCustomPaper JxDBRptScaleInch, 8.5, 14
        Case JxDBPaperSize.JxDBRptPaperLetter
            SetCustomPaper JxDBRptScaleInch, 8.5, 11
        Case JxDBPaperSize.JxDBRptPaperLetterSmall
            SetCustomPaper JxDBRptScaleInch, 8.5, 11
        Case JxDBPaperSize.JxDBRptPaperTabloid
            SetCustomPaper JxDBRptScaleInch, 11, 17
    End Select
End Property
Public Property Get PaperSize() As JxDBPaperSize
    PaperSize = mvarPaperSize
End Property

Public Property Let ReportType(ByVal vData As JxDBReportType)
    mvarReportType = vData
End Property
Public Property Get ReportType() As JxDBReportType
    ReportType = mvarReportType
End Property

Private Sub GetPaperSize()

    With ReportInfo
        .AvailHeight = .Height - .TopM - .BottomM - .FooterM - .HeaderM
        .AvailWidth = .Width - .LeftM - .RightM
    End With
    
End Sub
'##ModelId=3B3781B902EE
Public Sub Preview()
Attribute Preview.VB_Description = "This Method is called by the Client application for printing application to the Screen."
'************************************************************
'      Created by: Joyprakash saikia     Date: 23/06/2002
'
'          Description:
'               The Entry point for building A report
'                 This Routine is called by the client to preview the report.
'                 It will then call various sub to perform report operations
'                 and finally display the report
'************************************************************
On Error GoTo PreviewErr

    ValidateReport
    
    
    ReportInfo.CurrentPage = 0
    'check whether report file has been created or not
    If Not bmFileCreated Then
        If Not rst Is Nothing Then
          If rst.AbsolutePosition <> adPosBOF Or rst.AbsolutePosition <> adPosEOF Or rst.AbsolutePosition <> adPosUnknown Then rst.MoveFirst
        End If
        ' Open the File For Write Mode
        If Not OpenRptTempFile("w") Then
            Err.Raise vbObjectError + ERR_CREATING_RPT_FILE, MOD_NAME, LoadResString(ERR_CREATING_RPT_FILE)
            Exit Sub
        End If
        'Create the Report  File
        WriteReportToFile
     
    End If
    
    Set frmPreview.Document = Me
    Load frmPreview
    frmPreview.Visible = False 'Donot try to Print Immidiately
    
    frmPreview.PreviewPage 1, 1
    
    frmPreview.Show
    
    
PreviewEXIT:

Exit Sub
PreviewErr:
          Err.Raise Err.Number, SetErrSource(MOD_NAME, "Preview") & Err.Source, Err.Description
          
End Sub


Private Sub WriteReportHeader()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine responsible for writing each ReportHEader item
'                 It is called only once during report generation
'          NOTES:
'          This function is not currently used
'************************************************************
On Error GoTo WriteReportHeaderErr

    Dim i As Integer
    Dim oLabel As JxDBRptItem
    For i = 1 To ReportHeader.Count
        Set oLabel = ReportHeader.Item(i)
        If oLabel.PreAdvanceLine Then
            PrintTarget.Print
            ReportInfo.Height = PrintTarget.CurrentY 'ReportInfo.CurrentHeight + PrintTarget.TextHeight(oLabel.Caption)
        End If
'        With PrintTarget
'            .FontName = oLabel.FontName
'            .FontBold = oLabel.FontBold
'            .FontSize = oLabel.FontSize
'            .CurrentX = ReportInfo.LeftM + oLabel.XPosition
'        End With
        
        If oLabel.PrintAllign = JxDBAllignCenter Then
            'PrintTarget.CurrentX = CenterOnPage(oLabel.Value)
          '  ReportInfo
        End If
        PrintTarget.Print oLabel.Value
        ReportInfo.CurrentY = ReportInfo.CurrentY + PrintTarget.TextHeight(oLabel.Value)
    Next i
    
    PrintTarget.Print 'advance to next line
    'PrintTarget.CurrentX = CenterOnPage(Me.ReportTitle)
    'PrintTarget.Print Me.ReportTitle
    ReportInfo.CurrentY = ReportInfo.CurrentY + PrintTarget.TextHeight(Me.ReportTitle)
    'report.CurrentY = PrintTarget.CurrentY
WriteReportHeaderEXIT:

Exit Sub
WriteReportHeaderErr:
          Err.Raise Err.Number, Err.Source, Err.Description
          GoTo WriteReportHeaderEXIT
End Sub

Public Property Set Target(obj As Object)
    Set PrintTarget = obj
End Property
'Public Property Get PageCount() As Integer
'    PageCount = mvarMaxPage
'End Property
'##ModelId=3B3781BA008C
Private Function CheckGroupBreak(ByVal TotalGroup As Integer) As Integer
'************************************************************
'      Created by: joy     Date: 06/06/2002
'
'
'          Description:
'                 This Rutine compares each of the groups ID with the current record
'                 If the ID and current record are not the same, then break occurs
'                 For the first group that break occurs, tha group index will be returned
'                 and this is used for group printing loop
'
'
'          Input:
'
'              TotalGroup - Number of total groups in the report
'
'          OutPut: First group break index (or zero if no break occurs)
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo CheckGroupBreakErr

'    If bDEBUG Then
'        oTrace.LogProc "CheckGroupBreak", MOD_NAME
'    End If
    Dim i As Integer, j As Integer, k As Integer, numOfBreak As Integer
    Dim GroupBreak() As Integer
    Dim oItem As JxDBRptItem
    Dim ssPlitFld() As String
    Dim SFld As String, sGroupID As String
    Dim oGroup As JxDBRptGroup
    Dim bFirstBreak As Boolean 'first group that break occurs
    
    bFirstBreak = True
    CheckGroupBreak = 0

    
    For k = 1 To TotalGroup
        sGroupID = "" 'reset
        Set oGroup = Me.ReportGroups.Item(k)
        oGroup.GroupBreak = False
        ssPlitFld = Split(oGroup.GroupFields, ",")
        'build the group id from current record
        For j = 0 To UBound(ssPlitFld)
            SFld = ssPlitFld(j) 'get the field name
            sGroupID = sGroupID & "" & rst.Fields(Trim$(SFld)).Value
        Next j
        If IsEmpty(oGroup.GroupID) Then
            oGroup.GroupID = sGroupID
        Else '
            'Compare last groupID with current record, if different
            'then break occurs
            If oGroup.CheckForBreak Then
                If StrComp(oGroup.GroupID, sGroupID, vbTextCompare) Then
                    oGroup.GroupBreak = True
                    oGroup.GroupID = sGroupID
                    'if not initial and first group that break occurs
                    'we need to return that group for group looping printing
                    If Not bInitial Then
                        If bFirstBreak Then
                            CheckGroupBreak = k
                            bFirstBreak = False
                        End If
                    End If 'bInitial
                End If 'StrComp
            End If 'checkforbreak
        End If 'IsEmpty
        
    Next k
    Set oGroup = Nothing
CheckGroupBreakEXIT:

Exit Function
CheckGroupBreakErr:
          GoTo CheckGroupBreakEXIT
End Function
Private Sub AccumulateTotal()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine will accumulate/calculate for each FormulaType item
'                 in the report
'
'          Input:
'
'
'          OutPut:
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo AccumulateTotalErr

'    If bDEBUG Then
'        oTrace.LogProc "AccumulateTotal", MOD_NAME
'    End If
    Dim i As Integer, j As Integer
    Dim iTotalGroup As Integer
    Dim oGroup As JxDBRptGroup
    Dim oItem As JxDBRptItem
    Dim vntRstValue As Variant
    iTotalGroup = Me.ReportGroups.Count
    
    For i = 1 To iTotalGroup
        Set oGroup = Me.ReportGroups.Item(i)
        For j = 1 To oGroup.BreakItemCount
            Set oItem = oGroup.BreakItems(j)
            If oItem.ItemType = JxDBItemFormula Then
                vntRstValue = rst.Fields(oItem.FieldName).Value
                Select Case oItem.FormulaType
                    
                    Case JxDBRptFormulaType.JxDBRptFormulaSum
                        oItem.Value = oItem.Value + vntRstValue
                    Case JxDBRptFormulaType.JxDBRptFormulaAverage
                        oItem.Value = oItem.Value + vntRstValue
                        oItem.ValueExtra = oItem.ValueExtra + 1 'count the total item
                    Case JxDBRptFormulaType.JxDBRptFormulaHighest
                        'if for details groups, just compare the value
                        If i = iTotalGroup - 1 Then
                            If IsEmpty(oItem.ValueExtra) Then 'nothing yet, just assigned the value
                                'oItem.Value = Empty
                                oItem.ValueExtra = vntRstValue
                            Else
                                oItem.ValueExtra = IIf(oItem.ValueExtra < vntRstValue, vntRstValue, oItem.ValueExtra)
                            End If
                        Else
                            oItem.ValueExtra = oItem.ValueExtra + vntRstValue
                        End If
                    Case JxDBRptFormulaType.JxDBRptFormulaLowest
                        'if detail groups, just compare the value
                        If i = iTotalGroup - 1 Then
                            If IsEmpty(oItem.ValueExtra) Then 'nothing yet, just assigned the value
                                oItem.ValueExtra = vntRstValue
                            Else
                                oItem.ValueExtra = IIf(oItem.ValueExtra > vntRstValue, vntRstValue, oItem.ValueExtra)
                            End If
                        Else
                            'we accumulate this and compare during the group break
                            oItem.ValueExtra = oItem.ValueExtra + vntRstValue
                        End If
                    Case JxDBRptFormulaType.JxDBRptFormulaRunningTotal
                        oItem.Value = IIf(IsEmpty(oItem.Value), vntRstValue, oItem.Value + vntRstValue)

                    Case JxDBRptFormulaType.JxDBRptFormulaMedian
                    'TODO
                End Select
            End If
        Next j
    Next i
    Set oItem = Nothing
    Set oGroup = Nothing
AccumulateTotalEXIT:

Exit Sub
AccumulateTotalErr:
          Err.Raise Err.Number, Err.Source & SetErrSource(MOD_NAME, "AccumulateTotal"), Err.Description
          'GoTo AccumulateTotalEXIT
End Sub
'##ModelId=3B3781BA0320
Private Sub ResetCounterField(ByVal StartGroup As Integer, ByVal TotalGroup As Integer)
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine reset any counter field type in the StartGroup. Called
'                 when any group break occurs
'
'          Input:
'
'              TotalGroup -  Number of total groups in report
'              StartGroup - Starting index for group break
'
'          OutPut:
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo ResetCounterFieldErr
'
'    If bDEBUG Then
'        oTrace.LogProc "ResetCounterField", MOD_NAME
'    End If
    Dim i As Integer, j As Integer
    Dim oGroup As JxDBRptGroup, oGroupRst As JxDBRptGroup
    Dim oItem As JxDBRptItem
    
    For i = StartGroup To TotalGroup
        Set oGroup = Me.ReportGroups(i)
        If oGroup.GroupBreak Then
            Set oGroupRst = Me.ReportGroups(i + 1) 'get lower level group
            If oGroupRst Is Nothing Then Exit Sub ''no more lower level group
            For j = 1 To oGroupRst.PrintItemsCount
                Set oItem = oGroupRst.PrintItems(j)
                If oItem.ItemType = JxDBItemCounter Then
                    oItem.Value = 0
                End If
            
            Next j
        End If
    Next i
    
  
ResetCounterFieldEXIT:
'Release All Local Objects
        Set oGroup = Nothing: Set oGroupRst = Nothing: Set oItem = Nothing
Exit Sub
ResetCounterFieldErr:
          GoTo ResetCounterFieldEXIT
End Sub
Private Sub Class_Initialize()
    'set defaults
    bDEBUG = False
    ReportType = m_defReportType
    mvarReportScale = m_defReportScale
    Set ReportHeader = New JxDBRptHeaderFooter
    Set PageHeader = New JxDBRptHeaderFooter
    Set ReportGroups = New JxDBRptGroups
    Set PageFooter = New JxDBRptHeaderFooter
    PreviewRatio = 1
    Me.PageOrientation = JxDBRptPotrait
    Me.PaperSize = m_PaperSize
'    If bDEBUG Then
'        Set oTrace = New cTracer
'        oTrace.FileName = App.Path & "\trace.log"
'    End If
End Sub
'##ModelId=3B3781BC0118
Private Sub Class_Terminate()
    Set ReportHeader = Nothing
    Set PageHeader = Nothing
    Set PageFooter = Nothing
    Set ReportGroups = Nothing
End Sub
'##ModelId=3B3781BC01B8
Friend Sub PrintIt(Optional ByVal StartPage As Integer = 0, Optional ByVal EndPage As Integer = 0)
'************************************************************
'      Created by: joy     Date: 23/06/2002
'
'
'          Description:
'                 This Rutine is called by the preview window
'                 for printing the report
'
'          Input:
'
'              EndPage - Last page number to be printed
'              StartPage - Starting page number to be printed
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo PrintItErr

    
    DisplayReport StartPage, EndPage
        
PrintItEXIT:

Exit Sub
PrintItErr:
          Err.Raise Err.Number, Err.Source, Err.Description
          GoTo PrintItEXIT
End Sub
'##ModelId=3B3781BD0064
Friend Sub PreviewIt(Optional ByVal StartPage As Integer = 0, Optional ByVal EndPage As Integer = 0)
'************************************************************
'      Created by: joy     Date: 19/04/
'
'
'          Description:
'                 This Rutine is called by the preview windows
'                 for previewing
'
'          Input:
'
'              EndPage
'              StartPage
'
'          OutPut:
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo PreviewItErr

    
    DisplayReport StartPage, EndPage
PreviewItEXIT:

Exit Sub
PreviewItErr:
          GoTo PreviewItEXIT
End Sub

'##ModelId=3B3781BD0334
Private Sub DisplayReport(Optional ByVal StartPage As Integer = 0, Optional ByVal EndPage As Integer = 0)
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is called internally by Preview/PrintReport method
'                 It reads report data and index file created for requested page
'                 and send the item for printing/preview
'
'          Input:
'
'              EndPage - Starting Page No to be previewed
'              StartPage - Last page number to be previewed.
'
'          OutPut:
'
'
'          Globals Modfied
'
'          NOTES
'           Since preview only allow single page,EndPage parameter does not
'           affect preview. It might be useful in future when multiple pages preview
'           is supported
'
'************************************************************
On Error GoTo PreviewPageErr

    Dim lIndexOffset As Long 'offset for index file
    Dim lRet As Long, lRead As Long 'for API
    Dim lPageRead As Long, lPageSize As Long
    Dim lHeaderSize As Long
    Dim lPage As Long, bFinish As Boolean
    Dim bCancel As Boolean
    If Not OpenRptTempFile("r") Then
        Err.Raise ERR_OPENING_RPT_FILE, App.ProductName, LoadResString(ERR_OPENING_RPT_FILE - vbObjectError)
        Exit Sub
    End If
    '
    
    If StartPage = 0 Then StartPage = 1
    If EndPage = 0 Then EndPage = Me.TotalPages
    
    InitTarget
    'look index file
    lPage = StartPage
    Do While Not bFinish
        If Not bmPrinterFlag Then
            PrintTarget.Cls
        End If
        'get index offset
        lIndexOffset = (lPage - 1) * Len(tmPageIndex)
        'position the index file
        lRet = fseek(ReportInfo.hIndexFile, lIndexOffset, 0)
        
        'read page summary to get the offset & size for requested page
        lRet = ReadFile(ReportInfo.hIndexFile, ByVal tmPageIndex, Len(tmPageIndex), lRead, ByVal 0&)
       
        'position the report data file
        lRet = fseek(ReportInfo.hFile, tmPageIndex.PageOffset, 0)
        'read page summary
        lRet = ReadFile(ReportInfo.hFile, tmPageHeader, ByVal Len(tmPageHeader), lRead, ByVal 0&)
        ReportInfo.CurrentPage = tmPageHeader.PageNo 'just reuse this
        'raise event & allow user to cancel printing
        If bmPrinterFlag Then
            RaiseEvent Printing(lPage, bCancel)
            If bCancel Then
                fclose ReportInfo.hFile
                fclose ReportInfo.hIndexFile
                Printer.KillDoc
                Exit Sub
            End If
        End If
        
        'read individual item
        Do While lPageRead < tmPageIndex.PageSize
            lRet = ReadFile(ReportInfo.hFile, tmItemInfo, ByVal Len(tmItemInfo), lRead, ByVal 0&)
            If lRead = 0 Then Exit Do
            lPageRead = lPageRead + lRead
            PrintItem
        Loop
        lPage = lPage + 1
        lPageRead = 0 'reset size of page read
        If (lPage > EndPage) Or (lPage > Me.TotalPages) Then
            bFinish = True
        Else
            If bmPrinterFlag Then
                Printer.NewPage
            End If
        End If
    Loop
    If bmPrinterFlag Then
        Printer.EndDoc
    End If
    fclose (ReportInfo.hFile)
    fclose (ReportInfo.hIndexFile)

PreviewPageEXIT:

Exit Sub
PreviewPageErr:
        fclose (ReportInfo.hFile)
        fclose (ReportInfo.hIndexFile)
      
End Sub
'##ModelId=3B3781BE0212
Public Sub PrintReport(Optional ShowDialog As Boolean = False, Optional ByVal StartPage As Integer = 0, Optional ByVal EndPage As Integer = 0, Optional ByVal hWndOwner As Long = 0)
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine will be called by client application for direct printing
'                 the report without previwing first
'
'
'          Input:
'
'              EndPage - Last page to be printed
'              StartPage - Starting page to be printed
'              ShowDialog - If True, print dialog box will be displayed to allow user
'                           to change printer properties
'                           False - no dialog will be displayed and printing will used
'                           any default printer/properties
'              hWndOwner - dialog owner (if any)
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo PrintReportErr

    ValidateReport
   If Not bmFileCreated Then
        If Not rst Is Nothing Then
            If rst.AbsolutePosition <> 1 Then rst.MoveFirst
        End If
        If Not OpenRptTempFile("w") Then
            Err.Raise vbObjectError + ERR_CREATING_RPT_FILE, MOD_NAME, LoadResString(ERR_CREATING_RPT_FILE)
            Exit Sub
        End If
        WriteReportToFile
    End If
    If ShowDialog Then
        ShowPrinterDialog 1, hWndOwner
    End If
    
    ReportInfo.CurrentPage = 0
    
    Me.PreviewRatio = 1
    
PrintReportEXIT:

Exit Sub
PrintReportErr:
          Err.Raise Err.Number, Err.Source, Err.Description
          GoTo PrintReportEXIT
End Sub
'##ModelId=3B3781BF02A8
Private Function SetXpos(ByVal strText As String, ByVal XPos As Single, ByVal Allign As JxDBPrintAllign) As Single
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine calculates X position for printing based on supplied properties
'                 It is called during the actual preview/printing only
'                 and not during report file generation/creation
'
'          Input:
'
'              Allign - Printing allignment (Left,Right,Centre)
'              XPos - Item X position (used for calculation)
'              strText - Text to be printed (used for TextWidth)
'
'          OutPut: X Position for printtarget object
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo SetXposErr

    
    Dim x As Single
    'x = PrintTarget.ScaleWidth - PrintTarget.TextWidth(strText)
    x = XPos + ReportInfo.LeftM 'left allign
    strText = RTrim$(strText)
    If Allign = JxDBAllignRight Then
        If XPos = 0 Then
            x = (PrintTarget.ScaleWidth - ReportInfo.RightM) - PrintTarget.TextWidth(strText)
        Else
        'justify around xpos
            x = XPos - PrintTarget.TextWidth(strText) + ReportInfo.LeftM
           '  x = XPos + ReportInfo.LeftM
        End If
    End If
    If Allign = JxDBAllignCenter Then
        If XPos = 0 Then
            x = (PrintTarget.ScaleWidth - PrintTarget.TextWidth(strText)) / 2
        Else
            x = XPos - PrintTarget.TextWidth(strText) / 2 + ReportInfo.LeftM
        End If
    End If
    SetXpos = x
SetXposEXIT:

Exit Function
SetXposErr:
          GoTo SetXposEXIT
End Function
'##ModelId=3B3781C0029E
Friend Function GetTextHeight(ByVal sText As String) As Single
    GetTextHeight = PrintTarget.TextHeight(sText)
    'If Not bmPrinterFlag Then
     '   TextHeight = PrinterRatio.Y * TextHeight
    'End If
End Function

'##ModelId=3B3781C1006E

'##ModelId=3B3781C101F4
Private Sub InitTarget()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine set up the print target object for required size/scale
'
'
'          Input: None
'
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo InitTargetErr

    Dim lRet As Long
    Dim PrinterhDC As Long
    Static bFirst As Boolean
    bmPrinterFlag = False
    If TypeOf PrintTarget Is Printer Then
        bmPrinterFlag = True
    End If
    If bmPrinterFlag Then
     '   Printer.TrackDefault
        With Printer
        
    '  .ScaleMode = Me.ReportScale
            If Me.PaperSize = JxDBRptPaperCustom Then
                .ScaleMode = Me.ReportScale
                .PaperSize = vbPRPSUser
                .Width = ReportInfo.Width
                .Height = ReportInfo.Height
            Else
                .PaperSize = Me.PaperSize
            End If
            '
        End With
    End If


    PrintTarget.Scale (0, 0)-(ReportInfo.Width, ReportInfo.Height)
    
    'initialize the printer
    If bmPrinterFlag Then
        Printer.Print "";
    End If

InitTargetEXIT:

Exit Sub
InitTargetErr:
          Err.Raise Err.Number, Err.Source & SetErrSource(MOD_NAME, "InitTarget"), Err.Description
End Sub
'##ModelId=3B3781C102D0
Private Function CheckAdvanceLine(oItem As JxDBRptItem, Optional ByVal OnFooter As Boolean = False) As Boolean
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine checks for availbility of the printable height area.
'                 If current height plus item height exceeds printable area, it
'                 calls WritePageFooter and return True
'
'          Input:
'
'              oItem - item to be evaluated
'
'          OutPut: False if Current height + item height does not exceed printable height
'                  otherwise it returns True
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo CheckAdvanceLineErr

    Dim sText As String
    Dim fHeight As Single
'    If bDEBUG Then
 '       oTrace.LogProc "CheckAdvanceLine", MOD_NAME
  '  End If
    fHeight = oItem.Height
    'accumulate current height
    ReportInfo.CurrentY = ReportInfo.CurrentY + fHeight
    If Not OnFooter Then
        If ReportInfo.CurrentY >= ReportInfo.AvailHeight Then
            WritePageFooter
            CheckAdvanceLine = True
            Exit Function
        End If
    End If
    
CheckAdvanceLineEXIT:

Exit Function
CheckAdvanceLineErr:
          Err.Raise Err.Number, Err.Source & SetErrSource(MOD_NAME, "CheckAdvanceLine"), Err.Description
End Function
Private Function GetAnchorText(ByVal oItem As JxDBRptItem) As Variant
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is used to setup the item value/text to be popup During Display
'                 based on the item type
'
'          Input:
'
'              oItem - Report item to be setup
'
'
'************************************************************
Dim StrSplitFld() As String
Dim SFld As String, SAnchorId As String
Dim j As Long
On Error GoTo GetAnchorTextErr
            Select Case oItem.ItemType
                Case JxDBItemLabelAnchor
                        If Trim$(oItem.FieldName) <> "" Then
                        StrSplitFld = Split(oItem.FieldName, ",")
                        
                        For j = 0 To UBound(StrSplitFld)
                            SFld = StrSplitFld(j) 'get the field name
                            SAnchorId = SAnchorId & "" & rst.Fields(SFld).Value '& Chr(10)
                        Next j
                        
                        End If
            End Select
            GetAnchorText = SAnchorId
Exit Function
GetAnchorTextErr:
          If Err.Number = adErrItemNotFound Then 'field not found
            Err.Raise vbObjectError + ERR_NO_SUCH_FIELD, SetErrSource(MOD_NAME, "GetAnchorText"), LoadResString(ERR_NO_SUCH_FIELD) & " -" & oItem.FieldName
          Else
            Err.Raise Err.Number, Err.Source & SetErrSource(MOD_NAME, "GetAnchorText"), Err.Description
          End If

End Function
'##ModelId=3B3781C20172
Private Function GetPrintText(ByVal oItem As JxDBRptItem) As Variant
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is used to setup the item value/text to be printed
'                 based on the item type
'
'          Input:
'
'              oItem - Report item to be setup
'
'          OutPut:
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo GetPrintTextErr

    Dim vntData As Variant
    'oTrace.LogProc "GetPrintText", MOD_NAME
    
    Select Case oItem.ItemType
        Case JxDBRptItemType.JxDBItemLabel
        Case JxDBRptItemType.JxDBItemFormula
            Select Case oItem.FormulaType
                Case JxDBRptFormulaType.JxDBRptFormulaAverage
                    oItem.Value = oItem.Value / oItem.ValueExtra
                Case JxDBRptFormulaType.JxDBRptFormulaHighest
                Case JxDBRptFormulaType.JxDBRptFormulaLowest
            End Select
        Case JxDBRptItemType.JxDBItemCounter
            oItem.Value = oItem.Value + 1
        Case JxDBRptItemType.JxDBItemPageXofY
        Case JxDBRptItemType.JxDBItemCurrentDate
            oItem.Value = Now
        Case JxDBRptItemType.JxDBItemCurrentPage
            oItem.Value = ReportInfo.CurrentPage
        Case JxDBRptItemType.JxDBItemTotalPage
        Case JxDBRptItemType.JxDBItemDataField
        If rst.AbsolutePosition <> adPosBOF And rst.AbsolutePosition <> adPosEOF Then
            oItem.Value = rst.Fields(Trim$(oItem.FieldName)).Value
        End If
        Case JxDBRptItemType.JxDBItemRunningTotal
             oItem.Value = IIf(IsEmpty(oItem.Value), rst.Fields(oItem.FieldName).Value, oItem.Value + rst.Fields(oItem.FieldName).Value)
        Case JxDBRptItemType.JxDBItemCustom
    End Select
    vntData = oItem.Value
    If Len(oItem.FormatString) Then
        vntData = Format(vntData, oItem.FormatString)
    End If
    GetPrintText = vntData
GetPrintTextEXIT:

Exit Function
GetPrintTextErr:
          If Err.Number = adErrItemNotFound Then 'field not found
            Err.Raise vbObjectError + ERR_NO_SUCH_FIELD, SetErrSource(MOD_NAME, "GetPrintText"), LoadResString(ERR_NO_SUCH_FIELD) & " -" & oItem.FieldName
          Else
            Err.Raise Err.Number, Err.Source & SetErrSource(MOD_NAME, "GetPrintText"), Err.Description
          End If
End Function
'##ModelId=3B3781C20366
Private Function OpenRptTempFile(ByVal mode As String) As Boolean
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is responsible for opening/creating report data and index
'                 file. Opening/Creating depends on the supplied mode
'
'          Input:
'
'              mode - Open mode
'              "w" - if writing new report file
'              "r" - if reading existing report file
'
'          OutPut: Return True if no error occured during opening both data and index
'                  report files, otherwise it returns False
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo OpenRptTempFileErr

    ReportInfo.hFile = fopen(App.Path & "\temp\report.rpt", mode)
    ReportInfo.hIndexFile = fopen(App.Path & "\temp\report.idx", mode)
    If (ReportInfo.hFile <> INVALID_HANDLE_VALUE) And (ReportInfo.hIndexFile <> INVALID_HANDLE_VALUE) Then
        bmFileOpened = True
        OpenRptTempFile = True
    End If
OpenRptTempFileEXIT:

Exit Function
OpenRptTempFileErr:
          GoTo OpenRptTempFileEXIT
End Function
'##ModelId=3B3781C30168
Private Sub WriteReportToFile()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is the central for generating report either
'                 for preview or printing.
'
'
'          Input: None
'
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo WriteReportToFileErr

    Dim i As Integer
    Dim ssavePreviewRatio As Single
    Dim lRet As Long, lWritten As Long 'API return
    Dim lErr As Long, sErrSource As String, sErrDesc As String
    'Dim bFinish As Boolean ' flag to indicate printing is finished
    Static lCounter As Long
    If Not rst Is Nothing Then
        If rst.AbsolutePosition <> 1 Then
                 rst.MoveFirst 'reset rst position
        End If
    End If
    ReportInfo.NewPage = False
    
        
    'write the page signiture
    bInitial = True
    SetupNewPage
    WritePageHeader
    WriteGroup
    
    Call fclose(ReportInfo.hFile)
    Call fclose(ReportInfo.hIndexFile)
    'set the file flag
    bmFileCreated = True
WriteReportToFileEXIT:
    
Exit Sub
WriteReportToFileErr:
          SaveError Err.Number, Err.Source, Err.Description, Err.HelpContext
          'do some cleanup
          fclose (ReportInfo.hFile)
          fclose (ReportInfo.hIndexFile)
          
          Err.Raise rptError.ErrNumber, rptError.ErrSource, rptError.ErrDesc, rptError.ErrHelpFile, rptError.ErrHelpContext
          
End Sub
'##ModelId=3B3781C30212
Private Sub WritePageHeader()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is called for every new page to be generated/printing.
'                 It loops for all items in PageHeader and writes it to report data file
'
'          Input: None
'
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo WritePageHeaderErr
    Dim i As Integer
    Dim oItem As JxDBRptItem
    Dim lRet As Long, lWritten As Long 'for API
    If Me.PageHeader.Count = 0 Then Exit Sub
    
    If ReportInfo.CurrentPage > 1 Then
        'if not the first page and page header not to be repeated
        If Not PageHeader.RepeatHeader Then Exit Sub
    End If
    ReportInfo.CurrentY = ReportInfo.TopM
    
    
    
    For i = 1 To Me.PageHeader.Count
        Set oItem = PageHeader.Item(i)
        If oItem.PreAdvanceLine Then
            If CheckAdvanceLine(oItem) Then
                Exit Sub
            End If
        End If
        SetItemProperties oItem
        'check text for printing
'        tmItemInfo.XPos = oItem.XPosition 'SetXpos(oItem.Value, oItem.XPosition, oItem.PrintAllign)
                        
        lRet = WriteFile(ReportInfo.hFile, tmItemInfo, Len(tmItemInfo), lWritten, ByVal 0&)
        tmPageIndex.PageSize = tmPageIndex.PageSize + lWritten
        'ReportInfo.CurrentY = ReportInfo.CurrentY + PrintTarget.TextHeight(oLabel.Caption) * PageHeader.Count
        If oItem.PostAdvanceLine Then
            If CheckAdvanceLine(oItem) Then
                Exit Sub
            End If
        End If
        
    Next i
WritePageHeaderEXIT:

Exit Sub
WritePageHeaderErr:
          Err.Raise Err.Number, Err.Source, Err.Description
          GoTo WritePageHeaderEXIT
End Sub
'##ModelId=3B3781C302EE
Private Sub WriteGroup()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is the main loop for report group printing. This rutine
'                 is called only once during the whole report generating cycle.
'                 It loops for each record in the recordset(rst) and;
'
'                 1.Check for any group break, if any, it calls WriteGroupBreak sub
'                 2.If no break occurs, it calls WriteReportGroup for normal group printing
'                   by passing the most detail group(most break occured)
'                 3. If first time called, it calls WriteReportGroup with starting from
'                   less break occured group (index=1)
'
'          Input: None
'
'
'          OutPut:None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo WriteGroupErr

    'Dim oField As JxDBRptField
'    If bDEBUG Then
'        oTrace.LogProc "WriteGroup", MOD_NAME
'    End If
    Dim i As Integer, x As Integer, BreakGroup As Integer
    Dim retval As Long 'for API calls
    Dim sItem As String
    Dim oGroup As JxDBRptGroup
    Dim iTotalGroup As Integer
    'PrintTarget.CurrentX = ReportInfo.LeftM
    'On Error Resume Next
    iTotalGroup = Me.ReportGroups.Count
    If ReportInfo.CurrentPage = 1 Then
        Me.ReportGroups.ResetFormulaItem
    End If
    If iTotalGroup = 0 Then Exit Sub
    
    Do Until rst.EOF
            mvarRecordProcessed = mvarRecordProcessed + 1
            
            
            'Set oGroup = Me.ReportGroups.Item(i)
            'check for break
            BreakGroup = CheckGroupBreak(iTotalGroup)
            
            If bInitial Then
                WriteReportGroup 1, iTotalGroup
                bInitial = False 'set the initial flag
            ElseIf BreakGroup Then
            'if any break occurs and not initial
                WriteGroupBreak iTotalGroup
                ResetCounterField BreakGroup, iTotalGroup
                WriteReportGroup BreakGroup, iTotalGroup
            ElseIf Not BreakGroup Then
            'if no break occurs
                WriteReportGroup iTotalGroup, iTotalGroup
            End If
            AccumulateTotal
            
       
        rst.MoveNext
    Loop
    'no more records?
    If rst.EOF Then
        'all group break
        WriteGroupBreak iTotalGroup
        WritePageFooter
        EndOfReport
        
    End If
    
    
WriteGroupEXIT:

Exit Sub
WriteGroupErr:
          Err.Raise Err.Number, SetErrSource(MOD_NAME, "WriteGroup") & Err.Source, Err.Description
          'GoTo WriteGroupEXIT
End Sub

'##ModelId=3B3781C303CA
Private Sub WriteReportGroup(ByVal StartFromGroup As Integer, ByVal TotalGroup As Integer)
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine loop for group starting from index StartFromGroup until
'                 TotalGroup. It then calls PrintGroupItem for each group
'
'          Input:
'
'              TotalGroup -  Number of groups in report
'              StartFromGroup - Starting group index to be printed
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo WriteReportGroupErr

'    If bDEBUG Then
'        oTrace.LogProc "PrintReportGroup", MOD_NAME
'    End If
    Dim bBreakLevel As Boolean
    Dim i As Integer
    Dim oGroup As JxDBRptGroup
    
    If rst.EOF Then Exit Sub
    
    'If ReportInfo.NewPage Then Exit Sub
    For i = StartFromGroup To TotalGroup
        Set oGroup = Me.ReportGroups.Item(i)
        If oGroup.PrintGroup Then
            WriteGroupItem oGroup
        End If
    Next i
    
WriteReportGroupEXIT:

Exit Sub
WriteReportGroupErr:
          Err.Raise Err.Number, SetErrSource(MOD_NAME, "WriteReportGroup") & Err.Source, Err.Description
End Sub
'##ModelId=3B3781C402E4
Private Sub WriteGroupItem(ByVal oGroup As JxDBRptGroup)
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is loops for each print item in the group
'                 and write it to report data file
'
'
'          Input:
'
'              oGroup - Report group
'
'          OutPut:
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo WriteGroupItemErr

'    If bDEBUG Then
'        oTrace.LogProc "WriteGroupItem", MOD_NAME
'    End If
    Dim iIndex As Integer
    Dim oItem As JxDBRptItem
    Dim vntPrintItem As Variant
    Dim lRet As Long, lWritten As Long 'API
    
    For iIndex = 1 To oGroup.PrintItemsCount
        Set oItem = oGroup.PrintItems(iIndex)
        If oItem.PreAdvanceLine Then
            If CheckAdvanceLine(oItem) Then
                SetupNewPage
                WritePageHeader
            End If
        End If
        If oItem.ItemType = JxDBItemCustom Then
            RaiseEvent PrintCustomItem(oItem)
        End If
        
        SetItemProperties oItem
        lRet = WriteFile(ReportInfo.hFile, tmItemInfo, Len(tmItemInfo), lWritten, ByVal 0&)
        tmPageIndex.PageSize = tmPageIndex.PageSize + lWritten
        'Debug.Assert Len(ItemInfo.Value) > 50
        If oItem.PostAdvanceLine Then
            If CheckAdvanceLine(oItem) Then
                SetupNewPage
                WritePageHeader
            End If
        End If
    Next iIndex
WriteGroupItemEXIT:

Exit Sub
WriteGroupItemErr:
          Err.Raise Err.Number, SetErrSource(MOD_NAME, "WriteGroupItem") & Err.Source, Err.Description
End Sub
'##ModelId=3B3781C500E6
Private Sub SetFontItem(oFont As StdFont)
'************************************************************
'      Created by: joy     Date: 19/04/
'
'
'          Description:
'                 This Rutine
'
'
'          Input:
'
'              oFont
'
'          OutPut:
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo SetFontItemErr

    With tmItemInfo
        .FontBold = oFont.Bold
        .FontItalic = oFont.Italic
        .FontName = oFont.Name
        .FontSize = oFont.size
        .FontStrikethru = oFont.Strikethrough
        .FontUnderline = oFont.Underline
    End With
SetFontItemEXIT:

Exit Sub
SetFontItemErr:
          Err.Raise Err.Number, Err.Source & SetErrSource(MOD_NAME, "SetFontItem"), Err.Description
End Sub
'##ModelId=3B3781C502DA
Private Sub SetItemProperties(oItem As JxDBRptItem)
     'setup item prop
     With tmItemInfo
        .FontBold = oItem.Font.Bold
        .FontItalic = oItem.Font.Italic
        .FontName = oItem.Font.Name
        .FontSize = oItem.Font.size
        .FontStrikethru = oItem.Font.Strikethrough
        .FontUnderline = oItem.Font.Underline
        .FontColor = oItem.FontColor
        .ItemType = oItem.ItemType
    End With
     
    tmItemInfo.Value = GetPrintText(oItem)
    tmItemInfo.PopUpValue = GetAnchorText(oItem)
    tmItemInfo.Allign = oItem.PrintAllign
    tmItemInfo.XPos = oItem.XPosition
    tmItemInfo.YPos = ReportInfo.CurrentY

End Sub
'##ModelId=3B3781C600DC
Private Sub WriteGroupBreak(ByVal TotalGroup As Integer)
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is called when any group break occured.It loops
'                 for all groups and look for group that has these two properties
'                 set to true (PrintTotal and GroupBreak). If a group satisfy these
'                 conditions, it loops for each item in that group and write it to
'                 report data file.
'                 It also reset(set to 0) any formula type item for the group
'
'          Input:
'
'              TotalGroup - number of total groups in the report
'
'          OutPut: None
'
'
'          Globals 0Modfied
'
'
'************************************************************
On Error GoTo WriteGroupBreakErr
'    If bDEBUG Then
'        oTrace.LogProc "WriteGroupBreak", MOD_NAME
'    End If
    Dim oItem As JxDBRptItem
    Dim oGroup As JxDBRptGroup
    Dim i As Integer, j As Integer
    Dim vPrintItem As Variant
    Dim lRet As Long, lWritten As Long
    'start loop from detail group
    For i = TotalGroup To 1 Step -1
        Set oGroup = Me.ReportGroups.Item(i)
        If (oGroup.GroupBreak) Or (rst.EOF) Then
            If Not oGroup.UpperGroup Is Nothing Then
                oGroup.UpperGroup.LowerGroupBreak
            End If
        End If
        If oGroup.PrintGroupBreak And (oGroup.GroupBreak Or rst.EOF) Then
            For j = 1 To oGroup.BreakItemCount
                'get the group break items
                Set oItem = oGroup.BreakItems(j)
                '
                If oItem.PreAdvanceLine Then
                    If CheckAdvanceLine(oItem) Then
                        SetupNewPage
                        WritePageHeader
                    End If
                End If
                
                'setup print target
                SetItemProperties oItem
                'we need to compare these two formula during actual group break
                '
                If oItem.ItemType = JxDBItemFormula Then
                    Select Case oItem.FormulaType
                        Case JxDBRptFormulaType.JxDBRptFormulaHighest, JxDBRptFormulaType.JxDBRptFormulaLowest
                            If IsEmpty(oItem.Value) Then
                                oItem.Value = oItem.ValueExtra
                            End If
                            If oItem.FormulaType = JxDBRptFormulaHighest Then
                                oItem.Value = IIf(oItem.Value > oItem.ValueExtra, oItem.Value, oItem.ValueExtra)
                            Else
                                If Not IsEmpty(oItem.ValueExtra) Then
                                oItem.Value = IIf(oItem.Value < oItem.ValueExtra, oItem.Value, oItem.ValueExtra)
                                End If
                            End If
                            oItem.ValueExtra = Empty
                            tmItemInfo.Value = IIf(Len(oItem.FormatString), Format(oItem.Value, oItem.FormatString), oItem.Value)
                    End Select
                End If
                    
                
                
                'write it
                lRet = WriteFile(ReportInfo.hFile, tmItemInfo, Len(tmItemInfo), lWritten, ByVal 0&)
                tmPageIndex.PageSize = tmPageIndex.PageSize + lWritten
                'reset all formula type item
                If oItem.ItemType = JxDBItemFormula Then
                    oItem.Value = Empty
                End If
                
                
                'print advance line if any
                If oItem.PostAdvanceLine Then
                    If CheckAdvanceLine(oItem) Then
                        SetupNewPage
                        WritePageHeader
                    End If
                End If
            Next j
        End If
    Next i
    
WriteGroupBreakEXIT:

Exit Sub
WriteGroupBreakErr:
          Err.Raise Err.Number, Err.Source, Err.Description
          'GoTo WriteGroupBreakEXIT
End Sub
'##ModelId=3B3781C602C6
Private Sub WritePageFooter()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is responsible in writing report page footer
'                 to report data file. It is called when current page height
'                 about to exceed report printable area
'
'          Input: None
'
'
'          OutPut: None
'
'
'          Globals Modfied
'
'
'************************************************************
On Error GoTo WritePageFooterErr

      'oTrace.LogProc "PrintPageFooter", MOD_NAME
    
    Dim i As Integer
    Dim oItem As JxDBRptItem
    Dim lRet As Long, lWritten As Long
    
    If PageFooter.Count = 0 Then Exit Sub
    If ReportInfo.CurrentPage > 1 Then
        'if not the first page and page footer not to be repeated
        If Not PageFooter.RepeatHeader Then Exit Sub
    End If
    'set Y for page footer
    ReportInfo.CurrentY = ReportInfo.TopM + ReportInfo.HeaderM + ReportInfo.AvailHeight
    For i = 1 To PageFooter.Count
        Set oItem = PageFooter.Item(i)
        If oItem.PreAdvanceLine Then
            If CheckAdvanceLine(oItem, True) Then Exit Sub
        End If
        SetItemProperties oItem
        'SetFontItem oItem.Font
        
        'check for item type
        'get
        tmItemInfo.Value = GetPrintText(oItem)
        
        tmItemInfo.XPos = oItem.XPosition ' SetXpos(oItem.Value, oItem.XPosition, oItem.PrintAllign)
        tmItemInfo.YPos = ReportInfo.CurrentY
        tmItemInfo.Allign = oItem.PrintAllign
        lRet = WriteFile(ReportInfo.hFile, tmItemInfo, Len(tmItemInfo), lWritten, ByVal 0&)
        tmPageIndex.PageSize = tmPageIndex.PageSize + lWritten
        If oItem.PostAdvanceLine Then
            If CheckAdvanceLine(oItem, True) Then
                'Exit Sub
            End If
        End If
    Next i
WritePageFooterEXIT:

Exit Sub
WritePageFooterErr:
          GoTo WritePageFooterEXIT
End Sub
'##ModelId=3B3781C60370
Private Sub SetupNewPage()
'************************************************************
'      Created by: joy     Date: 22/06/2002
'
'
'          Description:
'                 This Rutine is called during initial report and during every page
'                 break (except for end of report).
'                It will
'                1. Increment current page number
'                2. Increment total page(s)
'                3. Write report index file (page#,pagesize,pageoffset)
'
'          Input: None
'
'
'          OutPut: None
'
'
'          Globals Modified
'
'
'************************************************************
On Error GoTo SetupNewPageErr

    Dim lRet As Long, lWritten As Long
    ReportInfo.CurrentY = 0
    ReportInfo.CurrentPage = ReportInfo.CurrentPage + 1
    ReportInfo.TotalPage = ReportInfo.CurrentPage
    
    'write page summary to index file
    If tmPageIndex.PageNo > 0 Then
        'write to report index file
        lRet = WriteFile(ReportInfo.hIndexFile, tmPageIndex, Len(tmPageIndex), lWritten, ByVal 0&)
        tmPageIndex.PageOffset = tmPageIndex.PageOffset + tmPageIndex.PageSize + Len(tmPageHeader)
        
    End If
    tmPageIndex.PageNo = ReportInfo.CurrentPage
    
    'reset page size
    tmPageIndex.PageSize = 0
    mvarTotalPages = ReportInfo.CurrentPage
    'setup page header in report data file
    tmPageHeader.NewPage = "PG#"
    tmPageHeader.PageNo = ReportInfo.CurrentPage
    
    lRet = WriteFile(ReportInfo.hFile, tmPageHeader, Len(tmPageHeader), lWritten, ByVal 0&)
    
SetupNewPageEXIT:

Exit Sub
SetupNewPageErr:
          GoTo SetupNewPageEXIT
End Sub
'##ModelId=3B3781C70064
Private Sub EndOfReport()
    Dim lRet As Long, lWritten As Long
    lRet = WriteFile(ReportInfo.hIndexFile, tmPageIndex, Len(tmPageIndex), lWritten, ByVal 0&)
End Sub
'##ModelId=3B3781C70104
Private Sub WriteRptIndexFile()
    Dim lRet As Long, lWritten As Long
    
End Sub
'##ModelId=3B3781C701E0
Private Sub PrintItem()
Dim StrSplitFld() As String
    With tmItemInfo
        'some formula type need to be taken care of
        'during actual printing/previewing
        Select Case .ItemType
            Case JxDBRptItemType.JxDBItemTotalPage
                .Value = Me.TotalPages
            Case JxDBRptItemType.JxDBItemPageXofY
                   .Value = ReportInfo.CurrentPage & " of " & Me.TotalPages
           Case JxDBRptItemType.JxDBItemLabelAnchor
'                  Call Controls.Add("VB.ctlLabel", "ctlLabel1" & CStr(iCount), PrintTarget)
                    iCount = iCount + 1
                  Load frmPreview.ctlLabel1(iCount)
                  Set frmPreview.ctlLabel1(iCount).Container = PrintTarget
                 frmPreview.ctlLabel1(iCount).ctl.Caption = tmItemInfo.Value
                 frmPreview.ctlLabel1(iCount).Left = SetXpos(Trim(.Value), .XPos, .Allign)
                 frmPreview.ctlLabel1(iCount).Top = .YPos
                  frmPreview.ctlLabel1(iCount).Tag = tmItemInfo.PopUpValue
                    Exit Sub
'                  strsPlitFld = Split(tmItemInfo., ",")
'
'                With PrintTarget.txtAnchor(iCount)
'                    .Font.Name = Trim(tmItemInfo.FontName)
'                    .Top = tmItemInfo.YPos
                    
 '               End With
        End Select
        PrintTarget.Font.Name = Trim(.FontName)

        PrintTarget.Font.size = .FontSize * Me.PreviewRatio
        PrintTarget.Font.Bold = .FontBold
        PrintTarget.Font.Italic = .FontItalic
        PrintTarget.Font.Strikethrough = .FontStrikethru
        PrintTarget.Font.Underline = .FontUnderline
        PrintTarget.ForeColor = .FontColor
        PrintTarget.CurrentY = .YPos
        PrintTarget.CurrentX = SetXpos(Trim(.Value), .XPos, .Allign)
        
        PrintTarget.Print Trim(.Value);
       
    End With
    If bmPrinterFlag Then
        Debug.Print Printer.CurrentY
    End If
    'PrintTarget.Print Trim(tmItemInfo.Value);
    
            
End Sub
'##ModelId=3B3781C702BC
Friend Function ShowPrinterDialog(Optional ByVal CurrentPage As Long = 0, Optional ByVal hWndOwner As Long = 0) As Boolean
    Dim lPrintStart As Integer, lPrintEnd As Integer
    Dim lRet As Long
    Dim lStartPage As Long, lEndPage As Long, lCopy As Long
    Dim sDeviceName As String, sDefPrinter As String
    Dim oPrinter As Printer
    Dim i As Integer
    'get default printer
    sDefPrinter = Printer.DeviceName
    
    lCopy = 1
    lStartPage = 1
    lEndPage = Me.TotalPages
    sDeviceName = Printer.DeviceName

    lRet = ShowPrintDialog(hWndOwner, sDeviceName, lCopy, lStartPage, lEndPage)
    If lRet = 0 Then
        Exit Function
    End If
    'user selects selection
    If lStartPage = -1 Then
        lStartPage = CurrentPage
        lEndPage = CurrentPage
    End If
    
    sDeviceName = GetString(sDeviceName)
    'set printer to default
    For Each oPrinter In Printers
        If oPrinter.DeviceName = sDeviceName Then
            Set Printer = oPrinter
            Exit For
        End If
    Next
    
    Set Me.Target = Printer
    Me.PrintIt lStartPage, lEndPage
    'restore printer
    For Each oPrinter In Printers
        If oPrinter.DeviceName = sDefPrinter Then
            Set Printer = oPrinter
            Exit For
        End If
    Next
End Function
'##ModelId=3B3781C801A4
Private Sub ValidateReport()
    'reset initial report info
    mvarRecordProcessed = 0
    If Printers.Count = 0 Then
        Err.Raise vbObjectError + ERR_NO_PRINTER, MOD_NAME, LoadResString(ERR_NO_PRINTER)
        Exit Sub
    End If
    If rst Is Nothing Then
        Err.Raise vbObjectError + ERR_NO_DATASOURCE, MOD_NAME, LoadResString(ERR_NO_DATASOURCE)
        Exit Sub
    End If
    If rst.AbsolutePosition <> 1 Then rst.MoveFirst
    If (rst.EOF) Or (rst.BOF) Then
        Err.Raise vbObjectError + ERR_NO_RECORD, MOD_NAME, LoadResString(ERR_NO_RECORD)
        Exit Sub
    End If
    If (ReportInfo.Width = 0) Or (ReportInfo.Height = 0) Then
        Err.Raise vbObjectError + ERR_REPORT_SIZE_UNDEFINED, MOD_NAME, LoadResString(ERR_REPORT_SIZE_UNDEFINED)
        Exit Sub
    End If
    
End Sub

'##ModelId=3B3781C80244
Public Sub ShowAbout()
    frmAbout.Show vbModal
End Sub

Public Property Let DiscardSavedData(Data As Boolean)
    ShowwithPrevData = Data
End Property
Public Property Get DiscardSavedData() As Boolean
    DiscardSavedData = ShowwithPrevData

End Property

Public Sub DestroyHeader()
    Set PageHeader = Nothing
End Sub
Public Sub DestroyFooter()
    Set ReportFooter = Nothing
End Sub

