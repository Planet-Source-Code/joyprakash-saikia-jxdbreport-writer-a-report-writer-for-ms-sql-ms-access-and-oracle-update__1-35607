VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JxDBRptGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This Class is the Main and Most Infortant Class on the Heirarchy "
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PrintFields"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781DB02BC"
'**********************************************************************
'
'           Module Name: JxDBRptGroup.cls
'
'           Purpose    : Hold information about report group and a colletion
'                        class for Report Item (JxDBRptItem)
'
'           Author     : Joyprakash Saikia
'           Created On : 16/03/2002
'
'**********************************************************************
Option Explicit
'This Constant is Used for Debugging Purpose
'This is not used on this submission
Private Const MOD_NAME = "JxDBReport.JxDBRptGroup"

Private mvarFields As Variant 'list of group fields
Private mvarLevel As Integer 'local copy of GroupLevel "NOT USED"
Private mvarColGroupFields As Collection ' To Hold All Data Field Name in a group
Private mvarGroupId As Variant 'local copy of group id (checking for group break)
Private mvarPrintItems As JxDBRptItems 'hold list of items To be printed
Private mvarPrintGroupBreak As Boolean 'local copy of Print Group Break
Private mvarGroupBreakItems As JxDBRptItems 'hold list of items in group break
Private mvarGroupBreak As Boolean 'local copy of GroupBreak
Private mvarCheckForBreak As Boolean 'local copy of Check Group Break
Private mvarGroupName As String 'local copy of Group Name
Private mvarPrintGroup As Boolean 'local copy
Private mvarHasFormulaItem As Boolean 'flag wheter a group has formula type item
Private Enum JxDBGroupReset
    JxDBResetAll = 0 ' reset all formula type value
    JxDBResetValue = 1 ' reset value only
    JxDBResetExtra = 2 ' reset extra value only
End Enum
'trying using linked list
Private oGroupUpper As JxDBRptGroup 'upper group (less break occurs)
Private oGroupLower As JxDBRptGroup 'lower group (more break occurs)
Friend Function SetUpperGroup(oGroup As JxDBRptGroup)
    '************************************************************
    ' Description :
    ' Add pointer to a Group which upper to Current One
    ' This functionality is Similar with Linked List
    ' Example to the person having Knowledge of C/C++
    '             node=root;
    '             node->next=Node;
    '
    '************************************************************
    Set oGroupUpper = oGroup
End Function
Friend Function SetLowerGroup(oGroup As JxDBRptGroup)
    '************************************************************
    ' Description :
    ' Add pointer to a Group which Lower to Current One
    ' This functionality is Similar with Linked List in C/C++
    ' Example to the person having Knowledge of C/C++
    '             node=root;
    '             node->previous=node;
    '
    '************************************************************
    Set oGroupLower = oGroup
End Function
Friend Property Get UpperGroup() As JxDBRptGroup
    'to Retrive Upper Group of Current One
    Set UpperGroup = oGroupUpper
End Property
Friend Property Get LowerGroup() As JxDBRptGroup
    'to Retrive Lower Group of Current One
    Set LowerGroup = oGroupLower
End Property
Public Sub ReleaseLinkedGroup(Optional Group As JxDBRptGroup)
    '************************************************************
    ' Routine to Destroy all the Items in a group
    ' This Function is Recursive
    '************************************************************
         If Not Group Is Nothing Then  ' And oGroupUpper Is Nothing Then
                 On Error Resume Next
                 If BreakItemCount <> 0 Or PrintItemsCount <> 0 Or GroupFieldsCount <> 0 Then
                 If Err.Number = 91 Then Exit Sub
                    Set mvarPrintItems = Nothing
                    Set mvarColGroupFields = Nothing
                    Set mvarGroupBreakItems = Nothing
                    ReleaseLinkedGroup oGroupLower
                Else
                    Set Group = Nothing
                End If
           End If


End Sub
Friend Property Let HasFormulaItem(ByVal vData As Boolean)
    ' If the Group Has a item of Formula
    mvarHasFormulaItem = vData
End Property
Friend Property Get HasFormulaItem() As Boolean
     'Returns The Property
    HasFormulaItem = mvarHasFormulaItem
End Property

Public Property Let PrintGroup(ByVal vData As Boolean)
    '************************************************************
    'A group should be displayed or printed or not
    ' Default is True
    '************************************************************
    mvarPrintGroup = vData
End Property
Public Property Get PrintGroup() As Boolean
    ' Returns the Property
    PrintGroup = mvarPrintGroup
End Property
Public Property Let GroupName(ByVal vData As String)
    '************************************************************
    'To assign an Identification to Group by name
    'This Value Should be unique
    '************************************************************
    mvarGroupName = vData
End Property
Public Property Get GroupName() As String
    ' Returns the Property
    GroupName = mvarGroupName
End Property
'##ModelId=
Public Property Let CheckForBreak(ByVal vData As Boolean)
    '************************************************************
    'Sets whether the Group is to be verified for Break (Line Break)
    ' Used on Creation of report
    '************************************************************
    mvarCheckForBreak = vData
End Property


Public Property Get CheckForBreak() As Boolean
    ' Returns the Property
    CheckForBreak = mvarCheckForBreak
End Property

Public Property Let GroupBreak(ByVal vData As Boolean)
    '************************************************************
    ' Sets whether the Group is to be verified as a Break Item
    ' Used on Creation of report
    '************************************************************
    mvarGroupBreak = vData
End Property
Public Property Get GroupBreak() As Boolean
    ' Returns the Property
    GroupBreak = mvarGroupBreak
End Property

Public Property Let PrintGroupBreak(ByVal vData As Boolean)
    '************************************************************
    ' Whether a Line break should appear on not on the Item
    '************************************************************
    mvarPrintGroupBreak = vData
End Property


Public Property Get PrintGroupBreak() As Boolean
    ' Returns the Property
    PrintGroupBreak = mvarPrintGroupBreak
End Property
Public Property Get PrintItemsCount() As Integer
    '************************************************************
    'Description :
    'Returns Total No. of Items, that are printed on the Report for this Group
    '************************************************************
    PrintItemsCount = mvarPrintItems.Count
End Property
Public Property Get BreakItems() As JxDBRptItems

    If mvarGroupBreakItems Is Nothing Then Set mvarGroupBreakItems = New JxDBRptItems
    
    Set BreakItems = mvarGroupBreakItems
End Property
Public Property Get PrintItems() As JxDBRptItems

    If mvarPrintItems Is Nothing Then Set mvarPrintItems = New JxDBRptItems
    Set PrintItems = mvarPrintItems
End Property

Public Sub AddGroupField(ByVal FieldName As String)
    '************************************************************
    'description :
    '       Add  a database Column Name or Field Name to the Fields Collection
    '************************************************************
    mvarColGroupFields.Add FieldName
End Sub


Public Sub AddGroupFields(ParamArray args1())
    '************************************************************
    'description :
    '       Add  a List of database Column Name or Field Name to the Fields Collection
    '       Column Names Should Be Separated By Comma
    '************************************************************
    Dim sSplit() As String
    Dim i As Integer
    Dim sField As String
    
    For i = 0 To UBound(args1)
        sField = args1(i)
        mvarColGroupFields.Add sField
    Next
End Sub


Public Property Get GroupFieldsCount() As Integer
    GroupFieldsCount = mvarColGroupFields.Count
End Property

Friend Property Let GroupID(ByVal vData As Variant)
    mvarGroupId = vData
End Property
Friend Property Get GroupID() As Variant
    GroupID = mvarGroupId
End Property
Public Sub AddPrintItem(RptItem As JxDBRptItem)
'************************************************************
    'description :
    'create a new Item Object and
    ' And Add it to the Collection for Printed Items
'************************************************************
    Dim objNewMember As JxDBRptItem
    Set objNewMember = RptItem
    
    mvarPrintItems.AddItem objNewMember
End Sub
Public Function AddBreakItem(RptItem As JxDBRptItem, Optional ByVal sKey As String) As JxDBRptItem
Attribute AddBreakItem.VB_Description = "Add An Item to the group as A Breakitem. Parameters : an Object of JxDBRptItem  , Key (optional) \r\nA Break Item is neednot be a printed item on the report"
Attribute AddBreakItem.VB_HelpID = 20000
        '************************************************************
        '          Description:
        '                 This Routine add the item for Printing at the Time of Group Break
        '          Input:
        '              sKey - the Key for the Collection
        '              RptItem - the Item to be Added as break Item
        '
        '          OutPut: RptItem
        '       The Collection is Changed
        '
        '************************************************************

On Error GoTo AddBreakItemErr

    Dim objNewMember As JxDBRptItem
    Set objNewMember = RptItem
    If RptItem.ItemType = JxDBItemFormula Then
       Me.HasFormulaItem = True
    End If
    If Len(sKey) Then
         mvarGroupBreakItems.AddItem objNewMember, sKey
    Else
        mvarGroupBreakItems.AddItem objNewMember
    End If
    Set AddBreakItem = objNewMember
    Set objNewMember = Nothing
AddBreakItemEXIT:

Exit Function
AddBreakItemErr:
          GoTo AddBreakItemEXIT
End Function
Public Sub RemoveBreakItem(vntIndexKey As Variant)
    '************************************************************
    'Description:
    '       used when removing an element from the collection
    '************************************************************


      mvarGroupBreakItems.Remove vntIndexKey
End Sub
Public Property Get BreakItemCount() As Integer
    BreakItemCount = mvarGroupBreakItems.Count
End Property
Friend Property Get GroupFields() As Variant
        '************************************************************
        '          Description:
        ' Returns Group Fields separated By comma as a String
        '************************************************************


    Dim i As Integer
    Dim vData As Variant
    For i = 1 To mvarColGroupFields.Count
        vData = vData & mvarColGroupFields.Item(i)
        If Not i = mvarColGroupFields.Count Then
            vData = vData & ","
        End If
    Next i
    GroupFields = vData
End Property
Public Property Let GroupLevel(ByVal vData As Integer)
    mvarLevel = vData
End Property
Public Property Get GroupLevel() As Integer
    GroupLevel = mvarLevel
End Property
Friend Sub ResetFormulaItem()
'************************************************************
'          Description:
'                 This Rutine reset all formula type item in the group and this
'                 include RunningTotal and Counter type item
'
'          Input: NOne
'          OutPut:None
'
'************************************************************
On Error GoTo ResetFormulaItemErr
    Dim intBrkItemCount As Long
    Dim intMaxItemCount As Long
    Dim oItem As JxDBRptItem
    intMaxItemCount = Me.BreakItemCount 'Retrive the Property For first Time
    'Reset the Break Formula
    For intBrkItemCount = 1 To intMaxItemCount
        Set oItem = Me.BreakItems(intBrkItemCount)
        With oItem
            If .ItemType = JxDBItemFormula Then
                .Value = Empty
                .ValueExtra = Empty
            End If
        End With
    Next intBrkItemCount
    intBrkItemCount = Me.PrintItemsCount 'Retrive the Property For first Time
    'Reset the Printable Formula
    For intBrkItemCount = 1 To intMaxItemCount
        Set oItem = Me.PrintItems(intBrkItemCount)
        With oItem
            If (.ItemType = JxDBItemCounter) Or (.ItemType = JxDBItemRunningTotal) Then
                .Value = Empty
            End If
            If .ItemType = JxDBItemFormula Then
                .Value = Empty
                .ValueExtra = Empty
            End If
        End With
    Next intBrkItemCount
ResetFormulaItemEXIT:

Exit Sub
ResetFormulaItemErr:
          GoTo ResetFormulaItemEXIT
End Sub
Friend Sub ResetBreakFormulaItem(ByVal Flag As JxDBGroupReset)
        '************************************************************
        '   Description:
        '       Reinitialize the Value of Formula Items
        '       Input : Flag
        '       Output : None
        '************************************************************

    Dim intBrkItemCount As Long
    Dim intMaxBrkItemCount As Long
    Dim oItem As JxDBRptItem
    intMaxBrkItemCount = Me.BreakItemCount 'Retrive the Property For first Time
    For intBrkItemCount = 1 To intMaxBrkItemCount
        Set oItem = Me.BreakItems(intBrkItemCount)
        If oItem.ItemType = JxDBItemFormula Then
            With oItem
                If Flag = JxDBResetAll Then
                    .Value = Empty
                    .ValueExtra = Empty
                ElseIf Flag = JxDBResetExtra Then
                    .ValueExtra = Empty
                ElseIf Flag = JxDBResetValue Then
                    .Value = Empty
                End If
            End With
        End If
    Next intBrkItemCount
End Sub
Friend Sub LowerGroupBreak()
        '************************************************************
        '          Description:
        '                 This Routine is used to handle when any break occur to lower group.
        '                 It checks for formula type Highest & Lowest, move the value from
        '                 ValueExtra to Value and empty the ValueExtra.
        '                 It need to be done to avoid unnecessary summation
        '           Result:
        '                  The Collections are Modified
        '
        '
        '************************************************************
On Error GoTo LowerGroupBreakErr

    Dim intBrkItemCount As Long
    Dim intMaxBrkItemCount As Long
    Dim oItem As JxDBRptItem
    intMaxBrkItemCount = Me.BreakItemCount 'Retrive the Property For first Time
    For intBrkItemCount = 1 To intMaxBrkItemCount
        Set oItem = Me.BreakItems(intBrkItemCount)
        If oItem.ItemType = JxDBItemFormula Then
            Select Case oItem.FormulaType
                Case JxDBRptFormulaType.JxDBRptFormulaHighest, JxDBRptFormulaType.JxDBRptFormulaLowest
                    'move extra to value
                    If IsEmpty(oItem.Value) Then
                        oItem.Value = oItem.ValueExtra
                    End If
                    If oItem.FormulaType = JxDBRptFormulaHighest Then
                        oItem.Value = IIf(oItem.Value > oItem.ValueExtra, oItem.Value, oItem.ValueExtra)
                    Else
                        'we need this isEmpty coz empty is always less than anything
                        If Not IsEmpty(oItem.ValueExtra) Then
                            oItem.Value = IIf(oItem.Value < oItem.ValueExtra, oItem.Value, oItem.ValueExtra)
                        End If
                    End If
                    oItem.ValueExtra = Empty 'empty the extra Value
            End Select
        End If
    Next intBrkItemCount
LowerGroupBreakEXIT:

Exit Sub
LowerGroupBreakErr:
          GoTo LowerGroupBreakEXIT
End Sub
Public Sub ClearBreakItems()
    mvarGroupBreakItems.ClearItems
End Sub
Public Sub ClearPrintItems()
    mvarPrintItems.ClearItems
End Sub

Private Sub Class_Initialize()
        '************************************************************
        '          Description:
        'Initializes All Collection and Groups
        '************************************************************

        
    Set mvarColGroupFields = New Collection
    Set mvarPrintItems = New JxDBRptItems
    Set mvarGroupBreakItems = New JxDBRptItems
    Me.PrintGroup = True 'as Default Value
    Me.CheckForBreak = True 'as Default Value

End Sub
Private Sub Class_Terminate()
    '************************************************************
    '          Description:
    ' Release All the Collections
    ' Which In turn Release the Objects within themself
    '************************************************************

    
    Set mvarPrintItems = Nothing
    Set mvarColGroupFields = Nothing
    Set mvarGroupBreakItems = Nothing
End Sub

