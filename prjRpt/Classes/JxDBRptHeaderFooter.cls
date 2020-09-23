VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JxDBRptHeaderFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"iabRptLabel"
Attribute VB_Ext_KEY = "Member0" ,"iabRptLabel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781D800F0"
'**********************************************************************
'
'           Module Name: JxDBRptHeaderFooter.cls
'
'           Purpose    : Hold information about report Header and Footer
'                        and a collection class for Report Item (JxDBRptItem)
'
'           Author     : Joyprakash saikia
'
'**********************************************************************
Option Explicit
Private Const MOD_NAME = "JxDBReport.JxDBRptHeaderFooter"
Private oItems As JxDBRptItems 'Collection of JxDBRptItem
Private mvarRepeatHeader As Boolean

Public Property Let RepeatHeader(ByVal vData As Boolean)
       '************************************************************
       'Description:
       ' Sets the Value whether  this Object (header or Footer ) is to Be displayed on
       ' Each Page or Only on the First Page.
       '************************************************************

       
    mvarRepeatHeader = vData
End Property


Public Property Get RepeatHeader() As Boolean
    RepeatHeader = mvarRepeatHeader
End Property

Public Function AddItem(ReportItem As JxDBRptItem, Optional ByVal sKey As String) As JxDBRptItem
       
       '************************************************************
       'Description:
       '  Add an Item to ReportItems Collection for printing
       '************************************************************

       
    Dim objNewMember As JxDBRptItem
    Set objNewMember = ReportItem
    oItems.AddItem objNewMember, sKey
    Set AddItem = objNewMember
End Function
Public Property Get Item(vntIndexKey As Variant) As JxDBRptItem
Attribute Item.VB_UserMemId = 0
       '************************************************************
       'Description:
       'used when referencing an element in the collection
       '************************************************************


      Set Item = oItems(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = oItems.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    oItems.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        '************************************************************
        'Description:
        'this property allows you to enumerate
        'this collection with the For...Each syntax
        '************************************************************
       Set NewEnum = oItems.NewEnum
End Property


Private Sub Class_Initialize()
       '************************************************************
       'Description:
       'creates the collection of Items when this class is created
       '************************************************************

       
    Set oItems = New JxDBRptItems
End Sub

Private Sub Class_Terminate()
       '************************************************************
       'Description:
       '    destroys collection when this class is terminated
       '    Which in turn Destoys Each Item object
       '************************************************************
    Set oItems = Nothing
End Sub

