VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JxDBRptItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"iabRptItem"
Attribute VB_Ext_KEY = "Member0" ,"iabRptItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781E900D2"
'**********************************************************************
'
'           Module Name: JxDBRptItems.cls
'
'           Purpose    : A Collection class for Report Item (JxDBRptItem)
'                        It Acts same as the VB collections it also has the hidden
'                         _newEnum property.
'           Author     : Joyprakash saikia
'
'**********************************************************************
Option Explicit
'This Constant is Used for Debugging Purpose
'This is not used on this submission

Private Const MOD_NAME = "JxDBReport.JxDBRptItems"
'local variable to hold collection
Private mCol As Collection
Public Property Get Item(vntIndexKey As Variant) As JxDBRptItem
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
       '************************************************************
       'Description:
       'used when removing an element from the collection
       '************************************************************
       
        
    On Error Resume Next
    mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
       '************************************************************
       'Description:
       'this property allows you to enumerate
       'this collection with the For...Each syntax
       '************************************************************
       'Description:
        
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
       '************************************************************
       'Description:
       'creates the collection when this class is created
       '************************************************************

       
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
       '************************************************************
       'Description:
       '    destroys collection when this class is terminated
       '    Which in turn Destoys Each Item object
       '************************************************************
        Set mCol = Nothing
End Sub


Public Function AddItem(RptItem As JxDBRptItem, Optional sKey As String) As JxDBRptItem
        '************************************************************
        'Description:
        '  This Routine Add Report Item to the collection
        '  Input:
        '              sKey - Item's Key (Optional)
        '              RptItem - Object to be added
        '
        '          OutPut: Added Report Item
        '
        '************************************************************
On Error GoTo AddItemErr

    Dim objNewMember As JxDBRptItem
    Set objNewMember = RptItem 'Put a pointer to the Parameter


    If Len(Trim$(sKey)) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    Set AddItem = objNewMember
    Set objNewMember = Nothing 'Release the Current Reference
AddItemEXIT:

Exit Function
AddItemErr:
          If Err.Number = 457 Then 'item already exist,
            Err.Clear
          End If
          GoTo AddItemEXIT
End Function
Public Sub ClearItems()
    Dim inDex As Integer
    If mCol Is Nothing Then Exit Sub
    For inDex = mCol.Count - 1 To 1 Step -1
        'As you Remove a object from Collection , Index Decreases
        ' So the Loop is decremented from Max value
        mCol.Remove inDex
    Next inDex
End Sub
