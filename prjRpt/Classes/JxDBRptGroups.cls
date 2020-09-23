VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JxDBRptGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"iabGroupLevel"
Attribute VB_Ext_KEY = "Member0" ,"iabGroupLevel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781E00244"
'**********************************************************************
'
'           Module Name: JxDBRptGroups.cls
'           Purpose    : A colletion class for Report Group (JxDBRptGroup)
'           Author     : Joyprakash Saikia
'**********************************************************************
Option Explicit

Private Const MOD_NAME = "JxDBReport.JxDBRptGroups"

Private mCol As Collection



Public Function Add(ReportGroup As JxDBRptGroup, Optional sKey As String) As JxDBRptGroup
        '************************************************************
        'Description:
        '  This Routine Add Report Item to the collection
        '  Input:
        '              sKey - Item's Key (Optional)
        '              RptItem - Object to be added
        '
        ' OutPut:     create a new object and maintains a linked List
        '
        '************************************************************


    
    Dim i As Long
    Dim objNewMember As JxDBRptGroup
    Set objNewMember = ReportGroup

    
    If Len(sKey) Then
        mCol.Add objNewMember, sKey
    Else
        mCol.Add objNewMember
    End If
    
    
    i = mCol.Count
    'set upper & lower group
    If i > 1 Then
        ReportGroup.SetUpperGroup Me.Item(i - 1) 'the Previous group is Upper group
        Me.Item(i - 1).SetLowerGroup ReportGroup 'The Current One is Lower Group
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As JxDBRptGroup
Attribute Item.VB_UserMemId = 0
    If Val(vntIndexKey) > mCol.Count Then Exit Property
    Set Item = mCol.Item(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
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
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Public Sub DestroyGroups()
        '************************************************************
        'Description:
        ' Function to Release the Linked List for Each Group
        '************************************************************
    Dim IndexCnt As Long
    For IndexCnt = 1 To mCol.Count
      mCol.Item(IndexCnt).ReleaseLinkedGroup
    Next
End Sub


Private Sub Class_Terminate()
        '************************************************************
        'Description:
        'destroys collection when this class is terminated
        '************************************************************

        
    DestroyGroups
  
    Set mCol = Nothing
End Sub
Friend Sub ResetFormulaItem()
        '************************************************************
        'Description:
        'To reset  All formulas of all Groups in this Class
        ' This Routine Illustrates Use of
        ' this collection with the
        '       For ..Each ... in ...
        '       ...
        '       Next
        '       syntax
        ' Though it is Slower method I've given here as sample to use
        ' your own collection with above syntax
        '************************************************************

        
        Dim oGroup As JxDBRptGroup
        Dim oItem As JxDBRptItem
        Dim i As Integer
        For Each oGroup In Me
            oGroup.ResetFormulaItem
        Next
End Sub
