VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropertyListItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2070
Option Explicit

'==================================================================================================
'cPropertyListItems.cls                 7/5/04
'
'           GENERAL PURPOSE:
'               Creates and maintains a collection of cookie/name pairs for use with iPerPropertyBrowsing.
'
'==================================================================================================

Private mColl As Collection

Private Sub Class_Initialize()
    Set mColl = New Collection
End Sub
Public Property Get NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mColl.[_NewEnum]
End Property
Public Function Remove(ByVal vKeyOrIndex As Variant) As Boolean
    mColl.Remove vKeyOrIndex
End Function

Public Function Item(ByVal vKeyOrIndex As Variant) As cPropertyListItem
Attribute Item.VB_UserMemId = 0
    Set Item = mColl.Item(vKeyOrIndex)
End Function
Public Function Count() As Long
    Count = mColl.Count
End Function
Public Sub Clear()
    Set mColl = New Collection
End Sub
Public Function Add(ByVal DisplayName As String, ByVal Cookie As Long, Optional ByVal Key As Variant, Optional ByVal InsertBefore As Variant) As cPropertyListItem
    Dim oItem As cPropertyListItem
    Set oItem = New cPropertyListItem
    
    oItem.Cookie = Cookie
    oItem.DisplayName = DisplayName
    
    mColl.Add oItem, Key, InsertBefore
    
    Set Add = oItem
End Function

