VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cPopupMenuItems   - Exposes procedures to manage a collection of menu items.  This represents only
'                    one level of a menu; each item in this collection also has its own collection
'                    of subitems.
'
'Copyright free, use and abuse as you see fit.
'==================================================================================================

Implements iEnumerable

Private moOwner As cPopupMenus
Private mtPointer As tMenuPointer

'<Friendly Interface>
Friend Sub fInit(ByVal oOwner As cPopupMenus, ByRef tPointer As tMenuPointer)
    Set moOwner = oOwner
    LSet mtPointer = tPointer
End Sub
'</Friendly Interface>

'<Public Interface>
Public Function Add( _
            ByRef sCaption As String, _
   Optional ByRef sHelpText As String, _
   Optional ByRef sKey As String, _
   Optional ByVal iIconIndex As Long = -1&, _
   Optional ByVal iStyle As ePopupMenuItemStyle, _
   Optional ByVal iShortcutKey As KeyCodeConstants, _
   Optional ByVal iShortcutMask As ShiftConstants, _
   Optional ByVal iItemData As Long, _
   Optional ByRef vKeyOrIndexInsertBefore As Variant) _
                As Long
    iStyle = iStyle And &H3FF&
    Add = moOwner.fPopupMenuItems_Add(mtPointer, sCaption, sHelpText, sKey, iItemData, iIconIndex, iStyle, iShortcutKey, iShortcutMask, vKeyOrIndexInsertBefore)
End Function

Public Function Count() As Long
    Count = moOwner.fPopupMenuItems_Count(mtPointer)
End Function

Public Function Clear() As Long
    Clear = moOwner.fPopupMenuItems_Clear(mtPointer)
End Function

Public Function Exists(ByVal vKeyOrIndex As Variant) As Boolean
    Exists = moOwner.fPopupMenuItems_Exists(mtPointer, vKeyOrIndex)
End Function

Public Property Get Item(ByVal vKeyOrIndex As Variant) As cPopupMenuItem
    Set Item = moOwner.fPopupMenuItems_Item(mtPointer, vKeyOrIndex)
End Property

Public Sub Remove(ByVal vKeyOrIndex As Variant)
    moOwner.fPopupMenuItems_Remove mtPointer, vKeyOrIndex
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Set NewEnum = moOwner.fPopupMenuItems_NewEnum(mtPointer, Me)
End Property

Public Property Get ShowCheckAndIcon() As Boolean
    ShowCheckAndIcon = moOwner.fPopupMenuItems_ShowCheckAndIcon(mtPointer)
End Property

Public Property Let ShowCheckAndIcon(ByVal bNew As Boolean)
    moOwner.fPopupMenuItems_ShowCheckAndIcon(mtPointer) = bNew
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = moOwner.fPopupMenuItems_RightToLeft(mtPointer)
End Property

Public Property Let RightToLeft(ByVal bNew As Boolean)
    moOwner.fPopupMenuItems_RightToLeft(mtPointer) = bNew
End Property

Public Property Get Parent() As cPopupMenuItem
    Set Parent = moOwner.fPopupMenuItems_Parent(mtPointer)
End Property

Public Property Get Root() As cPopupMenu
    Set Root = moOwner.fPopupMenuItems_Root(mtPointer)
End Property
'</Public Interface>


Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fPopupMenuItems_NextItem tEnum, mtPointer, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = -1&
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
End Sub
