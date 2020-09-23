VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiWindows.cls                                7/4/04
'
'           GENERAL PURPOSE:
'               Exposes creation/destruction/manipulation of api windows in a collection interface.
'               Windows are created and destroyed by adding and removing from this collection, and
'               individual windows are accessed as items of the collection.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'           CLASSES CREATED BY THIS MODULE:
'               cEnumerator
'
'==================================================================================================

'1.  iEnumerable Interface  - for .. each enumeration
'2.  Friendly Interface     - Initializes the modular variable
'3.  Collection Interface   - Delegates to methods in mWindow
    
Implements iEnumerable
    
Private mtPointer As tPointer

'<Private Interface>
'</Private Interface>

'<iEnumerable Interface>
Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    mWindow.ApiWindows_NextItem mtPointer, tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = -1&
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
End Sub
'</iEnumerable Interface>

'<Friendly Interface>
Friend Sub fInit( _
        ByRef tPointer As tPointer)
    LSet mtPointer = tPointer
End Sub
'</Friendly Interface>

'<Public Interface>
'<Collection Interface>
Public Function Add( _
            Optional ByRef sClass As String, _
            Optional ByVal iClass As eWindowClass, _
            Optional ByVal iStyle As eWindowStyle, _
            Optional ByVal iExStyle As eWindowStyleEx, _
            Optional ByVal iLeft As Long, _
            Optional ByVal iTop As Long, _
            Optional ByVal iWidth As Long, _
            Optional ByVal iHeight As Long, _
            Optional ByRef sCaption As String, _
            Optional ByVal hWndParent As Long, _
            Optional ByVal hMenu As Long, _
            Optional ByVal lParam As Long) _
                As cApiWindow
Attribute Add.VB_Description = "Creates a new window of the specified class or predefined class.  If you need access to the message stream, you will need to use a subclass."
    Set Add = mWindow.ApiWindows_Add(mtPointer, sClass, iClass, iStyle, iExStyle, iLeft, iTop, iWidth, iHeight, sCaption, hWndParent, hMenu, lParam)
End Function

Public Sub Remove( _
            ByVal hWnd As Long)
Attribute Remove.VB_Description = "Destroys a window created in this collection."
    mWindow.ApiWindows_Remove mtPointer, hWnd
End Sub

Public Function Item( _
            ByVal hWnd As Long) _
                As cApiWindow
Attribute Item.VB_Description = "Returns an object representing the specified window, or Nothing if the window was not found."
Attribute Item.VB_UserMemId = 0
    Set Item = mWindow.ApiWindows_Item(mtPointer, hWnd)

End Function

Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of windows in the collection."
    Count = mWindow.ApiWindows_Count(mtPointer)
End Function

Public Function Clear() As Long
Attribute Clear.VB_Description = "Destroys all the windows in the collection and returns the number that were destroyed."
    Clear = mWindow.ApiWindows_Clear(mtPointer)
End Function

Public Function Exists( _
            hWnd As Long) _
                As Boolean
Attribute Exists.VB_Description = "Returns a value indicating whether a specific window  exists in the collection."
    Exists = mWindow.ApiWindows_Exists(mtPointer, hWnd)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Allows For ... Each Enumeration."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim loEnum As cEnumeration
    Set loEnum = New cEnumeration
    Set NewEnum = loEnum.GetEnum(Me, mWindow.ApiWindows_GetControl(mtPointer))
End Property
'</Collection Interface>
'</Public Interface>
