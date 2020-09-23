VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiClassWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiClassWindows.cls                   7/4/04
'
'           GENERAL PURPOSE:
'               Creates, destroys, accesses and enumerates api windows of a class registered with cApiWindowClasses.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'           CLASSES CREATED BY THIS MODULE:
'               cEnumeration
'
'==================================================================================================

'1.  iEnumerable Interface -  For .. Each Enumeration
'2.  Friendly Interface    -  Used internally to initialize the modular variables.
'3.  Collection Interface  -  The only public interface

Implements iEnumerable

Private msClass As String       'stores the class name this object was created with
Private mtPointer As tPointer

'<Private Interface>
'</Private Interface>

'<iEnumerable Interface>
Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    mWindow.ApiClassWindows_NextItem msClass, mtPointer, tEnum, vNextItem, bNoMoreItems
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
            ByRef tPointer As tPointer, _
            ByVal sClass As String)
    LSet mtPointer = tPointer
    msClass = sClass
End Sub
'</Friendly Interface>

'<Public Interface>
'<Collection Interface>
Public Function Add( _
            Optional ByVal iStyle As eWindowStyle, _
            Optional ByVal iExStyle As eWindowStyleEx, _
            Optional ByVal iLeft As Long = 0, _
            Optional ByVal iTop As Long = 0, _
            Optional ByVal iWidth As Long = 0, _
            Optional ByVal iHeight As Long = 0, _
            Optional ByRef sCaption As String = "", _
            Optional ByVal hWndParent As Long = 0, _
            Optional ByVal hMenu As Long = 0, _
            Optional ByVal lParam As Long = 0) _
                As cApiClassWindow
Attribute Add.VB_Description = "Creates a window and returns the hWnd.  If creation fails, return value is 0."

    Set Add = mWindow.ApiClassWindows_Add(msClass, mtPointer, iStyle, iExStyle, iLeft, iTop, iWidth, iHeight, sCaption, hWndParent, hMenu, lParam)
    
End Function

Public Sub Remove( _
                    ByVal hWnd As Long)
Attribute Remove.VB_Description = "Removes a window from the collection."
    
    mWindow.ApiClassWindows_Remove msClass, mtPointer, hWnd
    
End Sub

Public Function Item( _
            ByVal hWnd As Long) _
                As cApiClassWindow
Attribute Item.VB_Description = "Returns an object representing the specified item.  If the item cannot be found, returns Nothing."
Attribute Item.VB_UserMemId = 0
    Set Item = mWindow.ApiClassWindows_Item(msClass, mtPointer, hWnd)
End Function

Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of windows in the collection."
    Count = mWindow.ApiClassWindows_Count(msClass, mtPointer)
End Function

Public Function Exists( _
            ByVal hWnd As Long) _
                As Boolean
Attribute Exists.VB_Description = "Returns a value indicating whether the specified window exists in the collection."
    Exists = mWindow.ApiClassWindows_Exists(msClass, mtPointer, hWnd)
End Function

Public Function Clear() As Long
Attribute Clear.VB_Description = "Destroys all windows in the collection and returns the number of windows that were destroyed."
    Clear = mWindow.ApiClassWindows_Clear(mtPointer, msClass)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Provides For ... Each Enumeration Support"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim loEnum As cEnumeration
    Set loEnum = New cEnumeration
    Set NewEnum = loEnum.GetEnum(Me, mWindow.ApiClassWindows_GetControl(msClass, mtPointer))
End Property
'</Collection Interface>
'</Public Interface>
