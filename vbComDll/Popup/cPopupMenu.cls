VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cPopupMenu   - Exposes procedures to show an api popup menu and access the collection of menu items.
'
'Copyright free, use and abuse as you see fit.
'==================================================================================================

'1.  Friendly Interface  - Initialization
'2.  Public Interface

Private mtPointer As tMenuPointer
Private moOwner As cPopupMenus

'<Friendly Interface>
Friend Sub fInit(ByVal oOwner As cPopupMenus, ByRef tPointer As tMenuPointer)
    Set moOwner = oOwner
    LSet mtPointer = tPointer
End Sub
'</Friendly Interface>

'<Public Interface>
Public Function Items() As cPopupMenuItems
    Set Items = moOwner.fPopupMenu_Items(mtPointer)
End Function

Public Property Get hMenu() As Long
    hMenu = moOwner.fPopupMenu_hMenu(mtPointer)
End Property

Public Property Get Key() As String
    Key = moOwner.fPopupMenu_Key(mtPointer)
End Property

Public Property Get Index() As Long
    Index = moOwner.fPopupMenu_Index(mtPointer)
End Property

Public Property Get Sidebar() As StdPicture
    Set Sidebar = moOwner.fPopupMenu_Sidebar(mtPointer)
End Property

Public Property Set Sidebar(ByVal oNew As StdPicture)
    Set moOwner.fPopupMenu_Sidebar(mtPointer) = oNew
End Property

Public Function ShowAtWindow( _
            ByVal hWnd As Long, _
   Optional ByVal iFlags As ePopupMenuShow, _
   Optional ByVal vExclude As Variant) _
                As cPopupMenuItem
    Set ShowAtWindow = moOwner.fPopupMenu_Show(mtPointer, 0, 0, iFlags, False, hWnd, 0, vExclude)
End Function

Public Function ShowAtCursor( _
            ByVal iFlags As ePopupMenuShow) _
                As cPopupMenuItem
    Set ShowAtCursor = moOwner.fPopupMenu_Show(mtPointer, 0, 0, iFlags, True, 0, 0, Empty)
End Function
            
Public Function Show( _
            ByVal x As Long, _
            ByVal y As Long, _
   Optional ByVal iFlags As ePopupMenuShow, _
   Optional ByVal vExclude As Variant, _
   Optional ByVal hWndClientCoords As Long) _
                As cPopupMenuItem
    Set Show = moOwner.fPopupMenu_Show(mtPointer, x, y, iFlags, False, 0, hWndClientCoords, vExclude)
End Function
'</Public Interface>
