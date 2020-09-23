VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiWindow.cls                     7/4/04
'
'           GENERAL PURPOSE:
'               Represents an api window, whether created using the cApiWindows collection or not, and
'               delegates to many common api calls.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'==================================================================================================

'1.  Friendly Interface -  Used internally to initialize the modular variables.
'2.  Win Api Interface  -  Methods which delegate to Api calls you may find useful.

Private msClass  As String      'stores the class name of this window
Private mhWnd    As Long        'store the window's handle

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit(ByVal ihWnd As Long, ByRef sClass As String)
    mhWnd = ihWnd
    msClass = sClass
End Sub
'</Friendly Interface>

'<Public Interface>
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to the window."
Attribute hWnd.VB_UserMemId = 0
    hWnd = mhWnd
End Property

Public Property Get ClassName() As String
Attribute ClassName.VB_Description = "Returns the class name of the window."
    ClassName = msClass
End Property

'<Win Api Interface>
Public Function CallDefWndProc( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
Attribute CallDefWndProc.VB_Description = "Calls the default window procedure."
    CallDefWndProc = DefWindowProc(mhWnd, iMsg, wParam, lParam)
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/Returns whether the window is enabled for keyboard/mouse input."
    Enabled = WindowIsEnabled(mhWnd)
End Property

Public Property Let Enabled( _
            ByVal bVal As Boolean)
    WindowEnable mhWnd, bVal
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Sets/Returns the Window Text.  This is the same as calling the Api."
    Text = WindowText(mhWnd)
End Property

Public Property Let Text( _
            ByVal sVal As String)
    WindowText(mhWnd) = sVal
End Property

Public Function Move( _
            ByVal iLeft As Long, _
            ByVal iTop As Long, _
            ByVal iWidth As Long, _
            ByVal iHeight As Long, _
   Optional ByVal bRepaint As Boolean = True) _
                As Boolean
Attribute Move.VB_Description = "Calls the Api to move the window."
    Move = WindowMove(mhWnd, iLeft, iTop, iWidth, iHeight, bRepaint)
End Function

Public Function SetPos( _
   Optional ByVal iLeft As Long, _
   Optional ByVal iTop As Long, _
   Optional ByVal iWidth As Long, _
   Optional ByVal iHeight As Long, _
   Optional ByVal hWndInsertAfter As eSWPhWnd, _
   Optional ByVal iFlags As eSWPFlags) _
                As Long
Attribute SetPos.VB_Description = "Sets the position of the window using the Api"
    SetPos = WindowSetPos(mhWnd, hWndInsertAfter, iLeft, iTop, iWidth, iHeight, iFlags)
End Function

Public Function GetPos( _
   Optional ByRef iLeft As Long, _
   Optional ByRef iTop As Long, _
   Optional ByRef iWidth As Long, _
   Optional ByRef iHeight As Long) _
                As Boolean
Attribute GetPos.VB_Description = "Calls the Api to set the position of the window."
    GetPos = WindowGetPos(mhWnd, iLeft, iTop, iWidth, iHeight)
End Function

Public Function GetClientDimensions( _
   Optional ByRef iWidth As Long, _
   Optional ByRef iHeight As Long) _
                As Boolean
Attribute GetClientDimensions.VB_Description = "Retrieves the width and height of the client area, returning True if successful."
    GetClientDimensions = mWindow.WindowGetClientDimensions(mhWnd, iWidth, iHeight)
End Function

Public Function GetDc( _
   Optional ByVal bIncludeNonClient As Boolean) _
                As Long
Attribute GetDc.VB_Description = "Gets the handle to the display device context for the window, optionally including the non-client area."
    GetDc = mWindow.WindowGetDC(mhWnd, bIncludeNonClient)
End Function

Public Function Owner() As Long
Attribute Owner.VB_Description = "Returns a handle to the owner of the window."
    Owner = mWindow.WindowGetOwner(mhWnd)
End Function

Public Property Get Parent() As Long
Attribute Parent.VB_Description = "Sets/Returns the window's parent in the message chain."
    Parent = mWindow.WindowParent(mhWnd)
End Property

Public Property Let Parent( _
            ByVal iNew As Long)
    mWindow.WindowParent(mhWnd) = iNew
End Property

Public Property Get UserData() As Long
Attribute UserData.VB_Description = "Sets/Returns the application-defined long value for the window."
    UserData = mWindow.WindowGetLong(mhWnd, GWL_USERDATA)
End Property

Public Property Let UserData( _
            ByVal iNew As Long)
    mWindow.WindowSetLong mhWnd, GWL_USERDATA, iNew, &HFFFFFFFF
End Property

Public Property Get ID() As Long
Attribute ID.VB_Description = "Sets/Returns the application defined identifier for the window."
    ID = mWindow.WindowGetLong(mhWnd, GWL_ID)
End Property

Public Property Let ID( _
            ByVal iNew As Long)
    mWindow.WindowSetLong mhWnd, GWL_ID, iNew, &HFFFFFFFF
End Property

Public Function SetStyle( _
   Optional ByVal iStyleOr As eWindowStyle, _
   Optional ByVal iStyleAndNot As eWindowStyle) _
                As Boolean
Attribute SetStyle.VB_Description = "Sets the style property in the window memory."
    SetStyle = mWindow.WindowSetLong(mhWnd, GWL_STYLE, iStyleOr, iStyleAndNot)
End Function

Public Function GetStyle() As Long
Attribute GetStyle.VB_Description = "Gets the style value from the window memory."
    GetStyle = mWindow.WindowGetLong(mhWnd, GWL_STYLE)
End Function

Public Function SetStyleEx( _
   Optional ByVal iStyleExOr As eWindowStyleEx, _
   Optional ByVal iStyleExAndNot As eWindowStyleEx) _
                As Boolean
Attribute SetStyleEx.VB_Description = "Sets the extended style property in the window memory."
    SetStyleEx = mWindow.WindowSetLong(mhWnd, GWL_EXSTYLE, iStyleExOr, iStyleExAndNot)
End Function

Public Function GetStyleEx() As Long
Attribute GetStyleEx.VB_Description = "Gets the styleex value from the window memory."
    GetStyleEx = mWindow.WindowGetLong(mhWnd, GWL_EXSTYLE)
End Function

Public Property Get Property( _
            ByRef sPropName As String) _
                As Long
Attribute Property.VB_Description = "Sets/Returns a long value from the windows property database."
    Property = mWindow.WindowProp(mhWnd, sPropName)
End Property

Public Property Let Property( _
            ByRef sPropName As String, _
            ByVal iNew As Long)
    mWindow.WindowProp(mhWnd, sPropName) = iNew
End Property

Public Function PropertyRemove( _
            ByRef sPropName As String) _
                As Boolean
Attribute PropertyRemove.VB_Description = "Removes a property from the windows property database."
    PropertyRemove = mWindow.WindowRemoveProp(mhWnd, sPropName)
End Function

Public Function SetFocus() As Long
Attribute SetFocus.VB_Description = "Sets the keyboard focus to the window, returning the handle of the window previously in focus, or 0 if unsuccessful."
    SetFocus = mWindow.WindowSetFocus(mhWnd)
End Function

Public Function PeekMsg( _
            ByRef iMsg As tMsg, _
   Optional ByVal iFilterMin As Long, _
   Optional ByVal iFilterMax As Long, _
   Optional ByVal bRemove As Boolean) _
                As Long
Attribute PeekMsg.VB_Description = "Returns a message from the window's message queue, optionally removing the message."
    PeekMsg = WindowPeekMsg(mhWnd, iMsg, iFilterMin, iFilterMax, bRemove)
End Function

Public Function GetMsg( _
            ByRef iMsg As tMsg, _
   Optional ByVal iFilterMin As Long, _
   Optional ByVal iFilterMax As Long) _
                As Long
Attribute GetMsg.VB_Description = "Gets a message from the window's message queue."
    GetMsg = WindowGetMsg(mhWnd, iMsg, iFilterMin, iFilterMax)
End Function

Public Function SendMsg( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
Attribute SendMsg.VB_Description = "Sends a message to the window's wndproc."
    SendMsg = WindowSendMsg(mhWnd, iMsg, wParam, lParam)
End Function

Public Function PostMsg( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
Attribute PostMsg.VB_Description = "Posts a message to the window's message queue."
    PostMsg = WindowPostMsg(mhWnd, iMsg, wParam, lParam)
End Function

Public Sub ZOrder()
Attribute ZOrder.VB_Description = "Brings the window to the top of the z order."
    WindowZOrder mhWnd
End Sub
'</Win Api Interface>
'</Public Interface>
