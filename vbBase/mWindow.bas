Attribute VB_Name = "mWindow"
Option Explicit

'==================================================================================================
'mWindow.bas                            7/4/04
'
'           GENERAL PURPOSE:
'               Maintains collections of two kinds of windows; those created from classes registered
'               by this component and those created from other, pre-defined classes.  Uses pcWindowClassHub
'               objects to facilite the collections of windows from registered classes, and stores the
'               pre-defined windows in a private structure.
'
'           MODULES CALLED FROM THIS MODULE:
'               mVbBaseGeneral.bas
'
'           CLASSES CREATED BY THIS MODULE:
'               pcWindowClassHub
'
'               cApiClassWindow
'               cApiClassWindows
'               cApiWindow
'               cApiWindowClass
'
'==================================================================================================

'1.  Private Interface            - Utility procedures
'2.  cApiWindowClasses Interface  - procedures to manage collections of registered window classes
'3.  cApiWindowClass Interface    - procedures to set the default messages for each class and get a reference to the windows collection
'4.  cApiClassWindows Interface   - procedures to manage collections of windows created from registered classes
'5.  cApiClassWindow Interface    - procedures to work with a window created from a registered class
'6.  cApiWindows Interface        - procedures to manage collections of windows created from non-registered classes
'7.  General Api Interface        - utility apis that called from cApiWindow and cApiClassWindow

'<Utility Api's>
'public to be used by cApiWindow and cApiClassWindow
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'public to be used by pcWindowClassHub and this module
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
'public to be used by pcSubclass, cApiWindows and gVbBase
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'</Utility Api's>

'<Related to General Api Interface>
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As tRect) As Long
Private Declare Function GetDc Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As tMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As tRect) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As tMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SysAllocStringByteLen Lib "OLEAUT32" (ByVal pszStr As Long, ByVal lLenB As Long) As Long

Private Const GW_OWNER = 4
Private Const HWND_DESKTOP = 0

Private Const cApiWndClses = "cApiWindowClasses"
Private Const cApiWnds = "cApiWindows"
Private Const cApiClsWnds = "cApiClassWindows"
Private Const cApiClsWnd = "cApiClassWindow"
Private Const cApiWnd = "cApiWindow"
Private Const cApiWndCls = "cApiWindowClass"


'these consts used by cApiWindow and cApiClassWindow
Public Const GWL_EXSTYLE = -20&
Public Const GWL_ID = -12&
Public Const GWL_STYLE = -16&
Public Const GWL_USERDATA = -21&
'</Related to General Api Interface>

Private Type tWindowClient
    iPtr                As Long
    iControl            As Long
    iWindowCount        As Long
    iWindows()          As Long
End Type

Private mCollClasses    As Collection 'store the pcWindowClassHub objects with class names as keys
Private miClassesControl As Long

Private miClientCount   As Long
Private mtClients()     As tWindowClient

'<Private Interface>

Private Function pValidatePointer( _
            ByRef tPointer As tPointer, _
   Optional ByVal bForce As Boolean) _
                As Boolean
    If tPointer.iIndex > Undefined Then
        If miClientCount > tPointer.iIndex Then pValidatePointer = mtClients(tPointer.iIndex).iPtr = tPointer.iId
    Else
        tPointer.iIndex = pFindClient(tPointer.iId, bForce)
        If tPointer.iIndex > Undefined Then
            If miClientCount > tPointer.iIndex Then
                pValidatePointer = mtClients(tPointer.iIndex).iPtr = tPointer.iId
            End If
        End If
    End If
End Function

Private Function pValidateItemPointer( _
            ByRef tItemPointer As tItemPointer) _
                As Boolean
    If tItemPointer.iIndex > Undefined Then
        If tItemPointer.iIndex < miClientCount Then
            If tItemPointer.iItemIndex > Undefined Then
                If mtClients(tItemPointer.iIndex).iWindowCount > tItemPointer.iItemIndex Then
                    pValidateItemPointer = tItemPointer.iId = mtClients(tItemPointer.iIndex).iWindows(tItemPointer.iItemIndex)
                End If
            End If
        End If
    End If
End Function

Private Function pFindClient(ByVal iPtr As Long, Optional ByVal bForce As Boolean) As Long
    Dim liFirst As Long
    Dim liTempPtr As Long
    liFirst = Undefined
    For pFindClient = 0& To miClientCount - 1&
        liTempPtr = mtClients(pFindClient).iPtr
        If liTempPtr <> Undefined And liTempPtr <> 0& Then
            If liTempPtr = iPtr Then Exit For
        Else
            If liFirst = Undefined Then liFirst = pFindClient
        End If
    Next
    If pFindClient = miClientCount Then
        If bForce Then
            If liFirst = Undefined Then
                Dim liNewUbound As Long
                Dim liCurrentUbound As Long
                
                liNewUbound = ArrAdjustUbound(miClientCount)
                miClientCount = miClientCount + 1&
                
                On Error GoTo UndefinedArray
                liCurrentUbound = UBound(mtClients)
                            
                If liCurrentUbound < liNewUbound Then
UndefinedArray:
                    ReDim Preserve mtClients(0 To liNewUbound)
                End If
            Else
                pFindClient = liFirst
            End If
            mtClients(pFindClient).iPtr = iPtr
            mtClients(pFindClient).iWindowCount = 0&
        Else
            pFindClient = Undefined
        End If
    End If
End Function

Private Function pFindWindow(ByVal iClientIndex As Long, ByVal hWnd As Long) As Long
    With mtClients(iClientIndex)
        For pFindWindow = 0& To .iWindowCount - 1&
            If .iWindows(pFindWindow) = hWnd Then Exit Function
        Next
        pFindWindow = Undefined
    End With
End Function

Private Property Get ClassObject( _
            ByRef sClass As String) _
                As pcWindowClassHub
    
    If mCollClasses Is Nothing Then Set mCollClasses = New Collection       'init the collection
    
    On Error GoTo handler
    Set ClassObject = mCollClasses(sClass)                                  'try to find it by key
    
    If False Then
handler: gErr vbbItemDetached, cApiClsWnds
    End If
End Property
'</Private Interface>

'<Public Interface>
'<cApiWindowClasses Interface>
Public Function ApiWindowClasses_Register( _
            ByVal sClass As String, _
   Optional ByVal iBackColor As Long = &HFFFFFF, _
   Optional ByVal iStyle As eClassStyle = 0, _
   Optional ByVal hCursor As Long = 0, _
   Optional ByVal hIcon As Long = 0, _
   Optional ByVal hIconSm As Long = 0, _
   Optional ByVal cbClassExtra As Long = 0, _
   Optional ByVal cbWndExtra As Long = 0) _
                As cApiWindowClass
        
    Dim oThis As pcWindowClassHub
        
    On Error GoTo NotAlreadyThere
    Set oThis = ClassObject(sClass)
    gErr vbbKeyAlreadyExists, cApiWndClses

NotAlreadyThere:
    Set oThis = New pcWindowClassHub
    
    If Not oThis.Register(sClass, iBackColor, iStyle, hCursor, hIcon, hIconSm, cbClassExtra, cbWndExtra) _
        Then gErr vbbApiFailure, cApiWndClses
    
    mCollClasses.Add oThis, sClass
        
    Set ApiWindowClasses_Register = New cApiWindowClass
    ApiWindowClasses_Register.fInit sClass
    Incr miClassesControl
End Function

Public Sub ApiWindowClasses_Unregister( _
            ByRef sClass As String)
            
    With ClassObject(sClass)
        If Not .Unregister Then gErr vbbApiFailure, cApiWndClses
        If .Active = False Then mCollClasses.Remove sClass Else Debug.Assert False
    End With
  Incr miClassesControl
End Sub

Public Function ApiWindowClasses_Item(ByRef sClass As String) As cApiWindowClass
    On Error GoTo ItemNotThere
    IsObject mCollClasses(sClass)
    Set ApiWindowClasses_Item = New cApiWindowClass
    ApiWindowClasses_Item.fInit sClass
    Exit Function
ItemNotThere:
    gErr vbbKeyNotFound, cApiWndClses
End Function

Public Function ApiWindowClasses_GetControl() As Long
    If mCollClasses Is Nothing Then Set mCollClasses = New Collection
    ApiWindowClasses_GetControl = miClassesControl
End Function

Public Function ApiWindowClasses_NextItem(ByRef tEnum As tEnum, ByRef vNextItem As Variant, ByRef bNoMore As Boolean)
    If tEnum.iControl <> miClassesControl Then gErr vbbCollChangedDuringEnum, cApiWndClses
    tEnum.iIndex = tEnum.iIndex + 1&
    If mCollClasses.Count > tEnum.iIndex Then
        Dim loClass As cApiWindowClass
        Dim loHub As pcWindowClassHub
        
        Set loClass = New cApiWindowClass
        Set loHub = mCollClasses(tEnum.iIndex + 1&)
        loClass.fInit loHub.ClassName
        
        Set vNextItem = loClass
    Else
        bNoMore = True
    End If
End Function

Public Function ApiWindowClasses_Count() As Long
    If Not mCollClasses Is Nothing Then ApiWindowClasses_Count = mCollClasses.Count 'return the class count
End Function

Public Function ApiWindowClasses_Exists( _
            ByRef sClass As String) _
                As Boolean
    
    On Error Resume Next
    IsObject mCollClasses(sClass)
    ApiWindowClasses_Exists = (Err.Number = 0&) 'return whether the key was found
    On Error GoTo 0
    
End Function

'</cApiWindowClasses Interface>

'<cApiWindowClass Interface>
Public Function ApiWindowClass_AddDefMsg( _
            ByVal iMsg As eMsg, _
            ByRef sClass As String) _
                As Boolean
    
    ApiWindowClass_AddDefMsg = ClassObject(sClass).AddDefMsg(iMsg)
    
    'probably adding a message already there or window is not created
    Debug.Assert ApiWindowClass_AddDefMsg
End Function

Public Function ApiWindowClass_DelDefMsg( _
            ByVal iMsg As eMsg, _
            ByRef sClass As String) _
                As Boolean
                
    ApiWindowClass_DelDefMsg = ClassObject(sClass).DelDefMsg(iMsg)
    
    'probably deleting a message that isn't there or window class is not alive
    Debug.Assert ApiWindowClass_DelDefMsg

End Function

Public Function ApiWindowClass_DefMsgExists( _
            ByRef sClass As String, _
            ByVal iMsg As eMsg) _
                As Boolean
        
    ApiWindowClass_DefMsgExists = ClassObject(sClass).DefMsgExists(iMsg)
    
End Function

Public Function ApiWindowClass_DefMsgCount( _
            ByRef sClass As String) _
                As Long
    
    ApiWindowClass_DefMsgCount = ClassObject(sClass).DefMsgCount
    
End Function

Public Function ApiWindowClass_GetDefMessages( _
            ByRef iOutArray() As Long, _
            ByRef sClass As String) _
                 As Long
                 
    ApiWindowClass_GetDefMessages = ClassObject(sClass).GetDefMessages(iOutArray)
    
End Function

Public Function ApiWindowClass_WindowCount( _
            ByRef sClass As String) _
                As Long
    ApiWindowClass_WindowCount = ClassObject(sClass).AbsCount
End Function

Public Function ApiWindowClass_OwnedWindows( _
            ByRef sClass As String, _
            ByVal oObject As iWindow) _
                As cApiClassWindows
    
    Dim ltPtr As tPointer
    
    If oObject Is Nothing Then gErr vbbInvalidProcedureCall, "cApiWindowClass.OwnedWindows"
    
    Set ApiWindowClass_OwnedWindows = New cApiClassWindows
    ltPtr.iId = ObjPtr(oObject)
    ltPtr.iIndex = -1&
    ApiWindowClass_OwnedWindows.fInit ltPtr, sClass
    
End Function
'</cApiWindowClass Interface>

'<cApiClassWindows Interface>
Public Function ApiClassWindows_Clear( _
            ByRef tPointer As tPointer, _
            ByRef sClass As String) _
                As Long
    ApiClassWindows_Clear = ClassObject(sClass).Clear(tPointer)

End Function

Public Function ApiClassWindows_Exists( _
            ByRef sClass As String, _
            ByRef tPointer As tPointer, _
            ByVal hWnd As Long) _
                As Boolean
    ApiClassWindows_Exists = ClassObject(sClass).Exists(tPointer, hWnd)
End Function

Public Function ApiClassWindows_Count( _
            ByRef sClass As String, _
            ByRef tPointer As tPointer) _
                As Long
    ApiClassWindows_Count = ClassObject(sClass).Count(tPointer)
End Function

Public Function ApiClassWindows_Add( _
                    ByRef sClass As String, _
                    ByRef tPointer As tPointer, _
                    ByVal iStyle As eWindowStyle, _
                    ByVal iExStyle As eWindowStyleEx, _
                    ByVal iLeft As Long, _
                    ByVal iTop As Long, _
                    ByVal iWidth As Long, _
                    ByVal iHeight As Long, _
                    ByRef sCaption As String, _
                    ByVal hWndParent As Long, _
                    ByVal hMenu As Long, _
                    ByVal lParam As Long) _
                        As cApiClassWindow
    
    Set ApiClassWindows_Add = ClassObject(sClass).Create(tPointer, iExStyle, iStyle, iLeft, iTop, iWidth, iHeight, sCaption, hWndParent, hMenu, lParam)
    
End Function

Public Sub ApiClassWindows_Remove( _
                    ByRef sClass As String, _
                    ByRef tPointer As tPointer, _
                    ByVal hWnd As Long)
    ClassObject(sClass).Destroy tPointer, hWnd
End Sub

Public Function ApiClassWindows_Item( _
            ByRef sClass As String, _
            ByRef tPointer As tPointer, _
            ByVal hWnd As Long) _
                As cApiClassWindow
    Set ApiClassWindows_Item = ClassObject(sClass).Item(tPointer, hWnd)
End Function

Public Function ApiClassWindows_GetControl( _
            ByRef sClass As String, _
            ByRef tPointer As tPointer) _
                 As Long
   ApiClassWindows_GetControl = ClassObject(sClass).GetControl(tPointer)
End Function

Public Sub ApiClassWindows_NextItem( _
            ByRef sClass As String, _
            ByRef tPointer As tPointer, _
            ByRef tEnum As tEnum, _
            ByRef vNextItem As Variant, _
            ByRef bNoMore As Boolean)
        ClassObject(sClass).Enum_NextItem tPointer, tEnum, vNextItem, bNoMore
End Sub

'</cApiClassWindows Interface>

'<cApiClassWindow Interface>
Public Function ApiClassWindow_AddMsg( _
            ByRef sClass As String, _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean
    
    
    ApiClassWindow_AddMsg = ClassObject(sClass).AddMsg(tPointer, iMsg)
    
    'probably adding a message already there or window is not created
    Debug.Assert ApiClassWindow_AddMsg
    
End Function

Public Function ApiClassWindow_DelMsg( _
            ByRef sClass As String, _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean

    ApiClassWindow_DelMsg = ClassObject(sClass).DelMsg(tPointer, iMsg)
    
    'probably deleting a message that isn't there or window is not alive
    Debug.Assert ApiClassWindow_DelMsg

End Function

Public Function ApiClassWindow_MsgExists( _
            ByRef sClass As String, _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean
    ApiClassWindow_MsgExists = ClassObject(sClass).MsgExists(tPointer, iMsg)
    
End Function

Public Function ApiClassWindow_MsgCount( _
            ByRef sClass As String, _
            ByRef tPointer As tItemPointer) _
                As Long
    
    ApiClassWindow_MsgCount = ClassObject(sClass).MsgCount(tPointer)
    
End Function

Public Function ApiClassWindow_GetMessages( _
            ByRef sClass As String, _
            ByRef iOutArray() As Long, _
            ByRef tPointer As tItemPointer) _
                 As Long

    ApiClassWindow_GetMessages = ClassObject(sClass).GetMessages(iOutArray, tPointer)

End Function

Public Property Let ApiClassWindow_DefMessages( _
            ByRef sClass As String, _
            ByRef tPointer As tItemPointer, _
            ByVal bVal As Boolean)
    ClassObject(sClass).DefMessages(tPointer) = bVal

End Property

Public Property Get ApiClassWindow_DefMessages( _
            ByRef sClass As String, _
            ByRef tPointer As tItemPointer) _
                As Boolean

    ApiClassWindow_DefMessages = ClassObject(sClass).DefMessages(tPointer)

End Property
'</cApiClassWindow Interface>

'<cApiWindows Interface>
Public Function ApiWindows_Count( _
            ByRef tPointer As tPointer) _
                As Long
    
    If pValidatePointer(tPointer) Then
        Dim i As Long
        With mtClients(tPointer.iIndex)
            For i = 0& To .iWindowCount - 1&
                If .iWindows(i) <> 0& Then ApiWindows_Count = ApiWindows_Count + 1&
            Next
        End With
    End If

End Function

Public Function ApiWindows_Clear( _
            ByRef tPointer As tPointer) _
                As Long

    If pValidatePointer(tPointer) Then
        Dim i As Long
        With mtClients(tPointer.iIndex)
            For i = 0& To .iWindowCount - 1&
                If .iWindows(i) <> 0& Then
                    If DestroyWindow(.iWindows(i)) <> 0& Then
                        .iWindows(i) = 0&
                        ApiWindows_Clear = ApiWindows_Clear + 1&
                    End If
                End If
            Next
            
            For i = i - 1& To 0& Step -1&
                If .iWindows(i) <> 0& Then Exit For
            Next
            .iWindowCount = i + 1&
            If .iWindowCount = 0& Then .iPtr = Undefined
            Incr .iControl
        End With
    End If

End Function

Public Function ApiWindows_Exists( _
            ByRef tPointer As tPointer, _
            ByVal hWnd As Long) _
                As Boolean
    If pValidatePointer(tPointer) Then ApiWindows_Exists = pFindWindow(tPointer.iIndex, hWnd) <> Undefined

End Function

Public Function ApiWindows_Add( _
                    ByRef tPointer As tPointer, _
                    ByRef sClass As String, _
                    ByVal iClass As eWindowClass, _
                    ByVal iStyle As eWindowStyle, _
                    ByVal iExStyle As eWindowStyleEx, _
                    ByVal iLeft As Long, _
                    ByVal iTop As Long, _
                    ByVal iWidth As Long, _
                    ByVal iHeight As Long, _
                    ByRef sCaption As String, _
                    ByVal hWndParent As Long, _
                    ByVal hMenu As Long, _
                    ByVal lParam As Long) _
                        As cApiWindow
    Dim hWnd As Long
    
    If pValidatePointer(tPointer, True) Then
        With mtClients(tPointer.iIndex)
            Dim i As Long
            For i = 0& To .iWindowCount - 1&
                If .iWindows(i) = 0& Then Exit For
            Next
            
            If i = .iWindowCount Then
                .iWindowCount = i + 1&
                ArrRedim .iWindows, .iWindowCount, True
            End If
            
            Select Case iClass
                Case eWindowClass.PREDEFINED_BUTTON:          sClass = "BUTTON"  'Predefined window classes
                Case eWindowClass.PREDEFINED_COMBOBOX:        sClass = "COMBOBOX"
                Case eWindowClass.PREDEFINED_EDIT:            sClass = "EDIT"
                Case eWindowClass.PREDEFINED_LISTBOX:         sClass = "LISTBOX"
                Case eWindowClass.PREDEFINED_MDICLIENT:       sClass = "MDICLIENT"
                Case eWindowClass.PREDEFINED_RICHEDIT:        sClass = "RichEdit"
                Case eWindowClass.PREDEFINED_RICHEDIT_CLASS:  sClass = "RICHEDIT_CLASS"
                Case eWindowClass.PREDEFINED_SCROLLBAR:       sClass = "SCROLLBAR"
                Case eWindowClass.PREDEFINED_STATIC:          sClass = "STATIC"
            End Select
            
            
            hWnd = CreateWindowEx(iExStyle, sClass, sCaption, iStyle, iLeft, iTop, iWidth, iHeight, hWndParent, hMenu, App.hInstance, lParam)
            
            If hWnd = 0& Then
                If i = .iWindowCount - 1& Then .iWindowCount = i
                gErr vbbApiFailure, cApiWnds
            End If
            
            .iWindows(i) = hWnd

            Set ApiWindows_Add = New cApiWindow
            ApiWindows_Add.fInit hWnd, sClass
            Incr .iControl
        End With
    Else
        'should never happen!
        Debug.Assert False
        gErr vbbKeyNotFound, cApiWnds
    End If
End Function

Public Sub ApiWindows_Remove( _
                    ByRef tPointer As tPointer, _
                    ByVal hWnd As Long)
    If pValidatePointer(tPointer) Then
        Dim liIndex As Long
        With mtClients(tPointer.iIndex)
            liIndex = ArrFindInt(.iWindows, .iWindowCount, hWnd)
            If liIndex = Undefined Then GoTo KeyNotFound
            If DestroyWindow(hWnd) = 0& Then gErr vbbApiFailure, cApiWnds
            .iWindows(liIndex) = 0&
            Incr .iControl
        End With
    Else
KeyNotFound:
        gErr vbbKeyNotFound, cApiWnds
    End If
End Sub

Public Function ApiWindows_Item( _
                    ByRef tPointer As tPointer, _
                    ByVal hWnd As Long) _
                        As cApiWindow
    If pValidatePointer(tPointer) Then
        With mtClients(tPointer.iIndex)
            If ArrFindInt(.iWindows, .iWindowCount, hWnd) = Undefined Then GoTo KeyNotFound
            
            Set ApiWindows_Item = New cApiWindow
            ApiWindows_Item.fInit hWnd, WindowClassName(hWnd)
            
        End With
    Else
KeyNotFound:
        gErr vbbKeyNotFound, cApiWnds
    End If
End Function
        
Public Function ApiWindows_GetControl( _
                    ByRef tPointer As tPointer) _
                        As Long
    If pValidatePointer(tPointer) Then ApiWindows_GetControl = mtClients(tPointer.iIndex).iControl
End Function

Public Sub ApiWindows_NextItem( _
            ByRef tPointer As tPointer, _
            ByRef tEnum As tEnum, _
            ByRef vNextItem As Variant, _
            ByRef bNoMore As Boolean)
    
    If pValidatePointer(tPointer) Then
        With mtClients(tPointer.iIndex)
            If .iControl <> tEnum.iControl Then gErr vbbCollChangedDuringEnum, cApiWnds
            
            tEnum.iIndex = tEnum.iIndex + 1&
            
            Do Until tEnum.iIndex = .iWindowCount Or .iWindows(tEnum.iIndex) <> 0&
                tEnum.iIndex = tEnum.iIndex + 1&
            Loop
            
            If tEnum.iIndex < .iWindowCount Then
                Dim loItem As cApiWindow
                Set loItem = New cApiWindow
                loItem.fInit .iWindows(tEnum.iIndex), WindowClassName(.iWindows(tEnum.iIndex))
                Set vNextItem = loItem
            Else
                bNoMore = True
            End If
        End With
    Else
        bNoMore = True
    End If
End Sub
'</cApiWindows Interface>

'<General Api Interface>
Public Function WindowClassName( _
            ByVal hWnd As Long) _
                As String
    Dim lsClassName As String * 256  'string buffer
    WindowClassName = Left$(lsClassName, GetClassName(hWnd, lsClassName, 256&)) 'get the class name
End Function

Public Function WindowEnable( _
            ByVal hWnd As Boolean, _
            ByVal bVal As Boolean) _
                As Boolean
    WindowEnable = EnableWindow(hWnd, Abs(bVal)) <> 0& 'enable the window
End Function

Public Function WindowGetClientDimensions( _
            ByVal hWnd As Long, _
            ByRef iWidth As Long, _
            ByRef iHeight As Long) _
                As Boolean
    Dim lR As tRect
    If GetClientRect(hWnd, lR) <> 0& Then       'if we can get the client rect
        WindowGetClientDimensions = True        'indicate success
        iWidth = lR.Right                       'store the width and height
        iHeight = lR.Bottom                     'left and top are always 0
    End If
End Function
            
Public Function WindowGetDC( _
            ByVal hWnd As Long, _
            ByVal bIncludeNonClient As Boolean) _
                As Long
    If Not bIncludeNonClient Then _
        WindowGetDC = GetDc(hWnd) _
    Else _
        WindowGetDC = GetWindowDC(hWnd)      'get the requested dc
End Function

Public Function WindowGetLong( _
            ByVal hWnd As Long, _
            ByVal iLong As Long) _
                 As Long
    WindowGetLong = GetWindowLong(hWnd, iLong)  'get the window long
End Function

Public Function WindowGetOwner( _
            ByVal hWnd As Long) _
                As Long
    WindowGetOwner = GetWindow(hWnd, GW_OWNER)   'get the owner
End Function

Public Function WindowGetPos( _
            ByVal hWnd As Long, _
            ByRef iLeft As Long, _
            ByRef iTop As Long, _
            ByRef iWidth As Long, _
            ByRef iHeight As Long) _
                As Boolean
    Dim lR As tRect
    
    If (GetWindowRect(hWnd, lR) <> 0&) Then 'if we can get the window rect
        WindowGetPos = True                 'indicate success
        MapWindowPoints HWND_DESKTOP, GetParent(hWnd), lR, 2
        With lR
            iLeft = .Left                   'store the position values
            iTop = .Top
            iWidth = .Right - iLeft
            iHeight = .Bottom - iTop
        End With
    End If
End Function

Public Function WindowIsEnabled( _
            ByVal hWnd As Long) _
                As Boolean
    WindowIsEnabled = IsWindowEnabled(hWnd) <> 0& 'check enabled status
End Function

Public Function WindowMove( _
            ByVal hWnd As Long, _
            ByVal iLeft As Long, _
            ByVal iTop As Long, _
            ByVal iWidth As Long, _
            ByVal iHeight As Long, _
            ByVal bRepaint As Boolean) _
                As Boolean
    WindowMove = MoveWindow(hWnd, iLeft, iTop, iWidth, iHeight, Abs(bRepaint)) <> 0& 'move the window
End Function

Public Property Get WindowParent( _
            ByVal hWnd As Long) _
                As Long
    WindowParent = GetParent(hWnd) 'get the parent
End Property

Public Property Let WindowParent( _
            ByVal hWnd As Long, _
            ByVal iNewParent As Long)
    SetParent hWnd, iNewParent     'set the parent
End Property
            
Public Property Get WindowProp( _
            ByVal hWnd As Long, _
            ByRef sPropName As String) _
                As Long
    WindowProp = GetProp(hWnd, sPropName)   'get the property
End Property

Public Property Let WindowProp( _
            ByVal hWnd As Long, _
            ByRef sPropName As String, _
            ByVal iNewVal As Long)
    SetProp hWnd, sPropName, iNewVal        'set the property
End Property
            
Public Function WindowRemoveProp( _
            ByVal hWnd As Long, _
            ByVal sPropName As String) _
                As Boolean
    WindowRemoveProp = RemoveProp(hWnd, sPropName) <> 0& 'remove the property
End Function
            
Public Function WindowSetFocus( _
            ByVal hWnd As Long) _
                As Long
    WindowSetFocus = SetFocus(hWnd)     'set kb focus
    'returns the hWnd that previously had the focus, or 0 if failed
End Function

Public Function WindowSetLong( _
            ByVal hWnd As Long, _
            ByVal iLong As Long, _
            ByVal iOr As Long, _
            ByVal iAndNot As Long) _
                As Boolean
    
    WindowSetLong = (SetWindowLong(hWnd, iLong, _
                        ((GetWindowLong(hWnd, iLong) _
                          And Not iAndNot) _
                          Or iOr)) _
                     <> 0&) 'set the new long
End Function

Public Function WindowSetPos( _
            ByVal hWnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) _
                As Long
    WindowSetPos = SetWindowPos(hWnd, hWndInsertAfter, x, y, cx, cy, wFlags)    'set position
End Function

Public Property Get WindowText( _
            ByVal hWnd As Long) _
                As String
    Dim liLen As Long                           'length
    liLen = GetWindowTextLength(hWnd)           'get the length
    Debug.Assert StrPtr(WindowText) = 0         'Must not point to a string, this string won't get deallocated
    If liLen > 0& Then                          'if there is some text
        CopyMemory WindowText, _
                   SysAllocStringByteLen(0&, liLen), _
                   4&                           'allocate the string
                                                'if the string was allocated then get the window text
        If LenB(WindowText) > 0& Then _
            GetWindowText hWnd, WindowText, liLen + 1& 'add one to account for trailing null char
    End If
End Property

Public Property Let WindowText( _
            ByVal hWnd As Long, _
            ByVal sVal As String)
    SetWindowText hWnd, sVal                    'set the window text
End Property

Public Function WindowPeekMsg( _
            ByVal hWnd As Long, _
            ByRef iMsg As tMsg, _
            ByVal iFilterMin As Long, _
            ByVal iFilterMax As Long, _
            ByVal bRemove As Boolean) _
                As Long
    'delegate to api
    WindowPeekMsg = PeekMessage(iMsg, hWnd, iFilterMin, iFilterMax, Abs(bRemove))
End Function

Public Function WindowGetMsg( _
            ByVal hWnd As Long, _
            ByRef iMsg As tMsg, _
            ByVal iFilterMin As Long, _
            ByVal iFilterMax As Long) _
                As Long
    'delegate to api
    WindowGetMsg = GetMessage(iMsg, hWnd, iFilterMin, iFilterMax)
End Function

Public Function WindowSendMsg( _
            ByVal hWnd As Long, _
            ByVal iMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    'delegate to api
    WindowSendMsg = SendMessage(hWnd, iMsg, wParam, lParam)
End Function

Public Function WindowPostMsg( _
            ByVal hWnd As Long, _
            ByVal iMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    'delegate to api
    WindowPostMsg = PostMessage(hWnd, iMsg, wParam, lParam)
End Function

Public Sub WindowZOrder( _
            ByVal hWnd As Long)
    'delegate to api
    BringWindowToTop hWnd
End Sub
'</General Api Interface>
'</Public Interface>
