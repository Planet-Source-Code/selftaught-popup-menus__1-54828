VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cPopupMenus   - Exposes procedures to manipulate a collection of api Popup menus, and relays procedure
'                calls from the child objects (cPopupMenu, cPopupMenuItems, and cPopupMenuItem) to mPopup
'                with the complete menu data that is stored in the modular variable mtMenus.
'
'Copyright free, use and abuse as you see fit.
'==================================================================================================

'1.  Private Interface   -  events received by this class and callbacks from the MsgHook component on the iSubclass & iTimer interfaces
'2.  Friendly Interface  -  Procedures called by child objects and mPopup.

Public Enum ePopupMenuItemStyle
    mnuChecked = 1
    mnuRadioChecked = 2
    mnuDisabled = 4
    mnuSeparator = 8
    mnuDefault = 16
    mnuInvisible = 32
    mnuInfrequent = 64
    mnuRedisplayOnClick = 128
    mnuNewVerticalLine = 256
End Enum

Public Enum ePopupMenuShow
    mnuCenterAlign = &H4&
    mnuVCenterAlign = &H10&
    mnuRightAlign = &H8&
    mnuRightButton = &H2&
    mnuBottomAlign = &H20&
    mnuPreserveVertAlign = &H40&
    mnuNoAnimation = &H4000&
    mnuAnimateLTR = &H400&
    mnuAnimateRTL = &H800&
    mnuAnimateTTB = &H1000&
    mnuAnimateBTT = &H2000&
    mnuRecurse = &H1&
End Enum

Public Enum ePopupMenuDrawStyle
    mnuGradientHighlight = 1&
    mnuButtonHighlightStyle = 2&
    mnuOfficeXPStyle = 4&
    mnuDrawSeparatorsAsHeaders = 8&
    mnuImageProcessBitmap = 16&
    mnuShowInfrequent = 32&
End Enum

Public Enum ePopupMenuErrors
    ccElemNotPartOfCollection = 35605
    ccCollectionChangedDuringEnum = 35606
End Enum

Implements iTimer                       'callback interface for timers
Implements iSubclass                    'callback interface for subclasses
Implements iEnumerable                  'we are a collection

Private mbSubclass As Boolean           'are we subclassing?

Private mtMenus As tMenus               'data store for unlimited popup menus
Private WithEvents moFont As StdFont    'we've got to know when the font changes so we can recalculate the menu item dimensions.
Attribute moFont.VB_VarHelpID = -1

Event Click(ByVal Item As cPopupMenuItem)
Event RightClick(ByVal Item As cPopupMenuItem)
Event ItemHighlight(ByVal Item As cPopupMenuItem)
Event InitPopupMenu(ByVal Items As cPopupMenuItems, ByVal ChevronAdded As cPopupMenuItem)
Event UnInitPopupMenu(ByVal Items As cPopupMenuItems)

'<Private Interface>
Private Sub Class_Initialize()
    PopupMenus_Initialize mtMenus
    Set moFont = mtMenus.tMenuDraw.oFont
End Sub

Private Sub Class_Terminate()
    PopupMenus_Terminate mtMenus
    Timers(Me).Clear
    Subclasses(Me).Clear
End Sub

Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    mPopup.Enum_NextMenu mtMenus, tEnum, vNextItem, bNoMoreItems, Me
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = -1&
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
End Sub

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hWnd As Long, iMsg As eMsg, wParam As Long, lParam As Long)
    PopupMenus_WndProc mtMenus, bBefore, bHandled, lReturn, iMsg, wParam, lParam, Me
End Sub

Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
    Dim bContinue As Boolean
    PopupMenu_CallBack lTimerID, lElapsedMS, bContinue
    If Not bContinue Then Timers(Me).Remove lTimerID
End Sub

Private Sub moFont_FontChanged(ByVal PropertyName As String)
    Set Me.Font = moFont
End Sub
'</Private Interface>

'<Friendly Interface>
Friend Sub fClick(ByVal oItem As cPopupMenuItem)
    RaiseEvent Click(oItem)
End Sub

Friend Sub fRightClick(ByVal oItem As cPopupMenuItem)
    RaiseEvent RightClick(oItem)
End Sub

Friend Sub fItemHighlight(ByVal oItem As cPopupMenuItem)
    RaiseEvent ItemHighlight(oItem)
End Sub

Friend Sub fMenuInit(ByVal oItems As cPopupMenuItems, ByVal bInit As Boolean, ByVal oItem As cPopupMenuItem)
    If bInit Then
        RaiseEvent InitPopupMenu(oItems, oItem)
    Else
        RaiseEvent UnInitPopupMenu(oItems)
    End If
End Sub

Friend Property Get fSubclass() _
                As Boolean
    fSubclass = mbSubclass
End Property

Friend Property Let fSubclass( _
            ByVal bNew As Boolean)
    
    If bNew Xor mbSubclass Then
        On Error GoTo handler:
        If mbSubclass Then Subclasses(Me).Remove mtMenus.hWndOwner
        If bNew And CBool(mtMenus.hWndOwner) Then
            With Subclasses(Me).Add(mtMenus.hWndOwner)
                .AddMsg WM_MENUSELECT, MSG_BEFORE
                .AddMsg WM_MEASUREITEM, MSG_BEFORE
                .AddMsg WM_DRAWITEM, MSG_BEFORE
                .AddMsg WM_MENUCHAR, MSG_BEFORE
                .AddMsg WM_INITMENUPOPUP, MSG_BEFORE
                .AddMsg WM_UNINITMENUPOPUP, MSG_BEFORE
                .AddMsg WM_MENURBUTTONUP, MSG_BEFORE
                .AddMsg WM_DESTROY, MSG_BEFORE
            End With
            mbSubclass = True
        Else
handler:
            mbSubclass = False
        End If
    End If
End Property

Friend Sub fCallback(ByVal iType As Long, ByVal iFreq As Long)
    On Error GoTo handler
    Timers(Me).Item(iType).Start iFreq, True
    Exit Sub
handler:
    On Error Resume Next
    Timers(Me).Add(iType, iFreq).Start
End Sub

'<cPopupMenu Interface>
Friend Function fPopupMenu_Items( _
            ByRef tPointer As tMenuPointer) _
                As cPopupMenuItems
    Set fPopupMenu_Items = mPopup.PopupMenu_Items(mtMenus, tPointer, Me)
End Function

Friend Function fPopupMenu_hMenu( _
            ByRef tPointer As tMenuPointer) _
                As Long
    fPopupMenu_hMenu = mPopup.PopupMenu_hMenu(mtMenus, tPointer)
End Function

Friend Function fPopupMenu_Key( _
            ByRef tPointer As tMenuPointer) _
                As String
    fPopupMenu_Key = mPopup.PopupMenu_Key(mtMenus, tPointer)
End Function

Friend Function fPopupMenu_Index( _
            ByRef tPointer As tMenuPointer) _
                As Long
    fPopupMenu_Index = mPopup.PopupMenu_Index(mtMenus, tPointer)
End Function

Friend Property Get fPopupMenu_Sidebar( _
            ByRef tPointer As tMenuPointer) _
                As StdPicture
    Set fPopupMenu_Sidebar = mPopup.PopupMenu_Sidebar(mtMenus, tPointer)
End Property

Friend Property Set fPopupMenu_Sidebar( _
            ByRef tPointer As tMenuPointer, _
            ByVal oNew As StdPicture)
    Set mPopup.PopupMenu_Sidebar(mtMenus, tPointer) = oNew
End Property

Friend Function fPopupMenu_Show( _
            ByRef tPointer As tMenuPointer, _
            ByVal xPixel As Long, _
            ByVal yPixel As Long, _
            ByVal iFlags As ePopupMenuShow, _
            ByVal bShowAtCursor As Boolean, _
            ByVal hWndShowAt As Long, _
            ByVal hWndClientCoords As Long, _
            ByRef Exclude As Variant) _
                As cPopupMenuItem
    Set fPopupMenu_Show = mPopup.PopupMenu_Show(mtMenus, tPointer, xPixel, yPixel, iFlags, bShowAtCursor, hWndShowAt, hWndClientCoords, Exclude, Me)
End Function
'</cPopupMenu Interface>

'<cPopupMenuItem Interface>
Friend Property Get fPopupMenuItem_IconIndex( _
            ByRef tPointer As tMenuItemPointer) _
                As Long
    fPopupMenuItem_IconIndex = mPopup.PopupMenuItem_IconIndex(mtMenus, tPointer)
End Property
Friend Property Let fPopupMenuItem_IconIndex( _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Long)
    mPopup.PopupMenuItem_IconIndex(mtMenus, tPointer) = iNew
End Property

Friend Property Get fPopupMenuItem_ShortCutShiftMask( _
            ByRef tPointer As tMenuItemPointer) _
                As Integer
    fPopupMenuItem_ShortCutShiftMask = mPopup.PopupMenuItem_ShortCutShiftMask(mtMenus, tPointer)
End Property
Friend Property Let fPopupMenuItem_ShortCutShiftMask( _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Integer)
    mPopup.PopupMenuItem_ShortCutShiftMask(mtMenus, tPointer) = iNew
End Property

Friend Property Get fPopupMenuItem_ShortCutShiftKey( _
            ByRef tPointer As tMenuItemPointer) _
                As Integer
    fPopupMenuItem_ShortCutShiftKey = mPopup.PopupMenuItem_ShortCutShiftKey(mtMenus, tPointer)
End Property
Friend Property Let fPopupMenuItem_ShortCutShiftKey( _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Integer)
    mPopup.PopupMenuItem_ShortCutShiftKey(mtMenus, tPointer) = iNew
End Property

Friend Property Get fPopupMenuItem_Style( _
            ByRef tPointer As tMenuItemPointer) _
                As ePopupMenuItemStyle
    fPopupMenuItem_Style = mPopup.PopupMenuItem_Style(mtMenus, tPointer)
End Property

Friend Property Let fPopupMenuItem_Style( _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As ePopupMenuItemStyle)
    mPopup.PopupMenuItem_Style(mtMenus, tPointer) = iNew
End Property

Friend Property Get fPopupMenuItem_Caption( _
            ByRef tPointer As tMenuItemPointer) _
                As String
    fPopupMenuItem_Caption = mPopup.PopupMenuItem_Caption(mtMenus, tPointer)
End Property
Friend Property Let fPopupMenuItem_Caption( _
            ByRef tPointer As tMenuItemPointer, _
            ByRef sNew As String)
    mPopup.PopupMenuItem_Caption(mtMenus, tPointer) = sNew
End Property

Friend Property Get fPopupMenuItem_Index( _
            ByRef tPointer As tMenuItemPointer) _
                As Long
    fPopupMenuItem_Index = mPopup.PopupMenuItem_Index(mtMenus, tPointer)
End Property

Friend Property Get fPopupMenuItem_Key( _
            ByRef tPointer As tMenuItemPointer) _
                As String
    fPopupMenuItem_Key = mPopup.PopupMenuItem_Key(mtMenus, tPointer)
End Property

Friend Property Get fPopupMenuItem_ShortCutDisplay( _
            ByRef tPointer As tMenuItemPointer) _
                As String
    fPopupMenuItem_ShortCutDisplay = mPopup.PopupMenuItem_ShortCutDisplay(mtMenus, tPointer)
End Property
Friend Property Let fPopupMenuItem_ShortCutDisplay( _
            ByRef tPointer As tMenuItemPointer, _
            ByRef sNew As String)
    mPopup.PopupMenuItem_ShortCutDisplay(mtMenus, tPointer) = sNew
End Property

Friend Property Get fPopupMenuItem_SubItems( _
            ByRef tPointer As tMenuItemPointer) _
                As cPopupMenuItems
    Set fPopupMenuItem_SubItems = mPopup.PopupMenuItem_SubItems(mtMenus, tPointer, Me)
End Property
'</cPopupMenuItem Interface>

'<cPopupMenuItems Interface>
Friend Function fPopupMenuItems_Add( _
            ByRef tPointer As tMenuPointer, _
            ByRef sCaption As String, _
            ByRef sHelpText As String, _
            ByRef sKey As String, _
            ByVal iItemData As Long, _
            ByVal iIconIndex As Long, _
            ByVal iStyle As ePopupMenuItemStyle, _
            ByVal iShortcutKey As KeyCodeConstants, _
            ByVal iShortcutMask As ShiftConstants, _
            ByRef vKeyOrIndexInsertBefore As Variant) _
                As Long

    fPopupMenuItems_Add = mPopup.PopupMenuItems_Add(mtMenus, tPointer, sCaption, sHelpText, sKey, iItemData, iIconIndex, iStyle, iShortcutKey, iShortcutMask, vKeyOrIndexInsertBefore)
End Function

Friend Function fPopupMenuItems_Count( _
            ByRef tPointer As tMenuPointer) _
                As Long
    fPopupMenuItems_Count = mPopup.PopupMenuItems_Count(mtMenus, tPointer)
End Function

Friend Function fPopupMenuItems_Clear( _
            ByRef tPointer As tMenuPointer) _
                As Long
    fPopupMenuItems_Clear = mPopup.PopupMenuItems_Clear(mtMenus, tPointer)
End Function

Friend Function fPopupMenuItems_Exists( _
            ByRef tPointer As tMenuPointer, _
            ByRef vKeyOrIndex As Variant) _
                As Boolean
    fPopupMenuItems_Exists = mPopup.PopupMenuItems_Exists(mtMenus, tPointer, vKeyOrIndex)
End Function

Friend Function fPopupMenuItems_Item( _
            ByRef tPointer As tMenuPointer, _
            ByRef vKeyOrIndex As Variant) _
                As cPopupMenuItem
    Set fPopupMenuItems_Item = mPopup.PopupMenuItems_Item(mtMenus, tPointer, vKeyOrIndex, Me)
End Function

Friend Sub fPopupMenuItems_Remove( _
            ByRef tPointer As tMenuPointer, _
            ByRef vKeyOrIndex As Variant)
    mPopup.PopupMenuItems_Remove mtMenus, tPointer, vKeyOrIndex
End Sub

Friend Property Get fPopupMenuItems_ShowCheckAndIcon( _
            ByRef tPointer As tMenuPointer) _
                As Boolean
    fPopupMenuItems_ShowCheckAndIcon = mPopup.PopupMenuItems_ShowCheckAndIcon(mtMenus, tPointer)
End Property

Friend Property Let fPopupMenuItems_ShowCheckAndIcon( _
            ByRef tPointer As tMenuPointer, _
            ByVal bNew As Boolean)
    mPopup.PopupMenuItems_ShowCheckAndIcon(mtMenus, tPointer) = bNew
End Property

Friend Property Get fPopupMenuItems_RightToLeft( _
            ByRef tPointer As tMenuPointer) _
                As Boolean
    fPopupMenuItems_RightToLeft = mPopup.PopupMenuItems_RightToLeft(mtMenus, tPointer)
End Property

Friend Property Let fPopupMenuItems_RightToLeft( _
            ByRef tPointer As tMenuPointer, _
            ByVal bNew As Boolean)
    mPopup.PopupMenuItems_RightToLeft(mtMenus, tPointer) = bNew
End Property

Friend Property Get fPopupMenuItems_Parent( _
            ByRef tPointer As tMenuPointer) _
                As cPopupMenuItem
    Set fPopupMenuItems_Parent = mPopup.PopupMenuItems_Parent(mtMenus, tPointer, Me)
End Property

Friend Property Get fPopupMenuItems_Root( _
            ByRef tPointer As tMenuPointer) _
                As cPopupMenu
    Set fPopupMenuItems_Root = mPopup.PopupMenuItems_Root(mtMenus, tPointer, Me)
End Property

Friend Property Get fPopupMenuItems_NewEnum( _
            ByRef tPointer As tMenuPointer, _
            ByVal oOwner As cPopupMenuItems) _
                As IUnknown
    Set fPopupMenuItems_NewEnum = mPopup.PopupMenuItems_NewEnum(mtMenus, tPointer, oOwner)
End Property

Friend Sub fPopupMenuItems_NextItem( _
            ByRef tEnum As tEnum, _
            ByRef tPointer As tMenuPointer, _
            ByRef vNextItem As Variant, _
            ByRef bNoMoreItems As Boolean)
    mPopup.Enum_NextItem mtMenus, tPointer, tEnum, vNextItem, bNoMoreItems, Me
End Sub


'</cPopupMenuItems Interface>
'</Friendly Interface>

'<Public Interface>
Public Function Add( _
            ByRef sKey As String, _
   Optional ByVal vKeyOrIndexInsertBefore As Variant) _
                As cPopupMenu
    Set Add = mPopup.PopupMenus_Add(mtMenus, sKey, vKeyOrIndexInsertBefore, Me)
    Incr mtMenus.iControl
End Function

Public Function Count() As Long
    Count = mPopup.PopupMenus_Count(mtMenus)
End Function

Public Function Clear() As Long
    Clear = mPopup.PopupMenus_Clear(mtMenus)
    Incr mtMenus.iControl
End Function

Public Function Exists(ByVal vKeyOrIndex As Variant) As Boolean
    Exists = mPopup.PopupMenus_Exists(mtMenus, vKeyOrIndex)
End Function

Public Property Get Item(ByVal vKeyOrIndex As Variant) As cPopupMenu
    Set Item = mPopup.PopupMenus_Item(mtMenus, vKeyOrIndex, Me)
End Property

Public Property Get Font() As StdFont
    Set Font = mtMenus.tMenuDraw.oFont
End Property

Public Property Set Font(ByVal oNew As StdFont)
    Set mPopup.PopupMenus_Font(mtMenus) = oNew
    Set moFont = oNew
End Property

Public Sub Remove(ByVal vKeyOrIndex As Variant)
    mPopup.PopupMenus_Remove mtMenus, vKeyOrIndex
    Incr mtMenus.iControl
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Dim loEnum As cEnumeration
    Set loEnum = New cEnumeration
    Set NewEnum = loEnum.GetEnum(Me, mtMenus.iControl)
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = mtMenus.hWndOwner
End Property
Public Property Let hWndOwner(ByVal iNew As Long)
    If iNew <> mtMenus.hWndOwner Then
        If mbSubclass Then fSubclass = False
        mtMenus.hWndOwner = iNew
    End If
End Property

Public Property Set ImageList(ByVal oNew As Object)
    mPopup.PopupMenus_SetImageList mtMenus, oNew
End Property

Public Property Let ImageList(ByVal vNew As Variant)
    mPopup.PopupMenus_SetImageList mtMenus, vNew
End Property

Public Property Get ImageList() As Variant
    If mtMenus.tMenuDraw.oVB6ImageList Is Nothing Then ImageList = mtMenus.tMenuDraw.hIml Else Set ImageList = mtMenus.tMenuDraw.oVB6ImageList
End Property

Public Sub SetDrawStyle(Optional ByVal iStateOr As ePopupMenuDrawStyle, Optional ByVal iStateAndNot As ePopupMenuDrawStyle)
    DrawStyle = ((mtMenus.tMenuDraw.iFlags Or iStateOr) And Not iStateAndNot)
End Sub

Public Property Get DrawStyle() As ePopupMenuDrawStyle
    DrawStyle = mtMenus.tMenuDraw.iFlags
End Property

Public Property Let DrawStyle(ByVal iNew As ePopupMenuDrawStyle)
    mPopup.PopupMenus_SetDrawStyle mtMenus, iNew
End Property

Public Property Get GradientHighlight() As Boolean
    GradientHighlight = CBool(mtMenus.tMenuDraw.iFlags And mnuGradientHighlight)
End Property

Public Property Let GradientHighlight(ByVal bNew As Boolean)
    If bNew Then
        SetDrawStyle mnuGradientHighlight
    Else
        SetDrawStyle , mnuGradientHighlight
    End If
End Property

Public Property Get ButtonHighlight() As Boolean
    ButtonHighlight = CBool(mtMenus.tMenuDraw.iFlags And mnuButtonHighlightStyle)
End Property

Public Property Let ButtonHighlight(ByVal bNew As Boolean)
    If bNew Then
        SetDrawStyle mnuButtonHighlightStyle
    Else
        SetDrawStyle , mnuButtonHighlightStyle
    End If
End Property

Public Property Get DrawSeparatorsAsHeaders() As Boolean
    DrawSeparatorsAsHeaders = CBool(mtMenus.tMenuDraw.iFlags And mnuDrawSeparatorsAsHeaders)
End Property

Public Property Let DrawSeparatorsAsHeaders(ByVal bNew As Boolean)
    If bNew Then
        SetDrawStyle mnuDrawSeparatorsAsHeaders
    Else
        SetDrawStyle , mnuDrawSeparatorsAsHeaders
    End If
End Property

Public Property Get ImageProcessBitmap() As Boolean
    ImageProcessBitmap = CBool(mtMenus.tMenuDraw.iFlags And mnuImageProcessBitmap)
End Property

Public Property Let ImageProcessBitmap(ByVal bNew As Boolean)
    If bNew Then
        SetDrawStyle mnuImageProcessBitmap
    Else
        SetDrawStyle , mnuImageProcessBitmap
    End If
End Property

Public Property Get ShowInfrequent() As Boolean
    ShowInfrequent = CBool(mtMenus.tMenuDraw.iFlags And mnuShowInfrequent)
End Property

Public Property Let ShowInfrequent(ByVal bNew As Boolean)
    If bNew Then
        SetDrawStyle mnuShowInfrequent
    Else
        SetDrawStyle , mnuShowInfrequent
    End If
End Property

Public Property Get OfficeXPStyle() As Boolean
    OfficeXPStyle = CBool(mtMenus.tMenuDraw.iFlags And mnuOfficeXPStyle)
End Property

Public Property Let OfficeXPStyle(ByVal bNew As Boolean)
    If bNew Then
        SetDrawStyle mnuOfficeXPStyle
    Else
        SetDrawStyle , mnuOfficeXPStyle
    End If
End Property

Public Property Get ActiveForeColor() As OLE_COLOR
    ActiveForeColor = mtMenus.tMenuDraw.iActiveForeColor
    If ActiveForeColor = -1& Then ActiveForeColor = vbHighlightText
End Property

Public Property Let ActiveForeColor(ByVal iNew As OLE_COLOR)
    mtMenus.tMenuDraw.iActiveForeColor = iNew
End Property

Public Property Get InActiveForeColor() As OLE_COLOR
    InActiveForeColor = mtMenus.tMenuDraw.iInActiveForeColor
    If InActiveForeColor = -1& Then InActiveForeColor = vbMenuText
End Property

Public Property Let InActiveForeColor(ByVal iNew As OLE_COLOR)
    mtMenus.tMenuDraw.iInActiveForeColor = iNew
End Property

Public Property Get ActiveBackColor() As OLE_COLOR
    ActiveBackColor = mtMenus.tMenuDraw.iActiveBackColor
    If ActiveBackColor = -1& Then ActiveBackColor = vbHighlight
End Property

Public Property Let ActiveBackColor(ByVal iNew As OLE_COLOR)
    mtMenus.tMenuDraw.iActiveBackColor = iNew
End Property

Public Property Get InActiveBackColor() As OLE_COLOR
    InActiveBackColor = mtMenus.tMenuDraw.iInActiveBackColor
    If InActiveBackColor = -1& Then InActiveBackColor = IIf(CBool(mtMenus.tMenuDraw.iFlags And mnuOfficeXPStyle), vbWindowBackground, vbMenuBar)
End Property

Public Property Let InActiveBackColor(ByVal iNew As OLE_COLOR)
    mtMenus.tMenuDraw.iInActiveBackColor = iNew
End Property

Public Property Get BackgroundPicture() As StdPicture
    Set BackgroundPicture = mtMenus.tMenuDraw.oBitmap
End Property

Public Property Set BackgroundPicture(ByVal oNew As StdPicture)
    mPopup.PopupMenus_SetPicture mtMenus, oNew
End Property

Public Function AcceleratorPress(ByVal iKey As KeyCodeConstants, ByVal iShift As ShiftConstants) As cPopupMenuItem
    Set AcceleratorPress = mPopup.PopupMenus_AcceleratorPress(mtMenus, iKey, iShift, Me)
End Function

Public Property Get AutoShowInfreqDelay() As Long
    AutoShowInfreqDelay = mtMenus.tMenuDraw.iInfreqShowDelay
End Property

Public Property Let AutoShowInfreqDelay(ByVal iNew As Long)
    mtMenus.tMenuDraw.iInfreqShowDelay = iNew
End Property
'<Public Interface>
