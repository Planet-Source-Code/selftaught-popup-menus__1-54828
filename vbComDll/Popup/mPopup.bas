Attribute VB_Name = "mPopup"
Option Explicit

'==================================================================================================
'mPopup   - Handles all the dirty work for creating/maintaining popup menus, with the exception of
'           the custom drawing.  This module is called only from cPopupMenus.
'
'Copyright free, use and abuse as you see fit.
'==================================================================================================

'1.  Private Interface      - utility functions for the public procedures
'2.  Public Interface       - procedures called by cPopupMenus for to create/maintain menus


Private Const MIIM_STATE         As Long = &H1&
Private Const MIIM_ID            As Long = &H2&
Private Const MIIM_SUBMENU       As Long = &H4&
'Private Const MIIM_CHECKMARKS    As Long = &H8&
'Private Const MIIM_TYPE          As Long = &H10&
Private Const MIIM_DATA          As Long = &H20&
'
'Private Const MF_APPEND          As Long = &H100&
'Private Const MF_BITMAP          As Long = &H4&
Private Const MF_BYCOMMAND       As Long = &H0&
Private Const MF_BYPOSITION      As Long = &H400&
'Private Const MF_CALLBACKS       As Long = &H8000000
'Private Const MF_CHANGE          As Long = &H80&
Private Const MF_CHECKED         As Long = &H8&
'Private Const MF_CONV            As Long = &H40000000
'Private Const MF_DELETE          As Long = &H200&
'Private Const MF_DISABLED        As Long = &H2&
Private Const MF_ENABLED         As Long = &H0&
'Private Const MF_END             As Long = &H80&
'Private Const MF_ERRORS          As Long = &H10000000
Private Const MF_GRAYED          As Long = &H1&
'Private Const MF_HELP            As Long = &H4000&
Private Const MF_HILITE          As Long = &H80&
'Private Const MF_HSZ_INFO        As Long = &H1000000
'Private Const MF_INSERT          As Long = &H0&
'Private Const MF_LINKS           As Long = &H20000000
'Private Const MF_MASK            As Long = &HFF000000
Private Const MF_MENUBARBREAK    As Long = &H20&
'Private Const MF_MENUBREAK       As Long = &H40&
'Private Const MF_MOUSESELECT     As Long = &H8000&
Private Const MF_OWNERDRAW       As Long = &H100&
Private Const MF_POPUP           As Long = &H10&
'Private Const MF_POSTMSGS        As Long = &H4000000
'Private Const MF_REMOVE          As Long = &H1000&
'Private Const MF_SENDMSGS        As Long = &H2000000
Private Const MF_SEPARATOR       As Long = &H800&
Private Const MF_STRING          As Long = &H0&
Private Const MF_SYSMENU         As Long = &H2000&
Private Const MF_UNCHECKED       As Long = &H0&
'Private Const MF_UNHILITE        As Long = &H0&
'Private Const MF_USECHECKBITMAPS As Long = &H200&
'Private Const MF_DEFAULT         As Long = &H1000&
Private Const MF_RIGHTORDER    As Long = &H2000&
'Private Const MFT_RADIOCHECK     As Long = &H200&

'Private Const TPM_HORIZONTAL = &H0&          '/* Horz alignment matters more */
'Private Const TPM_RIGHTBUTTON = &H2&
'Private Const TPM_CENTERALIGN = &H4&
'Private Const TPM_VCENTERALIGN = &H10&
'Private Const TPM_RIGHTALIGN = &H8&
'Private Const TPM_BOTTOMALIGN = &H20&
'Private Const TPM_VERTICAL = &H40&           '/* Vert alignment matters more */
'Private Const TPM_NONOTIFY = &H80&           '/* Don't send any notification msgs */
Private Const TPM_RETURNCMD = &H100&
'Private Const TPM_HORPOSANIMATION = &H400&
'Private Const TPM_HORNEGANIMATION = &H800&
'Private Const TPM_VERPOSANIMATION = &H1000&
'Private Const TPM_VERNEGANIMATION = &H2000&
'Private Const TPM_NOANIMATION = &H4000&

Private Const ODT_MENU = 1&

'Private Const MNC_IGNORE As Long = 0&
'Private Const MNC_CLOSE As Long = 1&
Private Const MNC_EXECUTE As Long = 2&
Private Const MNC_SELECT As Long = 3&

Private Const KEYEVENTF_EXTENDEDKEY = &H1&
Private Const KEYEVENTF_KEYUP = &H2&
Private Const VK_CONTROL = &H11&

Private Const MOUSEEVENTF_ABSOLUTE = &H8000& '  absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2& '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4& '  left button up
'Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
'Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_MOVE = &H1& '  mouse move
'Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
'Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up

Private Const SPI_GETNONCLIENTMETRICS = 41&

Private Const LF_FACESIZE = 32&

Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Private Const FW_REGULAR = 400

Private Type MENUITEMINFO
    cbSize                      As Long
    fMask                       As Long
    fType                       As Long
    fState                      As Long
    wID                         As Long
    hSubMenu                    As Long
    hbmpChecked                 As Long
    hbmpUnchecked               As Long
    dwItemData                  As Long
    dwTypeData                  As Long
    cch                         As Long
End Type

Private Type TPMPARAMS
    cbSize As Long
    rcExclude As tRect
End Type

Private Type NMLOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 4) As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As NMLOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As NMLOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As NMLOGFONT
    lfStatusFont As NMLOGFONT
    lfMessageFont As NMLOGFONT
End Type

'User32 Declares:
    'Menu
Private Declare Function AppendMenuByLong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long
    'Other
Private Declare Function ClientToScreenAny Lib "user32" Alias "ClientToScreen" (ByVal hWnd As Long, lpPoint As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tPoint) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As tRect) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Any, ByVal tPoint As Currency) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'gdi32
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'comctl32
Private Declare Function ImageList_GetImageRect Lib "comctl32.DLL" (ByVal hIml As Long, ByVal i As Long, prcImage As tRect) As Long
'kernel32
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const Undefined As Long = -1

Public Const mnuChevron = 1024

Public Type tMenuItemPointer
    iId As Long
    iIndex As Long
    iItemIndex As Long
End Type

Public Type tMenuPointer
    iId As Long
    iIndex As Long
End Type

Private Const tMenuItemLen As Long = 44
Public Type tMenuItem
    iStyle              As Long
    
    iApiID              As Long
    iIconIndex          As Long
    iId                 As Long
    iItemData           As Long
    iWidth              As Integer
    iHeight             As Integer
    
    iShortcutMask       As Integer
    iShortcutShiftKey   As Integer
    iAccelerator        As Integer
    bMeasured           As Boolean
    
    
    sCaption            As String
    sHelpText           As String
    sKey                As String
    sShortCutDisplay    As String
    
    tChild              As tMenuPointer
End Type

Private Const tMenuLen As Long = 30
Public Type tMenu
    hMenu               As Long
    iId                 As Long
    iItemCount          As Long
    iControl            As Long
    sKey                As String
    tParent             As tMenuItemPointer
    bShowCheckAndIcon   As Boolean
    bRightToLeft        As Boolean
    oSideBar            As pcMemDC
    oSideBarPic         As StdPicture
    tMenuItems()        As tMenuItem
End Type

Public Type tMenuDraw
    hIml                As Long
    iIconSize           As Long
    iInfreqShowDelay    As Long
    iItemHeight         As Long
    iActiveForeColor    As OLE_COLOR
    iInActiveForeColor  As OLE_COLOR
    iInActiveBackColor  As OLE_COLOR
    iActiveBackColor    As OLE_COLOR
    iFlags              As Long
    oMemDC              As pcMemDC
    oBrush              As pcDottedBrush
    oFont               As StdFont
    oFontSymbol         As StdFont
    oBitmap             As pcMemDC
    oBitmapLight        As pcMemDC
    oBitmapSuperLight   As pcMemDC
    oVB6ImageList       As ImageList
    oPic                As StdPicture
End Type

Public Type tRootMenuLookup
    iMenuCount          As Long
    iMenuIndices()      As Long
End Type

Public Type tMenus
    hWndOwner           As Long
    iMenuCount          As Long
    iControl            As Long
    tMenus()            As tMenu
    tRootMenuLookup     As tRootMenuLookup
    tMenuDraw           As tMenuDraw
End Type

Public Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type

Public Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As tRect
   ItemData As Long
End Type

Private miIDs()                     As Long 'we've got to keep track of ID's so that when we
Private miIDCount                   As Long 'create a sub menu we know that it's a unique ID.

Private miRedisplayHierarchy()      As Long 'if the item clicked is marked for redisplay or if it
Private miRedisplayHierarchyCount   As Long 'is a chevron, then these members will tell us how to
Private miRedisplayHierarchyIndex   As Long 'get from the root menu to the item that was clicked
Private miRedisplayHierarchyItem    As Long 'the item clicked to highlight again

Private mhWndChevronMenu            As Long 'When the mouse hovers over a chevron, a callback is set
Private miChevronIndex              As Long 'and these members are used in the callback procedure to
Private miInfreqShowDelay           As Long 'if and when to click the chevron automatically.
Private miHoverTime                 As Long 'time to hover over a chevron before clicking automatically
Private mbChevronSelected           As Boolean


'<Private Interface>
Private Sub pResizeArray( _
            ByRef iArray() As Long, _
            ByVal iItemCount As Long)
    
    Dim liNewUbound As Long
    
    liNewUbound = pArrAdjustUbound(iItemCount - 1&)
    If liNewUbound > pArrUbound(iArray) Then _
        ReDim Preserve iArray(0 To liNewUbound)

End Sub

Private Sub pResizeMenuItems( _
            ByRef tArray() As tMenuItem, _
            ByVal iItemCount As Long)
    
    Dim liNewUbound As Long
    
    liNewUbound = pArrAdjustUbound(iItemCount - 1&)
    If liNewUbound > pArrUboundT(tArray) Then _
        ReDim Preserve tArray(0 To liNewUbound)

End Sub

Private Sub pResizeMenus( _
            ByRef tArray() As tMenu, _
            ByVal iMenuCount As Long)
    
    Dim liNewUbound As Long
    
    liNewUbound = pArrAdjustUbound(iMenuCount - 1&)
    If liNewUbound > pArrUboundT2(tArray) Then _
        ReDim Preserve tArray(0 To liNewUbound)

End Sub

Private Function pArrUbound( _
            ByRef iArray() As Long) _
                As Long
    On Error Resume Next
    pArrUbound = UBound(iArray)
    If Err.Number <> 0& Then pArrUbound = Undefined
End Function

Private Function pArrUboundT( _
            ByRef tArray() As tMenuItem) _
                As Long
    On Error Resume Next
    pArrUboundT = UBound(tArray)
    If Err.Number <> 0& Then pArrUboundT = Undefined
End Function

Private Function pArrUboundT2( _
            ByRef tArray() As tMenu) _
                As Long
    On Error Resume Next
    pArrUboundT2 = UBound(tArray)
    If Err.Number <> 0& Then pArrUboundT2 = Undefined
End Function

Private Function pArrAdjustUbound( _
            ByVal iBound As Long) _
                As Long
    'Adjusts a Ubound to the next increment of the blocksize
    Const ArrBlockSize As Long = 8&
    
    'if ibound < 0 then ibound = 0
    If CBool(iBound And &H80000000) Then iBound = 0&
    
    Dim liMod As Long
    
    liMod = iBound Mod ArrBlockSize
    
    If Not (liMod = 0) Then
        'If the bound is not an even multiple, then round it up
        pArrAdjustUbound = iBound + ArrBlockSize - liMod
    Else
        'If it is an even multiple, then keep it the same,
        'unless it's zero, then make it equal to ArrBlockSize
        If iBound Then _
            pArrAdjustUbound = iBound _
        Else _
            pArrAdjustUbound = ArrBlockSize
    End If
End Function

Private Sub pSetItemData( _
            ByVal hMenu As Long, _
            ByVal iId As Long, _
            ByVal iItemData As Long)
    Dim tMII As MENUITEMINFO
    tMII.cbSize = Len(tMII)
    tMII.fMask = MIIM_DATA
    GetMenuItemInfo hMenu, iId, False, tMII
    tMII.dwItemData = iItemData
    SetMenuItemInfo hMenu, iId, False, tMII
End Sub

Private Function pParseMenuChar( _
        ByVal bInfrequentShown As Boolean, _
        ByRef tMenu As tMenu, _
        ByVal iChar As Integer) _
            As Long

Dim liAccel() As Long
Dim liFirstChar() As Long
Dim liAccelCount As Long
Dim liFirstCharCount As Long
Dim liPos As Long
Dim liHilite As Long
Dim liReturn As Long
Dim tMI As MENUITEMINFO
Dim lbFirst As Boolean

Dim lR As Long

    Dim lhWndRedisplay As Long

    ReDim liAccel(0 To tMenu.iItemCount - 1&)
    ReDim liFirstChar(0 To tMenu.iItemCount - 1&)

    liHilite = Undefined

    tMI.cbSize = LenB(tMI)
    tMI.fMask = MIIM_STATE
    'Debug.Assert miRedisplayHierarchyIndex > 0
    If miRedisplayHierarchyIndex > Undefined& Then
        lhWndRedisplay = miRedisplayHierarchy(miRedisplayHierarchyIndex)
        miRedisplayHierarchyIndex = miRedisplayHierarchyIndex - 1
        tMI.fMask = MIIM_SUBMENU
    ElseIf miRedisplayHierarchyItem > Undefined Then
        pParseMenuChar = miRedisplayHierarchyItem Or (MNC_SELECT * &H10000)
        miRedisplayHierarchyItem = Undefined
        Exit Function
    End If

    For liPos = 0& To tMenu.iItemCount - 1&
        If GetMenuItemInfo(tMenu.hMenu, liPos, True, tMI) Then
            If (lhWndRedisplay = 0&) Then
                If CBool(tMI.fState And MF_HILITE) Then liHilite = liPos
    
                If Not CBool(tMenu.tMenuItems(liPos).iStyle And mnuDisabled) _
                   And Not CBool(tMenu.tMenuItems(liPos).iStyle And mnuInvisible) _
                   And Not CBool(tMenu.tMenuItems(liPos).iStyle And mnuSeparator) _
                   And (bInfrequentShown Or Not CBool(tMenu.tMenuItems(liPos).iStyle And mnuInfrequent)) Then         'CBool(tMI.fState And MF_ENABLED) And Not CBool(tMI.fState And MF_SEPARATOR) Then
    
                    If (tMenu.tMenuItems(liPos).iAccelerator = iChar) Then
                        liAccel(liAccelCount) = liPos
                        liAccelCount = liAccelCount + 1&
                    End If
                    
                    If Len(tMenu.tMenuItems(liPos).sCaption) Then
                        If Asc(LCase$(tMenu.tMenuItems(liPos).sCaption)) = iChar Then
                            liFirstChar(liFirstCharCount) = liPos
                            liFirstCharCount = liFirstCharCount + 1&
                        End If
                    End If
                End If
            Else
                If tMI.hSubMenu = lhWndRedisplay Or tMI.hSubMenu <> 0 Then
                    pParseMenuChar = liPos Or (MNC_EXECUTE * &H10000)
                    Exit Function
                End If
            End If
        Else
            Debug.Assert False
        End If

    Next
    
    liReturn = Undefined
    
    If liAccelCount > 0 Then
        If liHilite > Undefined Then
            lbFirst = True
            For liPos = liAccelCount - 1& To 0& Step -1&
                If liAccel(liPos) > liHilite Then
                    liReturn = liAccel(liPos)
                Else
                    If liReturn = Undefined Then liReturn = liAccel(0)
                    Exit For
                End If
            Next
        Else
            liReturn = liAccel(0)
        End If
        
        If liAccelCount = 1& Then
            liReturn = liReturn Or (MNC_EXECUTE * &H10000)
        Else
            liReturn = liReturn Or (MNC_SELECT * &H10000)
        End If
        
    ElseIf liFirstCharCount > 0 Then
        If liHilite > Undefined Then
            lbFirst = True
            For liPos = liFirstCharCount - 1& To 0& Step -1&
                If liFirstChar(liPos) > liHilite Then
                    liReturn = liFirstChar(liPos)
                Else
                    If liReturn = Undefined Then liReturn = liFirstChar(0)
                    Exit For
                End If
            Next
        Else
            liReturn = liFirstChar(0)
        End If
        
        liReturn = liReturn Or (MNC_SELECT * &H10000)
        
    End If
    
    If liReturn <> Undefined Then pParseMenuChar = liReturn
    
End Function

Private Sub pClickMouse(ByRef tp As tPoint)

Dim xl As Double
Dim yl As Double
Dim xMax As Long
Dim yMax As Long
   
   ' mouse_event ABSOLUTE coords run from 0 to 65535:
   xMax = Screen.Width \ Screen.TwipsPerPixelX
   yMax = Screen.Height \ Screen.TwipsPerPixelY
   xl = (tp.x * 65535# / CDbl(xMax))
   yl = (tp.y * 65535# / CDbl(yMax))
   ' Move the mouse:
   mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_ABSOLUTE, xl, yl, 0, 0
   mouse_event MOUSEEVENTF_LEFTUP Or MOUSEEVENTF_ABSOLUTE, xl, yl, 0, 0
   
   ' Move the mouse:
   mouse_event MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE, xl + 1, yl + 1, 0, 0

End Sub


Private Sub pSetMenuCaption( _
        ByRef tMenus As tMenus, _
        ByVal iMenuIndex As Long, _
        ByVal iItemIndex As Long, _
        ByRef sCaption As String, _
        ByVal bRecalc As Boolean)

Dim iPos As Long
Dim liWidth As Long

    iPos = pParseCaption(sCaption, "&")
    With tMenus.tMenus(iMenuIndex).tMenuItems(iItemIndex)
        If iPos Then
            .iAccelerator = Asc(LCase$(Mid$(sCaption, iPos + 1&, 1)))
        Else
            .iAccelerator = 0
        End If
    
        .sCaption = sCaption
    
        liWidth = .iWidth
    
        Draw_GetItemExtent tMenus, iMenuIndex, iItemIndex
        
        If bRecalc Then
            If .iWidth > liWidth Then
                pModifyMenuItem tMenus, iMenuIndex, iItemIndex
            End If
        End If
    End With
    
End Sub

Private Sub pModifyMenuItem(ByRef tMenus As tMenus, ByVal iMenuIndex As Long, ByVal iItemIndex As Long)

'commented out is all we would have to do on older windows versions.  On XP it's much less efficient.

'    Dim lR As Long
'    Dim hMenu As Long
'    With tMenus.tMenus(iMenuIndex)
'        hMenu = .hMenu
'        If hMenu Then
'            With .tMenuItems(iItemIndex)
'                lR = ModifyMenuByLong(hMenu, iItemIndex, pGetMenuFlags(tMenus.tMenuDraw.iFlags, .iStyle) Or MF_BYPOSITION Or MF_OWNERDRAW Or CLng(IIf(.tChild.iId <> 0, MF_POPUP, 0)), .iApiID, .iItemData)
'            End With
'        End If
'    End With
'    Debug.Assert lR


'remove and reinsert the item...

    Dim lR As Long
    Dim hMenu As Long
    Dim liFlags As Long
    Dim bRTL As Boolean
    With tMenus.tMenus(iMenuIndex)
        hMenu = .hMenu
        If hMenu Then
            bRTL = .bRightToLeft
            With .tMenuItems(iItemIndex)
                lR = RemoveMenu(hMenu, iItemIndex, MF_BYPOSITION)
                Debug.Assert lR
                liFlags = pGetMenuFlags(tMenus.tMenuDraw.iFlags, .iStyle, bRTL) Or MF_BYPOSITION Or MF_OWNERDRAW
                lR = InsertMenuByLong(hMenu, iItemIndex, liFlags, .iApiID, .iItemData)
                Debug.Assert lR
                If (.tChild.iIndex <> Undefined) Then
                   ' If we had a submenu then put that back again:
                   liFlags = liFlags Or MF_POPUP
                   lR = ModifyMenuByLong(hMenu, iItemIndex, liFlags, tMenus.tMenus(.tChild.iIndex).hMenu, .iItemData)
                   Debug.Assert lR
                End If
            End With
        End If
    End With
End Sub

Private Function pParseCaption( _
            ByRef sCaption As String, _
            ByVal sToken As String, _
   Optional ByVal bRemoveTokens As Boolean) _
                As Long
    
    pParseCaption = InStr(1, sCaption, sToken)
    If bRemoveTokens Then
        Do While pParseCaption
            sCaption = Left$(sCaption, pParseCaption - 1&) & Mid$(sCaption, pParseCaption + 1&)
            If (StrComp(Mid$(sCaption, pParseCaption, 1), sToken) <> 0) Then Exit Do
            pParseCaption = InStr(pParseCaption + 1&, sCaption, sToken)
        Loop
    Else
        Do While pParseCaption
            If (StrComp(Mid$(sCaption, pParseCaption + 1&, 1), sToken) <> 0) Then Exit Do
            pParseCaption = InStr(pParseCaption + 2&, sCaption, sToken)
        Loop
    End If
End Function

Private Sub pAddNewMenuItem( _
            ByRef tMenus As tMenus, _
            ByVal iMenuIndex As Long, _
            ByVal iMenuItemIndex As Long)
        
Dim hMenu As Long
Dim hMenuParent As Long
Dim liParentMenuIndex As Long
Dim liParentItemIndex As Long
Dim liNewId As Long

    With tMenus.tMenus(iMenuIndex)
        hMenu = .hMenu
        If hMenu = 0& Then
            hMenu = CreatePopupMenu()
            'Debug.Print "Created " & hMenu
            .hMenu = hMenu
            If .tParent.iIndex > Undefined Then
                
                'this menu must become a unique ID
                Do While pAddThisId(hMenu) = Undefined
                    DestroyMenu hMenu
                    'Debug.Print "Destroyed " & hMenu
                    hMenu = CreatePopupMenu()
                    'Debug.Print "Created " & hMenu
                Loop
                
                .hMenu = hMenu
                liNewId = .iId
                liParentItemIndex = .tParent.iItemIndex
                liParentMenuIndex = .tParent.iIndex
                ' Now set the parent item so it has a popup menu:
                With tMenus.tMenus(liParentMenuIndex)
                    hMenuParent = .hMenu
                    With .tMenuItems(liParentItemIndex)
                        ' When you add a sub menu to an item, its id becomes the sub menu handle:
                        .iApiID = hMenu
                    
                        With .tChild
                            .iId = liNewId
                            .iIndex = iMenuIndex
                        End With
                        
                        pModifyMenuItem tMenus, liParentMenuIndex, liParentItemIndex
                        
                        pSetItemData hMenuParent, .iApiID, .iItemData
                    End With
                
                End With
            End If
        End If
        
        Dim bRTL As Boolean
        
        If (hMenu <> 0&) Then
            Dim lR As Long
            If iMenuItemIndex < .iItemCount - 1& Then
                bRTL = .bRightToLeft
                With .tMenuItems(iMenuItemIndex)
                    
                    lR = InsertMenuByLong(hMenu, iMenuItemIndex, (pGetMenuFlags(tMenus.tMenuDraw.iFlags, .iStyle, bRTL) Or MF_OWNERDRAW Or MF_BYPOSITION) _
                                                    And Not (MF_STRING Or MF_BYCOMMAND), _
                                          .iApiID, .iItemData)
                End With
            Else
                If GetMenuItemCount(hMenu) > .iItemCount Then
                    With .tMenuItems(iMenuItemIndex)
                        
                        lR = InsertMenuByLong(hMenu, iMenuItemIndex, (pGetMenuFlags(tMenus.tMenuDraw.iFlags, .iStyle, bRTL) Or MF_OWNERDRAW Or MF_BYPOSITION) _
                                                        And Not (MF_STRING Or MF_BYCOMMAND), _
                                              .iApiID, .iItemData)
                    End With
                Else
                    With .tMenuItems(iMenuItemIndex)
                        lR = AppendMenuByLong(hMenu, (pGetMenuFlags(tMenus.tMenuDraw.iFlags, .iStyle, bRTL) Or MF_OWNERDRAW Or MF_BYCOMMAND) _
                                                And Not (MF_STRING Or MF_BYPOSITION), _
                                      .iApiID, .iItemData)
                    End With
                End If
            End If
            Debug.Assert lR
        Else
            Debug.Assert False
        End If
    End With
    
End Sub

Private Function pGetNewID() As Long
    
    Const ONE As Long = 1
    
    Static ID As Long
    Dim liIndex As Long
    
    liIndex = Undefined
    Do While liIndex = Undefined
        
        If ID = &H7FFFFFFF Then 'wrap after a few billion items...
            ID = &H80000000
        ElseIf ID >= -ONE& And ID < &H800 Then 'make sure never to return -1& to 800h as an ID
            ID = &H801&
        Else
            ID = ID + ONE
        End If
        
        liIndex = pAddThisId(ID)
    Loop
    
    pGetNewID = ID
    
End Function

Private Function pAddThisId(ByVal iThis As Long) As Long

    Dim i As Long
    Dim liIndex As Long
    
    liIndex = Undefined
    pAddThisId = Undefined
    
    For i = 0& To miIDCount - 1&
        If miIDs(i) = 0& Then
            If liIndex = Undefined Then liIndex = i
        ElseIf miIDs(i) = iThis Then
            Exit Function
        End If
    Next
    
    If liIndex = Undefined Then
        liIndex = miIDCount
        miIDCount = miIDCount + 1
        Dim liUbound As Long
        Dim liNewUbound As Long
        On Error Resume Next
        liUbound = UBound(miIDs)
        If Err.Number Then liUbound = -1&
        liNewUbound = pArrAdjustUbound(miIDCount)
        If liNewUbound > liUbound Then ReDim Preserve miIDs(0 To liNewUbound)
        
    End If
    
    miIDs(liIndex) = iThis
    
    pAddThisId = liIndex
    
End Function

Private Sub pRemoveId(ByVal iId As Long)
    Dim i As Long
    For i = 0 To miIDCount - 1&
        If miIDs(i) = iId Then
            miIDs(i) = Undefined
            Exit For
        End If
    Next
    If i = miIDCount - 1& Then
        For i = i - 1& To 0& Step -1&
            If miIDs(i) <> Undefined Then Exit For
        Next
        miIDCount = i + 1&
    End If
End Sub

Private Function pGetMenuFlags( _
                    ByVal iDrawStyle As Long, _
                    ByVal iStyle As Long, _
                    ByVal bRTL As Boolean) _
                        As Long
      If CBool(iStyle And mnuChecked) Then
          pGetMenuFlags = pGetMenuFlags Or MF_CHECKED
      Else
          pGetMenuFlags = pGetMenuFlags Or MF_UNCHECKED
      End If
      If CBool(iStyle And mnuDisabled) Then
          pGetMenuFlags = pGetMenuFlags Or MF_GRAYED
      Else
          pGetMenuFlags = pGetMenuFlags Or MF_ENABLED
      End If
      If CBool(iStyle And mnuSeparator) Or CBool(iStyle And mnuInvisible) Or (CBool(iStyle And mnuInfrequent) And Not CBool(iDrawStyle And mnuShowInfrequent)) Then
         'Debug.Assert Not CBool(iStyle And mnuInfrequent)
         pGetMenuFlags = pGetMenuFlags Or MF_SEPARATOR
      End If
      If (iStyle And mnuNewVerticalLine) Then
         pGetMenuFlags = pGetMenuFlags Or MF_MENUBARBREAK
      End If
      
      If (bRTL) Then
         pGetMenuFlags = pGetMenuFlags Or MF_RIGHTORDER
      End If
      
      'If (iStyle and ) Then
      '   pGetMenuFlags = pGetMenuFlags Or MF_MENUBREAK
      'End If
   
End Function

Private Sub pGetExclude(ByRef vExclude As Variant, ByRef tr As tRect)
    Dim liHwnd As Long
    If Not IsMissing(vExclude) Then
        
        On Error GoTo catch
        If VarType(vExclude) = vbLong Then
            liHwnd = vExclude
        Else
            If InStr(1, TypeName(vExclude), "rect", vbTextCompare) = 0 Then
                liHwnd = vExclude.hWnd
            Else
                tr.Left = vExclude.Left
                tr.Right = vExclude.Right
                tr.Bottom = vExclude.Bottom
                tr.Top = vExclude.Top
                Exit Sub
            End If
        End If
        
        If IsWindow(liHwnd) = 0& Then
            On Error GoTo 0
catch:
            pErr 5, "cPopupMenu.Show"
        End If
        
        GetWindowRect liHwnd, tr
        
    End If
End Sub

Private Function pGetIndex(ByRef tMenu As tMenu, ByRef vKey As Variant) As Long
    On Error GoTo catch
    Dim i As Long
    i = tMenu.iItemCount
    If pIsNumericVarType(VarType(vKey)) Then
        pGetIndex = CLng(vKey) - 1& 'this will cause an error if vKey doesn't fit in a long
        If (pGetIndex < i And pGetIndex > Undefined) Then GoTo exitproc
        'if the index is invalid, pass control to the error handler
    Else
        Dim lsKey As String
        lsKey = CStr(vKey)
        
        With tMenu
            For pGetIndex = 0& To i - 1&
                If StrComp(.tMenuItems(pGetIndex).sKey, lsKey) = 0& Then GoTo exitproc
            Next
            'if the index is not found, control is passed to the error handler
        End With
    End If
    On Error GoTo 0 'in case we got here w/o an error
catch:
    'pGetIndex = Undefined
    
    pErr 5, "cPopupMenus.GetMenuIndex"
exitproc:
End Function

Private Function pGetRootMenuIndex(ByRef tMenus As tMenus, ByRef vKey As Variant) As Long
    On Error GoTo catch
    Dim i As Long
    i = tMenus.tRootMenuLookup.iMenuCount
    If pIsNumericVarType(VarType(vKey)) Then
        pGetRootMenuIndex = CLng(vKey) - 1& 'this will cause an error if vKey doesn't fit in a long
        If (i > pGetRootMenuIndex And i >= 0&) Then GoTo exitproc
        'if the index is invalid, pass control to the error handler
    Else
        Dim lsKey As String
        lsKey = CStr(vKey)
        
        With tMenus
            For pGetRootMenuIndex = 0& To i - 1&
                If StrComp(.tMenus(.tRootMenuLookup.iMenuIndices(pGetRootMenuIndex)).sKey, lsKey) = 0& Then GoTo exitproc
            Next
            'if the index is not found, control is passed to the error handler
        End With
    End If
    On Error GoTo 0 'in case we got here w/o an error
catch:
    'pGetIndex = Undefined
    pErr 5, "cPopupMenus.GetRootMenuIndex"
exitproc:
End Function

Private Function pGetNewMenuIndex(ByRef tMenus As tMenus) As Long
    Dim liCount As Long
    
    With tMenus
        liCount = .iMenuCount
        For pGetNewMenuIndex = 0 To liCount - 1&
            If .tMenus(pGetNewMenuIndex).iId = 0& Then Exit For
        Next
        
        If pGetNewMenuIndex = liCount Then
            liCount = liCount + 1&
            pResizeMenus .tMenus, liCount
            .iMenuCount = liCount
        End If
    End With
End Function

Private Function pIsNumericVarType(ByVal iType As VariantTypeConstants) As Boolean
    pIsNumericVarType = (iType = vbVCurrency Or iType = vbVDouble Or iType = vbVInteger Or iType = vbVLong Or iType = vbVSingle)
End Function

Private Sub pDetachMenuFromParent(ByRef tMenus As tMenus, ByVal iIndex As Long)
    Dim hMenu As Long
    Dim liItemIndex As Long
    Dim bRTL As Boolean
    With tMenus.tMenus(iIndex)
        If .hMenu Then
            With .tParent
                liItemIndex = .iItemIndex
                If .iId <> 0& Then
                    hMenu = tMenus.tMenus(.iIndex).hMenu
                    bRTL = tMenus.tMenus(.iIndex).bRightToLeft
                    With tMenus.tMenus(.iIndex).tMenuItems(liItemIndex)
                        
                        Dim ltMI As MENUITEMINFO
                        Dim liFlags As Long
                        
                        ' remove it from the menu:
                        ltMI.fMask = MIIM_ID
                        RemoveMenu hMenu, liItemIndex, MF_BYPOSITION

                        ' Insert it back again at the corect position with the same ID etc:
                        liFlags = pGetMenuFlags(tMenus.tMenuDraw.iFlags, .iStyle, bRTL) _
                                        Or (MF_OWNERDRAW Or MF_BYPOSITION) _
                                        And Not (MF_STRING Or MF_BYCOMMAND)

                        InsertMenuByLong hMenu, liItemIndex, liFlags, .iApiID, .iItemData
                        .tChild.iId = 0&
                        .tChild.iIndex = Undefined
                    End With
                End If
                .iId = 0&
                .iIndex = Undefined
                .iItemIndex = Undefined
            End With
        End If
    End With
End Sub

Private Sub pRemoveMenu(ByRef tMenus As tMenus, ByVal iIndex As Long)
    Dim liEach As Long
    Dim liTemp As Long
    With tMenus
        With .tMenus(iIndex)
            For liEach = 0 To .iItemCount - 1&
                liTemp = .tMenuItems(liEach).tChild.iIndex
                If liTemp > Undefined Then pRemoveMenu tMenus, liTemp
                pRemoveId .tMenuItems(liEach).iId
                .iId = 0&
            Next
            liTemp = .hMenu
            If liTemp Then DestroyMenu liTemp
            'If liTemp Then Debug.Print "Destroyed " & liTemp
            pRemoveId .iId
            .iId = 0&
            .hMenu = 0&
            Erase .tMenuItems
        End With
        For liEach = .iMenuCount - 1& To 0& Step -1&
            If .tMenus(liEach).hMenu Then Exit For
        Next
        .iMenuCount = liEach + 1&
    End With
End Sub

Private Function pValidatePointer( _
                    ByRef tMenus As tMenus, _
                    ByRef tPointer As tMenuPointer) _
                        As Boolean
    If tPointer.iIndex > Undefined Then
        If tMenus.iMenuCount > tPointer.iIndex Then
            pValidatePointer = tMenus.tMenus(tPointer.iIndex).iId = tPointer.iId
        End If
    End If

    If Not pValidatePointer Then pErr ccElemNotPartOfCollection, "PopupMenus.ValidateMenu"
End Function

Private Function pValidateItemPointer( _
                    ByRef tMenus As tMenus, _
                    ByRef tPointer As tMenuItemPointer) _
                        As Boolean
    If tPointer.iIndex > Undefined Then
        If tMenus.iMenuCount > tPointer.iIndex Then
            If tPointer.iItemIndex > Undefined Then
                If tMenus.tMenus(tPointer.iIndex).iItemCount > tPointer.iItemIndex Then
                    pValidateItemPointer = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iId = tPointer.iId
                End If
            End If
            
            If Not pValidateItemPointer Then
                With tMenus.tMenus(tPointer.iIndex)
                    For tPointer.iItemIndex = 0& To .iItemCount - 1&
                        If .tMenuItems(tPointer.iItemIndex).iId = tPointer.iId Then
                            pValidateItemPointer = True
                            Exit For
                        End If
                    Next
                End With
            End If

        End If
    End If
    
    If Not pValidateItemPointer Then pErr ccElemNotPartOfCollection, "PopupMenus.ValidateMenuItem"

End Function

Private Sub pErr(ByVal iNum As Long, ByRef sSource As String)
    Dim lsDesc As String
    Select Case iNum
        Case 5, 13
            Err.Raise iNum, sSource
        Case ccElemNotPartOfCollection
            lsDesc = "Item is not part of a collection."
        Case ccCollectionChangedDuringEnum
            lsDesc = "Collection changed during enumeration."
    End Select
    Err.Raise iNum, sSource, lsDesc
End Sub

Private Sub pSetFont(ByRef tMenus As tMenus)
    
    Dim hFntOld As Long
    Dim oIFont As IFont
    
    With tMenus.tMenuDraw
        If Not .oFont Is Nothing Then
            Set oIFont = .oFontSymbol
            oIFont.Size = .oFont.Size * 1.3
            
            Set oIFont = .oFont
            hFntOld = SelectObject(.oMemDC.hdc, oIFont.hFont)
        End If
        
        
        Draw_GetTextExtent .oMemDC.hdc, "Aa", , .iItemHeight

        SelectObject .oMemDC.hdc, hFntOld
        
        .iItemHeight = .iItemHeight + 4
        If .iItemHeight < .iIconSize + 6& Then .iItemHeight = .iIconSize + 6&
    
        pResetForRecalc tMenus
    End With
   
End Sub

Private Sub pResetForRecalc(ByRef tMenus As tMenus, Optional ByVal iIndex As Long = Undefined, Optional ByVal iItemIndex As Long = Undefined)
    Dim liDrawStyle As Long
    liDrawStyle = tMenus.tMenuDraw.iFlags
    If iIndex > Undefined And iIndex < tMenus.iMenuCount Then
        If iItemIndex > Undefined And iItemIndex < tMenus.tMenus(iIndex).iItemCount Then
            With tMenus.tMenus(iIndex).tMenuItems(iItemIndex)
                pModifyMenuItem tMenus, iIndex, iItemIndex
                pSetWidth tMenus, iIndex, iItemIndex
            End With
        Else
            pResetMenuForRecalc tMenus, iIndex
            pSetWidth tMenus, iIndex, -1
        End If
    Else
        For iIndex = 0& To tMenus.iMenuCount - 1&
            pResetMenuForRecalc tMenus, iIndex
        Next
        pSetWidth tMenus, -1, -1
    End If
End Sub

Private Sub pResetMenuForRecalc(ByRef tMenus As tMenus, ByVal iIndex As Long)

    Dim liItemIndex As Long
    
    With tMenus.tMenus(iIndex)
        For liItemIndex = 0& To .iItemCount - 1&
            pModifyMenuItem tMenus, iIndex, liItemIndex
        Next
    End With

End Sub

Private Sub pSetWidth(ByRef tMenus As tMenus, ByVal iIndex As Long, ByVal iItemIndex As Long)
    
    If iIndex > Undefined And iIndex < tMenus.iMenuCount Then
        With tMenus.tMenus(iIndex)
            If iItemIndex > Undefined And iItemIndex < .iItemCount Then
                Draw_GetItemExtent tMenus, iIndex, iItemIndex
            Else
                For iItemIndex = 0& To .iItemCount - 1&
                    Draw_GetItemExtent tMenus, iIndex, iItemIndex
                Next
            End If
        End With
    Else
        For iIndex = 0 To tMenus.iMenuCount - 1&
            With tMenus.tMenus(iIndex)
                For iItemIndex = 0 To .iItemCount - 1&
                    Draw_GetItemExtent tMenus, iIndex, iItemIndex
                Next
            End With
        Next
    End If

End Sub

Private Sub pSetShortcut(ByRef tItem As tMenuItem)
    Dim liMask As ShiftConstants
    Dim liKey As KeyCodeConstants
    Dim lsKey As String
    Dim lsMask As String
    
    Const sAlt = "Alt"
    Const sControl = "Ctrl"
    Const sShift = "Shift"
    Const sPlus = "+"
    
    liMask = tItem.iShortcutMask
    liKey = tItem.iShortcutShiftKey
    
    If CBool(liMask And (vbCtrlMask Or vbShiftMask Or vbAltMask)) Then
        
        Select Case liKey
            Case vbKeyHome
                lsKey = "Home"
            Case vbKeyEnd
                lsKey = "End"
            Case vbKeyLeft
                lsKey = "Left"
            Case vbKeyRight
                lsKey = "Right"
            Case vbKeyUp
                lsKey = "Up"
            Case vbKeyDown
                lsKey = "Down"
            Case vbKeyClear
                lsKey = "Clear"
            Case vbKeyPageUp
                lsKey = "Pg Up"
            Case vbKeyPageDown
                lsKey = "Pg Dn"
            Case vbKeyDelete
                lsKey = "Del"
            Case vbKeyEscape
                lsKey = "Esc"
            Case vbKeyTab
                lsKey = "Tab"
            Case vbKeyReturn
                lsKey = "Return"
            Case vbKeyAdd
                lsKey = "Plus"
            Case vbKeySubtract
                lsKey = "Minus"
            Case vbKeyBack
                lsKey = "Bkspc"
            Case vbKeyDivide
                lsKey = "Divide"
            Case vbKeyMultiply
                lsKey = "Multiply"
            Case vbKeyInsert
                lsKey = "Ins"
            Case vbKeySpace
                lsKey = "Space"
            Case vbKeyF1 To vbKeyF16
                lsKey = "F" & vbKeyF1 - liKey + 1&
            Case Else
                lsKey = UCase$(Chr$(liKey))
        End Select
        
        tItem.sShortCutDisplay = IIf(CBool(liMask And vbCtrlMask), sControl & sPlus, vbNullString) & _
                IIf(CBool(liMask And vbShiftMask), sShift & sPlus, vbNullString) & _
                IIf(CBool(liMask And vbAltMask), sAlt & sPlus, vbNullString) & _
                lsKey
    End If
End Sub

Private Function pIndexForId(ByRef tMenus As tMenus, ByVal iId As Long, ByRef iMenuIndex As Long, ByRef iItemIndex As Long) As Boolean
    Dim liCount As Long
    liCount = tMenus.iMenuCount
    Dim liItem As Long
    Dim liMenu As Long
    
    liMenu = iMenuIndex
    
    If liMenu > Undefined And liMenu < liCount Then
        
        With tMenus.tMenus(liMenu)
            For liItem = 0& To .iItemCount - 1&
                With .tMenuItems(liItem)
                    pIndexForId = CBool(.iApiID = iId)
                    If pIndexForId Then
                        iMenuIndex = liMenu
                        iItemIndex = liItem
                    ElseIf .tChild.iIndex > Undefined Then
                        iMenuIndex = .tChild.iIndex
                        pIndexForId = pIndexForId(tMenus, iId, iMenuIndex, iItemIndex)
                    End If
                    If pIndexForId Then Exit For
                End With
            Next
        End With
    Else
        With tMenus
            For iMenuIndex = 0 To liCount - 1&
                With .tMenus(iMenuIndex)
                    For liItem = 0 To .iItemCount - 1&
                        pIndexForId = (.tMenuItems(liItem).iApiID = iId)
                        If pIndexForId Then Exit For
                    Next
                End With
                If pIndexForId Then Exit For
            Next
            If pIndexForId Then
                iItemIndex = liItem
            Else
                iItemIndex = Undefined
                iMenuIndex = Undefined
            End If
        End With
    End If
End Function

Private Function pIndexForHandle(ByRef tMenus As tMenus, ByVal hMenu As Long, ByRef iIndex As Long) As Boolean
    If hMenu Then
        For iIndex = 0 To tMenus.iMenuCount - 1&
            If tMenus.tMenus(iIndex).hMenu = hMenu Then Exit For
        Next
        If Not CBool(iIndex = tMenus.iMenuCount) Then pIndexForHandle = True Else iIndex = -1&
    End If
End Function

Private Sub pRemoveChevrons(ByRef tMenus As tMenus, ByVal iIndex As Long)
    Dim liIndex As Long
    Dim lR As Long
    Dim ltPointer As tMenuPointer
    
    With tMenus.tMenus(iIndex)
        For liIndex = .iItemCount - 1 To 0& Step -1&
            If CBool(.tMenuItems(liIndex).iStyle And mnuChevron) Then
                ltPointer.iIndex = iIndex
                ltPointer.iId = .iId
                mPopup.PopupMenuItems_Remove tMenus, ltPointer, liIndex + 1&
            Else
                If .tMenuItems(liIndex).tChild.iId <> 0 _
                Then pRemoveChevrons tMenus, .tMenuItems(liIndex).tChild.iIndex
            End If
        Next
    End With
End Sub

Private Function pGetMissing(Optional ByVal v As Variant) As Variant
    pGetMissing = v
End Function

Private Sub pProcessBitmap(ByRef tDraw As tMenuDraw)
    With tDraw
    
        If Not (tDraw.oPic Is Nothing) Then
            Dim hdc As Long
            Dim liHeight As Long
            Dim liWidth As Long
    

            Set .oBitmap = New pcMemDC
            With .oBitmap
                .CreateFromPicture tDraw.oPic
                hdc = .hdc
                liHeight = .Height
                liWidth = .Width
            End With
                
            If CBool(tDraw.iFlags And mnuImageProcessBitmap) Then
            
                Dim cDib As New pcDibSection
                Set cDib = New pcDibSection
          
                If cDib.Create(liWidth, liHeight) Then
                    ' create a lighter version:
                    cDib.LoadPictureBlt hdc
                    cDib.Lighten 20
             
                    Set .oBitmapLight = New pcMemDC
                    With .oBitmapLight
                        .Width = liWidth
                        .Height = liHeight
                        cDib.PaintPicture .hdc
                    End With
                    
                    'cDib.LoadPictureBlt hDc
                    cDib.Lighten 20
             
                    Set .oBitmapSuperLight = New pcMemDC
                    With .oBitmapSuperLight
                        .Width = liWidth
                        .Height = liHeight
                        cDib.PaintPicture .hdc
                    End With
                Else
                    Set .oBitmapLight = Nothing
                    Set .oBitmapSuperLight = Nothing
                End If
            Else
                Set .oBitmapLight = Nothing
                Set .oBitmapSuperLight = Nothing
            End If
        Else
            Set .oBitmap = Nothing
            Set .oBitmapLight = Nothing
            Set .oBitmapSuperLight = Nothing
        End If

    End With
End Sub

Private Sub pGetHierarchy(ByRef tMenus As tMenus, ByVal iMenuIndex As Long)
    miRedisplayHierarchyCount = 0&
    With tMenus
        Do While .tMenus(iMenuIndex).tParent.iIndex > Undefined
            ReDim Preserve miRedisplayHierarchy(0 To miRedisplayHierarchyCount)
            miRedisplayHierarchy(miRedisplayHierarchyCount) = .tMenus(iMenuIndex).hMenu
            miRedisplayHierarchyCount = miRedisplayHierarchyCount + 1&
            iMenuIndex = .tMenus(iMenuIndex).tParent.iIndex
        Loop
    End With
End Sub

Private Function pGetMenuFont(ByVal hdc As Long) As StdFont
    Dim tNCM As NONCLIENTMETRICS
    tNCM.cbSize = 340 'LenB(m_tNCM) - why doesn't this go?
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, tNCM, 0
    Set pGetMenuFont = pLogFontToStdFont(tNCM.lfMenuFont, hdc)
End Function

Private Function pLogFontToStdFont(ByRef tLF As NMLOGFONT, ByVal hdc As Long) As StdFont
   Set pLogFontToStdFont = New StdFont
   With pLogFontToStdFont
     .Name = StrConv(tLF.lfFaceName, vbUnicode)
     If tLF.lfHeight < 1 Then
       .Size = Abs((72# / GetDeviceCaps(hdc, LOGPIXELSY)) * tLF.lfHeight)
     Else
       .Size = tLF.lfHeight
     End If
     .Charset = tLF.lfCharSet
     .Italic = Not (tLF.lfItalic = 0)
     .Underline = Not (tLF.lfUnderline = 0)
     .Strikethrough = Not (tLF.lfStrikeOut = 0)
     .Bold = (tLF.lfWeight > FW_REGULAR)
   End With
End Function
'</Private Interface>

'<cPopupMenus Interface>
Public Function PopupMenus_Add( _
            ByRef tMenus As tMenus, _
            ByRef sKey As String, _
            ByRef vKeyOrIndexInsertBefore As Variant, _
            ByVal oOwner As cPopupMenus) _
                As cPopupMenu
    
    Dim liCount As Long
    Dim liRootMenuIndex As Long
    Dim liNewIndex As Long
    Dim liRealIndex As Long
    Dim liLen As Long
    Dim tPointer As tMenuPointer
    
    With tMenus.tRootMenuLookup
    
        liCount = .iMenuCount
        
        If IsMissing(vKeyOrIndexInsertBefore) Then
            liNewIndex = liCount
        Else
            liRootMenuIndex = pGetRootMenuIndex(tMenus, vKeyOrIndexInsertBefore)
        End If

        liLen = (liCount - liNewIndex) * 4&
        liCount = liCount + 1&
        .iMenuCount = liCount
        pResizeArray .iMenuIndices, liCount
        
        If liLen Then CopyMemory .iMenuIndices(liNewIndex + 1&), .iMenuIndices(liNewIndex), liLen
        
        liRealIndex = pGetNewMenuIndex(tMenus)
        
        .iMenuIndices(liNewIndex) = liRealIndex
        
        With tMenus.tMenus(liRealIndex)
            .hMenu = 0&
            .iId = pGetNewID()
            .iItemCount = 0&
            .sKey = sKey
            
            With .tParent
                .iId = 0
                .iIndex = Undefined
                .iItemIndex = Undefined
            End With
            
            tPointer.iId = .iId
            tPointer.iIndex = liRealIndex
        End With
        
        Set PopupMenus_Add = New cPopupMenu
        PopupMenus_Add.fInit oOwner, tPointer
        
    End With

End Function

Public Function PopupMenus_Count( _
            ByRef tMenus As tMenus) _
                As Long
    PopupMenus_Count = tMenus.tRootMenuLookup.iMenuCount
End Function

Public Function PopupMenus_Clear( _
            ByRef tMenus As tMenus) _
                As Long
    With tMenus.tRootMenuLookup
        For PopupMenus_Clear = 0& To .iMenuCount - 1&
            pRemoveMenu tMenus, .iMenuIndices(PopupMenus_Clear)
        Next
    End With
End Function

Public Function PopupMenus_Exists( _
            ByRef tMenus As tMenus, _
            ByVal vKeyOrIndex As Variant) _
                As Boolean
    On Error GoTo catch
    pGetRootMenuIndex tMenus, vKeyOrIndex
    PopupMenus_Exists = True
catch:
End Function

Public Property Get PopupMenus_Item( _
            ByRef tMenus As tMenus, _
            ByVal vKeyOrIndex As Variant, _
            ByVal oOwner As cPopupMenus) _
                As cPopupMenu
    Dim tPointer As tMenuPointer
    Dim liIndex As Long
        
    liIndex = tMenus.tRootMenuLookup.iMenuIndices(pGetRootMenuIndex(tMenus, vKeyOrIndex))
    tPointer.iIndex = liIndex
    tPointer.iId = tMenus.tMenus(liIndex).iId
    
    Set PopupMenus_Item = New cPopupMenu
    PopupMenus_Item.fInit oOwner, tPointer
    
End Property

Public Sub PopupMenus_Remove( _
            ByRef tMenus As tMenus, _
            ByVal vKeyOrIndex As Variant)

    Dim liIndex As Long
    
    With tMenus.tRootMenuLookup
        liIndex = .iMenuIndices(pGetRootMenuIndex(tMenus, vKeyOrIndex))
        pRemoveMenu tMenus, liIndex
        pDetachMenuFromParent tMenus, liIndex
    End With
    
End Sub

Public Property Set PopupMenus_Font(ByRef tMenus As tMenus, ByVal oNew As StdFont)
    If oNew Is Nothing Then Set oNew = pGetMenuFont(tMenus.tMenuDraw.oMemDC.hdc)
    Set tMenus.tMenuDraw.oFont = oNew
    pSetFont tMenus
End Property
'</cPopupMenus Interface>

'<cPopupMenus Interface>
Public Sub PopupMenus_WndProc( _
            ByRef tMenus As tMenus, _
            ByVal bBefore As Boolean, _
            ByRef bHandled As Boolean, _
            ByRef lReturn As Long, _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            ByVal oOwner As cPopupMenus)
            
Dim liID As Long
Dim liIndex As Long
Dim liItemIndex As Long
Dim hMenu As Long
Dim liChar As Integer
Dim ltPointer As tMenuItemPointer
Dim ltMenuPointer As tMenuPointer
Dim loItem As cPopupMenuItem
Dim loItems As cPopupMenuItems
Dim loMenu As cPopupMenu

Static liLastHighlight As Long

    Select Case iMsg
    
    ' Handle Menu Select events:
    Case WM_MENUSELECT
        ' Extract the menu id and flags for the selected
        ' menu item:
        liID = wParam And &HFFFF&

        ' Menu handle is passed in as lParam:
        hMenu = lParam
        
        ltPointer.iId = liID
        ltPointer.iIndex = -1&

        If pIndexForHandle(tMenus, hMenu, ltPointer.iIndex) Then
            If pIndexForId(tMenus, liID, ltPointer.iIndex, ltPointer.iItemIndex) Then
                Set loItem = New cPopupMenuItem
                loItem.fInit oOwner, ltPointer
                liLastHighlight = ltPointer.iIndex
                oOwner.fItemHighlight loItem
                
                If CBool(tMenus.tMenus(ltPointer.iIndex).tMenuItems(ltPointer.iItemIndex).iStyle And mnuChevron) And miInfreqShowDelay > 0& Then  'And (CBool((wParam And Not &HFFFF&) \ &H1000&) And MF_MOUSESELECT)
                    miChevronIndex = ltPointer.iItemIndex
                    mhWndChevronMenu = hMenu
                    mbChevronSelected = True
                    miHoverTime = 0
                    oOwner.fCallback 1, 50
                Else
                    mhWndChevronMenu = 0&
                    miChevronIndex = -2&
                    mbChevronSelected = False
                End If
                
            Else
                If liLastHighlight > Undefined Then
                    mhWndChevronMenu = 0&
                    miChevronIndex = -2&
                    mbChevronSelected = False
                    liLastHighlight = Undefined
                    oOwner.fItemHighlight Nothing
                End If
            End If
        Else
            If liLastHighlight > Undefined Then
                mhWndChevronMenu = 0&
                miChevronIndex = -2&
                mbChevronSelected = False
                liLastHighlight = Undefined
                oOwner.fItemHighlight Nothing
            End If
            'menu not found!
            'Debug.Assert False
        End If
    
    ' Draw Menu items:
    Case WM_DRAWITEM
        Dim tDIS As DRAWITEMSTRUCT
        
        CopyMemory tDIS, ByVal lParam, Len(tDIS)
        
        If tDIS.CtlType = ODT_MENU Then

          If pIndexForHandle(tMenus, tDIS.hwndItem, liIndex) Then
            If pIndexForId(tMenus, tDIS.itemID, liIndex, liItemIndex) Then
                Draw_DrawItem tDIS, tMenus.tMenuDraw, tMenus, tMenus.tMenus(liIndex), tMenus.tMenus(liIndex).tMenuItems(liItemIndex), liIndex, liItemIndex, wParam, lParam, oOwner
                CopyMemory ByVal lParam, tDIS, Len(tDIS)
                lReturn = 1
                bHandled = True
            Else
                Debug.Assert False
            End If
          Else
            Debug.Assert False
          End If
        End If
        
    ' Measure Menu items prior to drawing them:
    Case WM_MEASUREITEM
        Dim tMIS As MEASUREITEMSTRUCT
           CopyMemory tMIS, ByVal lParam, LenB(tMIS)
           If tMIS.CtlType = ODT_MENU Then
              liIndex = Undefined
              If pIndexForId(tMenus, tMIS.itemID, liIndex, liItemIndex) Then
                  With tMenus.tMenus(liIndex).tMenuItems(liItemIndex)
                    tMIS.itemHeight = .iHeight
                    tMIS.itemWidth = .iWidth - (xSystemMetric(SM_CXMENUCHECK) - 1&)
                  End With
                  CopyMemory ByVal lParam, tMIS, LenB(tMIS)
                  lReturn = 1
                  bHandled = True
              Else
                  Debug.Assert False
              End If
           End If
    
    ' Handle accelerator (&key) messages in the menu:
    Case WM_MENUCHAR
        ' Check that this is my menu:
        If Not CBool((wParam \ &H10000) And MF_SYSMENU) Then
            hMenu = lParam
            liChar = (wParam And &HFFFF&)
            ' See if this corresponds to an accelerator on the menu:
            If pIndexForHandle(tMenus, hMenu, liIndex) Then
                'Debug.Assert liChar <> vbKeyEscape
                lReturn = pParseMenuChar(tMenus.tMenuDraw.iFlags And mnuShowInfrequent, tMenus.tMenus(liIndex), liChar)
                bHandled = CBool(lReturn)
            Else
                'menu not found!
                Debug.Assert False
            End If
        End If
        
    Case WM_INITMENUPOPUP, WM_UNINITMENUPOPUP
      ' Check the sys menu flag:
        If (lParam \ &H10000) > 0 Then
            ' System menu.
        Else
            hMenu = wParam
            ' Find the item which is the parent
            ' of this popup menu:
            If pIndexForHandle(tMenus, hMenu, ltMenuPointer.iIndex) Then
                ltMenuPointer.iId = tMenus.tMenus(ltMenuPointer.iIndex).iId
                Set loItems = New cPopupMenuItems
                loItems.fInit oOwner, ltMenuPointer
                
                With tMenus.tMenus(ltMenuPointer.iIndex)
                    If iMsg = WM_INITMENUPOPUP Then
                        If Not CBool(tMenus.tMenuDraw.iFlags And mnuShowInfrequent) Then
                            For liItemIndex = 0 To .iItemCount - 1&
                                If CBool(.tMenuItems(liItemIndex).iStyle And mnuInfrequent) Then Exit For
                            Next

                            If liItemIndex < .iItemCount Then
                                For liItemIndex = .iItemCount - 1& To 0& Step -1&
                                    If CBool(.tMenuItems(liItemIndex).iStyle And mnuChevron) Then Exit For
                                Next
                                If liItemIndex = Undefined Then
                                    liItemIndex = mPopup.PopupMenuItems_Add(tMenus, ltMenuPointer, vbNullString, vbNullString, vbNullString, 0, -1&, mnuChevron, 0, 0, pGetMissing()) - 1&
                                    If liItemIndex > Undefined Then Set loItem = loItems.Item(liItemIndex + 1)
                                End If
                            End If
                        End If
                    End If
                End With
            Else
                'menu not found!
                Debug.Assert False
            End If
            oOwner.fMenuInit loItems, (iMsg = WM_INITMENUPOPUP), loItem
      End If
   Case WM_MENURBUTTONUP
      If pIndexForHandle(tMenus, lParam, ltPointer.iIndex) Then
        ltPointer.iItemIndex = wParam
        ltPointer.iId = tMenus.tMenus(ltPointer.iIndex).tMenuItems(wParam).iId
        Set loItem = New cPopupMenuItem
        loItem.fInit oOwner, ltPointer
        oOwner.fRightClick loItem
      End If
   Case WM_DESTROY
      oOwner.fSubclass = False
      
   End Select
            
End Sub

Public Sub PopupMenus_SetPicture(ByRef tMenus As tMenus, ByVal oPic As StdPicture)
    If Not (tMenus.tMenuDraw.oPic Is oPic) Then
        Set tMenus.tMenuDraw.oPic = oPic
        pProcessBitmap tMenus.tMenuDraw
    End If
End Sub
            
Public Function PopupMenus_AcceleratorPress(ByRef tMenus As tMenus, ByVal iKey As KeyCodeConstants, ByVal iMask As ShiftConstants, ByVal oOwner As cPopupMenus) As cPopupMenuItem
    Dim liIndex As Long
    Dim liItem As Long
    If iKey <> 0& And (iMask And (vbAltMask Or vbShiftMask Or vbCtrlMask)) Then
        For liIndex = 0& To tMenus.iMenuCount - 1&
            With tMenus.tMenus(liIndex)
                For liItem = 0& To .iItemCount - 1&
                    If .tMenuItems(liItem).tChild.iIndex = Undefined Then
                        If .tMenuItems(liItem).iShortcutShiftKey = iKey Then
                            If .tMenuItems(liItem).iShortcutMask = iMask Then
                                Dim ltPointer As tMenuItemPointer
                                With ltPointer
                                    .iIndex = liIndex
                                    .iItemIndex = liItem
                                End With
                                ltPointer.iId = .tMenuItems(liItem).iId
                                
                                Set PopupMenus_AcceleratorPress = New cPopupMenuItem
                                PopupMenus_AcceleratorPress.fInit oOwner, ltPointer
                                
                                oOwner.fClick PopupMenus_AcceleratorPress
                            End If
                        End If
                    End If
                Next
            End With
        Next
    End If
End Function
            
Public Sub PopupMenus_SetImageList(ByRef tMenus As tMenus, ByRef vIml As Variant)
    On Error GoTo catch
    With tMenus.tMenuDraw
        .hIml = 0&
        Set .oVB6ImageList = Nothing
        If VarType(vIml) = vbLong Then
            .hIml = vIml
            If .hIml <> 0& Then
                Dim tr As tRect
                ImageList_GetImageRect .hIml, 0, tr
                .iIconSize = tr.Bottom - tr.Top + 6
            Else
                .iIconSize = 0
            End If
        ElseIf VarType(vIml) = vbObject Then
            Set .oVB6ImageList = vIml
            If Not (.oVB6ImageList Is Nothing) Then
                .oVB6ImageList.ListImages(1).Draw 0, 0, 0, 1
                .iIconSize = .oVB6ImageList.ImageHeight
            Else
                .iIconSize = 0
            End If
        End If
        If .iItemHeight < .iIconSize + 6 Then .iItemHeight = .iIconSize + 6
    End With
    
    If False Then
catch:
        pErr 13, "cPopupMenu.SetImageList"
    End If
    
End Sub

Public Sub PopupMenus_SetDrawStyle(ByRef tMenus As tMenus, ByVal iStyle As Long)
    Dim liOriginalState As Long
    Dim lbRecalc As Boolean
    Dim lbProcessBitmap As Boolean
    
    'mask out any invalid flags
    iStyle = iStyle And &H3F&
    
    liOriginalState = tMenus.tMenuDraw.iFlags

    If (liOriginalState And mnuDrawSeparatorsAsHeaders) Xor (iStyle And mnuDrawSeparatorsAsHeaders) Then
        lbRecalc = True
    End If
    
    If (liOriginalState And mnuShowInfrequent) Xor (iStyle And mnuShowInfrequent) Then
        lbRecalc = True
    End If
    
    If (liOriginalState And mnuImageProcessBitmap) Xor (iStyle And mnuImageProcessBitmap) Then
        lbProcessBitmap = True
    End If

    tMenus.tMenuDraw.iFlags = iStyle
    
    If lbRecalc Then
        Dim iIndex As Long
        Dim iItem As Long
        Dim lR As Long

        With tMenus
            For iIndex = 0 To .iMenuCount - 1&
                lbRecalc = False
                With .tMenus(iIndex)
                    For iItem = 0 To .iItemCount - 1&
                        If CBool(.tMenuItems(iItem).iStyle And mnuInfrequent) Then
                            pModifyMenuItem tMenus, iIndex, iItem
                            lbRecalc = True
                        End If
                    Next
                End With
                If lbRecalc Then pSetWidth tMenus, iIndex, -1&
            Next
        End With
    End If
    If lbProcessBitmap Then pProcessBitmap tMenus.tMenuDraw
End Sub
          
Public Sub PopupMenus_Initialize(ByRef tMenus As tMenus)
    'm_lLastMaxId = &H800

    ' Stuff for drawing:
    With tMenus.tMenuDraw
        Set .oMemDC = New pcMemDC
        With .oMemDC
            .Width = Screen.Width \ Screen.TwipsPerPixelY
            .Height = 24&
        End With
        .iActiveForeColor = -1&
        .iInActiveForeColor = -1&
        .iInActiveBackColor = -1&
        .iActiveBackColor = -1&
        
        Set .oFontSymbol = New StdFont
        .oFontSymbol.Name = "Marlett"
        
        Set .oBrush = New pcDottedBrush
        .oBrush.Create

        Set .oFont = pGetMenuFont(.oMemDC.hdc)

        pSetFont tMenus
        
        .iInfreqShowDelay = 1500&

        .iFlags = mnuImageProcessBitmap Or mnuShowInfrequent
    
    End With
End Sub

Public Sub PopupMenus_Terminate(ByRef tMenus As tMenus)
    PopupMenus_Clear tMenus
End Sub
'</cPopupMenus Interface>

'<cPopupMenu Interface>
Public Function PopupMenu_Items( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByVal oOwner As cPopupMenus) _
                As cPopupMenuItems
    If pValidatePointer(tMenus, tPointer) Then
        Set PopupMenu_Items = New cPopupMenuItems
        PopupMenu_Items.fInit oOwner, tPointer
    End If
End Function

Public Function PopupMenu_hMenu( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer) _
                As Long
    If pValidatePointer(tMenus, tPointer) Then PopupMenu_hMenu = tMenus.tMenus(tPointer.iIndex).hMenu
End Function

Public Function PopupMenu_Key( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer) _
                As String
    If pValidatePointer(tMenus, tPointer) Then PopupMenu_Key = tMenus.tMenus(tPointer.iIndex).sKey
End Function

Public Function PopupMenu_Index( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer) _
                As Long
    
    If pValidatePointer(tMenus, tPointer) Then
        Dim liIndex As Long
        With tMenus.tRootMenuLookup
            liIndex = tPointer.iIndex
            For PopupMenu_Index = 0& To .iMenuCount - 1&
                If .iMenuIndices(PopupMenu_Index) = liIndex Then
                    PopupMenu_Index = PopupMenu_Index + 1&
                    Exit Function
                End If
            Next
            pErr ccElemNotPartOfCollection, "cPopupMenu.Index"
        End With
    End If
End Function

Public Property Get PopupMenu_Sidebar( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer) _
                As StdPicture
    If pValidatePointer(tMenus, tPointer) Then
        Set PopupMenu_Sidebar = tMenus.tMenus(tPointer.iIndex).oSideBarPic
    End If
End Property

Public Property Set PopupMenu_Sidebar( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByVal oNew As StdPicture)
    If pValidatePointer(tMenus, tPointer) Then
        With tMenus.tMenus(tPointer.iIndex)
            Set .oSideBarPic = oNew
            If oNew Is Nothing Then
                Set .oSideBar = Nothing
            Else
                Set .oSideBar = New pcMemDC
                .oSideBar.CreateFromPicture oNew
            End If
        End With
        pResetForRecalc tMenus, tPointer.iIndex
        pSetWidth tMenus, tPointer.iIndex, Undefined
    End If
End Property

Public Function PopupMenu_Show( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByVal ixPixel As Long, _
            ByVal iyPixel As Long, _
            ByVal iFlags As Long, _
            ByVal bShowAtCursor As Boolean, _
            ByVal hWndShowAt As Long, _
            ByVal hWndClientCoords As Long, _
            ByRef vExclude As Variant, _
            ByVal oOwner As cPopupMenus) _
                As cPopupMenuItem
    
    If pValidatePointer(tMenus, tPointer) Then
        'On Error Resume Next
        Dim tp As tPoint
        Dim tPM As TPMPARAMS
        Dim hMenu As Long
        Dim lR As Long
        Dim tr As tRect
        Dim bTempShowInfrequent As Boolean
        Dim liDblClickTime As Long
        
        Dim liMenuIndex As Long
        Dim liItemIndex As Long
        
        liMenuIndex = Undefined
        liItemIndex = Undefined
        
        hMenu = tMenus.tMenus(tPointer.iIndex).hMenu
        
        'liDblClickTime = GetDoubleClickTime()
        'SetDoubleClickTime 1
        
        If Not (IsWindow(tMenus.hWndOwner) = 0& Or hMenu = 0&) Then
        
            tPM.cbSize = LenB(tPM)
            pGetExclude vExclude, tPM.rcExclude
            
            If IsWindow(hWndClientCoords) <> 0& Then
                tp.x = ixPixel
                tp.y = iyPixel
                MapWindowPoints hWndClientCoords, 0&, tp, 1
                ixPixel = tp.x
                iyPixel = tp.y
            End If
        
            If bShowAtCursor Then
                GetCursorPos tp
                ixPixel = tp.x
                iyPixel = tp.y
            End If
            
            If hWndShowAt Then
                GetWindowRect hWndShowAt, tr
                ixPixel = tr.Left
                iyPixel = tr.Bottom
            End If
            
            If ixPixel < 0 Then ixPixel = 0
            If iyPixel < 0 Then iyPixel = 0
            
            iFlags = (iFlags Or TPM_RETURNCMD)
            
            miInfreqShowDelay = tMenus.tMenuDraw.iInfreqShowDelay
            
            Do
                miRedisplayHierarchyIndex = Undefined
                
                SendMessageLong tMenus.hWndOwner, WM_ENTERMENULOOP, 1&, 0&
                oOwner.fSubclass = True
                
                If liMenuIndex > Undefined Then
                    pGetHierarchy tMenus, liMenuIndex
                    miRedisplayHierarchyItem = liItemIndex
                    oOwner.fCallback 0, 1
                    liMenuIndex = tPointer.iIndex
                End If
                
                lR = TrackPopupMenuEx(tMenus.tMenus(tPointer.iIndex).hMenu, iFlags, ixPixel, iyPixel, tMenus.hWndOwner, tPM)
                oOwner.fSubclass = False
                
                Set PopupMenu_Show = Nothing
                
                ' Find the index of the item with id lR within the menu:
                If lR > 0 Then
                    If pIndexForId(tMenus, lR, liMenuIndex, liItemIndex) Then
                        With tMenus.tMenus(liMenuIndex).tMenuItems(liItemIndex)
                            If Not CBool(.iStyle And mnuChevron) Then
                                Dim tItemPointer As tMenuItemPointer
                                tItemPointer.iId = .iId
                                tItemPointer.iIndex = liMenuIndex
                                tItemPointer.iItemIndex = liItemIndex
                                
                                Set PopupMenu_Show = New cPopupMenuItem
                                PopupMenu_Show.fInit oOwner, tItemPointer
                                
                                If Not CBool(.iStyle And mnuRedisplayOnClick) Then liItemIndex = -1&
                            Else
                                bTempShowInfrequent = True
                                oOwner.ShowInfrequent = True
                            End If
                        End With
                        If Not PopupMenu_Show Is Nothing Then oOwner.fClick PopupMenu_Show
                    Else
                        'got a click return value, but can't find the appropriate menu item from it's ID!
                        Debug.Assert False
                        liItemIndex = Undefined
                        liMenuIndex = tPointer.iIndex
                    End If
                Else
                    'popup dismissed
                    liItemIndex = Undefined
                    liMenuIndex = tPointer.iIndex
                End If
                
                SendMessageLong tMenus.hWndOwner, WM_EXITMENULOOP, 1&, 0&
                
                pRemoveChevrons tMenus, tPointer.iIndex
            
            Loop While liItemIndex > Undefined
            If bTempShowInfrequent Then oOwner.ShowInfrequent = False
            
        End If
    
        'SetDoubleClickTime liDblClickTime
    End If
    
End Function

Public Sub PopupMenu_CallBack(ByVal iType As Long, ByVal iElapsed As Long, ByRef bContinue As Boolean)
    If iType = 0 Then
        miRedisplayHierarchyIndex = miRedisplayHierarchyCount - 1&
        Const VK_LeftCurlyBracket = &HDB&

        For iType = 0& To miRedisplayHierarchyCount
            keybd_event VK_CONTROL, 0, 0, 0
            
            
            keybd_event VK_LeftCurlyBracket, 0, 0, 0
            keybd_event VK_LeftCurlyBracket, 0, KEYEVENTF_KEYUP, 0
            keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
        Next
        
    ElseIf iType = 1 Then
        If Not mbChevronSelected Then Exit Sub
       
        Dim tp As tPoint
        Dim lC As Currency
        GetCursorPos tp
        
        CopyMemory lC, tp, 8&
    
        If MenuItemFromPoint(0&, mhWndChevronMenu, lC) = miChevronIndex Then
            If (iElapsed - miHoverTime) >= miInfreqShowDelay Then
                pClickMouse tp
                Exit Sub
            End If
        Else
            miHoverTime = iElapsed
        End If
        bContinue = True
    End If
End Sub
'</cPopupMenu Interface>

'<cPopupMenuItem Interface>
Public Property Get PopupMenuItem_IconIndex( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As Long
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_IconIndex = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iIconIndex
    End If
End Property
Public Property Let PopupMenuItem_IconIndex( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Long)
    If pValidateItemPointer(tMenus, tPointer) Then
        tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iIconIndex = iNew
    End If
End Property

Public Property Get PopupMenuItem_ShortCutShiftMask( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As Integer
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_ShortCutShiftMask = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iShortcutMask
    End If
End Property
Public Property Let PopupMenuItem_ShortCutShiftMask( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Integer)
    If pValidateItemPointer(tMenus, tPointer) Then
        tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iShortcutMask = iNew
        pSetShortcut tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex)
        Dim liWidth As Long
        With tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex)
            liWidth = .iWidth
            pSetWidth tMenus, tPointer.iIndex, tPointer.iItemIndex
            If .iWidth > liWidth Then pSetWidth tMenus, tPointer.iIndex, tPointer.iItemIndex
        End With
    End If
End Property

Public Property Get PopupMenuItem_ShortCutShiftKey( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As Integer
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_ShortCutShiftKey = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iShortcutShiftKey
    End If
End Property
Public Property Let PopupMenuItem_ShortCutShiftKey( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Integer)
    If pValidateItemPointer(tMenus, tPointer) Then
        tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iShortcutShiftKey = iNew
        pSetShortcut tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex)
        Dim liWidth As Long
        With tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex)
            liWidth = .iWidth
            pSetWidth tMenus, tPointer.iIndex, tPointer.iItemIndex
            If .iAccelerator > liWidth Then pSetWidth tMenus, tPointer.iIndex, tPointer.iItemIndex
        End With
    End If
End Property

Public Property Get PopupMenuItem_Style( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As Long
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_Style = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iStyle
    End If
End Property

Public Property Let PopupMenuItem_Style( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByVal iNew As Long)
    If pValidateItemPointer(tMenus, tPointer) Then
        Dim liOriginalState As Long
        
        'mask out any invalid states
        iNew = iNew And &H3FF&
        
        liOriginalState = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).iStyle
        
        Dim lbResetForRecalc As Boolean
        Dim lbResetMenu As Boolean
        
        'If (liOriginalState And mnuChecked) Xor (iNew And mnuChecked) Then
            
        'End If
        
        'If (liOriginalState And mnuDefault) Xor (iNew And mnuDefault) Then
        
        'End If
        
        If (liOriginalState And mnuDisabled) Xor (iNew And mnuDisabled) Then
            lbResetForRecalc = True
        End If
        
        If (liOriginalState And mnuInfrequent) Xor (iNew And mnuInfrequent) Then
            lbResetForRecalc = True
        End If
        
        If (liOriginalState And mnuInvisible) Xor (iNew And mnuInvisible) Then
            lbResetForRecalc = True
        End If
        
        If (liOriginalState And mnuNewVerticalLine) Xor (iNew And mnuNewVerticalLine) Then
            lbResetMenu = True
        End If
        
        'If (liOriginalState And mnuRadioChecked) Xor (iNew And mnuRadioChecked) Then
            
        'End If
        
        'If (liOriginalState And mnuRedisplayOnClick) Xor (iNew And mnuRedisplayOnClick) Then
        
        'End If
        
        If (liOriginalState And mnuSeparator) Xor (iNew And mnuSeparator) Then
            lbResetForRecalc = True
        End If
        
        With tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex)
            .iStyle = iNew Or (.iStyle And mnuChevron)
        End With
        
        If lbResetMenu Then
            pResetForRecalc tMenus, tPointer.iIndex
        ElseIf lbResetForRecalc Then
            pResetForRecalc tMenus, tPointer.iIndex, tPointer.iItemIndex
        End If
        
    End If
End Property

Public Property Get PopupMenuItem_Caption( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As String
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_Caption = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).sCaption
    End If
End Property
Public Property Let PopupMenuItem_Caption( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByRef sNew As String)
    If pValidateItemPointer(tMenus, tPointer) Then
        pSetMenuCaption tMenus, tPointer.iIndex, tPointer.iItemIndex, sNew, True
    End If
End Property

Public Property Get PopupMenuItem_Index( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As Long
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_Index = tPointer.iItemIndex + 1&
    End If
End Property

Public Property Get PopupMenuItem_Key( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As String
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_Key = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).sKey
    End If
End Property

Public Property Get PopupMenuItem_ShortCutDisplay( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer) _
                As String
    If pValidateItemPointer(tMenus, tPointer) Then
        PopupMenuItem_ShortCutDisplay = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).sShortCutDisplay
    End If
End Property
Public Property Let PopupMenuItem_ShortCutDisplay( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByRef sNew As String)
    If pValidateItemPointer(tMenus, tPointer) Then
        tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).sShortCutDisplay = sNew
    End If
End Property

Public Property Get PopupMenuItem_SubItems( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuItemPointer, _
            ByVal oOwner As cPopupMenus) _
                As cPopupMenuItems
    If pValidateItemPointer(tMenus, tPointer) Then
        Dim tMenuPointer As tMenuPointer
        Dim liChild As Long
        liChild = tMenus.tMenus(tPointer.iIndex).tMenuItems(tPointer.iItemIndex).tChild.iIndex

        If liChild = Undefined Then
            liChild = pGetNewMenuIndex(tMenus)
            With tMenus.tMenus(liChild)
                .iId = pGetNewID()
                .iItemCount = 0&
                .sKey = vbNullString
                LSet .tParent = tPointer
            End With
        End If
        
        tMenuPointer.iId = tMenus.tMenus(liChild).iId
        tMenuPointer.iIndex = liChild
        
        Set PopupMenuItem_SubItems = New cPopupMenuItems
        PopupMenuItem_SubItems.fInit oOwner, tMenuPointer
        
    End If
End Property
'</cPopupMenuItem Interface>

'<cPopupMenuItems Interface>
Public Function PopupMenuItems_Add( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByRef sCaption As String, _
            ByRef sHelpText As String, _
            ByRef sKey As String, _
            ByVal iItemData As Long, _
            ByVal iIconIndex As Long, _
            ByVal iStyle As Long, _
            ByVal iShortcutKey As KeyCodeConstants, _
            ByVal iShortcutMask As ShiftConstants, _
            ByRef vKeyOrIndexInsertBefore As Variant) _
                As Long

    PopupMenuItems_Add = Undefined
    
    Dim liCount As Long
    Dim liLen As Long
    
    If pValidatePointer(tMenus, tPointer) Then
        liCount = tMenus.tMenus(tPointer.iIndex).iItemCount
        If IsMissing(vKeyOrIndexInsertBefore) Then
            PopupMenuItems_Add = liCount
        Else
            PopupMenuItems_Add = pGetIndex(tMenus.tMenus(tPointer.iIndex), vKeyOrIndexInsertBefore)
        End If
        
        liLen = (liCount - PopupMenuItems_Add) * tMenuItemLen
        liCount = liCount + 1&
        
        pResizeMenuItems tMenus.tMenus(tPointer.iIndex).tMenuItems, liCount
        tMenus.tMenus(tPointer.iIndex).iItemCount = liCount
        
        If liLen > 0& Then
            With tMenus.tMenus(tPointer.iIndex)
                CopyMemory .tMenuItems(PopupMenuItems_Add + 1&), .tMenuItems(PopupMenuItems_Add), liLen
            End With
        End If
        
        With tMenus.tMenus(tPointer.iIndex)
            With .tMenuItems(PopupMenuItems_Add)
                With .tChild
                    .iId = 0&
                    .iIndex = Undefined
                End With
                .iHeight = tMenus.tMenuDraw.iItemHeight
                .iIconIndex = iIconIndex
                .iId = pGetNewID()
                .iApiID = .iId
                .iItemData = iItemData
                .iStyle = iStyle
                .sKey = sKey
                
                .iShortcutShiftKey = iShortcutKey
                .iShortcutMask = iShortcutMask
            
                With tMenus.tMenus(tPointer.iIndex)
                    pSetShortcut .tMenuItems(PopupMenuItems_Add)
'            pSetShortcut takes care of .sShortCutDisplay
                    pSetMenuCaption tMenus, tPointer.iIndex, PopupMenuItems_Add, sCaption, False
'            pSetMenuCaption takes care of these elements:
'                .sCaption
'                .iAccelerator
                End With
            End With
            Incr .iControl
        End With

        pAddNewMenuItem tMenus, tPointer.iIndex, PopupMenuItems_Add

        With tMenus.tMenus(tPointer.iIndex)
            pSetItemData .hMenu, .tMenuItems(PopupMenuItems_Add).iId, .tMenuItems(PopupMenuItems_Add).iItemData
        End With
        
        PopupMenuItems_Add = PopupMenuItems_Add + 1&
        
    End If
    
End Function

Public Function PopupMenuItems_Count( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer) _
                As Long
    If pValidatePointer(tMenus, tPointer) Then
        PopupMenuItems_Count = tMenus.tMenus(tPointer.iIndex).iItemCount
    End If
End Function

Public Function PopupMenuItems_Clear( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer) _
                As Long
    If pValidatePointer(tMenus, tPointer) Then
        pRemoveMenu tMenus, tPointer.iIndex
        pDetachMenuFromParent tMenus, tPointer.iIndex
        Incr tMenus.tMenus(tPointer.iIndex).iControl
    End If
End Function

Public Function PopupMenuItems_Exists( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByRef vKeyOrIndex As Variant) _
                As Boolean
    If pValidatePointer(tMenus, tPointer) Then
        On Error GoTo catch
        pGetIndex tMenus.tMenus(tPointer.iIndex), vKeyOrIndex
        PopupMenuItems_Exists = True
    Else
catch:
        Err.Clear
    End If
End Function

Public Function PopupMenuItems_Item( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByRef vKeyOrIndex As Variant, _
            ByVal oOwner As cPopupMenus) _
                As cPopupMenuItem
    
    If pValidatePointer(tMenus, tPointer) Then
        Dim ltItemPointer As tMenuItemPointer
        
        ltItemPointer.iItemIndex = pGetIndex(tMenus.tMenus(tPointer.iIndex), vKeyOrIndex)
        ltItemPointer.iId = tMenus.tMenus(tPointer.iIndex).tMenuItems(ltItemPointer.iItemIndex).iId
        ltItemPointer.iIndex = tPointer.iIndex
        
        Set PopupMenuItems_Item = New cPopupMenuItem
        PopupMenuItems_Item.fInit oOwner, ltItemPointer
    Else
catch:
        Err.Clear
    End If
    
End Function

Public Sub PopupMenuItems_Remove( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByRef vKeyOrIndex As Variant)
    
    Dim liCount As Long
    Dim liLen As Long
    Dim liIndex As Long
    Dim lR As Long
    
    If pValidatePointer(tMenus, tPointer) Then
        liIndex = pGetIndex(tMenus.tMenus(tPointer.iIndex), vKeyOrIndex)
        With tMenus.tMenus(tPointer.iIndex)
            liCount = .iItemCount - 1&
            .iItemCount = liCount
            liLen = (liCount - liIndex) * tMenuLen
            
            lR = RemoveMenu(.hMenu, .tMenuItems(liIndex).iApiID, MF_BYCOMMAND)
            Debug.Assert lR
            pRemoveId .tMenuItems(liIndex).iId
            If liLen Then
                CopyMemory .tMenuItems(liIndex), .tMenuItems(liIndex + 1&), liLen
            End If
            Incr .iControl
        End With
    End If
End Sub

Public Property Get PopupMenuItems_ShowCheckAndIcon(ByRef tMenus As tMenus, ByRef tPointer As tMenuPointer) As Boolean
    If pValidatePointer(tMenus, tPointer) Then
        PopupMenuItems_ShowCheckAndIcon = tMenus.tMenus(tPointer.iIndex).bShowCheckAndIcon
    End If
End Property

Public Property Let PopupMenuItems_ShowCheckAndIcon(ByRef tMenus As tMenus, ByRef tPointer As tMenuPointer, ByVal bNew As Boolean)
    If pValidatePointer(tMenus, tPointer) Then
        If bNew Xor tMenus.tMenus(tPointer.iIndex).bShowCheckAndIcon Then
            tMenus.tMenus(tPointer.iIndex).bShowCheckAndIcon = bNew
            pResetForRecalc tMenus, tPointer.iIndex
        End If
    End If
End Property

Public Property Get PopupMenuItems_RightToLeft(ByRef tMenus As tMenus, ByRef tPointer As tMenuPointer) As Boolean
    If pValidatePointer(tMenus, tPointer) Then
        PopupMenuItems_RightToLeft = tMenus.tMenus(tPointer.iIndex).bRightToLeft
    End If
End Property

Public Property Let PopupMenuItems_RightToLeft(ByRef tMenus As tMenus, ByRef tPointer As tMenuPointer, ByVal bNew As Boolean)
    If pValidatePointer(tMenus, tPointer) Then
        If bNew Xor tMenus.tMenus(tPointer.iIndex).bRightToLeft Then
            tMenus.tMenus(tPointer.iIndex).bRightToLeft = bNew
            'pResetForRecalc tMenus, tPointer.iIndex
        End If
    End If
End Property

Public Property Get PopupMenuItems_Parent(ByRef tMenus As tMenus, ByRef tPointer As tMenuPointer, ByVal oOwner As cPopupMenus) As cPopupMenuItem
    Dim ltItemPointer As tMenuItemPointer
    If pValidatePointer(tMenus, tPointer) Then
        With tMenus.tMenus(tPointer.iIndex).tParent
            If .iId <> 0& Then
                ltItemPointer.iId = .iId
                ltItemPointer.iIndex = .iIndex
                ltItemPointer.iItemIndex = .iItemIndex
                
                Set PopupMenuItems_Parent = New cPopupMenuItem
                
                PopupMenuItems_Parent.fInit oOwner, ltItemPointer
            End If
        End With
    End If
End Property

Public Property Get PopupMenuItems_Root(ByRef tMenus As tMenus, ByRef tPointer As tMenuPointer, ByVal oOwner As cPopupMenus) As cPopupMenu
    Dim ltPointer As tMenuPointer
    If pValidatePointer(tMenus, tPointer) Then
        Dim liIndex As Long
        liIndex = tPointer.iIndex
        
        With tMenus
            Do While .tMenus(liIndex).tParent.iId <> 0
                liIndex = .tMenus(liIndex).tParent.iIndex
            Loop
        
            ltPointer.iId = .tMenus(liIndex).iId
            ltPointer.iIndex = liIndex
        End With
        
        Set PopupMenuItems_Root = New cPopupMenu
        
        PopupMenuItems_Root.fInit oOwner, ltPointer
        
    End If
End Property

Public Property Get PopupMenuItems_NewEnum( _
        ByRef tMenus As tMenus, _
        ByRef tPointer As tMenuPointer, _
        ByVal oOwner As cPopupMenuItems) _
            As IUnknown
    Dim loEnum As cEnumeration
    If pValidatePointer(tMenus, tPointer) Then
        Set loEnum = New cEnumeration
        Set PopupMenuItems_NewEnum = loEnum.GetEnum(oOwner, tMenus.tMenus(tPointer.iIndex).iControl, tPointer.iIndex)
    End If
End Property
'</cPopupMenuItems Interface>

'<Enumeration Stuff>
Public Sub Enum_NextItem( _
            ByRef tMenus As tMenus, _
            ByRef tPointer As tMenuPointer, _
            ByRef tEnum As tEnum, _
            ByRef vItem As Variant, _
            ByRef bNoMore As Boolean, _
            ByVal oOwner As cPopupMenus)
    If pValidatePointer(tMenus, tPointer) Then
        With tMenus.tMenus(tPointer.iIndex)
            If .iControl <> tEnum.iControl Then pErr ccCollectionChangedDuringEnum, "cPopupMenuItems.Enum_NextItem"
        
            Dim liIndex As Long
            liIndex = tEnum.iIndex
            liIndex = liIndex + 1&
            
            If liIndex < .iItemCount Then
                Dim loItem As cPopupMenuItem
                Dim ltPointer As tMenuItemPointer
                With .tMenuItems(liIndex)
                    ltPointer.iId = .iId
                    ltPointer.iIndex = tPointer.iIndex
                    ltPointer.iItemIndex = liIndex
                End With
                Set loItem = New cPopupMenuItem
                loItem.fInit oOwner, ltPointer
                Set vItem = loItem
            Else
                bNoMore = True
            End If
            tEnum.iIndex = liIndex
        End With
        
    End If
End Sub

Public Sub Enum_NextMenu( _
            ByRef tMenus As tMenus, _
            ByRef tEnum As tEnum, _
            ByRef vItem As Variant, _
            ByRef bNoMore As Boolean, _
            ByVal oOwner As cPopupMenus)
    
    Dim loMenu As cPopupMenu
    Dim liIndex As Long
    Dim ltPointer As tMenuPointer
    
    'liIndex = tEnum
    
    If tMenus.iControl <> tEnum.iControl Then pErr ccCollectionChangedDuringEnum, "cPopupMenus.Enum_NextItem"
    
    liIndex = tEnum.iIndex
    liIndex = liIndex + 1&
    With tMenus.tRootMenuLookup
        If liIndex < .iMenuCount Then
            liIndex = .iMenuIndices(liIndex)
            ltPointer.iId = tMenus.tMenus(liIndex).iId
            ltPointer.iIndex = liIndex
            Set loMenu = New cPopupMenu
            loMenu.fInit oOwner, ltPointer
            Set vItem = loMenu
        Else
            bNoMore = True
        End If
    End With
    tEnum.iIndex = liIndex
End Sub
'</Enumeration Stuff>


Public Function Incr(ByRef i As Long)
    If i = &H7FFFFFFF Then
        i = &H80000000
    Else
        i = i + 1&
    End If
End Function
