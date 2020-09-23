Attribute VB_Name = "mPopupDraw"
Option Explicit

'==================================================================================================
'mPopupDraw  - Encapsulates all the fun stuff involved in drawing menu items that look really cool.
'
'Copyright free, use and abuse as you see fit.
'==================================================================================================

'gdi32
    'resource creation
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    'resource destruction
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    'other
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As tPoint) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'user32
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As tRect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStyleProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As tRect, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As tRect, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As tRect, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As tRect, ByVal x As Long, ByVal y As Long) As Long

'comctl32
Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hIml As Long, ByVal hBmp As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32" (ByVal cx As Long, ByVal cy As Long, ByVal fMask As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hIml As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hIml As Long, ByVal i As Long, ByVal diIgnore As Long) As Long

'olepro32
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

    'Private Const DT_BOTTOM = &H8
    Private Const DT_CENTER = &H1
    Private Const DT_LEFT = &H0
    Private Const DT_CALCRECT = &H400
    'Private Const DT_WORDBREAK = &H10
    Private Const DT_VCENTER = &H4
    'Private Const DT_TOP = &H0
    'Private Const DT_TABSTOP = &H80
    Private Const DT_SINGLELINE = &H20
    Private Const DT_RIGHT = &H2
    'Private Const DT_NOCLIP = &H100
    'Private Const DT_INTERNAL = &H1000
    'Private Const DT_EXTERNALLEADING = &H200
    'Private Const DT_EXPANDTABS = &H40
    'Private Const DT_CHARSTREAM = 4
    'Private Const DT_EDITCONTROL = &H2000&
    'Private Const DT_PATH_ELLIPSIS = &H4000&
    'Private Const DT_END_ELLIPSIS = &H8000&
    'Private Const DT_MODIFYSTRING = &H10000
    'Private Const DT_RTLREADING = &H20000
    'Private Const DT_WORD_ELLIPSIS = &H40000

    Private Const BITSPIXEL = 12
    Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

    Private Const CLR_INVALID = -1
    Private Const OPAQUE = 2
    Private Const TRANSPARENT = 1
    
    'Private Const PS_DASH = 1
    'Private Const PS_DASHDOT = 3
    'Private Const PS_DASHDOTDOT = 4
    'Private Const PS_DOT = 2
    Private Const PS_SOLID = 0
    'Private Const PS_NULL = 5
    
    Private Const BF_LEFT = &H1
    Private Const BF_BOTTOM = &H8
    Private Const BF_RIGHT = &H4
    Private Const BF_TOP = &H2
    Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'    Private Const BDR_INNER = &HC
'    Private Const BDR_OUTER = &H3
'    Private Const BDR_RAISED = &H5
    Private Const BDR_RAISEDINNER = &H4
    Private Const BDR_RAISEDOUTER = &H1
'    Private Const BDR_SUNKEN = &HA
'    Private Const BDR_SUNKENINNER = &H8
    Private Const BDR_SUNKENOUTER = &H2
'    Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
'    Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
'    Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)


    Private Const SRCCOPY = &HCC0020
'    Private Const SRCINVERT = &H660046
'    Private Const BLACKNESS = &H42
'    Private Const WHITENESS = &HFF0062
'    Private Const SRCAND = &H8800C6
'    Private Const SRCERASE = &H440328
'    Private Const SRCPAINT = &HEE0086
    
'Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
'Private Const ILD_FOCUS = 4
'Private Const ILD_MASK = &H10&
'Private Const ILD_IMAGE = &H20&
'Private Const ILD_ROP = &H40&
'Private Const ILD_OVERLAYMASK = 3840



'/* Image type */
'Private Const DST_COMPLEX = &H0
'Private Const DST_TEXT = &H1
'Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
'Private Const DST_BITMAP = &H4
'
'Private Const DSS_NORMAL = &H0
'Private Const DSS_UNION = &H10         ' /* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
'Private Const DSS_RIGHT = &H8000
    


Private Const ILC_MASK = 1&
'Private Const ILC_COLOR = 0&
'Private Const ILC_COLORDDB = &HFE&
'Private Const ILC_COLOR4 = &H4&
'Private Const ILC_COLOR8 = &H8&
'Private Const ILC_COLOR16 = &H10&
'Private Const ILC_COLOR24 = &H18&
Private Const ILC_COLOR32 = &H20&
'Private Const ILC_PALETTE = &H800&

'DrawItem declares:
'modular because we don't want to allocate all these variables over and over
'again for each item drawn only to have to pass them to all the private procedures.
Private mbRadioCheck        As Boolean
Private mbDisabled          As Boolean
Private mbChecked           As Boolean
Private mbHighlighted       As Boolean
Private mbHeader            As Boolean
Private mbSeparator         As Boolean
Private mbDefault           As Boolean
Private mbInfrequent        As Boolean
Private mbChevron           As Boolean
Private mbOfficeXPStyle     As Boolean
Private mtRectItem          As tRect
Private mtRectLeft          As tRect
Private mtRectRight         As tRect
Private mtRectIcon1         As tRect
Private mtRectIcon2         As tRect
Private mtRectCaption       As tRect
Private mtRectSideBar       As tRect
Private mtRectTemp          As tRect
Private miState             As ePopupMenuItemStyle
Private miFlags             As ePopupMenuDrawStyle
Private moItem              As cPopupMenuItem
Private mhPen               As Long
Private mhPenOld            As Long
Private mhFont              As Long
Private mhFontOld           As Long
Private mhBrush             As Long
Private miYOffset           As Long
Private miIconIndex         As Long
Private miWidth             As Long
Private mhDc                As Long
Private miActiveForeColor   As Long
Private miInActiveForeColor As Long
Private miActiveBackColor   As Long
Private miInActiveBackColor As Long
Private mtJunk              As tPoint

Public Sub Draw_DrawItem(ByRef tDIS As DRAWITEMSTRUCT, ByRef tDraw As tMenuDraw, ByRef tMenus As tMenus, ByRef tMenu As tMenu, ByRef tItem As tMenuItem, ByVal iMenuIndex As Long, ByVal iItemIndex As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal oOwner As cPopupMenus)
    
    miState = tItem.iStyle
    miIconIndex = tItem.iIconIndex
    
    If Not (CBool(miState And mnuInvisible) Or (CBool(miState And mnuInfrequent) And Not CBool(tDraw.iFlags And mnuShowInfrequent))) Then
        miFlags = tDraw.iFlags
        ' Get info about the menu item:
        With tItem
            mbRadioCheck = CBool(miState And mnuRadioChecked)
            mbDisabled = CBool(miState And mnuDisabled)
            mbChecked = CBool(miState And mnuChecked) Or mbRadioCheck
            mbHighlighted = CBool(tDIS.itemState And 1&) And Not mbDisabled
            
            mbSeparator = CBool(miState And mnuSeparator)
            mbHeader = mbSeparator And CBool(tDraw.iFlags And mnuDrawSeparatorsAsHeaders)
            
            mbDefault = CBool(miState And mnuDefault)
            mbInfrequent = CBool(miState And mnuInfrequent)
            mbChevron = CBool(tItem.iStyle And mnuChevron)
            mbOfficeXPStyle = CBool(tMenus.tMenuDraw.iFlags And mnuOfficeXPStyle)
        End With
        
        With tDraw
            miActiveBackColor = .iActiveBackColor
            miActiveForeColor = .iActiveForeColor
            miInActiveForeColor = .iInActiveForeColor
            miInActiveBackColor = .iInActiveBackColor
        End With
        
        If miActiveBackColor = -1& Then miActiveBackColor = vbHighlight
        If miActiveForeColor = -1& Then miActiveForeColor = vbHighlightText
        If miInActiveBackColor = -1& Then miInActiveBackColor = IIf(mbOfficeXPStyle, vbWindowBackground, vbMenuBar)
        If miInActiveForeColor = -1& Then miInActiveForeColor = vbMenuText

        With tDIS.rcItem
            mtRectItem.Top = 0&
            mtRectItem.Left = .Left
            mtRectItem.Right = .Right - .Left
            mtRectItem.Bottom = .Bottom - .Top
            miYOffset = .Top
        
            With tDraw.oMemDC
            ' ensure the memory dc is big enough:
                .Width = miWidth
                .Height = mtRectItem.Bottom - mtRectItem.Top
                mhDc = .hdc
            End With
        End With

        pGetRects tDraw, tMenu, tItem, mtRectItem, mtRectSideBar, mtRectLeft, mtRectRight, mtRectIcon1, mtRectIcon2, mtRectCaption, tMenu.bRightToLeft
        
        miWidth = mtRectItem.Right - mtRectItem.Left
        
        pDrawSidebar tDraw, tMenu, iItemIndex
        
        pDrawBackground tDraw, tMenu, iItemIndex
        
        If mbChevron Then
            pDrawChevron tDraw
        ElseIf mbSeparator Then
            pDrawSeparator tDraw, tItem
        Else
            pDrawIcon1 tDraw, tMenu
            
            If tMenu.bShowCheckAndIcon Then pDrawIcon tDraw, mtRectIcon2
            
            pDrawCaption tDraw, tItem
        End If
    
        'done and done
        With tDIS.rcItem
            BitBlt tDIS.hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, mhDc, 0&, 0&, vbSrcCopy
        End With
    
    End If
    
End Sub

Public Sub Draw_HLSforRGB( _
            ByVal r As Long, _
            ByVal g As Long, _
            ByVal b As Long, _
            ByRef h As Single, _
            ByRef s As Single, _
            ByRef l As Single)
            
 Dim Max As Single
 Dim Min As Single
 Dim delta As Single
 Dim rR As Single, rG As Single, rB As Single

     rR = r / 255: rG = g / 255: rB = b / 255

 '{Given: rgb each in [0,1].
 ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
         Max = pMax(rR, rG, rB)
         Min = pMin(rR, rG, rB)
             l = (Max + Min) / 2 '{This is the lightness}
         '{Next calculate saturation}
         If Max = Min Then
             'begin {Acrhomatic case}
             s = 0
             h = 0
             'end {Acrhomatic case}
         Else
             'begin {Chromatic case}
                 '{First calculate the saturation.}
             If l <= 0.5 Then
                 s = (Max - Min) / (Max + Min)
             Else
                 s = (Max - Min) / (2 - Max - Min)
             End If
             '{Next calculate the hue.}
             delta = Max - Min
             If rR = Max Then
                     h = (rG - rB) / delta '{Resulting color is between yellow and magenta}
             ElseIf rG = Max Then
                 h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
             ElseIf rB = Max Then
                 h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
             End If
         'end {Chromatic Case}
     End If
End Sub

Public Sub Draw_RGBforHLS( _
            ByVal h As Single, _
            ByVal s As Single, _
            ByVal l As Single, _
            ByRef r As Long, _
            ByRef g As Long, _
            ByRef b As Long)
            
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single

    If s = 0 Then
    ' Achromatic case:
    rR = l: rG = l: rB = l
    Else
    ' Chromatic case:
    ' delta = Max-Min
    If l <= 0.5 Then
        's = (Max - Min) / (Max + Min)
        ' Get Min value:
        Min = l * (1 - s)
    Else
        's = (Max - Min) / (2 - Max - Min)
        ' Get Min value:
        Min = l - s * (1 - l)
    End If
    ' Get the Max value:
    Max = 2 * l - Min
    
    ' Now depending on sector we can evaluate the h,l,s:
    If (h < 1) Then
        rR = Max
        If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
        Else
            rB = Min
            rG = h * (Max - Min) + rB
        End If
    ElseIf (h < 3) Then
        rG = Max
        If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
        Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
        End If
    Else
        rB = Max
        If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
        Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
        End If
        
    End If
            
    End If
    r = rR * 255: g = rG * 255: b = rB * 255
End Sub

Public Sub Draw_GetTextExtent(ByVal hdc As Long, ByRef sText As String, Optional ByRef iWidth As Long, Optional ByRef iHeight As Long)

    Dim tr As tRect
    
    If LenB(sText) Then
    
        DrawText hdc, sText, -1, tr, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
    
        iHeight = tr.Bottom - tr.Top
        iWidth = tr.Right - tr.Left
    Else
        iHeight = 0
        iWidth = 0
    End If

End Sub

Public Sub Draw_GetItemExtent(ByRef tMenus As tMenus, ByVal iMenuIndex As Long, ByVal iItemIndex As Long)
    Dim liWidth As Long
    Dim hdc As Long
    Dim liWidthCaption As Long
    Dim liWidthShortcut As Long
    Dim liIconSize As Long
    Dim hFontOld As Long
    
    
    hdc = tMenus.tMenuDraw.oMemDC.hdc
    
    hFontOld = SelectObject(hdc, pGetFontHandle(tMenus.tMenuDraw.oFont, CBool(tMenus.tMenus(iMenuIndex).tMenuItems(iItemIndex).iStyle And mnuDefault)))
    
    With tMenus.tMenus(iMenuIndex).tMenuItems(iItemIndex)
        Draw_GetTextExtent hdc, .sCaption, liWidthCaption
        Draw_GetTextExtent hdc, .sShortCutDisplay, liWidthShortcut
    End With
    
    SelectObject hdc, hFontOld
    
    liWidth = liWidthCaption + liWidthShortcut + 14&

    
    liIconSize = pGetIconSize(tMenus.tMenuDraw)
        
    If tMenus.tMenus(iMenuIndex).bShowCheckAndIcon Then
        liWidth = liWidth + liIconSize + liIconSize + 12&
    Else
        liWidth = liWidth + liIconSize + 8&
    End If
    
    
    With tMenus.tMenus(iMenuIndex)
        If Not .oSideBar Is Nothing Then liWidth = liWidth + .oSideBar.Width
        With .tMenuItems(iItemIndex)
            .iWidth = liWidth + 6&
        End With
        pCalcHeight tMenus, iMenuIndex, iItemIndex, .tMenuItems(iItemIndex)
    End With

End Sub

Private Sub pCalcHeight(ByRef tMenus As tMenus, ByVal iMenuIndex As Long, ByVal iItemIndex As Long, ByRef tItem As tMenuItem)
   
    Dim liState As ePopupMenuItemStyle
    liState = tItem.iStyle
    
    If CBool(liState And mnuInvisible) Or (CBool(liState And mnuInfrequent) And Not CBool(tMenus.tMenuDraw.iFlags And mnuShowInfrequent)) Then
        tItem.iHeight = 0
    ElseIf (liState And mnuSeparator) Then
        If Len(tItem.sCaption) = 0& Then
            If CBool(tMenus.tMenuDraw.iFlags And mnuOfficeXPStyle) Then
                tItem.iHeight = 3
            Else
                tItem.iHeight = 8
            End If
        Else
            Dim lhFont As Long
            Dim liHeight As Long
            lhFont = SelectObject(tMenus.tMenuDraw.oMemDC.hdc, pGetFontHandle(tMenus.tMenuDraw.oFont, True))
            
            Draw_GetTextExtent tMenus.tMenuDraw.oMemDC.hdc, tItem.sCaption, , liHeight
            tItem.iHeight = liHeight + 4
            
            SelectObject tMenus.tMenuDraw.oMemDC.hdc, lhFont
            
        End If
    ElseIf CBool(tItem.iStyle And mnuChevron) Then
        tItem.iHeight = 16
    Else
        tItem.iHeight = tMenus.tMenuDraw.iItemHeight
    End If
End Sub

Private Sub pDrawBackground(ByRef tDraw As tMenuDraw, ByRef tMenu As tMenu, ByVal iItemIndex As Long)
    'Debug.Assert Not (mbInfrequent And mbHighlighted)
    If Not mbInfrequent Or mbHighlighted Or mbHeader Then
        If Not mbHighlighted Or (mbHeader Or mbSeparator Or mbDisabled) Then
            If mbHeader Then
                pFillWithLighterSelectedColor tDraw, mhDc, mtRectItem, miYOffset
            ElseIf (mbOfficeXPStyle) Then
                pFillWithLighterControlColor tDraw, mhDc, mtRectLeft, miYOffset
                pFillWithNormalBackground tDraw, mhDc, mtRectRight, miYOffset
            Else
                pFillWithNormalBackground tDraw, mhDc, mtRectItem, miYOffset
            End If
        ElseIf Not mbInfrequent Or mbHighlighted Then
            If CBool(tDraw.iFlags And mnuGradientHighlight) Then
                If Not mbOfficeXPStyle Then
                    pDrawGradient mhDc, mtRectItem, pTranslateColor(miInActiveBackColor), pTranslateColor(miActiveBackColor), False
                    If CBool(tDraw.iFlags And mnuButtonHighlightStyle) Then pDrawEdge mhDc, mtRectItem, EDGE_RAISED, BF_RECT, mbOfficeXPStyle
                Else
                    pDrawGradient mhDc, mtRectItem, pTranslateColor(miInActiveBackColor), pTranslateColor(miActiveBackColor), False
                    pDrawEdge mhDc, mtRectItem, 0, 0, True
                End If
            ElseIf mbOfficeXPStyle Then
                pFillWithLighterSelectedColor tDraw, mhDc, mtRectItem, miYOffset
                pDrawEdge mhDc, mtRectItem, 0, 0, True
            Else
                If tDraw.oPic Is Nothing Then
                    If mbChevron Then
                        pFillWithNormalBackground tDraw, mhDc, mtRectItem, miYOffset
                    Else
                        If mbInfrequent Then
                            pFillWithLighterBackColor tDraw, mhDc, mtRectLeft, miYOffset, True
                        Else
                            pFillWithNormalBackground tDraw, mhDc, mtRectLeft, miYOffset
                        End If
                        pFillWithHighlightBackColor tDraw, mhDc, mtRectRight, miYOffset
                    End If
                Else
                    pFillWithHighlightBackColor tDraw, mhDc, mtRectItem, miYOffset
                End If
                If CBool(tDraw.iFlags And mnuButtonHighlightStyle) Or (tDraw.oPic Is Nothing And mbChevron) Then pDrawEdge mhDc, mtRectItem, EDGE_RAISED, BF_RECT, mbOfficeXPStyle
            End If
        End If
    Else
        If mbOfficeXPStyle Then
            pFillWithLighterControlColor tDraw, mhDc, mtRectLeft, miYOffset
            pFillWithLighterBackColor tDraw, mhDc, mtRectRight, miYOffset, True
        Else
            pFillWithLighterBackColor tDraw, mhDc, mtRectItem, miYOffset, True
        End If
    End If
    
    If Not (mbHighlighted And (CBool(tDraw.iFlags And mnuOfficeXPStyle) Or (tDraw.iFlags And mnuButtonHighlightStyle))) And CBool(tDraw.iFlags And mnuShowInfrequent) Then
        Dim lbNext As Boolean, lbPrior As Boolean
        pGetInfrequentStates tMenu, iItemIndex, lbPrior, lbNext, mbInfrequent
        If (mbInfrequent Xor lbPrior) Then
            mhPen = CreatePen(PS_SOLID, 1&, pTranslateColor(IIf(lbPrior, vbWhite, pSlightlyLighterColor(vbBlack))))
            mhPenOld = SelectObject(mhDc, mhPen)
            MoveToEx mhDc, mtRectItem.Left, mtRectItem.Top, mtJunk
            LineTo mhDc, mtRectItem.Right, mtRectItem.Top
            SelectObject mhDc, mhPenOld
            DeleteObject mhPen
        End If
        If (mbInfrequent Xor lbNext) Then
            mhPen = CreatePen(PS_SOLID, 1&, IIf(lbNext, pTranslateColor(vb3DShadow), pBlendColor(pTranslateColor(miInActiveBackColor), pTranslateColor(vb3DShadow))))
            mhPenOld = SelectObject(mhDc, mhPen)
            MoveToEx mhDc, mtRectItem.Left, mtRectItem.Bottom - 1&, mtJunk
            LineTo mhDc, mtRectItem.Right, mtRectItem.Bottom - 1&
            SelectObject mhDc, mhPenOld
            DeleteObject mhPen
        End If
    End If
End Sub

Private Sub pDrawSeparator(ByRef tDraw As tMenuDraw, tItem As tMenuItem)
    Dim ltTemp As tRect
    
    LSet ltTemp = mtRectItem
    
    With ltTemp
        .Top = (.Bottom - .Top - 2) \ 2 + .Top
        .Bottom = .Top + 2
    End With
    
    InflateRect ltTemp, -12, 0
    
    If (mbOfficeXPStyle) Then
        With ltTemp
            .Left = mtRectLeft.Right + 4
            .Right = .Right + 20
            .Top = .Top + 1
            .Bottom = .Top
        End With
    End If
    
    If Len(tItem.sCaption) = 0& Then
        pDrawEdge mhDc, ltTemp, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
    Else
        mhFontOld = SelectObject(mhDc, pGetFontHandle(tDraw.oFont, True))
        
        If mbHeader Then
            DrawText mhDc, tItem.sCaption, -1, mtRectItem, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
        Else
            Dim ltRectDraw As tRect
            Dim ltRectSep  As tRect
            Dim liWidth As Long
            
            DrawText mhDc, tItem.sCaption, -1, ltRectDraw, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
            
            liWidth = ltRectDraw.Right - ltRectDraw.Left
            
            If mbOfficeXPStyle Then
                LSet ltRectDraw = mtRectCaption
            Else
                LSet ltRectDraw = mtRectItem
            End If
            
            If (miWidth \ 2) - (liWidth \ 2) > ltRectDraw.Left Then ltRectDraw.Left = (miWidth \ 2) - (liWidth \ 2)
            ltRectDraw.Right = ltRectDraw.Left + liWidth
            
            LSet ltRectSep = ltTemp
            ltRectSep.Right = ltRectDraw.Left - 2

            If ltRectSep.Right > ltRectSep.Left Then pDrawEdge mhDc, ltRectSep, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
            
            ltRectSep.Right = mtRectItem.Right - 12
            ltRectSep.Left = ltRectDraw.Right + 2
            
            If ltRectSep.Right > ltRectSep.Left Then pDrawEdge mhDc, ltRectSep, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
            SetTextColor mhDc, pTranslateColor(miInActiveForeColor)
            pDrawText mhDc, tItem.sCaption, ltRectDraw, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE, mbDisabled, mbOfficeXPStyle
           
        End If
        SelectObject mhDc, mhFontOld
    End If
    

End Sub

Private Sub pDrawChevron(ByRef tDraw As tMenuDraw)
    Dim ltTemp As tRect
    LSet ltTemp = mtRectItem
    With ltTemp
        .Top = .Bottom - 14
        .Right = .Right - 2
        .Bottom = .Bottom - 1
    End With
    
    mhPen = CreatePen(PS_SOLID, 1, pTranslateColor(miInActiveForeColor))
    mhPenOld = SelectObject(mhDc, mhPen)

    ltTemp.Left = ((ltTemp.Right - ltTemp.Left) \ 2) - 3 + ltTemp.Left
    ltTemp.Top = ltTemp.Top + 2
    
    MoveToEx mhDc, ltTemp.Left, ltTemp.Top, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3
    MoveToEx mhDc, ltTemp.Left, ltTemp.Top + 1, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3 + 1
    
    MoveToEx mhDc, ltTemp.Left, ltTemp.Top + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3 + 4
    MoveToEx mhDc, ltTemp.Left, ltTemp.Top + 1 + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3 + 1 + 4
    
    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3
    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top + 1, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3 + 1
    
    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3 + 4
    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top + 1 + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3 + 1 + 4
    
    
    SelectObject mhDc, mhPenOld
    DeleteObject mhPen

End Sub

Private Sub pDrawCaption(ByRef tDraw As tMenuDraw, tItem As tMenuItem)
    
    SetBkMode mhDc, TRANSPARENT
    If mbHighlighted _
            Then SetTextColor mhDc, pTranslateColor(miActiveForeColor) _
            Else SetTextColor mhDc, miInActiveForeColor
            
            
        'And (Not ((CBool(tDraw.iFlags And mnuButtonHighlightStyle) And mbOfficeXPStyle)) _
        'Or CBool(tDraw.iFlags And mnuGradientHighlight)) _

    mhFont = pGetFontHandle(tDraw.oFont, mbDefault)

    mhFontOld = SelectObject(mhDc, mhFont)

    If Len(tItem.sCaption) Then
        pDrawText mhDc, tItem.sCaption, mtRectCaption, DT_SINGLELINE Or DT_VCENTER Or DT_LEFT, mbDisabled, mbOfficeXPStyle
    End If
    If Len(tItem.sShortCutDisplay) And tItem.tChild.iId = 0 Then
        'mtRectCaption.Right = mtRectCaption.Right - 10
        pDrawText mhDc, tItem.sShortCutDisplay, mtRectCaption, DT_SINGLELINE Or DT_VCENTER Or DT_RIGHT, mbDisabled, mbOfficeXPStyle
        'mtRectCaption.Right = mtRectCaption.Right + 10
    End If
    SelectObject mhDc, mhFontOld
    
End Sub

Private Sub pDrawIcon1(ByRef tDraw As tMenuDraw, ByRef tMenu As tMenu)
                  ' Check:
    If mbChecked Then
        ' Color in:
        If Not (mbHighlighted Or mbDisabled) Then
            If Not CBool(miFlags And mnuButtonHighlightStyle) Then
                If (mbOfficeXPStyle) Then
                    pFillWithLighterControlColor tDraw, mhDc, mtRectIcon1, miYOffset + mtRectIcon1.Top
                Else
                    pFillWithLighterBackColor tDraw, mhDc, mtRectIcon1, miYOffset + mtRectIcon1.Top, False
                End If
            End If
        ElseIf Not mbDisabled Then
            If Not mbOfficeXPStyle Then pFillWithNormalBackground tDraw, mhDc, mtRectIcon1, miYOffset + mtRectIcon1.Top
        End If

        pDrawEdge mhDc, mtRectIcon1, BDR_SUNKENOUTER, BF_RECT, mbOfficeXPStyle
    
        'If tMenu.bShowCheckAndIcon Then
           ' Draw the appropriate symbol:
           mhFontOld = SelectObject(mhDc, pGetFontHandle(tDraw.oFontSymbol, False))
           SetTextColor mhDc, miInActiveForeColor
           pDrawText mhDc, IIf(mbRadioCheck, "h", "b"), mtRectIcon1, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE, mbDisabled, False
           SelectObject mhDc, mhFontOld
        'End If
    ElseIf Not tMenu.bShowCheckAndIcon Then
        pDrawIcon tDraw, mtRectIcon1
    End If
        
End Sub

Private Sub pDrawSidebar(ByRef tDraw As tMenuDraw, ByRef tMenu As tMenu, ByVal iIndex As Long)
    If mtRectSideBar.Right > mtRectSideBar.Left Then
        If Not tMenu.oSideBar Is Nothing Then
            Dim liYOffset As Long
            For iIndex = iIndex To tMenu.iItemCount - 1
                liYOffset = liYOffset + tMenu.tMenuItems(iIndex).iHeight
            Next
            With mtRectSideBar
                BitBlt mhDc, .Left, .Top, .Right - .Left, .Bottom - .Top, tMenu.oSideBar.hdc, 0, tMenu.oSideBar.Height - liYOffset, vbSrcCopy
            End With
        End If
    End If
End Sub

Private Sub pDrawIcon(ByRef tDraw As tMenuDraw, ByRef tRect As tRect)
   If miIconIndex > -1& Then
        If mbDisabled Then
            pImageListDrawIconDisabled tDraw.oVB6ImageList, mhDc, tDraw.hIml, miIconIndex, tRect.Left, tRect.Top, tDraw.iIconSize
        Else
            If (mbHighlighted And mbOfficeXPStyle) Then
                pImageListDrawIconDisabled tDraw.oVB6ImageList, mhDc, tDraw.hIml, miIconIndex, tRect.Left + 1, tRect.Top + 1, tDraw.iIconSize, True
                pImageListDrawIcon tDraw.oVB6ImageList, mhDc, tDraw.hIml, miIconIndex, tRect.Left - 1, tRect.Top - 1
            Else
                pImageListDrawIcon tDraw.oVB6ImageList, mhDc, tDraw.hIml, miIconIndex, tRect.Left, tRect.Top
            End If
        End If
        If mbHighlighted And Not (mbOfficeXPStyle Or mbDisabled Or CBool(tDraw.iFlags And (mnuButtonHighlightStyle Or mnuGradientHighlight))) Then
            Dim ltTemp As tRect
            LSet ltTemp = tRect
            InflateRect ltTemp, 2, 2
            pDrawEdge mhDc, ltTemp, BDR_RAISEDINNER, BF_RECT, False
        End If
    End If
End Sub

Private Function pGetIconSize(ByRef tDraw As tMenuDraw) As Long
    pGetIconSize = tDraw.iIconSize
    If pGetIconSize <= 0& Then pGetIconSize = xSystemMetric(SM_CXMENUCHECK)
    If pGetIconSize > tDraw.iItemHeight - 4& Then pGetIconSize = tDraw.iItemHeight - 4&
End Function


Private Sub pGetRects(ByRef tDraw As tMenuDraw, ByRef tMenu As tMenu, ByRef tItem As tMenuItem, ByRef tRectItem As tRect, ByRef tSideBar As tRect, ByRef tLeft As tRect, ByRef tRight As tRect, ByRef tIcon1 As tRect, ByRef tIcon2 As tRect, ByRef tCaption As tRect, Optional ByVal bRightToLeft As Boolean)
    'IN: tRectItem
    'OUT: tRectItem and all other rects
    Const SPACING As Long = 4
    
    Dim liNumIcons As Long
    Dim liIconSize As Long
    Dim liWidth As Long
    Dim liHeight As Long
   
    'store the width and height that the rectangles are to fill.
    liWidth = tRectItem.Right - tRectItem.Left
    liHeight = tRectItem.Bottom - tRectItem.Top
    
    liIconSize = pGetIconSize(tDraw)
    'result is 1 or 2 items.
    liNumIcons = 1& + Abs(tMenu.bShowCheckAndIcon)
    
    'Side bar is always on the side!
    LSet tSideBar = tRectItem
    'if there is a picture to put on the sidebar then add the width
    If Not tMenu.oSideBar Is Nothing Then
        tSideBar.Right = tMenu.oSideBar.Width
        liWidth = liWidth + tMenu.oSideBar.Width
    Else
        tSideBar.Right = tSideBar.Left
    End If

    'the item (background, icons, caption) will be drawn next to the caption
    tRectItem.Left = tSideBar.Right
    
    'start out with the left and right sections equal to the whole item
    LSet tLeft = tRectItem
    LSet tRight = tRectItem
    
    'the left item extends only to cover the icons
    tLeft.Right = tLeft.Left + SPACING + (liIconSize + SPACING) * liNumIcons

    'the right item takes everything else
    tRight.Left = tLeft.Right
    
    'the left icon will be a square with sides = liiconsize, and will be centered
    'in the available height.  the left edge will be SPACING pixels to the right of the whole item
    With tIcon1
        .Top = tLeft.Top + ((tLeft.Bottom - tLeft.Top) \ 2&) - (liIconSize \ 2&)
        .Bottom = .Top + liIconSize
        .Left = SPACING + tLeft.Left
        .Right = .Left + liIconSize
    End With
    
    'if check and icon
    If liNumIcons = 2& Then
        'the next icon will also be centered vertically, but will be SPACING pixels to the right of the
        'first icon
        With tIcon2
            .Top = tLeft.Top + ((tLeft.Bottom - tLeft.Top) \ 2&) - (liIconSize \ 2&)
            .Bottom = .Top + liIconSize
            .Left = SPACING + tIcon1.Right
            .Right = .Left + liIconSize
        End With
    Else
        'no space for this icon
        tIcon2.Left = tIcon1.Right
        tIcon2.Right = tIcon1.Right
    End If
    
    'the caption will begin at the right item
    LSet tCaption = tRectItem
    'indent the text eight pixels
    tCaption.Left = tRight.Left + SPACING
    'shortcut captions are right aligned with a larger space from the right edge
    tCaption.Right = tCaption.Right - SPACING - SPACING
    
    If bRightToLeft Then
        liWidth = (tRectItem.Right - tSideBar.Left) \ 2&
        pMirrorRect tRectItem, liWidth
        pMirrorRect tSideBar, liWidth
        pMirrorRect tLeft, liWidth
        pMirrorRect tRight, liWidth
        pMirrorRect tIcon1, liWidth
        pMirrorRect tIcon2, liWidth
        pMirrorRect tCaption, liWidth
    End If
End Sub

Private Sub pMirrorRect(ByRef tRect As tRect, ByVal iMiddle As Long)
    Dim liTemp As Long
    With tRect
        liTemp = iMiddle + (iMiddle - .Left)
        .Left = iMiddle + (iMiddle - .Right)
        .Right = liTemp
    End With
End Sub

Private Function pMax(rR As Single, rG As Single, rB As Single) As Single
    If (rR > rG) Then
        If (rR > rB) Then
            pMax = rR
        Else
            pMax = rB
        End If
    Else
        If (rB > rG) Then
            pMax = rB
        Else
            pMax = rG
        End If
    End If
End Function
Private Function pMin(rR As Single, rG As Single, rB As Single) As Single
    If (rR < rG) Then
        If (rR < rB) Then
            pMin = rR
        Else
            pMin = rB
        End If
    Else
        If (rB < rG) Then
            pMin = rB
        Else
            pMin = rG
        End If
    End If
End Function

Private Function pGetFontHandle(ByVal oFont As IFont, ByVal bBold As Boolean) As Long
    Dim loFont As IFont
    
    If bBold Then
        oFont.Clone loFont
        loFont.Bold = True
        pGetFontHandle = loFont.hFont
    Else
        pGetFontHandle = oFont.hFont
    End If
End Function

Private Sub pGetInfrequentStates(ByRef tMenu As tMenu, ByVal iItemIndex As Long, ByRef bPrev As Boolean, ByRef bNext As Boolean, Optional ByVal bDefault As Boolean = True)
    Dim liNext As Long
    Dim liPrev As Long
    With tMenu
        'Debug.Assert iItemIndex <> 5
        For liNext = iItemIndex + 1& To .iItemCount - 1&
            If Not CBool(tMenu.tMenuItems(liNext).iStyle And mnuInvisible) Then Exit For
        Next
        
        For liPrev = iItemIndex - 1& To 0& Step -1&
            If Not CBool(tMenu.tMenuItems(liNext).iStyle And mnuInvisible) Then Exit For
        Next
        
        If liNext < .iItemCount Then
            bNext = (.tMenuItems(liNext).iStyle And mnuInfrequent)
        Else
            bNext = bDefault
        End If
        
        If liPrev < .iItemCount And liPrev > -1& Then
            bPrev = (.tMenuItems(liPrev).iStyle And mnuInfrequent)
        Else
            bPrev = bDefault
        End If
    End With
End Sub

Private Sub pFillWithLighterBackColor(ByRef tDraw As tMenuDraw, ByVal lHDC As Long, tr As tRect, ByVal lOffsetY As Long, ByVal bInfrequent As Boolean)
Dim hBrush As Long
   SetBkMode lHDC, OPAQUE
   If Not tDraw.oBitmapLight Is Nothing Then
      pTileArea lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, tDraw.oBitmapLight.hdc, tDraw.oBitmapLight.Width, tDraw.oBitmapLight.Height, lOffsetY
   Else
      If (pNoPalette) Then
         If bInfrequent Then
            If mbOfficeXPStyle Then
                hBrush = CreateSolidBrush(pSlightlyLighterColor(miInActiveBackColor))
            Else
                hBrush = CreateSolidBrush(pSlightlyLighterColor(pSlightlyLighterColor(miInActiveBackColor)))
            End If
         Else
            hBrush = CreateSolidBrush(pLighterColor(miInActiveBackColor))
         End If
         FillRect lHDC, tr, hBrush
         DeleteObject hBrush
      Else
         tDraw.oBrush.Rectangle lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, 1, PATCOPY, True, miActiveBackColor, vb3DHighlight
      End If
   End If
   SetBkMode lHDC, TRANSPARENT
End Sub
Private Sub pFillWithHighlightBackColor(ByRef tDraw As tMenuDraw, ByVal lHDC As Long, tr As tRect, ByVal lOffsetY As Long)
Dim hBr As Long
   If tDraw.oBitmapSuperLight Is Nothing Then
      hBr = CreateSolidBrush(pTranslateColor(miActiveBackColor))
      FillRect lHDC, tr, hBr
      DeleteObject hBr
   Else
      pTileArea lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, tDraw.oBitmapSuperLight.hdc, tDraw.oBitmapSuperLight.Width, tDraw.oBitmapSuperLight.Height, lOffsetY
   End If
End Sub
Private Sub pFillWithNormalBackground(ByRef tDraw As tMenuDraw, ByVal lHDC As Long, tr As tRect, ByVal lOffsetY As Long, Optional ByVal bIgnoreBitmap As Boolean)
Dim hBrush As Long
   If tDraw.oBitmap Is Nothing Or bIgnoreBitmap Then
      hBrush = CreateSolidBrush(pTranslateColor(miInActiveBackColor))
      FillRect lHDC, tr, hBrush
      DeleteObject hBrush
   Else
      pTileArea lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, tDraw.oBitmap.hdc, tDraw.oBitmap.Width, tDraw.oBitmap.Height, lOffsetY
   End If
End Sub
Private Sub pFillWithLighterControlColor(ByRef tDraw As tMenuDraw, ByVal lHDC As Long, tr As tRect, ByVal lOffsetY As Long)

Dim hBrush As Long
   SetBkMode lHDC, OPAQUE
   If Not tDraw.oBitmapLight Is Nothing Then
      pTileArea lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, tDraw.oBitmapLight.hdc, tDraw.oBitmapLight.Width, tDraw.oBitmapLight.Height, lOffsetY
   Else
      If (pNoPalette) Then
         hBrush = CreateSolidBrush(pBlendColor(vbButtonFace, miInActiveBackColor))
         FillRect lHDC, tr, hBrush
         DeleteObject hBrush
      Else
         tDraw.oBrush.Rectangle lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, 1, PATCOPY, True, miInActiveBackColor, vb3DHighlight
      End If
   End If
   SetBkMode lHDC, TRANSPARENT
End Sub

Private Sub pFillWithLighterSelectedColor(ByRef tDraw As tMenuDraw, ByVal lHDC As Long, tr As tRect, ByVal lOffsetY As Long)

Dim hBrush As Long
   SetBkMode lHDC, OPAQUE
   If Not tDraw.oBitmapSuperLight Is Nothing Then 'And mbInfrequent Then
        pTileArea lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, tDraw.oBitmapSuperLight.hdc, tDraw.oBitmapSuperLight.Width, tDraw.oBitmapSuperLight.Height, lOffsetY
   'ElseIf Not tDraw.oBitmapLight Is Nothing Then
        'pTileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, tDraw.oBitmapLight.hDc, tDraw.oBitmapLight.Width, tDraw.oBitmapLight.Height, lOffsetY
   Else
      If (pNoPalette) Then
         hBrush = CreateSolidBrush(pLighterColor(pLighterColor(pLighterColor(pBlendColor(vbHighlight, miInActiveBackColor)))))
         
         FillRect lHDC, tr, hBrush
         DeleteObject hBrush
      Else
         tDraw.oBrush.Rectangle lHDC, tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, 1, PATCOPY, True, pSlightlyLighterColor(pSlightlyLighterColor(pSlightlyLighterColor(pSlightlyLighterColor(pBlendColor(vbHighlight, miInActiveBackColor)))))
      End If
   End If
   SetBkMode lHDC, TRANSPARENT
End Sub

Private Function pDrawText(ByVal lHDC As Long, ByVal sText As String, tr As tRect, ByVal dtFlags As Long, ByVal bDisabled As Boolean, ByVal bOfficeXP As Boolean)
   If bDisabled Then
      If (bOfficeXP) Then
         SetTextColor lHDC, pTranslateColor(vb3DShadow)
      Else
         SetTextColor lHDC, pTranslateColor(vb3DHighlight)
         OffsetRect tr, 1, 1
      End If
   End If
   DrawText lHDC, sText, -1, tr, dtFlags
   If bDisabled Then
      If Not (bOfficeXP) Then
         OffsetRect tr, -1, -1
         SetTextColor lHDC, pTranslateColor(vbButtonShadow)
         DrawText lHDC, sText, -1, tr, dtFlags
      End If
   End If
End Function

Private Sub pTileArea( _
        ByVal hdcTo As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal hDcSrc As Long, _
        ByVal SrcWidth As Long, _
        ByVal SrcHeight As Long, _
        ByVal lOffsetY As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = (x Mod SrcWidth)
    lSrcStartY = ((y + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub

Private Function pNoPalette(Optional ByVal bForce As Boolean = False) As Boolean
Static bOnce As Boolean
Static bNoPalette As Boolean
Dim lHDC As Long
Dim lBits As Long
   If (bForce) Then
      bOnce = False
   End If
   If Not (bOnce) Then
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      If (lHDC <> 0) Then
         lBits = GetDeviceCaps(lHDC, BITSPIXEL)
         If (lBits <> 0) Then
            bOnce = True
         End If
         bNoPalette = (lBits > 8)
         DeleteDC lHDC
      End If
   End If
   pNoPalette = bNoPalette
End Function

Private Function pTranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, pTranslateColor) Then
        pTranslateColor = CLR_INVALID
    End If
End Function

Private Sub pDrawEdge( _
      ByVal hdc As Long, _
      qrc As tRect, _
      ByVal edge As Long, _
      ByVal grfFlags As Long, _
      ByVal bOfficeXPStyle As Boolean)
      
   If (bOfficeXPStyle) Then
      Dim junk As tPoint
      Dim hPenOld As Long
      Dim hPen As Long
      If (qrc.Bottom > qrc.Top) Then
         hPen = CreatePen(PS_SOLID, 1, pTranslateColor(vbHighlight))
      Else
         hPen = CreatePen(PS_SOLID, 1, pTranslateColor(vb3DShadow))
      End If
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, qrc.Left, qrc.Top, junk
      LineTo hdc, qrc.Right - 1, qrc.Top
      If (qrc.Bottom > qrc.Top) Then
         LineTo hdc, qrc.Right - 1, qrc.Bottom - 1
         LineTo hdc, qrc.Left, qrc.Bottom - 1
         LineTo hdc, qrc.Left, qrc.Top
      End If
      SelectObject hdc, hPenOld
      DeleteObject hPen
   Else
      DrawEdgeAPI hdc, qrc, edge, grfFlags
   End If
End Sub

Private Sub pDrawGradient( _
      ByVal hdc As Long, _
      ByRef rct As tRect, _
      ByVal lEndColor As Long, _
      ByVal lStartColor As Long, _
      ByVal bVertical As Boolean _
   )
Dim lStep As Long
Dim lPos As Long, lSize As Long
Dim bRGB(1 To 3) As Integer
Dim bRGBStart(1 To 3) As Integer
Dim dR(1 To 3) As Double
Dim dPos As Double, d As Double
Dim hBr As Long
Dim tr As tRect
   
   LSet tr = rct
   If bVertical Then
      lSize = (tr.Bottom - tr.Top)
   Else
      lSize = (tr.Right - tr.Left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
       lStep = 3
   End If
       
   bRGB(1) = lStartColor And &HFF&
   bRGB(2) = (lStartColor And &HFF00&) \ &H100&
   bRGB(3) = (lStartColor And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColor And &HFF&) - bRGB(1)
   dR(2) = ((lEndColor And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColor And &HFF0000) \ &H10000) - bRGB(3)
        
   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tr.Top = tr.Bottom - lStep
      Else
         tr.Left = tr.Right - lStep
      End If
      If tr.Top < rct.Top Then
         tr.Top = rct.Top
      End If
      If tr.Left < rct.Left Then
         tr.Left = rct.Left
      End If
      
      'Debug.Print tR.Right, tR.left, (bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1))
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tr, hBr
      DeleteObject hBr
            
      ' Adjust Color:
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tr.Bottom = tr.Top
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tr.Right = tr.Left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next

End Sub

Private Sub pImageListDrawIcon( _
        ByVal oVB6Iml As ImageList, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False _
    )
Dim lFlags As Long
Dim lR As Long

    lFlags = ILD_TRANSPARENT
    If (bSelected) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bBlend25) Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    If (Not oVB6Iml Is Nothing) Then
        'On Error Resume Next
        'oVB6Iml.ListImages(iIconIndex + 1).Draw hDc, lX * Screen.TwipsPerPixelX, lY * Screen.TwipsPerPixelY, lFlags
        lR = ImageList_Draw( _
                oVB6Iml.hImageList, _
                iIconIndex, _
                hdc, _
                lX, _
                lY, _
                lFlags)
        'On Error GoTo 0
    Else
        lR = ImageList_Draw( _
                hIml, _
                iIconIndex, _
                hdc, _
                lX, _
                lY, _
                lFlags)
        If (lR = 0) Then
            Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
        End If
    End If
End Sub

Private Sub pImageListDrawIconDisabled( _
        ByVal oVB6Iml As ImageList, _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long, _
        Optional ByVal asShadow As Boolean _
    )
Dim lR As Long
Dim hIcon As Long

   hIcon = 0
   If Not (oVB6Iml Is Nothing) Then
      On Error Resume Next

        Dim lhDCDisp As Long
        Dim lHDC As Long
        Dim lhBmp As Long
        Dim lhBmpOld As Long
        Dim lhIml As Long
                 
        lhDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        lHDC = CreateCompatibleDC(lhDCDisp)
        lhBmp = CreateCompatibleBitmap(lhDCDisp, oVB6Iml.ImageWidth, oVB6Iml.ImageHeight)
        
        DeleteDC lhDCDisp
        lhBmpOld = SelectObject(lHDC, lhBmp)
        oVB6Iml.ListImages.Item(iIconIndex + 1).Draw lHDC, 0, 0, 0
        SelectObject lHDC, lhBmpOld
        DeleteDC lHDC
        lhIml = ImageList_Create(oVB6Iml.ImageWidth, oVB6Iml.ImageHeight, ILC_MASK Or ILC_COLOR32, 1, 1)
        ImageList_AddMasked lhIml, lhBmp, pTranslateColor(oVB6Iml.BackColor)
        DeleteObject lhBmp
        hIcon = ImageList_GetIcon(lhIml, 0, 0)
        ImageList_Destroy lhIml
      On Error GoTo 0
   Else
      hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
   End If
   
   If (hIcon <> 0) Then
      If (asShadow) Then
         Dim hBr As Long
         hBr = GetSysColorBrush(vb3DShadow And &H1F)
         lR = DrawState(hdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO)
         DeleteObject hBr
      Else
         lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
      End If
      DestroyIcon hIcon
   End If
   
End Sub

Private Function pBlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = pTranslateColor(oColorFrom)
   lCTo = pTranslateColor(oColorTo)
Dim lCRetR As Long
Dim lCRetG As Long
Dim lCRetB As Long
   lCRetR = (lCFrom And &HFF) + ((lCTo And &HFF) - (lCFrom And &HFF)) \ 2
   If (lCRetR > 255) Then lCRetR = 255 Else If (lCRetR < 0) Then lCRetR = 0
   lCRetG = ((lCFrom \ &H100) And &HFF&) + (((lCTo \ &H100) And &HFF&) - ((lCFrom \ &H100) And &HFF&)) \ 2
   If (lCRetG > 255) Then lCRetG = 255 Else If (lCRetG < 0) Then lCRetG = 0
   lCRetB = ((lCFrom \ &H10000) And &HFF&) + (((lCTo \ &H10000) And &HFF&) - ((lCFrom \ &H10000) And &HFF&)) \ 2
   If (lCRetB > 255) Then lCRetB = 255 Else If (lCRetB < 0) Then lCRetB = 0
   pBlendColor = RGB(lCRetR, lCRetG, lCRetB)
End Function

Private Function pLighterColor(ByVal oColor As OLE_COLOR) As Long
Dim lC As Long
Dim h As Single, s As Single, l As Single
Dim lR As Long, lG As Long, lB As Long
Static s_lColLast As Long
Static s_lLightColLast As Long
   
   lC = pTranslateColor(oColor)
   If (lC <> s_lColLast) Then
      s_lColLast = lC
      Draw_HLSforRGB lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
      If (l > 0.99) Then
         l = l * 0.8
      Else
         l = l * 1.1
         If (l > 1) Then
            l = 1
         End If
      End If
      Draw_RGBforHLS h, s, l, lR, lG, lB
      s_lLightColLast = RGB(lR, lG, lB)
   End If
   pLighterColor = s_lLightColLast
End Function

Private Function pSlightlyLighterColor(ByVal oColor As OLE_COLOR) As Long
Dim lC As Long
Dim h As Single, s As Single, l As Single
Dim lR As Long, lG As Long, lB As Long
Static s_lColLast As Long
Static s_lLightColLast As Long
   
   lC = pTranslateColor(oColor)
   If (lC <> s_lColLast) Then
      s_lColLast = lC
      Draw_HLSforRGB lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
      If (l > 0.99) Then
         l = l * 0.95
      Else
         l = l * 1.05
         If (l > 1) Then
            l = 1
         End If
      End If
      Draw_RGBforHLS h, s, l, lR, lG, lB
      s_lLightColLast = RGB(lR, lG, lB)
   End If
   pSlightlyLighterColor = s_lLightColLast
End Function
