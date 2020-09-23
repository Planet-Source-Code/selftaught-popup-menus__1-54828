Attribute VB_Name = "mGeneral"
Option Explicit

Public giTestForms As Long

'Api declarations
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, lphModule As Any, cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, nSize As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function ExeFileName(ByVal hwnd As Long) As String

Const PROCESS_QUERY_INFORMATION As Long = &H400&
Const PROCESS_VM_READ           As Long = &H10&

Const opFlags       As Long = PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ
Const nMaxMods      As Long = 256
Const nBaseModule   As Long = 1
Const nBytesPerLong As Long = 4
Const MAX_PATH      As Long = 260
  
  Dim hModules()    As Long
  Dim hProcess      As Long
  Dim nProcessID    As Long
  Dim nBufferSize   As Long
  Dim nBytesNeeded  As Long
  Dim nRet          As Long
  Dim sBuffer       As String
  
  'Get the process ID from the window handle
  Call GetWindowThreadProcessId(hwnd, nProcessID)

  'Open the process so we can read some module info.
  hProcess = OpenProcess(opFlags, False, nProcessID)
  
  If hProcess Then
    'Get list of process modules.
    ReDim hModules(1 To nMaxMods) As Long
    nBufferSize = UBound(hModules) * nBytesPerLong
    nRet = EnumProcessModules(hProcess, hModules(nBaseModule), nBufferSize, nBytesNeeded)
    
    If nRet = False Then
      'Check to see if we need to allocate more space for results.
      If nBytesNeeded > nBufferSize Then
        ReDim hModules(nBaseModule To nBytesNeeded \ nBytesPerLong) As Long
        nBufferSize = nBytesNeeded
        nRet = EnumProcessModules(hProcess, hModules(nBaseModule), nBufferSize, nBytesNeeded)
      End If
    End If

    'Get the module name.
    sBuffer = Space$(MAX_PATH)
    nRet = GetModuleFileNameEx(hProcess, hModules(nBaseModule), sBuffer, MAX_PATH)
    
    If nRet Then
      ExeFileName = Left$(sBuffer, nRet)
    End If
    
    'Clean up
    Call CloseHandle(hProcess)
  End If
End Function

Function IsWindowLocal(ByVal hwnd As Long) As Boolean
    Dim idWnd As Long
    GetWindowThreadProcessId hwnd, idWnd
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function


Public Function GetHCName(ByVal iHC As eHookCode) As String
    Select Case iHC
    Case HC_ACTION:                 GetHCName = "HC_ACTION"
    Case HC_GETNEXT:                GetHCName = "HC_GETNEXT"
    Case HC_NOREM:                  GetHCName = "HC_NOREM"
    Case HC_NOREMOVE:               GetHCName = "HC_NOREMOVE"
    Case HC_SKIP:                   GetHCName = "HC_SKIP"
    Case HC_SYSMODALOFF:            GetHCName = "HC_SYSMODALOFF"
    Case HC_SYSMODALON:             GetHCName = "HC_SYSMODALON"
    Case HCBT_ACTIVATE:             GetHCName = "HCBT_ACTIVATE"
    Case HCBT_CLICKSKIPPED:         GetHCName = "HCBT_CLICKSKIPPED"
    Case HCBT_CREATEWND:            GetHCName = "HCBT_CREATEWND"
    Case HCBT_DESTROYWND:           GetHCName = "HCBT_DESTROYWND"
    Case HCBT_KEYSKIPPED:           GetHCName = "HCBT_KEYSKIPPED"
    Case HCBT_MINMAX:               GetHCName = "HCBT_MINMAX"
    Case HCBT_MOVESIZE:             GetHCName = "HCBT_MOVESIZE"
    Case HCBT_QS:                   GetHCName = "HCBT_QS"
    Case HCBT_SETFOCUS:             GetHCName = "HCBT_SETFOCUS"
    Case HCBT_SYSCOMMAND:           GetHCName = "HCBT_SYSCOMMAND"
    Case HSHELL_ACTIVATESHELLWINDOW: GetHCName = "HSHELL_ACTIVATESHELLWINDOW"
    Case HSHELL_GETMINRECT:         GetHCName = "HSHELL_GETMINRECT"
    Case HSHELL_LANGUAGE:           GetHCName = "HSHELL_LANGUAGE"
    Case HSHELL_REDRAW:             GetHCName = "HSHELL_REDRAW"
    Case HSHELL_TASKMAN:            GetHCName = "HSHELL_TASKMAN"
    Case HSHELL_WINDOWACTIVATED:    GetHCName = "HSHELL_WINDOWACTIVATED"
    Case HSHELL_WINDOWCREATED:      GetHCName = "HSHELL_WINDOWCREATED"
    Case HSHELL_WINDOWDESTROYED:    GetHCName = "HSHELL_WINDOWDESTROYED"
    Case MSGF_DDEMGR:               GetHCName = "MSGF_DDEMGR"
    Case MSGF_DIALOGBOX:            GetHCName = "MSGF_DIALOGBOX"
    Case MSGF_MAX:                  GetHCName = "MSGF_MAX"
    Case MSGF_MENU:                 GetHCName = "MSGF_MENU"
    Case MSGF_MESSAGEBOX:           GetHCName = "MSGF_MESSAGEBOX"
    Case MSGF_NEXTWINDOW:           GetHCName = "MSGF_NEXTWINDOW"
    Case MSGF_SCROLLBAR:            GetHCName = "MSGF_SCROLLBAR"
    Case MSGF_USER:                 GetHCName = "MSGF_USER"
    Case PM_NOREMOVE:               GetHCName = "PM_NOREMOVE"
    Case PM_NOYIELD:                GetHCName = "PM_NOYIELD"
    Case PM_REMOVE:                 GetHCName = "PM_REMOVE"
    End Select
End Function

'Return the name of the passed message number
Public Function GetMsgName(ByVal uMsg As eMsg) As String
    Select Case uMsg
    Case WM_ACTIVATE:            GetMsgName = "WM_ACTIVATE"
    Case WM_ACTIVATEAPP:         GetMsgName = "WM_ACTIVATEAPP"
    Case WM_ASKCBFORMATNAME:     GetMsgName = "WM_ASKCBFORMATNAME"
    Case WM_CANCELJOURNAL:       GetMsgName = "WM_CANCELJOURNAL"
    Case WM_CANCELMODE:          GetMsgName = "WM_CANCELMODE"
    Case WM_CAPTURECHANGED:      GetMsgName = "WM_CAPTURECHANGED"
    Case WM_CHANGECBCHAIN:       GetMsgName = "WM_CHANGECBCHAIN"
    Case WM_CHAR:                GetMsgName = "WM_CHAR"
    Case WM_CHARTOITEM:          GetMsgName = "WM_CHARTOITEM"
    Case WM_CHILDACTIVATE:       GetMsgName = "WM_CHILDACTIVATE"
    Case WM_CLEAR:               GetMsgName = "WM_CLEAR"
    Case WM_CLOSE:               GetMsgName = "WM_CLOSE"
    Case WM_COMMAND:             GetMsgName = "WM_COMMAND"
    Case WM_COMPACTING:          GetMsgName = "WM_COMPACTING"
    Case WM_COMPAREITEM:         GetMsgName = "WM_COMPAREITEM"
    Case WM_COPY:                GetMsgName = "WM_COPY"
    Case WM_COPYDATA:            GetMsgName = "WM_COPYDATA"
    Case WM_CREATE:              GetMsgName = "WM_CREATE"
    Case WM_CTLCOLORBTN:         GetMsgName = "WM_CTLCOLORBTN"
    Case WM_CTLCOLORDLG:         GetMsgName = "WM_CTLCOLORDLG"
    Case WM_CTLCOLOREDIT:        GetMsgName = "WM_CTLCOLOREDIT"
    Case WM_CTLCOLORLISTBOX:     GetMsgName = "WM_CTLCOLORLISTBOX"
    Case WM_CTLCOLORMSGBOX:      GetMsgName = "WM_CTLCOLORMSGBOX"
    Case WM_CTLCOLORSCROLLBAR:   GetMsgName = "WM_CTLCOLORSCROLLBAR"
    Case WM_CTLCOLORSTATIC:      GetMsgName = "WM_CTLCOLORSTATIC"
    Case WM_CUT:                 GetMsgName = "WM_CUT"
    Case WM_DEADCHAR:            GetMsgName = "WM_DEADCHAR"
    Case WM_DELETEITEM:          GetMsgName = "WM_DELETEITEM"
    Case WM_DESTROY:             GetMsgName = "WM_DESTROY"
    Case WM_DESTROYCLIPBOARD:    GetMsgName = "WM_DESTROYCLIPBOARD"
    Case WM_DRAWCLIPBOARD:       GetMsgName = "WM_DRAWCLIPBOARD"
    Case WM_DRAWITEM:            GetMsgName = "WM_DRAWITEM"
    Case WM_DROPFILES:           GetMsgName = "WM_DROPFILES"
    Case WM_ENABLE:              GetMsgName = "WM_ENABLE"
    Case WM_ENDSESSION:          GetMsgName = "WM_ENDSESSION"
    Case WM_ENTERIDLE:           GetMsgName = "WM_ENTERIDLE"
    Case WM_ENTERMENULOOP:       GetMsgName = "WM_ENTERMENULOOP"
    Case WM_ENTERSIZEMOVE:       GetMsgName = "WM_ENTERSIZEMOVE"
    Case WM_ERASEBKGND:          GetMsgName = "WM_ERASEBKGND"
    Case WM_EXITMENULOOP:        GetMsgName = "WM_EXITMENULOOP"
    Case WM_EXITSIZEMOVE:        GetMsgName = "WM_EXITSIZEMOVE"
    Case WM_FONTCHANGE:          GetMsgName = "WM_FONTCHANGE"
    Case WM_GETDLGCODE:          GetMsgName = "WM_GETDLGCODE"
    Case WM_GETFONT:             GetMsgName = "WM_GETFONT"
    Case WM_GETHOTKEY:           GetMsgName = "WM_GETHOTKEY"
    Case WM_GETMINMAXINFO:       GetMsgName = "WM_GETMINMAXINFO"
    Case WM_GETTEXT:             GetMsgName = "WM_GETTEXT"
    Case WM_GETTEXTLENGTH:       GetMsgName = "WM_GETTEXTLENGTH"
    Case WM_HOTKEY:              GetMsgName = "WM_HOTKEY"
    Case WM_HSCROLL:             GetMsgName = "WM_HSCROLL"
    Case WM_HSCROLLCLIPBOARD:    GetMsgName = "WM_HSCROLLCLIPBOARD"
    Case WM_ICONERASEBKGND:      GetMsgName = "WM_ICONERASEBKGND"
    Case WM_IME_CHAR:            GetMsgName = "WM_IME_CHAR"
    Case WM_IME_COMPOSITION:     GetMsgName = "WM_IME_COMPOSITION"
    Case WM_IME_COMPOSITIONFULL: GetMsgName = "WM_IME_COMPOSITIONFULL"
    Case WM_IME_CONTROL:         GetMsgName = "WM_IME_CONTROL"
    Case WM_IME_ENDCOMPOSITION:  GetMsgName = "WM_IME_ENDCOMPOSITION"
    Case WM_IME_KEYDOWN:         GetMsgName = "WM_IME_KEYDOWN"
    Case WM_IME_KEYLAST:         GetMsgName = "WM_IME_KEYLAST"
    Case WM_IME_KEYUP:           GetMsgName = "WM_IME_KEYUP"
    Case WM_IME_NOTIFY:          GetMsgName = "WM_IME_NOTIFY"
    Case WM_IME_SELECT:          GetMsgName = "WM_IME_SELECT"
    Case WM_IME_SETCONTEXT:      GetMsgName = "WM_IME_SETCONTEXT"
    Case WM_IME_STARTCOMPOSITION: GetMsgName = "WM_IME_STARTCOMPOSITION"
    Case WM_INITDIALOG:          GetMsgName = "WM_INITDIALOG"
    Case WM_INITMENU:            GetMsgName = "WM_INITMENU"
    Case WM_INITMENUPOPUP:       GetMsgName = "WM_INITMENUPOPUP"
    Case WM_KEYDOWN:             GetMsgName = "WM_KEYDOWN"
    Case WM_KEYFIRST:            GetMsgName = "WM_KEYFIRST"
    Case WM_KEYLAST:             GetMsgName = "WM_KEYLAST"
    Case WM_KEYUP:               GetMsgName = "WM_KEYUP"
    Case WM_KILLFOCUS:           GetMsgName = "WM_KILLFOCUS"
    Case WM_LBUTTONDBLCLK:       GetMsgName = "WM_LBUTTONDBLCLK"
    Case WM_LBUTTONDOWN:         GetMsgName = "WM_LBUTTONDOWN"
    Case WM_LBUTTONUP:           GetMsgName = "WM_LBUTTONUP"
    Case WM_MBUTTONDBLCLK:       GetMsgName = "WM_MBUTTONDBLCLK"
    Case WM_MBUTTONDOWN:         GetMsgName = "WM_MBUTTONDOWN"
    Case WM_MBUTTONUP:           GetMsgName = "WM_MBUTTONUP"
    Case WM_MDIACTIVATE:         GetMsgName = "WM_MDIACTIVATE"
    Case WM_MDICASCADE:          GetMsgName = "WM_MDICASCADE"
    Case WM_MDICREATE:           GetMsgName = "WM_MDICREATE"
    Case WM_MDIDESTROY:          GetMsgName = "WM_MDIDESTROY"
    Case WM_MDIGETACTIVE:        GetMsgName = "WM_MDIGETACTIVE"
    Case WM_MDIICONARRANGE:      GetMsgName = "WM_MDIICONARRANGE"
    Case WM_MDIMAXIMIZE:         GetMsgName = "WM_MDIMAXIMIZE"
    Case WM_MDINEXT:             GetMsgName = "WM_MDINEXT"
    Case WM_MDIREFRESHMENU:      GetMsgName = "WM_MDIREFRESHMENU"
    Case WM_MDIRESTORE:          GetMsgName = "WM_MDIRESTORE"
    Case WM_MDISETMENU:          GetMsgName = "WM_MDISETMENU"
    Case WM_MDITILE:             GetMsgName = "WM_MDITILE"
    Case WM_MEASUREITEM:         GetMsgName = "WM_MEASUREITEM"
    Case WM_MENUCHAR:            GetMsgName = "WM_MENUCHAR"
    Case WM_MENUSELECT:          GetMsgName = "WM_MENUSELECT"
    Case WM_MOUSEACTIVATE:       GetMsgName = "WM_MOUSEACTIVATE"
    Case WM_MOUSEMOVE:           GetMsgName = "WM_MOUSEMOVE"
    Case WM_MOUSEWHEEL:          GetMsgName = "WM_MOUSEWHEEL"
    Case WM_MOVE:                GetMsgName = "WM_MOVE"
    Case WM_MOVING:              GetMsgName = "WM_MOVING"
    Case WM_NCACTIVATE:          GetMsgName = "WM_NCACTIVATE"
    Case WM_NCCALCSIZE:          GetMsgName = "WM_NCCALCSIZE"
    Case WM_NCCREATE:            GetMsgName = "WM_NCCREATE"
    Case WM_NCDESTROY:           GetMsgName = "WM_NCDESTROY"
    Case WM_NCHITTEST:           GetMsgName = "WM_NCHITTEST"
    Case WM_NCLBUTTONDBLCLK:     GetMsgName = "WM_NCLBUTTONDBLCLK"
    Case WM_NCLBUTTONDOWN:       GetMsgName = "WM_NCLBUTTONDOWN"
    Case WM_NCLBUTTONUP:         GetMsgName = "WM_NCLBUTTONUP"
    Case WM_NCMBUTTONDBLCLK:     GetMsgName = "WM_NCMBUTTONDBLCLK"
    Case WM_NCMBUTTONDOWN:       GetMsgName = "WM_NCMBUTTONDOWN"
    Case WM_NCMBUTTONUP:         GetMsgName = "WM_NCMBUTTONUP"
    Case WM_NCMOUSEMOVE:         GetMsgName = "WM_NCMOUSEMOVE"
    Case WM_NCPAINT:             GetMsgName = "WM_NCPAINT"
    Case WM_NCRBUTTONDBLCLK:     GetMsgName = "WM_NCRBUTTONDBLCLK"
    Case WM_NCRBUTTONDOWN:       GetMsgName = "WM_NCRBUTTONDOWN"
    Case WM_NCRBUTTONUP:         GetMsgName = "WM_NCRBUTTONUP"
    Case WM_NEXTDLGCTL:          GetMsgName = "WM_NEXTDLGCTL"
    Case WM_NULL:                GetMsgName = "WM_NULL"
    Case WM_PAINT:               GetMsgName = "WM_PAINT"
    Case WM_PAINTCLIPBOARD:      GetMsgName = "WM_PAINTCLIPBOARD"
    Case WM_PAINTICON:           GetMsgName = "WM_PAINTICON"
    Case WM_PALETTECHANGED:      GetMsgName = "WM_PALETTECHANGED"
    Case WM_PALETTEISCHANGING:   GetMsgName = "WM_PALETTEISCHANGING"
    Case WM_PARENTNOTIFY:        GetMsgName = "WM_PARENTNOTIFY"
    Case WM_PASTE:               GetMsgName = "WM_PASTE"
    Case WM_PENWINFIRST:         GetMsgName = "WM_PENWINFIRST"
    Case WM_PENWINLAST:          GetMsgName = "WM_PENWINLAST"
    Case WM_POWER:               GetMsgName = "WM_POWER"
    Case WM_QUERYDRAGICON:       GetMsgName = "WM_QUERYDRAGICON"
    Case WM_QUERYENDSESSION:     GetMsgName = "WM_QUERYENDSESSION"
    Case WM_QUERYNEWPALETTE:     GetMsgName = "WM_QUERYNEWPALETTE"
    Case WM_QUERYOPEN:           GetMsgName = "WM_QUERYOPEN"
    Case WM_QUEUESYNC:           GetMsgName = "WM_QUEUESYNC"
    Case WM_QUIT:                GetMsgName = "WM_QUIT"
    Case WM_RBUTTONDBLCLK:       GetMsgName = "WM_RBUTTONDBLCLK"
    Case WM_RBUTTONDOWN:         GetMsgName = "WM_RBUTTONDOWN"
    Case WM_RBUTTONUP:           GetMsgName = "WM_RBUTTONUP"
    Case WM_RENDERALLFORMATS:    GetMsgName = "WM_RENDERALLFORMATS"
    Case WM_RENDERFORMAT:        GetMsgName = "WM_RENDERFORMAT"
    Case WM_SETCURSOR:           GetMsgName = "WM_SETCURSOR"
    Case WM_SETFOCUS:            GetMsgName = "WM_SETFOCUS"
    Case WM_SETFONT:             GetMsgName = "WM_SETFONT"
    Case WM_SETHOTKEY:           GetMsgName = "WM_SETHOTKEY"
    Case WM_SETREDRAW:           GetMsgName = "WM_SETREDRAW"
    Case WM_SETTEXT:             GetMsgName = "WM_SETTEXT"
    Case WM_SHOWWINDOW:          GetMsgName = "WM_SHOWWINDOW"
    Case WM_SIZE:                GetMsgName = "WM_SIZE"
    Case WM_SIZING:              GetMsgName = "WM_SIZING"
    Case WM_SIZECLIPBOARD:       GetMsgName = "WM_SIZECLIPBOARD"
    Case WM_SPOOLERSTATUS:       GetMsgName = "WM_SPOOLERSTATUS"
    Case WM_SYSCHAR:             GetMsgName = "WM_SYSCHAR"
    Case WM_SYSCOLORCHANGE:      GetMsgName = "WM_SYSCOLORCHANGE"
    Case WM_SYSCOMMAND:          GetMsgName = "WM_SYSCOMMAND"
    Case WM_SYSDEADCHAR:         GetMsgName = "WM_SYSDEADCHAR"
    Case WM_SYSKEYDOWN:          GetMsgName = "WM_SYSKEYDOWN"
    Case WM_SYSKEYUP:            GetMsgName = "WM_SYSKEYUP"
    Case WM_TIMECHANGE:          GetMsgName = "WM_TIMECHANGE"
    Case WM_TIMER:               GetMsgName = "WM_TIMER"
    Case WM_UNDO:                GetMsgName = "WM_UNDO"
    Case WM_USER:                GetMsgName = "WM_USER"
    Case WM_VKEYTOITEM:          GetMsgName = "WM_VKEYTOITEM"
    Case WM_VSCROLL:             GetMsgName = "WM_VSCROLL"
    Case WM_VSCROLLCLIPBOARD:    GetMsgName = "WM_VSCROLL"
    Case WM_WINDOWPOSCHANGED:    GetMsgName = "WM_WINDOWPOSCHANGED"
    Case WM_WINDOWPOSCHANGING:   GetMsgName = "WM_WINDOWPOSCHANGING"
    Case WM_WININICHANGE:        GetMsgName = "WM_WININICHANGE"
    Case Else:                   GetMsgName = FmtHex(uMsg)
    End Select
End Function


Public Function HookName(ByVal iType As eHookType) As String
    Select Case iType
        Case WH_MSGFILTER:          HookName = "WH_MSGFILTER"
        Case WH_JOURNALRECORD:      HookName = "WH_JOURNALRECORD"
        Case WH_JOURNALPLAYBACK:    HookName = "WH_JOURNALPLAYBACK"
        Case WH_KEYBOARD:           HookName = "WH_KEYBOARD"
        Case WH_GETMESSAGE:         HookName = "WH_GETMESSAGE"
        Case WH_CALLWNDPROC:        HookName = "WH_CALLWNDPROC"
        Case WH_CBT:                HookName = "WH_CBT"
        Case WH_SYSMSGFILTER:       HookName = "WH_SYSMSGFILTER"
        Case WH_MOUSE:              HookName = "WH_MOUSE"
        Case WH_DEBUG:              HookName = "WH_DEBUG"
        Case WH_SHELL:              HookName = "WH_SHELL"
        Case WH_FOREGROUNDIDLE:     HookName = "WH_FOREGROUNDIDLE"
        Case WH_CALLWNDPROCRET:     HookName = "WH_CALLWNDPROCRET"
        Case WH_KEYBOARD_LL:        HookName = "WH_KEYBOARD_LL"
        Case WH_MOUSE_LL:           HookName = "WH_MOUSE_LL"
    End Select
End Function

'Return the passed Long value as a hex string with leading zeros, if required, to a width of eight characters, plus a trailing space
Public Function FmtHex(ByVal nValue As Long) As String
  Dim s As String
  s = Hex$(nValue)
  FmtHex = String$(8& - Len(s), "0") & s & " "
End Function

Public Function HexVal(ByVal sHex As String) As Long
    If Left$(sHex, 2) = "0x" Then Mid$(sHex, 1, 2) = "&H" Else sHex = "&H" & sHex
    HexVal = Val(sHex & "&")
End Function

Public Function lstOrAllItemData(ByVal oLst As ListBox) As Long
    Dim i As Long
    With oLst
        For i = 0 To .ListCount - 1&
            If .Selected(i) Then _
                lstOrAllItemData = lstOrAllItemData Or .ItemData(i)
        Next
    End With
End Function

Public Sub lstClear(ByVal oLst As ListBox)
    Dim i As Long
    With oLst
        For i = 0 To .ListCount - 1&
            .Selected(i) = False
        Next
    End With
End Sub

Public Sub lstSetChecksFromItemData(ByVal oLst As ListBox, ByVal liMask As Long)
    Dim i As Long
    With oLst
        For i = 0 To .ListCount - 1&
            .Selected(i) = CBool(liMask And .ItemData(i))
        Next
    End With
End Sub

Public Sub picShowOne(ByVal pics As Object, ByVal iIndex As Long)
    Dim loPic As PictureBox
    For Each loPic In pics
        loPic.Visible = loPic.Index = iIndex
    Next
End Sub

