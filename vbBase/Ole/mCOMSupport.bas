Attribute VB_Name = "mCOMSupport"
'////////////////////////////////////////////////////////////
'// Name : modCOMSupport.bas
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// ReplaceVTableEntry function adapted from the book
'// 'Hardcore Visual Basic' by Bruce McKinney (although I think it
'// was originally written by Mathew Curland for his 'Black Belt
'// Programming' articles in the VBPJ). This book is recommended
'// reading for anyone attempting VTable function address
'// overriding.
'/////////////////////////////////////////////////////////////
'// MapCOMErr & RShiftDWord adapted from 'VBCore' by Bruce McKinney
'/////////////////////////////////////////////////////////////
'// GetKeyModifiers & KeyIsPressed adapted from code written
'// by Steve McMahon for the Owner Draw Combo/ListBox available
'// at vbAccelerator.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

'win32 forward declarations
'constants
Public Const PAGE_EXECUTE_READWRITE& = &H40&

'standard COM return codes
Public Const S_OK = &H0&
Public Const S_FALSE = &H1&
Public Const E_NOTIMPL = &H80004001      '_HRESULT_TYPEDEF_=(0x80004001L&)
Public Const E_OUTOFMEMORY = &H8007000E  '_HRESULT_TYPEDEF_=(0x8007000EL&)
Public Const E_INVALIDARG = &H80070057   '_HRESULT_TYPEDEF_=(0x80070057L&)
Public Const E_NOINTERFACE = &H80004002  '_HRESULT_TYPEDEF_=(0x80004002L&)
Public Const E_POINTER = &H80004003      '_HRESULT_TYPEDEF_=(0x80004003L&)
Public Const E_HANDLE = &H80070006       '_HRESULT_TYPEDEF_=(0x80070006L&)
Public Const E_ABORT = &H80004004        '_HRESULT_TYPEDEF_=(0x80004004L&)
Public Const E_FAIL = &H80004005         '_HRESULT_TYPEDEF_=(0x80004005L&)
Public Const E_ACCESSDENIED = &H80070005 '_HRESULT_TYPEDEF_=(0x80070005L&)

'standard Dispatch ID constants
Public Const DISPID_UNKNOWN = (-1&)

Public Const DISPID_AMBIENT_BACKCOLOR = (-701&)
Public Const DISPID_AMBIENT_DISPLAYNAME = (-702&)
Public Const DISPID_AMBIENT_FONT = (-703&)
Public Const DISPID_AMBIENT_FORECOLOR = (-704&)
Public Const DISPID_AMBIENT_LOCALEID = (-705&)
Public Const DISPID_AMBIENT_MESSAGEREFLECT = (-706&)
Public Const DISPID_AMBIENT_SCALEUNITS = (-707&)
Public Const DISPID_AMBIENT_TEXTALIGN = (-708&)
Public Const DISPID_AMBIENT_USERMODE = (-709&)
Public Const DISPID_AMBIENT_UIDEAD = (-710&)
Public Const DISPID_AMBIENT_SHOWGRABHANDLES = (-711&)
Public Const DISPID_AMBIENT_SHOWHATCHING = (-712&)
Public Const DISPID_AMBIENT_DISPLAYASDEFAULT = (-713&)
Public Const DISPID_AMBIENT_SUPPORTSMNEMONICS = (-714&)
Public Const DISPID_AMBIENT_AUTOCLIP = (-715&)
Public Const DISPID_AMBIENT_APPEARANCE = (-716&)

Public Const DISPID_AMBIENT_CODEPAGE = (-725&)
Public Const DISPID_AMBIENT_PALETTE = (-726&)
Public Const DISPID_AMBIENT_CHARSET = (-727&)
Public Const DISPID_AMBIENT_TRANSFERPRIORITY = (-728&)

Public Const DISPID_AMBIENT_RIGHTTOLEFT = (-732&)
Public Const DISPID_AMBIENT_TOPTOBOTTOM = (-733&)

Public Const DISPID_Name = (-800&)
Public Const DISPID_Delete = (-801&)
Public Const DISPID_Object = (-802&)
Public Const DISPID_Parent = (-803&)

'accelerator flags (used with ACCEL structure)
Public Const FVIRTKEY = 1 '/* Assumed to be == TRUE */
Public Const FNOINVERT = &H2&
Public Const FSHIFT = &H4&
Public Const FCONTROL = &H8&
Public Const FALT = &H10&

'registry key flags
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const REG_SZ = 1

'OLE32
Private Declare Function IIDFromString Lib "OLE32.DLL" (ByVal lpsz As String, lpGuid As CLSID) As Long

'KERNEL32
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'USER32
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'ADVAPI32
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As Any, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long


Private Enum eInterfaces
    intOleControl
    intOleIPAO
    intOlePerPropertyBrowsing
End Enum

Private miClients() As Long
Private miClientCount As Long

Public Function FindGUIDForProgID(ByVal ProgID As String) As String
'find specified prog ID in registry & return GUID

    Dim hKey As Long
    Dim strCLSID As String
    Dim lngNullPos As Long
    Dim lngValueType As Long, lngStrLen As Long
    
    'open ProgID\CLSID registry key
    If RegOpenKey(HKEY_CLASSES_ROOT, ProgID & "\CLSID", hKey) <> ERROR_SUCCESS Then
        'attempt to open new version key for progid
        If RegOpenKey(HKEY_CLASSES_ROOT, ProgID & "\CurVer", hKey) <> ERROR_SUCCESS Then
            Exit Function
            
        Else
            'get ProgID string from key
            'get data type & size
            If RegQueryValueEx(hKey, 0&, 0&, lngValueType, ByVal 0&, lngStrLen) = ERROR_SUCCESS Then
                'if data type is string & size is > 0
                If lngValueType = REG_SZ And lngStrLen > 0 Then
                    ProgID = String$(lngStrLen, Chr$(0))
                    'get default value
                    If RegQueryValueEx(hKey, 0&, 0&, 0&, ByVal ProgID, lngStrLen) = ERROR_SUCCESS Then
                        'strip null terminator
                        lngNullPos = InStr(ProgID, Chr$(0))
                        If lngNullPos > 0 Then
                            ProgID = Left$(ProgID, lngNullPos - 1)
                            
                        End If
                        
                    End If
                
                End If
            
            End If
            'close ProgID\CurVer registry key
            RegCloseKey hKey
            'open ProgID\CLSID registry key
            If RegOpenKey(HKEY_CLASSES_ROOT, ProgID & "\CLSID", hKey) <> ERROR_SUCCESS Then
                Exit Function
                
            End If
            
        End If
    
    End If
    
    'get CLSID string from key
    'get data type & size
    If RegQueryValueEx(hKey, 0&, 0&, lngValueType, ByVal 0&, lngStrLen) = ERROR_SUCCESS Then
        'if data type is string & size is > 0
        If lngValueType = REG_SZ And lngStrLen > 0 Then
            strCLSID = String$(lngStrLen, Chr$(0))
            'get default value
            If RegQueryValueEx(hKey, 0&, 0&, 0&, ByVal strCLSID, lngStrLen) = ERROR_SUCCESS Then
                'strip null terminator
                lngNullPos = InStr(strCLSID, Chr$(0))
                If lngNullPos > 0 Then
                    strCLSID = Left$(strCLSID, lngNullPos - 1)
                    
                End If
                'store CLSID
                FindGUIDForProgID = strCLSID
                
            End If
        
        End If
    
    End If
    
    'close ProgID\CLSID registry key
    RegCloseKey hKey
End Function

Public Function GUIDFromString(ByVal Guid As String) As CLSID
'convert string GUID to *real* GUID

    Dim lpGuid As CLSID
    
    'convert string to unicode
    Guid = StrConv(Guid, vbUnicode)
    'convert string to GUID
    IIDFromString Guid, lpGuid
    'return *real* GUID
    GUIDFromString = lpGuid
End Function

Public Function MapCOMErr(ByVal ErrNumber As Long) As Long
'map vb error to COM error

    If ErrNumber <> 0 Then
        If (ErrNumber And &H80000000) Or (ErrNumber = 1) Then
            'Error HRESULT already set
            MapCOMErr = ErrNumber
            
        Else
            'Map back to a basic error number
            MapCOMErr = &H800A0000 Or ErrNumber
            
        End If
        
    End If
End Function

Public Function GetKeyModifiers() As Integer
'get pressed status of [SHIFT],[CONTROL], and [ALT] keys

    Dim nResult As Integer
    
    nResult = nResult Or (-1 * KeyIsPressed(vbKeyShift))
    nResult = nResult Or (-2 * KeyIsPressed(vbKeyMenu))
    nResult = nResult Or (-4 * KeyIsPressed(vbKeyControl))
    GetKeyModifiers = nResult
End Function

Public Function KeyIsPressed(ByVal VirtKeyCode As KeyCodeConstants) As Boolean
'poll windows to see if specified key is pressed

    Dim lngResult As Long
    
    lngResult = GetAsyncKeyState(VirtKeyCode)
    If (lngResult And &H8000&) = &H8000& Then
        KeyIsPressed = True
            
    End If
End Function

Public Function ReplaceVTableEntry(ByVal oObject As Long, _
                                   ByVal nEntry As Integer, _
                                   ByVal pFunc As Long) As Long
' Put the function address (callback) directly into the object v-table

    ' oObject - Pointer to object whose v-table will be modified
    ' nEntry - Index of v-table entry to be modified
    ' pFunc - Function pointer of new v-table method
                            
    Dim pFuncOld As Long, pVTableHead As Long
    Dim pFuncTmp As Long, lOldProtect As Long
    
    ' Object pointer contains a pointer to v-table--copy it to temporary
    CopyMemory pVTableHead, ByVal oObject, 4       ' pVTableHead = *oObject;
    ' Calculate pointer to specified entry
    pFuncTmp = pVTableHead + (nEntry - 1) * 4
    ' Save address of previous method for return
    CopyMemory pFuncOld, ByVal pFuncTmp, 4      ' pFuncOld = *pFuncTmp;
    ' Ignore if they're already the same
    If pFuncOld <> pFunc Then
        ' Need to change page protection to write to code
        VirtualProtect pFuncTmp, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        ' Write the new function address into the v-table
        CopyMemory ByVal pFuncTmp, pFunc, 4     ' *pFuncTmp = pfunc;
        ' Restore the previous page protection
        VirtualProtect pFuncTmp, 4, lOldProtect, lOldProtect 'Optional
        
    End If
    
    'return address of original proc
    ReplaceVTableEntry = pFuncOld
End Function

Public Function Attach(ByVal oObject As Object) As Boolean

    Dim liIndex As Long
    Dim liFirst As Long
    Dim liPtr As Long
    liPtr = ObjPtr(oObject)
    liIndex = ArrFindInt(miClients, miClientCount, liPtr, liFirst)

    If liIndex = Undefined Then

        If pSupports(oObject, intOleControl) Then
            ReplaceIOleControl oObject
            Attach = True
        End If

        If pSupports(oObject, intOleIPAO) Then
            ReplaceIOleInPlaceActiveObject oObject
            Attach = True
        End If

        If pSupports(oObject, intOlePerPropertyBrowsing) Then
            ReplaceIPerPropertyBrowsing oObject
            Attach = True
        End If

        If Attach Then
            If liFirst = Undefined Then
                liFirst = miClientCount
                miClientCount = miClientCount + 1&
                ArrRedim miClients, miClientCount, True
            End If
            miClients(liFirst) = liPtr
        End If

    End If
    If Not Attach Then gErr vbbTypeMismatch, "cOleHooks.Add"
End Function

Public Function Detach(ByVal oObject As Object) As Boolean

    Dim liIndex As Long
    Dim liPtr As Long
    
    liPtr = ObjPtr(oObject)
    
    liIndex = ArrFindInt(miClients, miClientCount, liPtr)

    If liIndex <> Undefined Then
        Detach = True
        If pSupports(oObject, intOleControl) Then RestoreIOleControl liPtr
        If pSupports(oObject, intOlePerPropertyBrowsing) Then RestoreIPerPropertyBrowsing liPtr
        If pSupports(oObject, intOleIPAO) Then RestoreIOleInPlaceActiveObject liPtr
        
        miClients(liIndex) = Undefined
        
        If liIndex = miClientCount - 1& Then
            For liIndex = liIndex - 1& To 0& Step -1&
                If miClients(liIndex) <> Undefined Then Exit For
            Next
            miClientCount = liIndex + 1&
        End If
        
    End If
    If Not Detach Then gErr vbbTypeMismatch, "cOleHooks.Remove"
End Function

Sub temp(c As Collection)
Debug.Print "ASDF"
End Sub

Public Function Exists(ByVal oObject As Object) As Boolean
    Exists = ArrFindInt(miClients, miClientCount, ObjPtr(oObject)) <> Undefined
End Function

Private Function pSupports(oUserControl As Object, iInterface As eInterfaces) As Boolean

    Dim oIPerPropertyBrowsing As vbACOMTLB.IPerPropertyBrowsing
    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB

    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject
    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB

    Dim oIOleControl As vbACOMTLB.IOleControl
    Dim oIOleControlVB As iOleControlVB

    On Error Resume Next

    Select Case iInterface
        Case intOlePerPropertyBrowsing
            Set oIPerPropertyBrowsing = oUserControl: Set oIPerPropertyBrowsingVB = oUserControl
            pSupports = Not (oIPerPropertyBrowsing Is Nothing Or oIPerPropertyBrowsingVB Is Nothing)
        Case intOleIPAO
            Set oIOleInPlaceActiveObject = oUserControl: Set oIOleInPlaceActiveObjectVB = oUserControl
            pSupports = Not (oIOleInPlaceActiveObject Is Nothing Or oIOleInPlaceActiveObjectVB Is Nothing)
        Case intOleControl
            Set oIOleControl = oUserControl: Set oIOleControlVB = oUserControl
            pSupports = Not (oIOleControl Is Nothing Or oIOleControlVB Is Nothing)
    End Select
    pSupports = pSupports And Err.Number = 0&
End Function
