Attribute VB_Name = "mPerPropertyBrowsing"
'////////////////////////////////////////////////////////////
'// Name : modIPerPropertyBrowsing.bas
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Custom implementation to make the IPerPropertyBrowsing
'// interface more 'VB Friendly'
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

Private Declare Function CoTaskMemAlloc Lib "OLE32" (ByVal cBytes As Long) As Long
Private Declare Function SysAllocString Lib "OLEAUT32" (ByVal lpString As Long) As Long

'private members
Private m_lngObjRefCount As Long
Private m_lpfnOldGetDisplayString As Long
Private m_lpfnOldGetPredefinedStrings As Long
Private m_lpfnOldGetPredefinedValue As Long
Private m_lpfnOldMapPropertyToPage As Long

Public Function IPerPropertyBrowsing_GetDisplayString(ByVal oThis As Object, ByVal DispID As Long, ByVal lpDisplayName As Long) As Long
'new vtable method for IPerPropertyBrowsing::GetDisplayString

    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    Dim strDisplayName As String
    Dim lpString As Long
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate passed pointer
    If VarPtr(lpDisplayName) = 0 Then
        IPerPropertyBrowsing_GetDisplayString = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetDisplayString bNoDefault, DispID, strDisplayName
        
    'if no param set by user
    If Not bNoDefault Then
        'return 'unimplemented' so container displays default
        IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
        
    Else
        'copy display string to passed ptr (caller should free the memory allocated)
        lpString = SysAllocString(StrPtr(strDisplayName))
        
        CopyMemory ByVal lpDisplayName, lpString, 4
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'return 'unimplemented' so container displays default
    IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
    
End Function
Public Function IPerPropertyBrowsing_MapPropertyToPage(ByVal oThis As Object, ByVal DispID As Long, lpCLSID As CLSID) As Long
'new vtable method for IPerPropertyBrowsing::MapPropertyToPage

    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    Dim strGUID As String
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate passed pointer
    If VarPtr(lpCLSID) = 0 Then
        IPerPropertyBrowsing_MapPropertyToPage = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.MapPropertyToPage bNoDefault, DispID, strGUID
        
    'if no param set by user
    If Not bNoDefault Then
        'return 'unimplemented' so container displays default
        IPerPropertyBrowsing_MapPropertyToPage = E_NOTIMPL
        
    Else
        'if valid string
        If Len(strGUID) > 2 Then
            'if not a GUID
            If Not (Left$(strGUID, 1) = "{" And Right$(strGUID, 1) = "}") Then
                'get CLSID from ProgID
                strGUID = FindGUIDForProgID(strGUID)
                
            End If
            'convert string CLSID to *real* CLSID
            lpCLSID = GUIDFromString(strGUID)
            
        End If
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'return 'unimplemented' so container displays default
    IPerPropertyBrowsing_MapPropertyToPage = E_NOTIMPL
    
End Function
Public Function IPerPropertyBrowsing_GetPredefinedStrings(ByVal oThis As Object, ByVal DispID As Long, pCaStringsOut As CALPOLESTR, pCaCookiesOut As CADWORD) As Long
'new vtable method for IPerPropertyBrowsing::GetPredefinedStrings

    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    
    Dim cElems As Long
    Dim pElems As Long
    Dim nElemCount As Integer
    Dim lpString As Long
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate passed pointers
    If VarPtr(pCaStringsOut) = 0 Or VarPtr(pCaCookiesOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedStrings = E_POINTER
        Exit Function
        
    End If
    
    'create & initialise cPropertyListItems collection
    Dim loProps As cPropertyListItems
    Set loProps = New cPropertyListItems
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetPredefinedStrings bNoDefault, DispID, loProps
    
    'if no param set by user
    If (Not bNoDefault) Or (loProps.Count = 0) Then
        'return 'unimplemented' so container displays default
        IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
        
    Else
        'initialise CALPOLESTR struct
        cElems = loProps.Count
        pElems = CoTaskMemAlloc(cElems * 4)
        pCaStringsOut.cElems = cElems
        pCaStringsOut.pElems = pElems
        
        
        Dim loItem As cPropertyListItem
        Dim lsTemp As String
        
        For Each loItem In loProps
            lpString = loItem.lpDisplayName
            
            CopyMemory ByVal pElems, lpString, 4
            'incr the element count
            pElems = pElems + 4
        Next
        
        'initialise CADWORD struct
        pElems = CoTaskMemAlloc(cElems * 4)
        pCaCookiesOut.cElems = cElems
        pCaCookiesOut.pElems = pElems
        
        'copy dwords to CADWORD struct
        For Each loItem In loProps
            CopyMemory ByVal pElems, loItem.Cookie, 4
            pElems = pElems + 4
        Next
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'return 'unimplemented' so container displays default
    IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
    
End Function
Public Function IPerPropertyBrowsing_GetPredefinedValue(ByVal oThis As Object, ByVal DispID As Long, ByVal dwCookie As Long, pVarOut As Variant) As Long
'new vtable method for IPerPropertyBrowsing::GetPredefinedValue

    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate passed pointers
    If VarPtr(dwCookie) = 0 Or VarPtr(pVarOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedValue = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetPredefinedValue bNoDefault, DispID, dwCookie, pVarOut
        
    'if no param set by user
    If Not bNoDefault Then
        'return 'unimplemented' so container displays default
        IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'return 'unimplemented' so container displays default
    IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
    
End Function
Public Sub ReplaceIPerPropertyBrowsing(ByVal pObject As Object)
'replace vtable for IPerPropertyBrowsing interface

    Dim oIPerPropertyBrowsing As vbACOMTLB.IPerPropertyBrowsing

    'if already hooked IPerPropertyBrowsing interface then done
    If m_lngObjRefCount > 0 Then
        m_lngObjRefCount = m_lngObjRefCount + 1
        'Debug.Print m_lngObjRefCount
        Exit Sub
        
    Else
        m_lngObjRefCount = 1
        
    End If
    
    'get ref to OLE IPerPropertyBrowsing interface
    Set oIPerPropertyBrowsing = pObject
    
    'replace vtable methods with our subclass procs
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    m_lpfnOldGetDisplayString = ReplaceVTableEntry(ObjPtr(oIPerPropertyBrowsing), 4, AddressOf IPerPropertyBrowsing_GetDisplayString) 'GetDisplayString
    m_lpfnOldMapPropertyToPage = ReplaceVTableEntry(ObjPtr(oIPerPropertyBrowsing), 5, AddressOf IPerPropertyBrowsing_MapPropertyToPage) 'MapPropertyToPage
    m_lpfnOldGetPredefinedStrings = ReplaceVTableEntry(ObjPtr(oIPerPropertyBrowsing), 6, AddressOf IPerPropertyBrowsing_GetPredefinedStrings) 'GetPredefinedStrings
    m_lpfnOldGetPredefinedValue = ReplaceVTableEntry(ObjPtr(oIPerPropertyBrowsing), 7, AddressOf IPerPropertyBrowsing_GetPredefinedValue) 'GetPredefinedValue
    
    'Debug.Print "Replaced vtable methods IPerPropertyBrowsing"
End Sub
Public Sub RestoreIPerPropertyBrowsing(ByVal lpObject As Long)
'restore vtable for IPerPropertyBrowsing interface

    Dim oObject As Object
    Dim oIPerPropertyBrowsing As vbACOMTLB.IPerPropertyBrowsing

    'if not last ref count then done
    If m_lngObjRefCount > 1 Then
        m_lngObjRefCount = m_lngObjRefCount - 1
        'Debug.Print m_lngObjRefCount
        Exit Sub
        
    Else
        m_lngObjRefCount = 0
        
    End If
    
    'get ref to object from ptr (no AddRef so don't set to nothing !)
    CopyMemory oObject, lpObject, 4
    
    'get ref to OLE IPerPropertyBrowsing interface
    Set oIPerPropertyBrowsing = oObject
    
    'delete uncounted reference
    CopyMemory oObject, 0&, 4
    
    'restore vtable methods with original procs
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ReplaceVTableEntry ObjPtr(oIPerPropertyBrowsing), 4, m_lpfnOldGetDisplayString 'GetDisplayString
    ReplaceVTableEntry ObjPtr(oIPerPropertyBrowsing), 5, m_lpfnOldMapPropertyToPage 'MapPropertyToPage
    ReplaceVTableEntry ObjPtr(oIPerPropertyBrowsing), 6, m_lpfnOldGetPredefinedStrings  'GetPredefinedStrings
    ReplaceVTableEntry ObjPtr(oIPerPropertyBrowsing), 7, m_lpfnOldGetPredefinedValue  'GetPredefinedValue
    
    'Debug.Print "Restored vtable methods IPerPropertyBrowsing"
End Sub
