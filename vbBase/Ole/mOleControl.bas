Attribute VB_Name = "mOleControl"
'////////////////////////////////////////////////////////////
'// Name : modIOleControl.bas
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Custom implementation to make the IOleControl
'// interface more 'VB Friendly'
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

'private members
Private m_lngObjRefCount As Long
Private m_lpfnOldOnAmbientPropertyChange As Long
Private m_lpfnOldGetControlInfo As Long
Private m_lpfnOldOnMnemonic As Long
Private m_lpfnOldFreezeEvents As Long

Public Sub ReplaceIOleControl(ByVal pObject As Object)
'replace vtable for IOleControl interface

    Dim oIOleControl As vbACOMTLB.IOleControl

    'if already hooked IOleControl interface then done
    If m_lngObjRefCount > 0 Then
        m_lngObjRefCount = m_lngObjRefCount + 1
        Exit Sub
        
    Else
        m_lngObjRefCount = 1
        
    End If
    
    'get ref to OLE IOleControl interface
    Set oIOleControl = pObject
    
    'replace vtable methods with our subclass procs
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    m_lpfnOldGetControlInfo = ReplaceVTableEntry(ObjPtr(oIOleControl), 4, AddressOf IOleControl_GetControlInfo) 'GetControlInfo
    m_lpfnOldOnMnemonic = ReplaceVTableEntry(ObjPtr(oIOleControl), 5, AddressOf IOleControl_OnMnemonic) 'OnMnemonic
    m_lpfnOldOnAmbientPropertyChange = ReplaceVTableEntry(ObjPtr(oIOleControl), 6, AddressOf IOleControl_OnAmbientPropertyChange) 'OnAmbientPropertyChange
    m_lpfnOldFreezeEvents = ReplaceVTableEntry(ObjPtr(oIOleControl), 7, AddressOf IOleControl_FreezeEvents) 'FreezeEvents
    
    'Debug.Print "Replaced vtable methods IOleControl"
End Sub
Public Sub RestoreIOleControl(ByVal lpObject As Long)
'restore vtable for IOleControl interface

    Dim oObject As Object
    Dim oIOleControl As vbACOMTLB.IOleControl

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
    
    'get ref to OLE IOleControl interface
    Set oIOleControl = oObject
    
    'delete uncounted reference
    CopyMemory oObject, 0&, 4
    
    'restore vtable methods with original procs
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ReplaceVTableEntry ObjPtr(oIOleControl), 4, m_lpfnOldGetControlInfo 'GetControlInfo
    ReplaceVTableEntry ObjPtr(oIOleControl), 5, m_lpfnOldOnMnemonic 'OnMnemonic
    ReplaceVTableEntry ObjPtr(oIOleControl), 6, m_lpfnOldOnAmbientPropertyChange 'OnAmbientPropertyChange
    ReplaceVTableEntry ObjPtr(oIOleControl), 7, m_lpfnOldFreezeEvents 'FreezeEvents
    
    'Debug.Print "Restored vtable methods IOleControl"
End Sub
Public Function IOleControl_OnAmbientPropertyChange(ByVal oThis As Object, ByVal DispID As Long) As Long
'new vtable method for IOleControl::OnAmbientPropertyChange

    Dim oIOleControlVB As iOleControlVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'call custom implementation of 'OnAmbientPropertyChange'
    oIOleControlVB.OnAmbientPropertyChange bNoDefault, DispID
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_OnAmbientPropertyChange = Original_IOleControl_OnAmbientPropertyChange(oThis, DispID)
        
    Else
        'return 'OK'
        IOleControl_OnAmbientPropertyChange = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_OnAmbientPropertyChange = Original_IOleControl_OnAmbientPropertyChange(oThis, DispID)
    
End Function
Public Function IOleControl_FreezeEvents(ByVal oThis As Object, ByVal fFreeze As Long) As Long
'new vtable method for IOleControl::FreezeEvents

    Dim oIOleControlVB As iOleControlVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'call custom implementation of 'FreezeEvents'
    oIOleControlVB.FreezeEvents bNoDefault, IIf(fFreeze = 0, False, True)
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_FreezeEvents = Original_IOleControl_FreezeEvents(oThis, fFreeze)
        
    Else
        'return 'OK'
        IOleControl_FreezeEvents = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_FreezeEvents = Original_IOleControl_FreezeEvents(oThis, fFreeze)

End Function
Public Function IOleControl_OnMnemonic(ByVal oThis As Object, pMsg As tMsg) As Long
'new vtable method for IOleControl::OnMnemonic

    Dim oIOleControlVB As iOleControlVB
    Dim nShift As Integer
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate params
    If VarPtr(pMsg) = 0 Then
        IOleControl_OnMnemonic = E_POINTER
        Exit Function
        
    End If
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'get status of modifier keys
    nShift = GetKeyModifiers()
    
    'call custom implementation of 'OnMnemonic'
    oIOleControlVB.OnMnemonic bNoDefault, pMsg, nShift
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(oThis, pMsg)
    
    Else
        'return 'OK'
        IOleControl_OnMnemonic = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(oThis, pMsg)
    
End Function
Public Function IOleControl_GetControlInfo(ByVal oThis As Object, pCI As CONTROLINFO) As Long
'new vtable method for IOleControl::GetControlInfo

    Dim oIOleControlVB As iOleControlVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate params
    If VarPtr(pCI) = 0 Then
        IOleControl_GetControlInfo = E_POINTER
        Exit Function
        
    End If
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'call custom implementation of 'GetControlInfo'
    pCI.cb = LenB(pCI)
    
    oIOleControlVB.GetControlInfo bNoDefault, pCI.cAccel, pCI.hAccel, pCI.dwFlags
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(oThis, pCI)
    
    Else
        
        'if array contains items but mem handle is 0 then problem
        If pCI.cAccel > 0 And pCI.hAccel = 0 Then
            'return 'out of memory' error
            IOleControl_GetControlInfo = E_OUTOFMEMORY
        
        Else
            'return 'OK'
            IOleControl_GetControlInfo = S_OK
        
        End If
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(oThis, pCI)
    
End Function
Private Function Original_IOleControl_OnAmbientPropertyChange(ByVal oThis As Object, ByVal DispID As Long) As Long
'call original 'OnAmbientPropertyChange' method
    
    Dim oIOleControl As vbACOMTLB.IOleControl
    
    'do implicit QI to get IOleControl interface
    Set oIOleControl = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleControl), 6, m_lpfnOldOnAmbientPropertyChange 'OnAmbientPropertyChange

    'call the original method
    On Error Resume Next
    oIOleControl.OnAmbientPropertyChange DispID
    Original_IOleControl_OnAmbientPropertyChange = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldOnAmbientPropertyChange = ReplaceVTableEntry(ObjPtr(oIOleControl), 6, AddressOf IOleControl_OnAmbientPropertyChange) 'OnAmbientPropertyChange
End Function
Private Function Original_IOleControl_FreezeEvents(ByVal oThis As Object, ByVal fFreeze As Long) As Long
'call original 'FreezeEvents' method
    
    Dim oIOleControl As vbACOMTLB.IOleControl
    
    'do implicit QI to get IOleControl interface
    Set oIOleControl = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleControl), 7, m_lpfnOldFreezeEvents 'FreezeEvents

    'call the original method
    On Error Resume Next
    oIOleControl.FreezeEvents fFreeze
    Original_IOleControl_FreezeEvents = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldFreezeEvents = ReplaceVTableEntry(ObjPtr(oIOleControl), 7, AddressOf IOleControl_FreezeEvents) 'FreezeEvents
End Function
Private Function Original_IOleControl_GetControlInfo(ByVal oThis As Object, pCI As CONTROLINFO) As Long
'call original 'GetControlInfo' method
    
    Dim oIOleControl As vbACOMTLB.IOleControl
    
    'do implicit QI to get IOleControl interface
    Set oIOleControl = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleControl), 4, m_lpfnOldGetControlInfo 'GetControlInfo

    'call the original method
    On Error Resume Next
    oIOleControl.GetControlInfo pCI
    Original_IOleControl_GetControlInfo = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldGetControlInfo = ReplaceVTableEntry(ObjPtr(oIOleControl), 4, AddressOf IOleControl_GetControlInfo) 'GetControlInfo
End Function
Private Function Original_IOleControl_OnMnemonic(ByVal oThis As Object, pMsg As tMsg) As Long
'call original 'OnMnemonic' method
    
    Dim oIOleControl As vbACOMTLB.IOleControl
    
    'do implicit QI to get IOleControl interface
    Set oIOleControl = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleControl), 5, m_lpfnOldOnMnemonic 'OnMnemonic

    'call the original method
    On Error Resume Next
    Dim lMsg As vbACOMTLB.MSG
    
    LSet lMsg = pMsg
    oIOleControl.OnMnemonic lMsg
    LSet pMsg = lMsg
    
    Original_IOleControl_OnMnemonic = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldOnMnemonic = ReplaceVTableEntry(ObjPtr(oIOleControl), 5, AddressOf IOleControl_OnMnemonic) 'OnMnemonic
End Function
