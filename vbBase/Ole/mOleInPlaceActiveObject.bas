Attribute VB_Name = "mOleInPlaceActiveObject"
'////////////////////////////////////////////////////////////
'// Name : modIOleInPlaceActiveObject.bas
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Custom implementation to make the IOleInPlaceActiveObject
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
Private m_lpfnOldTranslateAccelerator As Long
Private m_lpfnOldOnFrameWindowActivate As Long
Private m_lpfnOldOnDocWindowActivate As Long
Private m_lpfnOldResizeBorder As Long
Private m_lpfnOldEnableModeless As Long
Public Function IOleInPlaceActiveObject_OnDocWindowActivate(ByVal oThis As Object, ByVal fActive As Long) As Long
'new vtable method for IOleInPlaceActiveObject::OnDocWindowActivate

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'call custom implementation of 'OnDocWindowActivate'
    oIOleInPlaceActiveObjectVB.OnDocWindowActivate bNoDefault, IIf(fActive = 0, False, True)
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleInPlaceActiveObject_OnDocWindowActivate = Original_IOleInPlaceActiveObject_OnDocWindowActivate(oThis, fActive)
        
    Else
        'return 'OK'
        IOleInPlaceActiveObject_OnDocWindowActivate = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_OnDocWindowActivate = Original_IOleInPlaceActiveObject_OnDocWindowActivate(oThis, fActive)
    
End Function
Public Function IOleInPlaceActiveObject_EnableModeless(ByVal oThis As Object, ByVal fActive As Long) As Long
'new vtable method for IOleInPlaceActiveObject::EnableModeless

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'call custom implementation of 'EnableModeless'
    oIOleInPlaceActiveObjectVB.EnableModeless bNoDefault, IIf(fActive = 0, False, True)
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleInPlaceActiveObject_EnableModeless = Original_IOleInPlaceActiveObject_EnableModeless(oThis, fActive)
        
    Else
        'return 'OK'
        IOleInPlaceActiveObject_EnableModeless = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_EnableModeless = Original_IOleInPlaceActiveObject_EnableModeless(oThis, fActive)
    
End Function
Public Function IOleInPlaceActiveObject_ResizeBorder(ByVal oThis As Object, prcBorder As tRECT, ByVal oUIWindow As vbACOMTLB.IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
'new vtable method for IOleInPlaceActiveObject::ResizeBorder

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim bNoDefault As Boolean
    
    Dim oOleInPlaceUIWindow As New cOleInPlaceUIWindow
   
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'create & initialise IOleInPlaceUIWindow wrapper object
    oOleInPlaceUIWindow.Attach oUIWindow
    
    'call custom implementation of 'ResizeBorder'
    oIOleInPlaceActiveObjectVB.ResizeBorder bNoDefault, prcBorder, oOleInPlaceUIWindow, IIf(fFrameWindow = 0, False, True)
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleInPlaceActiveObject_ResizeBorder = Original_IOleInPlaceActiveObject_ResizeBorder(oThis, prcBorder, oUIWindow, fFrameWindow)
        
    Else
        'return 'OK'
        IOleInPlaceActiveObject_ResizeBorder = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_ResizeBorder = Original_IOleInPlaceActiveObject_ResizeBorder(oThis, prcBorder, oUIWindow, fFrameWindow)
    
End Function
Public Function IOleInPlaceActiveObject_OnFrameWindowActivate(ByVal oThis As Object, ByVal fActive As Long) As Long
'new vtable method for IOleInPlaceActiveObject::OnFrameWindowActivate

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'call custom implementation of 'OnFrameWindowActivate'
    oIOleInPlaceActiveObjectVB.OnFrameWindowActivate bNoDefault, IIf(fActive = 0, False, True)
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleInPlaceActiveObject_OnFrameWindowActivate = Original_IOleInPlaceActiveObject_OnFrameWindowActivate(oThis, fActive)
        
    Else
        'return 'OK'
        IOleInPlaceActiveObject_OnFrameWindowActivate = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_OnFrameWindowActivate = Original_IOleInPlaceActiveObject_OnFrameWindowActivate(oThis, fActive)
    
End Function
Public Function IOleInPlaceActiveObject_TranslateAccelerator(ByVal oThis As Object, pMsg As tMsg) As Long
'new vtable method for IOleInPlaceActiveObject::TranslateAccelerator

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim nShift As Integer
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate params
    If VarPtr(pMsg) = 0 Then
        IOleInPlaceActiveObject_TranslateAccelerator = E_POINTER
        Exit Function
        
    End If
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'get status of modifier keys
    nShift = GetKeyModifiers()
    
    'call custom implementation of 'TranslateAccelerator'
    oIOleInPlaceActiveObjectVB.TranslateAccelerator bNoDefault, pMsg, nShift
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleInPlaceActiveObject_TranslateAccelerator = Original_IOleInPlaceActiveObject_TranslateAccelerator(oThis, pMsg)
        
    Else
        
        'return 'OK'
        IOleInPlaceActiveObject_TranslateAccelerator = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_TranslateAccelerator = Original_IOleInPlaceActiveObject_TranslateAccelerator(oThis, pMsg)
    
End Function
Private Function Original_IOleInPlaceActiveObject_TranslateAccelerator(ByVal oThis As Object, pMsg As tMsg) As Long
'call original 'TranslateAccelerator' method
    
    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject
    
    'do implicit QI to get IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 6, m_lpfnOldTranslateAccelerator 'TranslateAccelerator

    'call the original method
    On Error Resume Next
    Dim lMsg As vbACOMTLB.MSG
    
    LSet lMsg = pMsg
    
    oIOleInPlaceActiveObject.TranslateAccelerator lMsg
    
    LSet pMsg = lMsg
    
    Original_IOleInPlaceActiveObject_TranslateAccelerator = MapCOMErr(Err.Number)
    
    
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldTranslateAccelerator = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 6, AddressOf IOleInPlaceActiveObject_TranslateAccelerator) 'TranslateAccelerator
End Function
Private Function Original_IOleInPlaceActiveObject_OnDocWindowActivate(ByVal oThis As Object, ByVal fActive As Long) As Long
'call original 'OnDocWindowActivate' method
    
    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject
    
    'do implicit QI to get IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 8, m_lpfnOldOnDocWindowActivate 'OnDocWindowActivate

    'call the original method
    On Error Resume Next
    oIOleInPlaceActiveObject.OnDocWindowActivate fActive
    Original_IOleInPlaceActiveObject_OnDocWindowActivate = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldOnDocWindowActivate = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 8, AddressOf IOleInPlaceActiveObject_OnDocWindowActivate) 'OnDocWindowActivate
End Function
Private Function Original_IOleInPlaceActiveObject_EnableModeless(ByVal oThis As Object, ByVal fActive As Long) As Long
'call original 'EnableModeless' method
    
    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject
    
    'do implicit QI to get IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 10, m_lpfnOldEnableModeless 'EnableModeless

    'call the original method
    On Error Resume Next
    oIOleInPlaceActiveObject.EnableModeless fActive
    Original_IOleInPlaceActiveObject_EnableModeless = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldEnableModeless = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 10, AddressOf IOleInPlaceActiveObject_EnableModeless) 'EnableModeless
End Function
Private Function Original_IOleInPlaceActiveObject_ResizeBorder(ByVal oThis As Object, prcBorder As tRECT, ByVal oUIWindow As vbACOMTLB.IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
'call original 'ResizeBorder' method
    
    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject
    
    'do implicit QI to get IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 9, m_lpfnOldResizeBorder 'ResizeBorder

    'call the original method
    On Error Resume Next
    
    Dim lR As vbACOMTLB.RECT
    LSet lR = prcBorder
    
    oIOleInPlaceActiveObject.ResizeBorder lR, oUIWindow, fFrameWindow
    
    LSet prcBorder = lR
    
    Original_IOleInPlaceActiveObject_ResizeBorder = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldResizeBorder = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 9, AddressOf IOleInPlaceActiveObject_ResizeBorder) 'ResizeBorder
End Function
Private Function Original_IOleInPlaceActiveObject_OnFrameWindowActivate(ByVal oThis As Object, ByVal fActive As Long) As Long
'call original 'OnFrameWindowActivate' method
    
    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject
    
    'do implicit QI to get IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = oThis
    
    'temporarily unhook method so we can call the original
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 7, m_lpfnOldOnFrameWindowActivate 'OnFrameWindowActivate

    'call the original method
    On Error Resume Next
    oIOleInPlaceActiveObject.OnFrameWindowActivate fActive
    Original_IOleInPlaceActiveObject_OnFrameWindowActivate = MapCOMErr(Err.Number)
    On Error GoTo 0
    
    're-hook the method
    m_lpfnOldOnFrameWindowActivate = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 7, AddressOf IOleInPlaceActiveObject_OnFrameWindowActivate) 'OnFrameWindowActivate
End Function
Public Sub ReplaceIOleInPlaceActiveObject(ByVal pObject As Object)
'replace vtable for IOleInPlaceActiveObject interface

    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject

    'if already hooked IOleInPlaceActiveObject interface then done
    If m_lngObjRefCount > 0 Then
        m_lngObjRefCount = m_lngObjRefCount + 1
        'Debug.Print m_lngObjRefCount
        Exit Sub
        
    Else
        m_lngObjRefCount = 1
        
    End If
    
    'get ref to OLE IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = pObject
    
    'replace vtable methods with our subclass procs
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ' Ignore item 4: GetWindow
    ' Ignore item 5: ContextSensitiveHelp
    m_lpfnOldTranslateAccelerator = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 6, AddressOf IOleInPlaceActiveObject_TranslateAccelerator) 'TranslateAccelerator
    m_lpfnOldOnFrameWindowActivate = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 7, AddressOf IOleInPlaceActiveObject_OnFrameWindowActivate) 'OnFrameWindowActivate
    m_lpfnOldOnDocWindowActivate = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 8, AddressOf IOleInPlaceActiveObject_OnDocWindowActivate) 'OnDocWindowActivate
    m_lpfnOldResizeBorder = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 9, AddressOf IOleInPlaceActiveObject_ResizeBorder) 'ResizeBorder
    m_lpfnOldEnableModeless = ReplaceVTableEntry(ObjPtr(oIOleInPlaceActiveObject), 10, AddressOf IOleInPlaceActiveObject_EnableModeless) 'EnableModeless
    
    'Debug.Print "Replaced vtable methods IOleInPlaceActiveObject"
End Sub
Public Sub RestoreIOleInPlaceActiveObject(ByVal lpObject As Long)
'restore vtable for IOleInPlaceActiveObject interface

    Dim oObject As Object
    Dim oIOleInPlaceActiveObject As vbACOMTLB.IOleInPlaceActiveObject

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
    
    'get ref to OLE IOleInPlaceActiveObject interface
    Set oIOleInPlaceActiveObject = oObject
    
    'delete uncounted reference
    CopyMemory oObject, 0&, 4
    
    'restore vtable methods with original procs
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ' Ignore item 4: GetWindow
    ' Ignore item 5: ContextSensitiveHelp
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 6, m_lpfnOldTranslateAccelerator 'TranslateAccelerator
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 7, m_lpfnOldOnFrameWindowActivate 'OnFrameWindowActivate
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 8, m_lpfnOldOnDocWindowActivate 'OnDocWindowActivate
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 9, m_lpfnOldResizeBorder 'ResizeBorder
    ReplaceVTableEntry ObjPtr(oIOleInPlaceActiveObject), 10, m_lpfnOldEnableModeless 'EnableModeless
    
    'Debug.Print "Restored vtable methods IOleInPlaceActiveObject"
End Sub
