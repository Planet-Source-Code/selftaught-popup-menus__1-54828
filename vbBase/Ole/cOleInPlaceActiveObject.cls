VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOleInPlaceActiveObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 1500
'////////////////////////////////////////////////////////////
'// Name : cOleInPlaceActiveObject.cls
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Wrapper for the COM IOleInPlaceActiveObject interface
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

'private members
Private m_oInternal As vbACOMTLB.IOleInPlaceActiveObject
Public Sub EnableModeless(ByVal Enable As Boolean)
'Enables or disables modeless dialog boxes
'when the container creates or destroys a modal dialog box.

    'call internal object's EnableModeless method
    m_oInternal.EnableModeless -CLng(Enable)
End Sub
Public Sub OnDocWindowActivate(ByVal Activate As Boolean)
'Notifies the active in-place object when
'the container's document window is activated or deactivated.

    'call internal object's OnDocWindowActivate method
    m_oInternal.OnDocWindowActivate -CLng(Activate)
End Sub
Public Sub OnFrameWindowActivate(ByVal Activate As Boolean)
'Notifies the object when the container's top-level
'frame window is activated or deactivated.

    'call internal object's OnFrameWindowActivate method
    m_oInternal.OnFrameWindowActivate -CLng(Activate)
End Sub
Public Sub ResizeBorder(ByRef Border As tRECT, ByVal UIWindow As cOleInPlaceUIWindow, ByVal FrameWindow As Boolean)
'Alerts the object that it needs to resize its border space.
    
    Dim lR As vbACOMTLB.RECT
    LSet lR = Border
    
    'call internal object's ResizeBorder method
    m_oInternal.ResizeBorder lR, UIWindow.GetInternalInterface(), -CLng(FrameWindow)
    
    LSet Border = lR
End Sub
Public Sub TranslateAccelerator(ByRef MSG As tMsg)
'Processes menu accelerator-key messages from
'the container's message queue.

    Dim lMsg As vbACOMTLB.MSG

    LSet lMsg = MSG

    'call internal object's TranslateAccelerator method
    m_oInternal.TranslateAccelerator lMsg
    
    LSet MSG = lMsg
End Sub
Friend Sub Attach(Source As vbACOMTLB.IOleInPlaceActiveObject)
'attach to IOleInPlaceActiveObject interface

    Set m_oInternal = Source
End Sub
Friend Sub Detach()
'detach from internal IOleInPlaceActiveObject interface

    Set m_oInternal = Nothing
End Sub
Public Function GetUnknown() As stdole.IUnknown
'return IUnknown interface of internal object

    Set GetUnknown = m_oInternal
End Function
Friend Function GetInternalInterface() As vbACOMTLB.IOleInPlaceActiveObject
'return ref to internal object

    Set GetInternalInterface = m_oInternal
End Function
Private Sub Class_Terminate()

    'handle all errors because VB doesn't propagate errors
    'back to caller during Terminate events, and so any errors
    'thrown here will be fatal to the calling app
    On Error Resume Next

    'make sure object is detached
    Detach
End Sub

