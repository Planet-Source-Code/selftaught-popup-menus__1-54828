VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOleInPlaceUIWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 1750
'////////////////////////////////////////////////////////////
'// Name : cOleInPlaceUIWindow.cls
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Wrapper for the COM IOleInPlaceUIWindow interface
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

'private members
Private m_oInternal As vbACOMTLB.IOleInPlaceUIWindow
Public Function GetBorder() As tRect
'get RECT in which client can place toolbars etc. while active in place

    'call internal object's GetBorder method
    m_oInternal.GetBorder ByVal VarPtr(GetBorder)

End Function
Public Sub RequestBorderSpace(ByRef BorderWidths As tRect)
'ask if there is any available border space in which client can put in place active components

    'call internal object's RequestBorderSpace method
    m_oInternal.RequestBorderSpace ByVal VarPtr(BorderWidths)
End Sub
Public Sub SetBorderSpace(ByRef BorderWidths As tRect)
'Allocates space for the border requested in the call to RequestBorderSpace

    'call internal object's SetBorderSpace method
    m_oInternal.SetBorderSpace ByVal VarPtr(BorderWidths)

End Sub
Public Sub SetActiveObject(ByVal ActiveObject As cOleInPlaceActiveObject, ByVal ObjName As String)
'alert object of need to resize border space
    
    'call internal object's SetActiveObject method
    m_oInternal.SetActiveObject ActiveObject.GetInternalInterface(), ObjName
End Sub
Friend Sub Attach(Source As vbACOMTLB.IOleInPlaceUIWindow)
'attach to IOleInPlaceUIWindow interface

    Set m_oInternal = Source
End Sub
Friend Sub Detach()
'detach from internal IOleInPlaceUIWindow interface

    Set m_oInternal = Nothing
End Sub
Public Function GetUnknown() As stdole.IUnknown
'return IUnknown interface of internal object

    Set GetUnknown = m_oInternal
End Function
Friend Function GetInternalInterface() As vbACOMTLB.IOleInPlaceUIWindow
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

