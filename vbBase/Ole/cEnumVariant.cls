VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumVariantVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 400
'////////////////////////////////////////////////////////////
'// Name : pcEnumVariant.cls
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Generic enumerator object for variants
'// Used as a delegated object for clients that implement IEnumVARIANTVB
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

Implements vbACOMTLB.IEnumVARIANT

'private members
Private m_oClient As IEnumVariantVB


Public Sub Attach(Connect As IEnumVariantVB)
'attach this enumerator to client interface
    
    'make sure not already attached
    If Not m_oClient Is Nothing Then
        'detach first
        Detach
        
    End If
    
    'attach to client
    Set m_oClient = Connect
End Sub
Friend Sub GetNextItem(vNextItem As Variant, bNoMore As Boolean)
'call client's NextItem method
    
    If Not m_oClient Is Nothing Then
         m_oClient.GetNextItem vNextItem, bNoMore

    End If
End Sub
Friend Sub Skip(ByVal SkipCount As Long)
'call client's Skip method
    
    If Not m_oClient Is Nothing Then
        m_oClient.Skip SkipCount

    End If
End Sub
Public Sub Detach()
'detach this enumerator from client interface
    
    If Not m_oClient Is Nothing Then Set m_oClient = Nothing
    
End Sub

Private Sub Class_Initialize()

    'replace interface's vtable procs with our own
    ReplaceIEnumVARIANT Me
End Sub
Private Sub Class_Terminate()

    'handle all errors because VB doesn't propagate errors
    'back to caller during Terminate events, and so any errors
    'thrown here will be fatal to the calling app
    On Error Resume Next
    
    'make sure enumerator is detached from client interface
    Detach

    'restore interface's vtable procs
    RestoreIEnumVARIANT ObjPtr(Me)
End Sub
Private Sub IEnumVARIANT_Clone(ppenum As vbACOMTLB.IEnumVARIANT)
'not implemented

    Err.Raise E_NOTIMPL
End Sub
Private Sub IEnumVARIANT_Next(ByVal celt As Long, rgVar As Variant, ByVal pceltFetched As Long)
'dummy entry point, will be replaced in the
'vtable by IEnumVARIANT_Next
End Sub

Private Sub IEnumVARIANT_Reset()
'call client's Reset method
    
    If Not m_oClient Is Nothing Then
        m_oClient.Reset

    End If
End Sub

Private Sub IEnumVARIANT_Skip(ByVal celt As Long)
'call client's Skip method
    
    If Not m_oClient Is Nothing Then
        m_oClient.Skip celt

    End If
End Sub

