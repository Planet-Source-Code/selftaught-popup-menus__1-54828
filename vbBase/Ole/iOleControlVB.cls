VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iOleControlVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2480
'////////////////////////////////////////////////////////////
'// Name : IOleControlVB.cls
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// 'VB friendly' implementation of COM IOleControl interface
'// Use with Implements
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit


Public Sub FreezeEvents(ByRef bHandled As Boolean, ByVal bFreeze As Boolean)
End Sub
Public Sub GetControlInfo(ByRef bHandled As Boolean, ByRef iAccelCount As Long, ByRef hAccelTable As Long, ByRef iFlags As eControlInfoFlags)
End Sub
Public Sub OnAmbientPropertyChange(ByRef bHandled As Boolean, ByVal iAmbientID As eAmbientDispID)
End Sub
Public Sub OnMnemonic(ByRef bHandled As Boolean, ByRef tMsg As tMsg, ByVal iShift As ShiftConstants)
End Sub

