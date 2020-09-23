VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iOleInPlaceActiveObjectVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2530
'////////////////////////////////////////////////////////////
'// Name : IOleInPlaceActiveObjectVB.cls
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// 'VB friendly' implementation of COM IOleInPlaceActiveObject interface
'// Use with Implements
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit



Public Sub OnFrameWindowActivate(ByRef bHandled As Boolean, ByVal bActive As Boolean)
End Sub
Public Sub OnDocWindowActivate(ByRef bHandled As Boolean, ByVal bActive As Boolean)
End Sub
Public Sub EnableModeless(ByRef bHandled As Boolean, ByVal bEnable As Boolean)
End Sub
Public Sub ResizeBorder(ByRef bHandled As Boolean, ByRef tBorder As tRECT, ByVal oUIWindow As cOleInPlaceUIWindow, ByVal bFrameWindow As Boolean)
End Sub
Public Sub TranslateAccelerator(ByRef bHandled As Boolean, ByRef tMsg As tMsg, ByVal iShift As ShiftConstants)
End Sub

