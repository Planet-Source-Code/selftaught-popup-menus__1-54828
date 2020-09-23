VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iPerPropertyBrowsingVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2650
'////////////////////////////////////////////////////////////
'// Name : IPerPropertyBrowsingVB.cls
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// 'VB friendly' implementation of COM IPerPropertyBrowsing interface
'// Use with Implements
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

Public Sub GetPredefinedValue(ByRef bHandled As Boolean, ByVal iDispID As Long, ByVal iCookie As Long, vValue As Variant)
End Sub
Public Sub GetDisplayString(ByRef bHandled As Boolean, ByVal iDispID As Long, sDisplayName As String)
End Sub
Public Sub MapPropertyToPage(ByRef bHandled As Boolean, ByVal iDispID As Long, sClassID As String)
End Sub
Public Sub GetPredefinedStrings(ByRef bHandled As Boolean, ByVal iDispID As Long, ByVal oProperties As cPropertyListItems)
End Sub

