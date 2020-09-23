VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropertyListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2040
Option Explicit

'==================================================================================================
'cPropertyListItem.cls                  7/5/04
'
'           GENERAL PURPOSE:
'               R/W access to name and cookie value for Properties through IPerPropertyBrowsing
'
'==================================================================================================


Private Declare Function CoTaskMemAlloc Lib "OLE32" (ByVal cBytes As Long) As Long

Private msString As String
Private miCookie As Long

Public Property Get Cookie() As Long
    Cookie = miCookie
End Property
Public Property Let Cookie(ByVal NewVal As Long)
    miCookie = NewVal
End Property
Public Property Get DisplayName() As String
    DisplayName = msString
End Property
Public Property Let DisplayName(ByVal NewVal As String)
    msString = NewVal
End Property

Friend Function lpDisplayName() As Long
    Dim liLen As Long
    liLen = LenB(msString) + 1
    lpDisplayName = CoTaskMemAlloc(liLen)
    CopyMemory ByVal lpDisplayName, ByVal StrPtr(msString & vbNullChar), liLen
End Function

