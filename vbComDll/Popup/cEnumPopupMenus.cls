VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumPopupMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEnumVariantVB

Private mtEnum      As tEnum
Private moEnum      As cEnumVariantVB
Private moOwner     As cPopupMenus

Friend Sub fInit(ByVal oOwner As cPopupMenus, ByVal iControl As Long)
    Set moOwner = oOwner
    mtEnum.iIndex = -1&
    mtEnum.iControl = iControl
End Sub

Friend Function EnumObject() As IUnknown
    Set EnumObject = moEnum
End Function

Private Sub Class_Initialize()
    Set moEnum = New cEnumVariantVB
    moEnum.Attach Me
End Sub

Private Sub Class_Terminate()
    Set moEnum = Nothing
End Sub

Private Sub IEnumVariantVB_GetNextItem(vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.fEnum_NextMenu mtEnum, vNextItem, bNoMoreItems
End Sub

Private Sub IEnumVariantVB_Reset()
    mtEnum.iIndex = -1&
End Sub

Private Sub IEnumVariantVB_Skip(ByVal iSkipCount As Long)
    mtEnum.iIndex = mtEnum.iIndex + iSkipCount
End Sub
