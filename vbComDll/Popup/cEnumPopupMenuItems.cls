VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumPopupMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEnumVariantVB

Private mtPointer   As tMenuPointer
Private mtEnum      As tEnum
Private moEnum      As cEnumVariantVB
Private moOwner     As cPopupMenus



Friend Function fInit(ByVal oOwner As cPopupMenus, ByRef tPointer As tMenuPointer, ByVal iControl As Long)
    Set moOwner = oOwner
    LSet mtPointer = tPointer
    mtEnum.iControl = iControl
    mtEnum.iIndex = -1&
End Function

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
