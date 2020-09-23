VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cEnumeration.cls                       7/4/04
'
'           GENERAL PURPOSE:
'               Holds a placemarker and other client-specific data for an enumeration.
'
'==================================================================================================


Implements iEnumVariantVB

Public Type tEnum
    iIndex      As Long
    iControl    As Long
    iData       As Long
End Type

Private mtEnum As tEnum
Private moOwner As iEnumerable

Private moEnum As cEnumVariantVB

Public Function GetEnum( _
            ByVal oWho As iEnumerable, _
            ByVal iControl As Long, _
   Optional ByVal iData As Long, _
   Optional ByVal iIndex As Long = -1&) _
                As IUnknown
                
    With mtEnum
        .iIndex = iIndex
        .iControl = iControl
        .iData = iData
    End With
    Set moOwner = oWho
    Set GetEnum = moEnum
End Function


Private Sub Class_Initialize()
    Set moEnum = New cEnumVariantVB
    moEnum.Attach Me
End Sub

Private Sub Class_Terminate()
    Set moEnum = Nothing
End Sub

Private Sub iEnumVariantVB_GetNextItem(vNextItem As Variant, bNoMoreItems As Boolean)
    moOwner.GetNextItem mtEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumVariantVB_Reset()
    moOwner.Reset mtEnum
End Sub

Private Sub iEnumVariantVB_Skip(ByVal iSkipCount As Long)
    moOwner.Skip mtEnum, iSkipCount
End Sub
