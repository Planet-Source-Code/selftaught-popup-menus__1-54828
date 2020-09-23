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
'==================================================================================================
'cEnumVariantVB.cls                     7/5/04
'
'           LINEAGE:
'               Based on a class in the vbACOM component developed by Paul Wilde for vbaccelerator.com.
'
'           GENERAL PURPOSE:
'               Implements the iEnumVARIANT Interface so that clients can provide collections object
'               without referencing a type library.
'
'==================================================================================================

Option Explicit

Implements vbACOMTLB.IEnumVARIANT

'private members
Private moClient As iEnumVariantVB

Private Sub Class_Initialize()
    'replace interface's vtable procs with our own
    ReplaceIEnumVARIANT Me
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
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
    If Not moClient Is Nothing Then moClient.Reset
End Sub

Private Sub IEnumVARIANT_Skip(ByVal celt As Long)
    If Not moClient Is Nothing Then moClient.Skip celt
End Sub

Public Sub Attach(Connect As iEnumVariantVB)
    Set moClient = Connect
End Sub
Friend Sub GetNextItem(vNextItem As Variant, bNoMore As Boolean)
    If Not moClient Is Nothing Then moClient.GetNextItem vNextItem, bNoMore
End Sub
Friend Sub Skip(ByVal SkipCount As Long)
    If Not moClient Is Nothing Then moClient.Skip SkipCount
End Sub
