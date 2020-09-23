VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOleHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Add(UserControl As Object) As Boolean
    Add = mCOMSupport.Attach(UserControl)
End Function

Public Function Remove(UserControl As Object) As Boolean
    Remove = mCOMSupport.Detach(UserControl)
End Function

Public Function Exists(UserControl As Object) As Boolean
    Exists = mCOMSupport.Exists(UserControl)
End Function
