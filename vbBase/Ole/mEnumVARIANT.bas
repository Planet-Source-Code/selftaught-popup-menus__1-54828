Attribute VB_Name = "mEnumVariant"
#Const bVBVMTypeLib = True

'////////////////////////////////////////////////////////////
'// Name : modIEnumVARIANT.bas
'// Author : Paul R. Wilde
'// Created : 23rd April 1999
'/////////////////////////////////////////////////////////////
'// Copyright © Paul R. Wilde 1999. All Rights Reserved.
'/////////////////////////////////////////////////////////////
'// Bug reports, suggestions & comments should be emailed to :
'// prw.exponential@dial.pipex.com
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Custom implementation to make the IEnumVARIANT
'// interface more 'VB Friendly'
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Adapted from IEnumVARIANT code in 'VBCore' by Bruce McKinney
'/////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'// Revision history
'/////////////////////////////////////////////////////////////
'// 23/04/99
'// Initial development
'/////////////////////////////////////////////////////////////

Option Explicit

Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

'Now we love speed, don't we??
Private Type SafeArray1D
  cDims       As Integer
  fFeatures   As Integer
  cbElements  As Long
  cLocks      As Long
  pvData      As Long
  cElements   As Long
  lLBound     As Long
End Type

Private mtSAHeader      As SafeArray1D
Private mvArray()       As Variant 'never dimensioned, accesses memory already allocated

'private members
Public m_lngObjRefCount As Long
Private m_lpfnOldNext   As Long
Private m_lpfnOldSkip   As Long

Public Sub ReplaceIEnumVARIANT(ByVal oObject As Object)
'replace vtable for IEnumVARIANT interface

    Dim oIEnumVARIANT As vbACOMTLB.IEnumVARIANT

    'if already done IEnumVARIANT interface then done
    If m_lngObjRefCount > 0 Then
        m_lngObjRefCount = m_lngObjRefCount + 1
        'Debug.Print m_lngObjRefCount
        Exit Sub
        
    Else
        m_lngObjRefCount = 1
        
    End If
    
    'get ref to OLE IEnumVARIANT interface
    Set oIEnumVARIANT = oObject
    
'    'replace vtable methods with our subclass procs
'    ' Ignore item 1: QueryInterface
'    ' Ignore item 2: AddRef
'    ' Ignore item 3: Release
    m_lpfnOldNext = ReplaceVTableEntry(ObjPtr(oIEnumVARIANT), 4, AddressOf IEnumVARIANT_Next) 'Next
    m_lpfnOldSkip = ReplaceVTableEntry(ObjPtr(oIEnumVARIANT), 5, AddressOf IEnumVARIANT_Skip) 'Skip
'    m_lpfnOldReset = ReplaceVtableEntry(ObjPtr(oIEnumVARIANT), 6, AddressOf IEnumVARIANT_Reset) 'Reset
'    m_lpfnOldClone = ReplaceVTableEntry(ObjPtr(oIEnumVARIANT), 7, AddressOf IEnumVARIANT_Clone) 'Clone
    
    'Debug.Print "Replaced vtable methods IEnumVARIANT"
End Sub
Public Sub RestoreIEnumVARIANT(ByVal lpObject As Long)
'restore vtable for IEnumVARIANT interface

    Dim oObject As Object
    Dim oIEnumVARIANT As vbACOMTLB.IEnumVARIANT

    'if not last ref count then done
    If m_lngObjRefCount > 1 Then
        m_lngObjRefCount = m_lngObjRefCount - 1
        'Debug.Print m_lngObjRefCount
        Exit Sub
        
    Else
        m_lngObjRefCount = 0
        
    End If
    
    'get ref to object from ptr (no AddRef so don't set to nothing !)
    CopyMemory oObject, lpObject, 4
    
    'get ref to OLE IEnumVARIANT interface
    Set oIEnumVARIANT = oObject
    
    'delete uncounted reference
    CopyMemory oObject, 0&, 4
    
'    'restore vtable methods with original procs
'    ' Ignore item 1: QueryInterface
'    ' Ignore item 2: AddRef
'    ' Ignore item 3: Release
    ReplaceVTableEntry ObjPtr(oIEnumVARIANT), 4, m_lpfnOldNext 'Next
    ReplaceVTableEntry ObjPtr(oIEnumVARIANT), 5, m_lpfnOldSkip  'Skip
'    ReplaceVtableEntry ObjPtr(oIEnumVARIANT), 6, m_lpfnOldReset 'Reset
'    ReplaceVTableEntry ObjPtr(oIEnumVARIANT), 7, m_lpfnOldClone  'Clone
    
    'Debug.Print "Restored vtable methods IEnumVARIANT"
End Sub
Public Function IEnumVARIANT_Next(ByVal oThis As Object, ByVal lngVntCount As Long, vntArray As Variant, ByVal pcvFetched As Long) As Long
'new vtable method for IEnumVARIANT::Next

    Dim oEnumVARIANT As cEnumVariantVB
    Dim cvFetched As Long, bNoMore As Boolean
    Dim nCount As Integer
    
    'handle all errors so we don't crash caller
    On Error Resume Next
    
    pInitArray VarPtr(vntArray), lngVntCount
    
    'cast method to source interface
    Set oEnumVARIANT = oThis
    
    'loop through each requested variant
    For nCount = 0 To lngVntCount - 1&
        'call the class method that raises a Next event--it returns
        'true if the next value is fetched
        oEnumVARIANT.GetNextItem mvArray(nCount), bNoMore

        'if failure or nothing fetched, we're done
        If (Err) Or (bNoMore) Then Exit For

        ' Count the variant and point to the next one
        cvFetched = cvFetched + 1
    Next
    'ff error caused termination, undo what we did
    If Err.Number <> 0 Then
        'iterate back, emptying the invalid fetched variants
        For nCount = nCount To 0 Step -1
            mvArray(nCount) = Empty
        Next nCount
        'convert error to COM format
        IEnumVARIANT_Next = MapCOMErr(Err)
        'return 0 as the number fetched after error
        If pcvFetched Then
            #If bVBVMTypeLib Then
                MemLong(ByVal pcvFetched) = cvFetched
            #Else
                CopyMemory ByVal pcvFetched, 0&, 4
            #End If
        End If
        
    Else
        'if nothing fetched, break out of enumeration
        If cvFetched = 0 Then
            IEnumVARIANT_Next = 1
            
        End If
        'copy the actual number fetched to the pointer to fetched count
        If pcvFetched Then
            #If bVBVMTypeLib Then
                MemLong(ByVal pcvFetched) = cvFetched
            #Else
                CopyMemory ByVal pcvFetched, cvFetched, 4
            #End If
        End If
        
    End If
End Function
Public Function IEnumVARIANT_Skip(ByVal oThis As Object, ByVal cV As Long) As Long
'new vtable method for IEnumVARIANT::Skip

    Dim oEnumVARIANT As cEnumVariantVB
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'cast method to source interface
    Set oEnumVARIANT = oThis
    
    'call the class method that raises a Skip event
    oEnumVARIANT.Skip cV
    
    IEnumVARIANT_Skip = MapCOMErr(Err)
    Exit Function
    
CATCH_EXCEPTION:
    'return 'unimplemented'
    IEnumVARIANT_Skip = E_NOTIMPL
    
End Function

Private Sub pInitArray(ByVal iAddr As Long, icEl As Long)
    Const FADF_STATIC = &H2&      '// Array is statically allocated.
    Const FADF_FIXEDSIZE = &H10&  '// Array may not be resized or reallocated.
    Const FADF_VARIANT = &H800&   '// An array of VARIANTs.
    
    Const FADF_Flags = FADF_STATIC Or FADF_FIXEDSIZE Or FADF_VARIANT
    
    With mtSAHeader
        If .cDims = 0& Then
            .cbElements = 16
            .cDims = 1
            .fFeatures = FADF_Flags
            CopyMemory ByVal ArrPtr(mvArray), VarPtr(mtSAHeader), 4&
        End If
        .cElements = icEl
        .pvData = iAddr
    End With

End Sub
