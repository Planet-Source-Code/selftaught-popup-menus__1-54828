Attribute VB_Name = "mVbBaseGeneral"
Option Explicit

'==================================================================================================
'vbBase.prj                             7/5/04
'
'           GENERAL PURPOSE:
'               Subclasses, Api Windows, Registered window classes, api timers, windows hooks, and vtable
'               subclassing for ole interfaces, all in the familiar collection interface.
'
'           LINEAGE:
'               Inspired by and originally based on SSubTmr6.dll from vbaccelerator.com.
'               VTable subclassing is based on code from vbACOM10.dll also at vbaccelerator.com.
'               ASM and related utility procs are from Paul Caton's WinSubHook2 library at pscode.com.
'
'           COPYRIGHTS:
'               I won't sue you no matter how you use it provided that you don't sue me no matter how it fails to work.
'
'==================================================================================================


'==================================================================================================
'mVbBaseGeneral.bas                     7/4/04
'
'           GENERAL PURPOSE:
'               General utility functions for memory allocation, bit masks and arrays
'
'==================================================================================================

'1.  Bitmask Interface      - Public array of pre-calculated exponents of 2
'2.  ASM Resource Interface - a function to allocate memory, copy the ASM from the resource file and return it's address, and code patching functions.
'3.  Array Interface        - Procedures to redim long and byte arrays in blocks, find/add/delete long vals from arrays
'4.  Utility Interface      - a procedure to check if an objptr implements iWindow, getprocaddress, and InIDE

#Const bVBVMTypeLib = True  'Constant to allow easy switching between use of the VB Virtual Machine Type Library

Public Const Undefined = -1& 'Code Clarity

'<Public Interface>

Public Type tPointer
    iId As Long
    iIndex As Long
End Type

Public Type tItemPointer
    iId As Long
    iIndex As Long
    iItemIndex As Long
End Type

'<Utility API's>
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
'</Utility API's>

'<Related to ASM Resource Interface>
Public Enum eASMResources
    asmSubclass = 101
    asmHook = 102
    asmTimer = 103
    asmWindow = 104
End Enum

Private Type tASMResource
    yBytes()    As Byte    'ASM
    iLen        As Long    'Length
    bRetrieved  As Boolean 'Has it been retrieved from the resource file yet?
End Type

Private mtASMResource(asmSubclass To asmWindow) As tASMResource
'</Related to ASM Resource Interface>

'<Bitmask Interface>
Public Const BitMax = 31&           'bit masks to avoid constantly
Public BitMask(0 To BitMax) As Long 'performing exponential operations.

Private Sub Main()
    'initialize the bit mask once when starting up so
    'that it does not have to be checked before each use
    InitBitMask
    iCurrentHook = -100
End Sub

Private Sub InitBitMask()
    'Initialize the bit mask without using exponents or even multiplication!
    'Sure it could be done w/o addition by hardcoding the values, but I like this better for some reason.
    If BitMask(0) = 0& Then
        Dim liNum As Long
        Dim i As Long: i = 1&
        For liNum = 0& To BitMax
            BitMask(liNum) = i
            If liNum + 1& < BitMax Then i = i + i Else i = &H80000000
        Next
    End If
End Sub
'</Bitmask Interface>

'<ASM Resource Interface>
Public Function AllocASM( _
            ByVal iId As eASMResources, _
   Optional ByVal iAdditionalMem As Long) _
                As Long
    With mtASMResource(iId)
        If Not .bRetrieved Then
            .yBytes = LoadResData(iId, "ASSEMBLY")
            .iLen = UBound(.yBytes) + 1&
            .bRetrieved = True
        End If
        AllocASM = GlobalAlloc(0&, .iLen + iAdditionalMem)
        CopyMemory ByVal AllocASM, .yBytes(0), .iLen
    End With
End Function

Public Sub PatchValue( _
            ByVal iAddr As Long, _
            ByVal iOffset As Long, _
            ByVal iValue As Long)
    CopyMemory ByVal (iAddr + iOffset), iValue, 4&
End Sub

Public Sub PatchValueRelative( _
            ByVal iAddr As Long, _
            ByVal iOffset As Long, _
            ByVal iTarget As Long)
    CopyMemory ByVal (iAddr + iOffset), iTarget - iAddr - iOffset - 4&, 4&
End Sub
'</ASM Resource Interface>

'<Array Interface>
Public Sub ArrRedim( _
            ByRef iArray() As Long, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    'This sub will allocate arrays in blocks, saving constant reallocation
    'when elements need to be added.  It will only increase the size of the
    'array, it will never decrease it.
    
    'Arrays are dimensioned with upper bounds that are even multiples of
    'ArrBlockSize, not with a number of elements that are even multiples.

    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUbound(iArray) Then
        If bPreserve Then _
            ReDim Preserve iArray(0 To liNewUbound) _
        Else _
            ReDim iArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUbound( _
            ByRef iArray() As Long) _
                As Long
    On Error Resume Next
    ArrUbound = UBound(iArray)
    If Err.Number <> 0& Then ArrUbound = Undefined
End Function


Public Sub ArrRedimB( _
            ByRef yArray() As Byte, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundB(yArray) Then
        If bPreserve Then _
            ReDim Preserve yArray(0 To liNewUbound) _
        Else _
            ReDim yArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundB( _
            ByRef yArray() As Byte) _
                As Long
    On Error Resume Next
    ArrUboundB = UBound(yArray)
    If Err.Number <> 0& Then ArrUboundB = Undefined
End Function

Public Function ArrAdjustUbound( _
            ByVal iBound As Long) _
                As Long
    'Adjusts a Ubound to the next increment of the blocksize
    Const ArrBlockSize As Long = 10&
    
    'if ibound < 0 then ibound = 0
    If iBound And &H80000000 Then iBound = 0&
    
    Dim liMod As Long
    liMod = iBound Mod ArrBlockSize
    
    If Not (liMod = 0) Then
        'If the bound is not an even multiple, then round it up
        ArrAdjustUbound = iBound + ArrBlockSize - liMod
    Else
        'If it is an even multiple, then keep it the same,
        'unless it's zero, then make it equal to ArrBlockSize
        If Not (iBound = 0&) Then _
            ArrAdjustUbound = iBound _
        Else _
            ArrAdjustUbound = ArrBlockSize
    End If
End Function

Public Function ArrAddInt( _
            ByRef aTable() As Long, _
            ByRef iCount As Long, _
            ByVal iInt As Long) _
                As Long
    'Adds an integer to a table
    
    
    If ArrFindInt(aTable, iCount, iInt, ArrAddInt) _
                    = _
            Undefined Then                   'If the value is not already in the table
                
        If ArrAddInt = Undefined Then        'if there is not any available slot
            ArrAddInt = iCount                  'next index is current count
            iCount = iCount + 1&                'bump up the count
            ArrRedim aTable, iCount, True       'redim the array
        End If
        aTable(ArrAddInt) = iInt                'set the value
    Else
        'Value is already in the table
        'Debug.Assert False
    End If
                
End Function

Public Function ArrDelInt(ByRef aTable() As Long, _
            ByRef iCount As Long, _
            ByVal iInt As Long) _
                As Boolean
    
    iInt = ArrFindInt(aTable, iCount, iInt)         'Try to find the value in the table
    
    If iInt <> Undefined Then                    'if the value was found
        ArrDelInt = True                            'indicate success
        aTable(iInt) = Undefined                 'remove the value
        If iInt = iCount - 1& Then                  'if this was the last value
            For iCount = iInt - 1& To 0& Step -1&   'loop backwards to find lowest possible value for iCount
                If aTable(iCount) <> 0& And _
                   aTable(iCount) <> -1& Then Exit For
            Next
            iCount = iCount + 1&                    'store 1-based index instead of 0-based count
        End If
    End If
    
    'Value not found in table
    Debug.Assert ArrDelInt
    
End Function
                       

Public Function ArrFindInt( _
            ByRef aTable() As Long, _
            ByVal iCount As Long, _
            ByVal iInt As Long, _
   Optional ByRef iFirstAvailable As Long) _
                As Long
    'Find an integer in a table and get the index and/or the first available slot
    
    Dim liTemp As Long
    
    iFirstAvailable = Undefined                         'make sure the first available starts at nothing
    
    For ArrFindInt = 0& To iCount - 1&                  'loop through each index
        liTemp = aTable(ArrFindInt)                     'store the value of this slot
        If liTemp <> Undefined Then                     'if the slot contains a valid value
            If liTemp = iInt Then Exit Function         'if the value matches then bail
        Else
            If iFirstAvailable = Undefined Then _
                iFirstAvailable = ArrFindInt            'if the slot was not valid, it may be the first available
        End If
    Next
    
    ArrFindInt = Undefined                              'if we made it out here, the value was not found.
End Function
'</Array Interface>

'<Utility Interface>
Public Function AddrFunc( _
            ByRef sDLL As String, _
            ByRef sProc As String) _
                As Long
  AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
End Function

Public Function InIDE() As Boolean
    Debug.Assert SetTrue(InIDE)
End Function

Private Function SetTrue(ByRef bVal As Boolean) As Boolean
    bVal = True
    SetTrue = True
End Function

Public Function Incr(ByRef i As Long)
    If i = &H7FFFFFFF Then
        i = &H80000000
    Else
        i = i + 1&
    End If
End Function

Public Sub gErr(ByVal iNum As eVbBaseErrors, ByRef sSource As String)
    Dim lsDesc As String
    Select Case iNum
        Case iNum = vbbInvalidProcedureCall
            lsDesc = "Invalid procedure call or argument."
        Case vbbItemDetached
            lsDesc = "Item is not part of a collection."
        Case vbbCollChangedDuringEnum
            lsDesc = "Collection changed during enumeration."
        Case vbbKeyNotFound
            lsDesc = "Collection item was not found."
        Case vbbKeyAlreadyExists
            lsDesc = "Collection key already exists."
        Case vbbApiFailure
            lsDesc = "API Error: " & Err.LastDllError & vbNewLine & ErrorMessage(Err.LastDllError)
    End Select
    Err.Raise iNum, sSource, lsDesc
End Sub

Public Function ErrorMessage(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
    
    ' Return the error message associated with LastDLLError:
    sBuff = String$(256, 0)
    lCount = FormatMessage( _
    FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
    0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
    ErrorMessage = Left$(sBuff, lCount)
End If

End Function


'Public Function GetBaseAddress()
'
'    Const SixteenMB = 16777216
'    Const TwoGB = 2147483648#
'    Const Sixty4K = 65536
'    Dim nReturn
'    Dim nMultiple
'    Dim nSizeOf
'
'    ' Ask the User for the size in kilobytes of the component
'    nSizeOf = InputBox("Enter the Size of your component in Kilobytes.", "Base Address Generator")
'
'    ' Do some simple Error prevention.
'    If IsNumeric(nSizeOf) Then
'        If nSizeOf > 0 Then
'            nSizeOf = nSizeOf * 1024
'        Else
'            MsgBox "Your component must be larger than 0 kilobytes. Try again smarty-pants.", vbOKOnly + vbExclamation, "Base Address Generator"
'            Exit Function
'        End If
'    Else
'        MsgBox "Kilobytes are numbers jack ass!", vbOKOnly + vbExclamation, "Base Address Generator"
'        Exit Function
'    End If
'
'    ' Generate a random Number between 16 megabytes And two gigabytes minus the size
'    ' of the memory used by the component.
'    Randomize
'    nReturn = Rnd
'    nReturn = Int((((TwoGB - nSizeOf) - SixteenMB) + 1) * Rnd + SixteenMB)
'
'    ' The Number must be able to round up to a multiple of 64K
'    If nReturn > (TwoGB - Sixty4K) Then
'        While nReturn > (TwoGB - Sixty4K)
'            Randomize
'            nReturn = Rnd
'            nReturn = Int((((TwoGB - nSizeOf) - Sixty4K) + 1) * Rnd + SixteenMB)
'        Wend
'    End If
'    nMultiple = Int((nReturn / Sixty4K) + 1)
'    nReturn = Sixty4K * nMultiple
'
'    GetBaseAddress = "&H" & Hex(nReturn)
'
'End Function

'</Utility Interface>
'</Public Interface>












'==================================================================================================
'MODULE NAME                            DATE
'
'           GENERAL PURPOSE:
'
'           MODULES CALLED FROM THIS MODULE:
'
'           CLASSES CREATED BY THIS MODULE:
'
'==================================================================================================



