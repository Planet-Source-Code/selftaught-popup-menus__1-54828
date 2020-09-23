VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const bVBVMTypeLib = True

'==================================================================================================
'pcTimer - module-less, IDE safe, machine code timer thunk
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.00 20030107 First cut..........................................................................
'v1.01 20031118 Allow control over callback gating
'               Use global memory for the machine code buffer
'               Reform the assembler...............................................................
'v1.02 20040118 Use EbMode for breakpoint/stop detection rather than callback gating
'               Further reform the assembler for greater speed and smaller size
'               Made InIDE public..................................................................
'
'==================================================================================================

'<Utility Api's>
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'</Utility Api's>

Private nTimerID As Long                                    'Timer ID
Private nAddrTmr As Long                                    'The address of our timer thunk
Private iInterval As Long

'============================================
'Class creation/destruction
'============================================

'Build the timer thunk into allocated memory
Private Sub Class_Initialize()
Const PATCH_01 As Long = 3                                  'Code buffer offset to the location of the relative address to EbMode
Const PATCH_05 As Long = 52                                 'Code buffer offset to the location of the relative address to KillTimer
Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk know if the IDE has stopped or is on a breakpoint
Const FUNC_KIL As String = "KillTimer"                      'KillTimer allows the pcTimer machine code thunk to kill the timer itself if it detects via the EbMode function that the IDE has stopped
Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
Const MOD_USER As String = "user32"                         'Location of the KillTimer function
  Dim i        As Long                                      'Loop index
 
  nAddrTmr = AllocASM(asmTimer)

  If InIDE Then
    'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
    Call CopyMemory(ByVal nAddrTmr, &H9090, 2)

    i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll
    If i = 0 Then                                           'Found?
      i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
    End If

    Debug.Assert i                                          'Ensure the EbMode function was found
    Call PatchValueRelative(nAddrTmr, PATCH_01, i)          'Patch the relative address to the EbMode api function
  End If
  
  PatchValueRelative nAddrTmr, PATCH_05, _
                     AddrFunc(MOD_USER, FUNC_KIL)           'Patch the relative address to the KillTimer api function
End Sub

'Stop the timer and release the allocated memory
Private Sub Class_Terminate()
  If nTimerID <> 0& Then KillTimer 0&, nTimerID  'Kill the timer
  Call GlobalFree(nAddrTmr)                      'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

'4/10/04
Public Function Active() _
                As Boolean
    Active = nTimerID <> 0&
End Function

'Create the timer
Public Function Create( _
    Optional ByVal nInterval As Long = -1&) _
                As Boolean
Const PATCH_03 As Long = 26                                 'Code buffer offset to the location of the timer Create time

  If nInterval = -1& Then nInterval = iInterval
  iInterval = nInterval

  If nTimerID = 0 Then
    Call PatchValue(nAddrTmr, PATCH_03, GetTickCount)       'Set the Create time
    'Create the timer
    nTimerID = SetTimer(0&, 0&, nInterval, nAddrTmr)
    Create = (nTimerID <> 0&)                               'nTimer ID is non-zero on success
  End If

  Debug.Assert Create                                       'Let programmer know if Create failed
End Function

'Stop the timer
Public Function Destroy() _
                As Boolean
  If nTimerID <> 0& Then
    If KillTimer(0&, nTimerID) <> 0& Then                   'Kill the timer
      nTimerID = 0&                                         'Indicate the timer is inactive
      Destroy = True
    End If
  End If

  Debug.Assert Destroy                                      'Let programmer know if Stop failed
End Function

Public Function IsOwner( _
            ByVal iPtr As Long) _
                As Boolean
    Const PATCH_04 As Long = 32                             'Code buffer offset to the location of the owner object address
    #If bVBVMTypeLib Then
        IsOwner = (iPtr = MemLong(nAddrTmr + PATCH_04))
    #Else
        Dim i As Long
        CopyMemory i, ByVal (nAddrTmr + PATCH_04), 4&
        IsOwner = (iPtr = i)
    #End If
End Function

Public Property Let Owner( _
            ByVal iPtr As Long)
    Const PATCH_04 As Long = 32                             'Code buffer offset to the location of the owner object address
    #If bVBVMTypeLib Then
        MemLong(nAddrTmr + PATCH_04) = iPtr
    #Else
        CopyMemory ByVal (nAddrTmr + PATCH_04), iPtr, 4&
    #End If
End Property

Public Property Let Interval( _
            ByVal iInt As Long)
    iInterval = iInt
End Property

Public Property Get Interval() _
                As Long
    Interval = iInterval
End Property

Public Property Let ID( _
            ByVal iInt As Long)
    Const PATCH_02 As Long = 17                                 'Timer ID
    Call PatchValue(nAddrTmr, PATCH_02, iInt)
End Property

Public Property Get ID() As Long
    Const PATCH_02 As Long = 17                                 'Timer ID
    #If bVBVMTypeLib Then
        ID = MemLong(nAddrTmr + PATCH_02)
    #Else
        CopyMemory ID, ByVal nAddrTmr + PATCH_02, 4&
    #End If
End Property

