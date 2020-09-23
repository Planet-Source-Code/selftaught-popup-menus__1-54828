VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcHook - module-less, IDE safe, machine code hooking thunk
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.00 20030107 First cut..........................................................................
'v1.01 20030901 Changes to allow some global, system-wide hooks....................................
'v1.02 20031118 Allow control over callback gating
'               Use global memory for the machine code buffer
'               Reform the assembler...............................................................
'v1.03 20040118 Use EbMode for breakpoint/stop detection rather than callback gating
'               Further reform the assembler for greater speed and smaller size
'               Made InIDE public..................................................................
'
'==================================================================================================

'<Utility Api's>
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'</Utility Api's>

Private nAddrHook As Long                                   'Address of the hook thunk
Private hHook     As Long                                   'Hook handle

'============================================
'Class creation/destruction
'============================================

'Build the hook thunk into allocated memory
Private Sub Class_Initialize()
Const PATCH_01  As Long = 17                                'Code buffer offset to the location of the relative address to EbMode
Const PATCH_03  As Long = 71                                'Relative address of UnhookWindowsHookEx
Const PATCH_05  As Long = 101                               'Relative address of CallNextHookEx
Const FUNC_EBM  As String = "EbMode"                        'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_UWH  As String = "UnhookWindowsHookEx"           'We use CallWindowProc to call the original WndProc
Const FUNC_CNH  As String = "CallNextHookEx"                'Allows the thunk to call the next hook in the chain
Const MOD_VBA5  As String = "vba5"                          'Location of the EbMode function if running VB5
Const MOD_VBA6  As String = "vba6"                          'Location of the EbMode function if running VB6
Const MOD_USER  As String = "user32"                        'Location of the UnhookWindowsHookEx and CallNextHookEx functions
  Dim i         As Long                                     'Loop index
  
  nAddrHook = AllocASM(asmHook)
  
  If InIDE Then
    'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
    Call CopyMemory(ByVal nAddrHook + 14&, &H9090&, 2&)
    
    i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll
    If i = 0 Then                                           'Found?
      i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
    End If

    Debug.Assert i                                          'Ensure the EbMode function was found
    Call PatchValueRelative(nAddrHook, PATCH_01, i)         'Patch the relative address to the EbMode api function
  End If

  'Patch the the runtime values that are known
  Call PatchValueRelative(nAddrHook, PATCH_03, _
                          AddrFunc(MOD_USER, FUNC_UWH))     'Relative address of UnhookWindowsHookEx
  Call PatchValueRelative(nAddrHook, PATCH_05, _
                          AddrFunc(MOD_USER, FUNC_CNH))     'Relative address of CallNextHookEx
End Sub

'Unhook if required and release the allocated memory
Private Sub Class_Terminate()
  Call UnHook                                               'UnHook if the hook thunk is active
  Call GlobalFree(nAddrHook)                                'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

'Set the hook
Public Function Hook(ByVal Owner As iHook, ByVal HookType As eHookType, Optional ByVal bThread As Boolean = True) As Boolean
Const PATCH_02  As Long = 66                                'Hook handle for UnhookWindowsHookEx
Const PATCH_04  As Long = 96                                'Hook handle for CallNextHookEx
Const PATCH_06  As Long = 139                               'Address of the owner object
  Dim nThreadID As Long                                     'App.ThreadID
  
  If hHook = 0 Then
    If bThread Then                                         'Validate the parameters with regard to hook type vs thread or system hooking
    
      Select Case HookType
      Case WH_JOURNALPLAYBACK, WH_JOURNALRECORD, WH_SYSMSGFILTER
        Debug.Assert False                                  'Inappropriate thread hooks
      End Select
      
      nThreadID = App.ThreadID
    End If
    
    Call PatchValue(nAddrHook, PATCH_06, ObjPtr(Owner))     'Owner object address
    
    'Set the hook
    hHook = SetWindowsHookEx(HookType, nAddrHook, App.hInstance, nThreadID)
    
    If hHook <> 0 Then
      Hook = True
      Call PatchValue(nAddrHook, PATCH_02, hHook)           'Hook handle for UnhookWindowsHookEx
      Call PatchValue(nAddrHook, PATCH_04, hHook)           'Hook handle for CallNextHookEx
    End If
  End If
  
  Debug.Assert Hook
End Function

'Call this method to unhook
Public Function UnHook() As Boolean
  If hHook <> 0 Then
    If UnhookWindowsHookEx(hHook) <> 0 Then                 'Unhook the hook
      UnHook = True                                         'Success
      hHook = 0                                             'Hook inactive
    End If
  End If
End Function









'Added for utility purpose 4/1/04
Public Property Get Active() As Boolean
    Active = hHook <> 0&
End Property
