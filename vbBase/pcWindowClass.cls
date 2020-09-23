VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcWindowClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcWindowClass - module-less, IDE safe, machine code api window thunk
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
'               As well DestroyWindow also UnregisterClass on IDE stop
'               Made InIDE public..................................................................
'
'==================================================================================================
Private Const PATCH_02  As Long = 32                        'Table entry count
Private Const PATCH_0A  As Long = 148                       'Location of class name string

Private nAddrWnd As Long

'============================================
'Class creation/destruction
'============================================

'Build the window thunk into allocated memory
Private Sub Class_Initialize()
Const PATCH_01 As Long = 18                                 'Code buffer offset to the location of the relative address to EbMode
Const PATCH_05 As Long = 108                                'Relative address of DefWindowProc
Const PATCH_06 As Long = 127                                'Relative address of DestroyWindow
Const PATCH_07 As Long = 132                                'Application hInstance for UnregisterClass
Const PATCH_08 As Long = 137                                'Address of the class name for UnregisterClass
Const PATCH_09 As Long = 142                                'Relative address of UnregisterClass
Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_DWP As String = "DefWindowProcA"                 'DefWindowProc is used to call the default window proc
Const FUNC_DW  As String = "DestroyWindow"                  'DestroyWindow is used if the thunk detects that that the IDE has stopped
Const FUNC_URC As String = "UnregisterClassA"               'UnregisterClass is used if the thunk detects that that the IDE has stopped
Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
Const MOD_USER As String = "user32"                         'Location of the SetWindowLong & CallWindowProc functions
  Dim i        As Long                                      'Loop index

  nAddrWnd = AllocASM(asmWindow, 255)
      
  If InIDE Then
    'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
    Call CopyMemory(ByVal nAddrWnd + 15, &H9090, 2)
    
    i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll
    If i = 0 Then                                           'Found?
      i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
    End If

    Debug.Assert i                                          'Ensure the EbMode function was found
    Call PatchValueRelative(nAddrWnd, PATCH_01, i)          'Patch the relative address to the EbMode api function
  End If
  
  'Patch the WndProc thunk with runtime values
  Call PatchValue(nAddrWnd, PATCH_02, 0)                    'Initialize the MsgCount
  Call PatchValue(nAddrWnd, PATCH_07, App.hInstance)        'Application hInstance for UnregisterClass
  Call PatchValue(nAddrWnd, PATCH_08, nAddrWnd + PATCH_0A)  'Address of the class name for UnregisterClass
  
  Call PatchValueRelative(nAddrWnd, PATCH_05, _
                          AddrFunc(MOD_USER, FUNC_DWP))     'Address of the DefWindowProc api function
  Call PatchValueRelative(nAddrWnd, PATCH_06, _
                          AddrFunc(MOD_USER, FUNC_DW))      'Address of the DestroyWindow api function
  Call PatchValueRelative(nAddrWnd, PATCH_09, _
                          AddrFunc(MOD_USER, FUNC_URC))     'Address of the UnregisterClass function)
  
End Sub

'Destroy window(s) and release the allocated memory
Private Sub Class_Terminate()
  Call PatchValue(nAddrWnd, PATCH_02, 0)                    'Patch the code to ensure no further iWindow_Proc callbacks
  Call GlobalFree(nAddrWnd)                                 'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

Public Sub PatchTable(ByRef aTable() As Long, ByVal iCount As Long)
Const PATCH_03 As Long = 43                          'Table address
  
On Error Resume Next
  
  Dim iAddr As Long
  
  iAddr = VarPtr(aTable(0))
  
  Call PatchValue(nAddrWnd, PATCH_02, iCount)        'Patch the table entry count
  Call PatchValue(nAddrWnd, PATCH_03, iAddr)         'Patch the address of the message table. We need do this because there's no guarantee that the table existed at WndCreate time, the table only gets created if a message number is added.

End Sub

Public Property Get AddrWndProc() As Long
    AddrWndProc = nAddrWnd
End Property

'Set the window class owner, Form/Class/UserControl
Public Property Set Owner(ByVal NewOwner As iWindow)
Const PATCH_04 As Long = 79
  
  'Set m_Owner = NewOwner
  Call PatchValue(nAddrWnd, PATCH_04, ObjPtr(NewOwner)) 'Owner object address for iWindow_WndProc
End Property

Public Property Let ClassName(ByRef sClass As String)
    Dim lsTemp As String
    lsTemp = StrConv(Left$(sClass, 255), vbFromUnicode) & vbNullChar
    CopyMemory ByVal (nAddrWnd + PATCH_0A), ByVal StrPtr(lsTemp), LenB(lsTemp)
End Property
