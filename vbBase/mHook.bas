Attribute VB_Name = "mHook"
Option Explicit

'==================================================================================================
'mHook.bas                              7/5/04
'
'           GENERAL PURPOSE:
'               Maintains windows hooks using a separate pcHookHub object for each type of hook.  The
'               pcHookHub object is responsible for delivering the callbacks to each requesting object.
'
'           CLASSES CREATED BY THIS MODULE:
'               pcHookHub
'
'==================================================================================================

'1.  Private Interface  - Methods to return the collections maintained by this component.
'2.  Hooks Interface    - Methods which are called by the cHooks object

Const MaxHook = 14                                  'this is the maximum value for the eHookType enum

Global iCurrentHook As eHookType            'this variable is set before making a hook callback,
                                            'and can be read by the dll client object though the
                                            'cHooks object in order to distinguish between
                                            'different hooks it may have.


Private moHooks(0 To MaxHook)        As pcHookHub   'Store the objects that will relay hook notifications.
Private moHooksThread(0 To MaxHook)  As pcHookHub


Private Const miMissingNumber = 8&                  'This number is missing from the HookType constants, so adjust
                                                    'all numbers below it up one to get a nice and easy array index.

'<Private Interface>
Private Function pAddHook( _
            ByVal iClient As Long, _
            ByRef oHooks() As pcHookHub, _
            ByVal iType As eHookType, _
            ByVal bThread As Boolean) _
                As Boolean
    
    If iType < miMissingNumber Then iType = iType + 1&  'Adjust for array index
    
    Dim loTemp As pcHookHub: Set loTemp = oHooks(iType) 'Store the object from the array index into a local variable
    
    If loTemp Is Nothing Then
        Set loTemp = New pcHookHub                      'If the array object was nothing, then create it
        Set oHooks(iType) = loTemp
    End If

    pAddHook = loTemp.AddClient(iClient)                'Add the client who is requesting hook notifications
    
    If Not pAddHook Then gErr vbbKeyAlreadyExists, "cHooks.Add"
    
    If Not loTemp.Active Then                           'If the hook is not already active
        If iType <= miMissingNumber Then iType = iType - 1& 'Adjust back to the WH code from the array index
        pAddHook = loTemp.Hook(iType, bThread)          'start the hook
        If Not pAddHook Then gErr vbbApiFailure, "cHooks.Add"
    End If
    
    
End Function

Private Function pDelHook( _
            ByVal iClient As Long, _
            ByRef oHooks() As pcHookHub, _
            ByVal iType As eHookType) _
                As Boolean
    
    If iType < miMissingNumber Then iType = iType + 1&      'Adjust for array index
        
    Dim loTemp As pcHookHub: Set loTemp = oHooks(iType)     'Get the object from the array index
    
    If Not loTemp Is Nothing Then _
        pDelHook = loTemp.DelClient(iClient)                'If the object is valid, ask it to remove this client
        
    If Not pDelHook Then gErr vbbKeyNotFound, "cHooks.Remove"

End Function

Private Function pIsValid( _
            ByVal iType As eHookType) _
                As Boolean
    pIsValid = iType > -2& And iType < 16& And iType <> 8&
End Function
'</Private Interface>


'<Public Interface>
'<Hooks Interface>
Public Sub AddHook( _
           ByVal iWho As Long, _
           ByVal iType As eHookType, _
  Optional ByVal bThread As Boolean = True)
  
    If Not pIsValid(iType) Then gErr vbbInvalidProcedureCall, "cHooks.Add"
    'Delegate the task to the private helper function with the appropriate
    'modular variables.
    If bThread _
        Then pAddHook iWho, moHooksThread, iType, bThread _
        Else pAddHook iWho, moHooks, iType, bThread

End Sub

Public Function RemoveHook( _
           ByVal iWho As Long, _
           ByVal iType As eHookType, _
  Optional ByVal bThread As Boolean = True) _
                As Boolean
    If Not pIsValid(iType) Then gErr vbbInvalidProcedureCall, "cHooks.Remove"
    'Delegate the task to the private helper function with the appropriate
    'modular variables.
    If bThread _
        Then pDelHook iWho, moHooksThread, iType _
        Else pDelHook iWho, moHooks, iType

End Function

Public Function HookExists( _
            ByVal iWho As Long, _
            ByVal iType As eHookType, _
            ByVal bThread As Boolean) _
                As Boolean
    If Not pIsValid(iType) Then gErr vbbInvalidProcedureCall, "cHooks.Exists"
    
    If iType < miMissingNumber Then iType = iType + 1&  'Adjust for array index
    
    Dim loHook As pcHookHub
    
    If bThread Then _
        Set loHook = moHooksThread(iType) _
    Else _
        Set loHook = moHooks(iType)                     'get the hook object
    
    If Not loHook Is Nothing Then _
        HookExists = loHook.HookExists(iWho)            'true if this object is using this hook

End Function

Public Function HookCount( _
            ByVal iWho As Long, _
   Optional ByVal bThread As Boolean) _
                As Long
    Dim i As Long
    Dim loHook As pcHookHub
    
    For i = 0 To MaxHook                    'loop through each hook
        If bThread Then _
            Set loHook = moHooksThread(i) _
        Else _
            Set loHook = moHooks(i)         'get the hook object

        If Not loHook Is Nothing Then       'if there is a hook object
            If loHook.HookExists(iWho) Then _
                HookCount = HookCount + 1&  'if this client is using this hook, increment the count
        End If
    Next
End Function

Public Function HookClear( _
            ByVal iWho As Long, _
   Optional ByVal bThread As Boolean) _
                As Long
    Dim i As Long
    Dim loHook As pcHookHub
    
    For i = 0& To MaxHook                       'loop through each hook
        If bThread Then _
            Set loHook = moHooksThread(i) _
        Else _
            Set loHook = moHooks(i)             'get the hook object
            
        If Not loHook Is Nothing Then           'if there is a hook object
            If loHook.HookExists(iWho) Then     'if this client is using this hook
                RemoveHook iWho, i, bThread     'remove the client from this hook
                HookClear = HookClear + 1&      'increment the count
            End If
        End If
    Next
End Function
'</Hooks Interface>
'</Public Interface>
