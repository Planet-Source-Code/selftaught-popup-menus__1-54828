VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcHookHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcHookHub.cls                          7/5/04
'
'           GENERAL PURPOSE:
'               Maintain a single windows hook and it's callback function, and deliver the
'               notifications to the requesting objects.
'
'           MODULES CALLED FROM THIS MODULE:
'               mVbBaseGeneral.bas
'
'           CLASSES CREATED BY THIS MODULE:
'               pcHook.cls
'
'==================================================================================================

'1.  Private Interface           - General procedures
'2.  Public Interface            - General procedures
'3.  iHook Interface             - callback from the asm thunk

'Allow easy toggle between use of the VBVM Typelib
#Const bVBVMTypeLib = True

Implements iHook

Private moHook As pcHook                    'assembler thunk

Private iHookType As eHookType              'hook type

'Store the client object pointers and the number of clients
'Clients are stored with the most recent request at the lowest index
'to the first request at the highest index
Private miClients() As Long         'store the objptr's of the objects to callback
Private miClientCount As Long       'number of objects

'<Private Interface>
Private Sub Class_Initialize()
    Set moHook = New pcHook
End Sub

Private Sub Class_Terminate()
    Set moHook = Nothing
End Sub
'</Private Interface>

'<Public Interface>
Public Property Get Active() As Boolean
    Active = moHook.Active 'Return whether the hook is active
End Property

Public Function AddClient( _
            ByVal iPtr As Long) _
                As Boolean
    
    Dim liFirst As Long
    Dim liIndex As Long
    
    liIndex = ArrFindInt(miClients, miClientCount, iPtr, liFirst)   'See if we already have this client
    
    If liIndex = Undefined Then                                     'if we do not already have this client
        If liFirst = Undefined Then                                 'if there is not any open slot
            liIndex = miClientCount                                 'new index is old count
            miClientCount = miClientCount + 1&                      'bump up the count
            ArrRedim miClients, miClientCount, True                 'redim the array
        Else
            liIndex = liFirst                                       'Use the first open slot
        End If
        If liIndex > 0 Then _
            CopyMemory miClients(1), miClients(0), liIndex * 4&     'Move everything up on index to store this client at index 0
        miClients(0) = iPtr                                         'store the client pointer
        AddClient = True                                            'indicate success
    Else
        'Adding a client that is already there!
        Debug.Assert False
    End If
End Function

Public Function DelClient( _
            ByVal iPtr As Long) _
                As Boolean
    
    
    Dim liIndex As Long
    liIndex = ArrFindInt(miClients, miClientCount, iPtr)            'try to find the client
    
    If Not (liIndex = Undefined) Then                               'If we found the client
        miClients(liIndex) = Undefined                              'mark the slot as open
        If liIndex = miClientCount - 1& Then                        'If this is the last client
            For liIndex = liIndex To 0 Step -1&                     'count back from the top to find the
                If miClients(liIndex) <> Undefined Then Exit For    'smallest possible value for miClientCount
            Next
            miClientCount = liIndex + 1&                            'store 1-based count, not 0-based index
            If miClientCount = 0& Then moHook.UnHook                'If nobody's left listening, unhook
        End If
        DelClient = True                                            'indicate success
    Else
        'Deleting a client that isn't there!
        Debug.Assert False
    End If
    
End Function

Public Function HookExists( _
            ByVal iPtr As Long) _
                As Boolean
    HookExists = ArrFindInt(miClients, miClientCount, iPtr) > Undefined 'return true if the client was found
End Function

Public Function Hook( _
            ByVal iType As eHookType, _
            ByVal bThread As Boolean) _
                As Boolean
    iHookType = iType                       'store the hook type
    Hook = moHook.Hook(Me, iType, bThread)  'set the hook
End Function
'</Public Interface>

'<iHook Interface>
Private Sub iHook_Proc( _
            ByVal bBefore As Boolean, _
            ByRef bHandled As Boolean, _
            ByRef lReturn As Long, _
            ByRef nCode As eHookCode, _
            ByRef wParam As Long, _
            ByRef lParam As Long)
    On Error Resume Next
    Dim i           As Long
    Dim liClient    As Long
    Dim iHook       As iHook
    
    
    For i = 0 To miClientCount - 1& Step 1&                                     'Loop through each client
        liClient = miClients(i)
        If liClient <> 0& And liClient <> Undefined Then                        'If the slot is being used
            #If bVBVMTypeLib Then                                               'then get an object reference
                ObjectPtr(iHook) = liClient
            #Else
                CopyMemory iHook, liClient, 4&
            #End If
            
            If Not iHook Is Nothing Then                                        'if the reference is valid
                iCurrentHook = iHookType                                        'set the current hook enum value
                iHook.Proc bBefore, bHandled, lReturn, nCode, wParam, lParam    'call the hook proc
                iCurrentHook = -100
                If bHandled Then Exit For                                       'if handled, then cease notifications
            Else
                'Slot appears to be valid, but the object reference is invalid.
                Debug.Assert False
            End If
        End If
    Next

    'Delete the illegal object reference.
    #If bVBVMTypeLib Then
        ObjectPtr(iHook) = 0&
    #Else
        CopyMemory iHook, 0&, 4&
    #End If

End Sub
'</iHook Interface>

