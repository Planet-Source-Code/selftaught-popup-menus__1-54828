VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcWindowClassHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcWindowClassHub.cls                   7/4/04
'
'           GENERAL PURPOSE:
'               Creates windows from a registered class and manages the messages delivered to each window.
'
'           MODULES CALLED FROM THIS MODULE:
'               mVbBaseGeneral.bas
'
'           CLASSES CREATED BY THIS MODULE:
'               cApiClassWindow
'
'==================================================================================================

'1.  Private Interface           - General procedures
'2.  Public Interface            - procedures to add and remove messages for specific clients and to create/destroy windows
'3.  iWindow Interface           - callback from the asm thunk

#Const bVBVMTypeLib = True

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Implements iWindow

'<Utility Api's>
Private Type tWndClassEx
    cbSize          As Long
    Style           As Long
    lpfnWndProc     As Long
    cbClsExtra      As Long
    cbWndExtra      As Long
    hInstance       As Long
    hIcon           As Long
    hCursor         As Long
    hbrBackground   As Long
    lpszMenuName    As Long
    lpszClassName   As Long
    hIconSm         As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (ByRef ltWNDCLASSEX As tWndClassEx) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

'</Utility Api's>


Private Const DWordBits                 As Long = 32    'don't adjust this one - 32 bits in a long

Private Const MsgMaskCount              As Long = 4     'Adjust this constant up/down to change the limit on messages
                                                        'normally it should be an even power of 2 and not larger than 16
                                                        
Private Const MaxMsgCount               As Long _
                     = MsgMaskCount * DWordBits         'no more than this number of messages will be delivered collectively to
                                                        'all windows created from this class.
                                                        
Private Const AllMessages               As Long = 1
Private Const CustomMessages            As Long = 2

Private Const ClientIndexPropName       As String = "WndCIProp070404"

Private Type tClassWindow
    hWnd                                As Long
    Flags                               As Long         'bit 1 is for all messages, and bit 2 indicates that the window receives the messages based on it's own msgmask
    iMsgMask(0 To MsgMaskCount - 1&)    As Long         'masks indicating which messages are received
End Type

Private Type tClassClient
    iPtr                                As Long
    iControl                            As Long
    iWindowCount                        As Long
    tWindows()                          As tClassWindow
End Type

Private miDefMsgMask(0 To MsgMaskCount - 1&) As Long
Private mbDefAllMsgs                    As Boolean      'if default is all messages

Private miMsgTable()                    As Long         'this array will never grow beyond MaxMsgCount # of elements
Private miMsgCount                      As Long
Private miAllMsgCount                   As Long         'count the number of windows that have requested all messages

Private mtClients()                     As tClassClient 'client data store
Private miClientCount                   As Long         'number of clients defined in the array

Private miRegisterCount                 As Long         'Outstanding calls to Register

Private msClass                         As String       'store the class name

Private moWindow                        As pcWindowClass 'worker - asm thunk

Private mbCreating                      As Boolean
Private miCreatingPtr                   As Long

Private Const cApiClsWnds = "cApiClassWindows"
Private Const cApiClsWnd = "cApiClassWindow"

'<Private Interface>
Private Sub Class_Initialize()
    Set moWindow = New pcWindowClass
    Set moWindow.Owner = Me
End Sub

Private Sub Class_Terminate()
    Set moWindow = Nothing
    If Len(msClass) > 0 Then                                 'If a class was registered
      Call UnregisterClass(msClass, App.hInstance)           'Unregister the window class
    End If
End Sub

Private Sub PatchTable()
    
    Dim iPatchVal As Long
    
    If miAllMsgCount = 0& Then                  'If no one wants all messages
        iPatchVal = miMsgCount                  'patch the real count
    Else                                        'else (Someone wants all messages)
        'ALL_MESSAGES
        iPatchVal = -1&                         'patch -1, telling the asm thunk we want all messages
    End If
    
    moWindow.PatchTable miMsgTable, iPatchVal   'patch the values into the thunk
End Sub

Private Function pFindClient( _
            ByVal iPtr As Long, _
   Optional ByRef iFirstAvailable As Long) _
            As Long
    Dim liTemp As Long
    
    iFirstAvailable = Undefined
    
    For pFindClient = 0& To miClientCount - 1&
        liTemp = mtClients(pFindClient).iPtr
        If liTemp <> 0& And liTemp <> Undefined Then
            If liTemp = iPtr Then Exit Function
        Else
            If iFirstAvailable = Undefined Then iFirstAvailable = pFindClient
        End If
    Next
    
    pFindClient = Undefined
    'Debug.Assert False
End Function

Private Function pFindWindow( _
            ByVal iClientIndex As Long, _
            ByVal hWnd As Long, _
   Optional ByRef iFirstAvailable As Long) _
            As Long
    
    Dim liTemp As Long
    
    iFirstAvailable = Undefined
    
    With mtClients(iClientIndex)
        For pFindWindow = 0 To .iWindowCount - 1&
            liTemp = .tWindows(pFindWindow).hWnd
            If liTemp <> 0& Then
                If liTemp = hWnd Then Exit Function
            Else
                If iFirstAvailable = Undefined Then iFirstAvailable = pFindWindow
            End If
        Next
    End With
    
    pFindWindow = Undefined

End Function

Private Function pValidatePointer( _
            ByRef tPointer As tPointer) _
                As Boolean
    If tPointer.iIndex = Undefined Then tPointer.iIndex = pFindClient(tPointer.iId)
    If tPointer.iIndex > Undefined Then
        If tPointer.iIndex < miClientCount Then _
            pValidatePointer = (mtClients(tPointer.iIndex).iPtr = tPointer.iId)
    End If
    'Debug.Assert pValidatePointer
End Function

Private Function pValidateItemPointer( _
            ByRef tItemPointer As tItemPointer) _
                As Boolean
    
    If tItemPointer.iIndex < miClientCount Then
        With mtClients(tItemPointer.iIndex)
            If tItemPointer.iItemIndex < .iWindowCount Then
                pValidateItemPointer = (.tWindows(tItemPointer.iItemIndex).hWnd = tItemPointer.iId)
            End If
        End With
    End If
    
    Debug.Assert pValidateItemPointer
            
End Function
            

'private implementation of ArrRedim to allow string typing
Private Sub ArrRedimT( _
            ByRef tArray() As tClassWindow, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundT(tArray) Then
        If bPreserve Then _
            ReDim Preserve tArray(0 To liNewUbound) _
        Else _
            ReDim tArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundT( _
            ByRef tArray() As tClassWindow _
        ) As Long
    On Error Resume Next
    ArrUboundT = UBound(tArray)
    If Err.Number <> 0& Then ArrUboundT = Undefined
End Function

Private Sub ArrRedimT2( _
            ByRef tArray() As tClassClient, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundT2(tArray) Then
        If bPreserve Then _
            ReDim Preserve tArray(0 To liNewUbound) _
        Else _
            ReDim tArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundT2( _
            ByRef tArray() As tClassClient _
        ) As Long
    On Error Resume Next
    ArrUboundT2 = UBound(tArray)
    If Err.Number <> 0& Then ArrUboundT2 = Undefined
End Function

Private Function pMsgIsDead( _
                ByVal iTableIndex As Long, _
                ByVal iMask As Long) _
                    As Boolean

    Dim liEachClient As Long
    Dim liEachWindow As Long
    
    For liEachClient = 0& To miClientCount - 1&                               'loop through each client
        With mtClients(liEachClient)
            For liEachWindow = 0& To mtClients(liEachClient).iWindowCount - 1&
                If mtClients(liEachClient).iPtr <> Undefined Then                     'if the client is defined
                    If CBool(.tWindows(iTableIndex).iMsgMask(iTableIndex) And iMask) _
                        Then Exit Function                                          'if the client wants this message, bail
                End If
            Next
        End With
    Next
    
    If CBool(miDefMsgMask(iTableIndex) And iMask) _
        Then Exit Function
    
    pMsgIsDead = True                                                           'if we made it here, nobody cares about this message
End Function

Private Sub pTranslateMsgIndex( _
            ByVal iMsgTableIndex As Long, _
            ByRef iClientTableIndex As Long, _
            ByRef iMask As Long)

    iClientTableIndex = iMsgTableIndex \ DWordBits                      'get the index in the tSubclassClient msg array
    iMask = BitMask(iMsgTableIndex Mod DWordBits)                       'get the mask that identifies this message
    
    'internal error
    Debug.Assert iClientTableIndex < MsgMaskCount                       'alert developer in case we've exceeded the max
End Sub

Private Function pDelAllMessages( _
                ByRef iMsgMask() As Long) _
                    As Long

    Dim liMask     As Long      'a mask containing each bit one at a time
    Dim liTempMask As Long      'the current msgmask
    Dim iEachBit   As Long      'counter
    Dim iEachIndex As Long      'counter
    
    For iEachBit = 0 To BitMax                              'loop through each bit
        liMask = BitMask(iEachBit)                          'store the bitmask for this bit
        For iEachIndex = 0 To MsgMaskCount - 1&             'for each bit, loop through each index
            If ((iEachBit * DWordBits) + iEachIndex) = miMsgCount Then Exit Function
            liTempMask = iMsgMask(iEachIndex)              'store the msgmask at this index
            If CBool(liTempMask And liMask) Then            'if the client has requested this message
                pDelAllMessages = pDelAllMessages + 1&      'inc the count of messages removed
                iMsgMask(iEachIndex) = (liTempMask And Not liMask) 'remove the message from the client's table
                If pMsgIsDead(iEachIndex, liMask) _
                    Then miMsgTable(iEachIndex * DWordBits + iEachBit) = Undefined 'If nobody wants the message anymore, remove it from the table
            End If
        Next
    Next

End Function

Private Function pAddMsg( _
            ByVal iClientIndex As Long, _
            ByVal iWindowIndex As Long, _
            ByVal iMsg As eMsg) _
                As Boolean
    
    Dim liFirst As Long
    Dim liIndex As Long
    
    If iMsg <> ALL_MESSAGES Then
        
        liIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg, liFirst)
        If liIndex = Undefined Then
            If liFirst = Undefined Then
                If miMsgCount = MaxMsgCount Then Exit Function
                liIndex = miMsgCount
                miMsgCount = miMsgCount + 1&
                ArrRedim miMsgTable, miMsgCount, True
            Else
                liIndex = liFirst
            End If
            miMsgTable(liIndex) = iMsg
        End If
        
        Dim liClientTableIndex As Long
        Dim liMask As Long
        Dim liCurrent As Long
        
        pTranslateMsgIndex liIndex, liClientTableIndex, liMask
        
        If iClientIndex <> Undefined Then 'custom message
            With mtClients(iClientIndex).tWindows(iWindowIndex)
                If Not (.Flags And CustomMessages) Then .Flags = .Flags Or CustomMessages
                liCurrent = .iMsgMask(liClientTableIndex)
                pAddMsg = Not CBool(liCurrent And liMask)
                If pAddMsg Then .iMsgMask(liClientTableIndex) = (liCurrent Or liMask)
            End With
        Else 'default message
            liCurrent = miDefMsgMask(liClientTableIndex)
            pAddMsg = Not CBool(liCurrent And liMask)
            If pAddMsg Then miDefMsgMask(liClientTableIndex) = (liCurrent Or liMask)
        End If
        
    Else 'ALL_MESSAGES
        
        If iClientIndex <> Undefined Then 'custom message
            With mtClients(iClientIndex).tWindows(iWindowIndex)
                If Not CBool(.Flags And AllMessages) Then
                    pAddMsg = True
                    .Flags = .Flags Or AllMessages
                End If
            End With
        Else 'default message
            If Not mbDefAllMsgs Then
                pAddMsg = True
                mbDefAllMsgs = True
            End If
        End If
        
        If pAddMsg Then miAllMsgCount = miAllMsgCount + 1&
        
    End If
    
    PatchTable
    
End Function

Private Function pDelMsg( _
            ByVal iClientIndex As Long, _
            ByVal iWindowIndex As Long, _
            ByVal iMsg As eMsg) _
                As Long
                
    Dim liMsgIndex          As Long
    Dim liClientMsgIndex    As Long
    Dim liClientMsgMask     As Long
    Dim liTempMask          As Long
    
    If iMsg <> ALL_MESSAGES Then                            'If we're adding a single message
        liMsgIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)     'See if the message exists in the table
            
        If liMsgIndex <> Undefined Then                     'if the message exists
                                                            'get the array index and bit mask for our MsgMask
            pTranslateMsgIndex liMsgIndex, liClientMsgIndex, liClientMsgMask
            
            If iClientIndex <> Undefined Then 'custom message
                With mtClients(iClientIndex).tWindows(iWindowIndex)
                    If CBool(.Flags And CustomMessages) Then
                        liTempMask = .iMsgMask(liClientMsgIndex) 'Store the current mask
                        pDelMsg = CBool(liTempMask And liClientMsgMask) 'succeed only if subscribed to this msg
                        .iMsgMask(liClientMsgIndex) = _
                            liTempMask And Not liClientMsgMask   'remove the bit from the MsgMask
                    End If
                End With
            Else 'default message
                liTempMask = miDefMsgMask(liClientMsgIndex) 'Store the current mask
                pDelMsg = CBool(liTempMask And liClientMsgMask) 'succeed only if subscribed to this msg
                miDefMsgMask(liClientMsgIndex) = _
                    liTempMask And Not liClientMsgMask   'remove the bit from the MsgMask
            End If
            
            If pDelMsg Then                                  'if we succeeded then we may have a message that's not delivered to any object
                If pMsgIsDead(liClientMsgIndex, liClientMsgMask) Then _
                    miMsgTable(liMsgIndex) = Undefined
            End If
            
        End If
    Else
        If iClientIndex <> Undefined Then
            With mtClients(iClientIndex).tWindows(iWindowIndex)
                pDelMsg = CBool(.Flags And AllMessages) _
                                Or CBool(pDelAllMessages(.iMsgMask))
                .Flags = .Flags And Not AllMessages
            End With
        Else
            pDelMsg = mbDefAllMsgs Or CBool(pDelAllMessages(miDefMsgMask))
            mbDefAllMsgs = False
        End If

        If pDelMsg Then miAllMsgCount = miAllMsgCount - 1&
    End If
    
    For liMsgIndex = miMsgCount - 1& To 0& Step -1&
        If miMsgTable(liMsgIndex) <> Undefined Then Exit For
    Next
    
    If liMsgIndex < miMsgCount - 1& Then
        miMsgCount = liMsgIndex + 1&
        PatchTable
    End If
    
End Function

'</Private Interface>

'<Public Interface>
'declared friend b/c VB won't let udts declared in std modules to be passed in public procedures, even
'if the objects are private!
Friend Property Get Active() As Boolean
    Active = (miRegisterCount > 0&)
End Property

Friend Function AddMsg( _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean
 
    If pValidateItemPointer(tPointer) _
        Then AddMsg = pAddMsg(tPointer.iIndex, tPointer.iItemIndex, iMsg) _
        Else gErr vbbItemDetached, cApiClsWnd
    
    'client or window not found or message already there
    Debug.Assert AddMsg
    
    
End Function

Friend Function DelMsg( _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean
                
    If pValidateItemPointer(tPointer) Then _
        DelMsg = pDelMsg(tPointer.iIndex, tPointer.iItemIndex, iMsg) _
        Else gErr vbbItemDetached, cApiClsWnd
    
    'client, window or message not found
    Debug.Assert DelMsg
End Function

Friend Function AddDefMsg( _
                    ByVal iMsg As eMsg _
                ) As Boolean
 
    AddDefMsg = pAddMsg(-1, -1, iMsg)
 
    'message already there?
    Debug.Assert AddDefMsg
  
End Function

Friend Function DelDefMsg( _
                       ByVal iMsg As eMsg _
                ) As Boolean
                
    DelDefMsg = pDelMsg(-1, -1, iMsg)
    
    'message not there?
    Debug.Assert DelDefMsg
  
End Function

Friend Property Get DefMessages( _
            ByRef tPointer As tItemPointer) _
                As Boolean
    
    If pValidateItemPointer(tPointer) _
        Then DefMessages = Not CBool(mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex).Flags And CustomMessages) _
        Else gErr vbbItemDetached, cApiClsWnd
    
End Property

Friend Property Let DefMessages( _
            ByRef tPointer As tItemPointer, _
            ByVal bVal As Boolean)

    If pValidateItemPointer(tPointer) Then
        With mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex)
            If bVal _
            Then .Flags = (.Flags And Not CustomMessages) _
            Else .Flags = (.Flags Or CustomMessages)
        End With
    Else
        gErr vbbItemDetached, cApiClsWnd
    End If

End Property

Friend Function Register( _
                     ByRef sClass As String, _
            Optional ByVal colBackground As Long = &HFFFFFF, _
            Optional ByVal Style As eClassStyle, _
            Optional ByVal hCursor As Long, _
            Optional ByVal hIcon As Long, _
            Optional ByVal hIconSm As Long, _
            Optional ByVal cbClassExtra As Long, _
            Optional ByVal cbWndExtra As Long) _
                 As Boolean

    If miRegisterCount = 0& Then
        Dim wc      As tWndClassEx
        Dim sAscii  As String
        
        msClass = sClass                                         'Store the class name
        sAscii = StrConv(msClass & vbNullChar, vbFromUnicode)    'Get the ascii (rather than unicode) version of the class name
          
        With wc
            .cbSize = Len(wc)                                       'Size of the window class type
            .cbClsExtra = cbClassExtra                              'Number of class extra bytes
            .cbWndExtra = cbWndExtra                                'Number of window extra bytes
            .hbrBackground = CreateSolidBrush(colBackground)        'Class background
            .hCursor = hCursor                                      'Class cursor
            .hIcon = hIcon                                          'Class icon
            .hIconSm = hIconSm                                      'Class small icon
            .hInstance = App.hInstance                              'Application instance handle
            .lpfnWndProc = moWindow.AddrWndProc                     'Class WndProc address
            .Style = Style                                          'Class style
            .lpszClassName = StrPtr(sAscii)                         'Class name
        End With

        Register = (RegisterClassEx(wc) <> 0)          'Register the window class
        
        moWindow.ClassName = msClass
        miRegisterCount = 1&
        
        'Windows Registration failed!
        Debug.Assert Register
    Else
        Debug.Assert False
        If msClass = sClass Then _
            Register = True: miRegisterCount = miRegisterCount + 1&
        'can't register different classes on one pcWindowClassHub object
        Debug.Assert Register
    End If
End Function

Friend Function Unregister() As Boolean
    
    miRegisterCount = miRegisterCount - 1&                    'decrement the count
    
    If miRegisterCount <= 0& Then                            'If we've unregistered the last register
        miRegisterCount = 0&                                 'make sure we don't go too low
        Unregister = _
            UnregisterClass(msClass, App.hInstance) <> 0&   'call the api
        If Unregister Then msClass = vbNullString           'delete the stored class name
    Else
        Debug.Assert False
        Unregister = True                                   'wait until the last one is released
    End If

    'unregister api failed
    Debug.Assert Unregister
    
End Function

Friend Function Create( _
                    ByRef tPointer As tPointer, _
           Optional ByVal dwExStyle As eWindowStyleEx, _
           Optional ByVal dwStyle As eWindowStyle, _
           Optional ByVal x As Long = 0, _
           Optional ByVal y As Long = 0, _
           Optional ByVal nWidth As Long = 0, _
           Optional ByVal nHeight As Long = 0, _
           Optional ByRef sCaption As String = "", _
           Optional ByVal hWndParent As Long = 0, _
           Optional ByVal hMenu As Long = 0, _
           Optional ByVal lParam As Long = 0 _
                ) As cApiClassWindow
  
    Debug.Assert (Len(msClass) > 0)                              'LOGIC ERROR! Class name not defined
 
    Dim liClientIndex As Long
    Dim liWindowIndex As Long
    Dim liFirst As Long
 
    If pValidatePointer(tPointer) Then
       liClientIndex = tPointer.iIndex
    Else
       liClientIndex = pFindClient(tPointer.iId, liFirst)
       If liClientIndex = Undefined Then
           If liFirst = Undefined Then
               liClientIndex = miClientCount
               miClientCount = miClientCount + 1&
               ArrRedimT2 mtClients, miClientCount, True
           Else
               liClientIndex = liFirst
           End If
           With mtClients(liClientIndex)
               .iPtr = tPointer.iId
               .iWindowCount = 0&
           End With
           tPointer.iIndex = liClientIndex
       End If
    End If
        
   'Create the window
   Dim ltPointer As tItemPointer
   
   mbCreating = True
   miCreatingPtr = tPointer.iId
   ltPointer.iId = CreateWindowEx(dwExStyle, msClass, sCaption, dwStyle, x, y, nWidth, nHeight, hWndParent, hMenu, App.hInstance, lParam)
   miCreatingPtr = 0&
   mbCreating = False
   
   Debug.Assert ltPointer.iId
   
   If ltPointer.iId Then
       SetProp ltPointer.iId, ClientIndexPropName, liClientIndex
       
       With mtClients(liClientIndex)
           For liWindowIndex = 0& To .iWindowCount - 1&
               If .tWindows(liWindowIndex).hWnd = 0& Then
                   .tWindows(liWindowIndex).hWnd = ltPointer.iId
                   Exit For
               End If
           Next
           
           If liWindowIndex = .iWindowCount Then
               .iWindowCount = liWindowIndex + 1&
               ArrRedimT .tWindows, .iWindowCount, True
               .tWindows(liWindowIndex).hWnd = ltPointer.iId
           End If
       
           ltPointer.iIndex = liClientIndex
           ltPointer.iItemIndex = liWindowIndex
       
           With .tWindows(liWindowIndex)
               .Flags = 0&
               ZeroMemory .iMsgMask(0), (4& * MsgMaskCount)
           End With
            
            Set Create = New cApiClassWindow
            Create.fInit ltPointer, msClass
       End With
   Else
        gErr vbbApiFailure, cApiClsWnds
   End If
End Function

Friend Sub Destroy(ByRef tPointer As tPointer, ByVal hWnd As Long)
    
    Dim liClientIndex As Long
    Dim liWindowIndex As Long
    

    If pValidatePointer(tPointer) Then
        With mtClients(tPointer.iIndex)
            liWindowIndex = pFindWindow(tPointer.iIndex, hWnd)
            
            If liWindowIndex = Undefined Then gErr vbbItemDetached, cApiClsWnds
            
            With .tWindows(liWindowIndex)
                RemoveProp .hWnd, ClientIndexPropName
                If CBool(DestroyWindow(.hWnd) = 0&) Then gErr vbbApiFailure, cApiClsWnds
                pDelAllMessages .iMsgMask
                .hWnd = 0&
                .Flags = 0&
            End With
            
            If .iWindowCount = liWindowIndex + 1& Then
                For liWindowIndex = liWindowIndex - 1& To 0& Step -1&
                    If .tWindows(liWindowIndex).hWnd <> 0& Then Exit For
                Next
                .iWindowCount = liWindowIndex + 1&
            End If
            
        End With
    Else
        gErr vbbItemDetached, cApiClsWnds
    End If
End Sub

Friend Function Item( _
            ByRef tPointer As tPointer, _
            ByVal hWnd As Long) _
                As cApiClassWindow
    If pValidatePointer(tPointer) Then
        Dim ltItemPointer As tItemPointer
        
        ltItemPointer.iItemIndex = pFindWindow(tPointer.iIndex, hWnd)
        
        If ltItemPointer.iItemIndex = Undefined Then gErr vbbKeyNotFound, cApiClsWnds
        
        ltItemPointer.iId = hWnd
        ltItemPointer.iIndex = tPointer.iIndex
        
        Set Item = New cApiClassWindow
        Item.fInit ltItemPointer, msClass
        
    Else
        gErr vbbKeyNotFound, cApiClsWnds
    End If
End Function

Friend Function Exists( _
                    ByRef tPointer As tPointer, _
                    ByVal hWnd As Long _
                ) As Boolean
    
    If pValidatePointer(tPointer) _
        Then Exists = (pFindWindow(tPointer.iIndex, hWnd) <> Undefined)
    
End Function

Friend Function Count( _
                    ByRef tPointer As tPointer _
                ) As Long
    
    If pValidatePointer(tPointer) Then
        
        Dim i As Long
        
        With mtClients(tPointer.iIndex)
            Count = .iWindowCount
            For i = 0 To .iWindowCount - 1&
                If .tWindows(i).hWnd = 0& Then
                    Count = Count - 1&
                End If
            Next
        End With
        
    End If
    
End Function

Friend Function GetControl(ByRef tPointer As tPointer) As Long
    If pValidatePointer(tPointer) Then GetControl = mtClients(tPointer.iIndex).iControl
End Function

Friend Sub Enum_NextItem( _
                    ByRef tPointer As tPointer, _
                    ByRef tEnum As tEnum, _
                    ByRef vNextItem As Variant, _
                    ByRef bNoMore As Boolean)
    
    If pValidatePointer(tPointer) Then
        tEnum.iIndex = tEnum.iIndex + 1&
        
        With mtClients(tPointer.iIndex)
            
            Do While tEnum.iIndex < .iWindowCount
                If .tWindows(tEnum.iIndex).hWnd <> 0& Then Exit Do
                tEnum.iIndex = tEnum.iIndex + 1&
            Loop
            
            If tEnum.iIndex < .iWindowCount Then
                Dim tItemPointer As tItemPointer
                tItemPointer.iId = .tWindows(tEnum.iIndex).hWnd
                tItemPointer.iIndex = tPointer.iIndex
                tItemPointer.iItemIndex = tEnum.iIndex
                
                Dim oTemp As cApiClassWindow
                Set oTemp = New cApiClassWindow               'create an object
                oTemp.fInit tItemPointer, msClass      'initialize it
                Set vNextItem = oTemp
            Else
                bNoMore = True
            End If
        
        End With
        
    Else
        bNoMore = True
    End If
    
    
End Sub

Friend Function AbsCount() As Long
    Dim liCount As Long
    Dim iEachClient As Long
    Dim iEachWindow As Long
    
    
    For iEachClient = 0 To miClientCount - 1&
        With mtClients(iEachClient)
            liCount = .iWindowCount
            For iEachWindow = 0& To .iWindowCount - 1&
                If .tWindows(iEachWindow).hWnd = 0& Then liCount = liCount - 1&
            Next
        End With
        AbsCount = AbsCount + liCount
    Next
    
End Function

Friend Function Clear( _
                    ByRef tPointer As tPointer _
                ) As Long
    Dim i As Long
    
    If pValidatePointer(tPointer) Then
        
        With mtClients(tPointer.iIndex)
            For i = 0& To .iWindowCount - 1&
                If .tWindows(i).hWnd Then
                    If DestroyWindow(.tWindows(i).hWnd) Then
                        .tWindows(i).hWnd = 0&
                        .tWindows(i).Flags = 0&
                        ZeroMemory .tWindows(i).iMsgMask(0), MsgMaskCount * 4&
                    End If
                End If
            Next
                
            For .iWindowCount = .iWindowCount - 1& To 0& Step -1&
                If .tWindows(.iWindowCount).hWnd Then
                    .iWindowCount = .iWindowCount + 1&
                    gErr vbbApiFailure, cApiClsWnds
                End If
            Next
            .iWindowCount = .iWindowCount + 1&
        End With
    
    End If
    
End Function

Friend Function MsgExists( _
                    ByRef tPointer As tItemPointer, _
                    ByVal iMsg As eMsg) _
                        As Boolean
    If pValidateItemPointer(tPointer) Then
        Dim liMsgIndex As Long
        Dim liClientTableIndex As Long
        Dim liMask As Long
        
        liMsgIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)
        
        If liMsgIndex <> Undefined Then
            pTranslateMsgIndex liMsgIndex, liClientTableIndex, liMask
            MsgExists = (mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex).iMsgMask(liClientTableIndex) And liMask)
        End If
        
    Else
        gErr vbbItemDetached, cApiClsWnd
    End If

End Function

Friend Function MsgCount( _
                    ByRef tPointer As tItemPointer) _
                        As Long
    
    If pValidateItemPointer(tPointer) Then
        Dim liIndex As Long
        Dim liBit As Long
        
        With mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex)
            For liIndex = 0& To MsgMaskCount - 1&
                For liBit = 0& To BitMax
                    If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function
                    If CBool(.iMsgMask(liIndex) And BitMask(liBit)) Then MsgCount = MsgCount + 1&
                Next
            Next
        End With
    Else
        gErr vbbItemDetached, cApiClsWnd
    End If
    
End Function

Friend Function GetMessages( _
                    ByRef iOutArray() As Long, _
                    ByRef tPointer As tItemPointer) _
                        As Long
                        
    If pValidateItemPointer(tPointer) Then
        If miMsgCount = 0& Then
            Erase iOutArray
            Exit Function
        End If
        
        Dim liIndex As Long
        Dim liBit As Long
        ReDim iOutArray(0 To miMsgCount - 1&)
        With mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex)
            For liIndex = 0& To MsgMaskCount - 1&
                For liBit = 0& To BitMax
                    If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function
                    If CBool(.iMsgMask(liIndex) And BitMask(liBit)) Then
                        iOutArray(GetMessages) = miMsgTable((liIndex * DWordBits) + liBit)
                        GetMessages = GetMessages + 1&
                    End If
                Next
            Next
            If GetMessages = 0& Then
                Erase iOutArray
                Exit Function
            End If
            ReDim Preserve iOutArray(0 To GetMessages - 1&)
        End With
    Else
        gErr vbbItemDetached, cApiClsWnd
    End If
    
End Function

Friend Function DefMsgExists( _
                    ByVal iMsg As eMsg _
                ) As Boolean
    Dim liIndex As Long
    liIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)
    
    If liIndex <> Undefined Then
        Dim liClientTableIndex As Long
        Dim liMask As Long
        pTranslateMsgIndex liIndex, liClientTableIndex, liMask
        DefMsgExists = CBool(miDefMsgMask(liClientTableIndex) And liMask)
    End If

End Function

Friend Function DefMsgCount() As Long
    
    Dim liIndex As Long
    Dim liBit As Long
    
    For liIndex = 0& To MsgMaskCount - 1&
        For liBit = 0& To BitMax
            If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function
            If CBool(miDefMsgMask(liIndex) And BitMask(liBit)) Then DefMsgCount = DefMsgCount + 1&
        Next
    Next
    
End Function

Friend Function GetDefMessages( _
            ByRef iOutArray() As Long) _
                As Long
    Dim liIndex As Long
    Dim liBit As Long
    If mbDefAllMsgs Then
        GetDefMessages = -1&
        Erase iOutArray
        Exit Function
    End If
    
    If miMsgCount = 0& Then
        Erase iOutArray
        Exit Function
    End If
    ReDim iOutArray(0 To miMsgCount - 1&)
        
    For liIndex = 0& To MsgMaskCount - 1&
        For liBit = 0& To BitMax
            If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function
            If CBool(miDefMsgMask(liIndex) And BitMask(liBit)) Then
                iOutArray(GetDefMessages) = miMsgTable((liIndex * DWordBits) + liBit)
                GetDefMessages = GetDefMessages + 1&
            End If
        Next
    Next
    
    If GetDefMessages = 0& Then
        Erase iOutArray
        Exit Function
    End If
    
    ReDim Preserve iOutArray(0 To GetDefMessages - 1&)
    
    
End Function
Friend Property Get ClassName() As String
    ClassName = msClass
End Property
'</Public Interface>

'<iWindow Interface>
Private Sub iWindow_Proc( _
                    ByRef bHandled As Boolean, _
                    ByRef lReturn As Long, _
                    ByRef hWnd As Long, _
                    ByRef iMsg As eMsg, _
                    ByRef wParam As Long, _
                    ByRef lParam As Long)
            
    Dim liClientIndex As Long
    Dim liWindowIndex As Long
    Dim liMsgIndex As Long
    Dim liClientTableIndex As Long
    Dim liMask As Long
    Dim iPtr As Long
    
    liMsgIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)
    
    If Not mbCreating Then
        liClientIndex = GetProp(hWnd, ClientIndexPropName)
        
        If liClientIndex < miClientCount Then
            
            liWindowIndex = pFindWindow(liClientIndex, hWnd)
            
            If liWindowIndex <> Undefined Then
                If liMsgIndex <> Undefined Then
                    pTranslateMsgIndex liMsgIndex, liClientTableIndex, liMask
                    iPtr = mtClients(liClientIndex).iPtr
                    With mtClients(liClientIndex).tWindows(liWindowIndex)
                        If Not CBool(.Flags And AllMessages) Then
                            If CBool(.Flags And CustomMessages) Then
                                If Not CBool(.iMsgMask(liClientTableIndex) And liMask) Then iPtr = 0&
                            Else
                                If Not mbDefAllMsgs Then
                                    If Not CBool(miDefMsgMask(liClientTableIndex) And liMask) Then iPtr = 0&
                                End If
                            End If
                        End If
                    End With
                Else
                    With mtClients(liClientIndex)
                        iPtr = .iPtr
                        With .tWindows(liWindowIndex)
                            If Not CBool(.Flags And AllMessages) Then
                                If CBool(.Flags And CustomMessages) Then
                                    Debug.Assert False
                                    iPtr = 0&
                                Else
                                    If Not mbDefAllMsgs Then iPtr = 0&
                                End If
                            End If
                        End With
                    End With
                End If
            Else
                Debug.Assert False
            End If
        Else
            Debug.Assert False
        End If
    Else
        If mbDefAllMsgs Then
            iPtr = miCreatingPtr
        Else
            If liMsgIndex <> Undefined Then
                pTranslateMsgIndex liMsgIndex, liClientTableIndex, liMask
                If CBool(miDefMsgMask(liClientTableIndex) And liMask) Then iPtr = miCreatingPtr
            End If
        End If
    End If
    
    Dim oThis As iWindow

    If iPtr <> 0& And iPtr <> Undefined Then        'If an object was found
        #If bVBVMTypeLib Then                          'If using the VBVM Type lib
            ObjectPtr(oThis) = iPtr                 'Set the object pointer directly
            oThis.Proc bHandled, lReturn, hWnd, _
                       iMsg, wParam, lParam         'call the object's procedure
            ObjectPtr(oThis) = 0&                   'remove the illegal reference
        #Else
            CopyMemory oThis, iPtr, 4&              'use the standard copymem method
            oThis.Proc bHandled, lReturn, hWnd, _
                       iMsg, wParam, lParam         'call the object's procedure
            CopyMemory oThis, 0&, 4&                'remove the illegal reference
        #End If
    Else
        'nobody wants this message
        'Debug.Assert False
    End If

End Sub
'</iWindow Interface>

