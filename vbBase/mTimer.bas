Attribute VB_Name = "mTimer"
Option Explicit

'==================================================================================================
'mTimer.bas                             7/5/04
'
'           GENERAL PURPOSE:
'               Uses a separate pcTimer object for each timer requested.  The pcTimer object is
'               responsible for delivering the callbacks to the requesting objects.
'
'           CLASSES CREATED BY THIS MODULE:
'               pcTimer
'               cTimer
'
'==================================================================================================

'1.  Private Interface      - Utility procedures - array redim and find a timer
'2.  cTimer Interface       - Procedures called by cTimer
'3.  cTimers Interface      - Procedures called by cTimers

Private Type tTimerClient
    iPtr As Long
    iTimerCount As Long
    oTimers() As pcTimer
    iControl As Long
End Type

Private mtClients() As tTimerClient
Private miClientCount As Long

Const msTimerObjectName = "cTimer"

'<Private Interface>
Private Function FindClient( _
            ByVal iClient As Long, _
   Optional ByRef iFirstAvailable As Long) _
                As Long
    iFirstAvailable = Undefined
    
    For FindClient = 0& To miClientCount - 1&                          'Loop through each timer
        If mtClients(FindClient).iPtr <> Undefined Then
            If mtClients(FindClient).iPtr = iClient Then Exit Function
        Else
            If iFirstAvailable = Undefined Then iFirstAvailable = FindClient
        End If
    Next
    FindClient = Undefined                                            'if we made it out here, then the timer was not found.
End Function

Private Function FindTimer( _
            ByVal iClientIndex As Long, _
            ByVal iId As Long, _
   Optional ByRef iFirstAvailable As Long) _
                As Long
    iFirstAvailable = Undefined
    With mtClients(iClientIndex)
        For FindTimer = 0 To .iTimerCount - 1&
            If Not (.oTimers(FindTimer) Is Nothing) Then
                If .oTimers(FindTimer).ID = iId Then Exit Function
            Else
                If iFirstAvailable = Undefined Then iFirstAvailable = FindTimer
            End If
        Next
    End With
    FindTimer = Undefined
End Function

'This is a private implementation of ArrRedim, to allow strong typing
'without cluttering the public interface of the general module.
Private Sub ArrRedimO( _
            ByRef oArray() As pcTimer, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundO(oArray) Then
        If bPreserve Then _
            ReDim Preserve oArray(0 To liNewUbound) _
        Else _
            ReDim oArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundO( _
            ByRef oArray() As pcTimer) _
                As Long
    On Error Resume Next
    ArrUboundO = UBound(oArray)
    If Err.Number <> 0& Then ArrUboundO = Undefined
End Function

Private Sub ArrRedimT( _
            ByRef tArray() As tTimerClient, _
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
            ByRef tArray() As tTimerClient) _
                As Long
    On Error Resume Next
    ArrUboundT = UBound(tArray)
    If Err.Number <> 0& Then ArrUboundT = Undefined
End Function
'</Private Interface>

'<Public Interface>
'<cTimer Interface>
Public Function Timer_Start( _
            ByVal iWho As Long, _
            ByVal iInterval As Long, _
            ByVal iId As Long, _
   Optional ByVal bRestartOK As Boolean = True) _
                As Boolean
            
    Dim liClientIndex As Long
    Dim liTimerIndex As Long
    
    liClientIndex = FindClient(iWho)
    
    If liClientIndex <> Undefined Then
        liTimerIndex = FindTimer(liClientIndex, iId)
        If liTimerIndex <> Undefined Then
            With mtClients(liClientIndex).oTimers(liTimerIndex)
                If .Active And bRestartOK Then .Destroy             'if restart OK, then stop it if it's already active
                If Not .Active Then Timer_Start = .Create(iInterval) 'create the timer
            End With
            Exit Function
        End If
        
    End If
    gErr vbbItemDetached, msTimerObjectName
End Function
        
Public Function Timer_Stop( _
            ByVal iWho As Long, _
            ByVal iId As Long) _
                As Boolean
    Dim liClientIndex As Long
    Dim liTimerIndex As Long
    
    liClientIndex = FindClient(iWho)
    
    If liClientIndex <> Undefined Then
        liTimerIndex = FindTimer(liClientIndex, iId)
        If liTimerIndex <> Undefined Then
            With mtClients(liClientIndex).oTimers(liTimerIndex)
                If .Active Then Timer_Stop = .Destroy 'if it's active, destroy it
            End With
            Exit Function
        End If
    End If
    gErr vbbItemDetached, msTimerObjectName
End Function
        
Public Function Timer_Active( _
            ByVal iWho As Long, _
            ByVal iId As Long) _
                As Boolean
    Dim liClientIndex As Long
    Dim liTimerIndex As Long
    
    liClientIndex = FindClient(iWho)
    
    If liClientIndex <> Undefined Then
        liTimerIndex = FindTimer(liClientIndex, iId)
        If liTimerIndex <> Undefined Then
            With mtClients(liClientIndex).oTimers(liTimerIndex)
                Timer_Active = .Active
            End With
            Exit Function
        End If
    End If
    gErr vbbItemDetached, msTimerObjectName
End Function

Public Property Get Timer_Interval( _
            ByVal iWho As Long, _
            ByVal iId As Long) _
                As Long
    Dim liClientIndex As Long
    Dim liTimerIndex As Long
    
    liClientIndex = FindClient(iWho)
    
    If liClientIndex <> Undefined Then
        liTimerIndex = FindTimer(liClientIndex, iId)
        If liTimerIndex <> Undefined Then
            With mtClients(liClientIndex).oTimers(liTimerIndex)
                Timer_Interval = .Interval  'if it was found query it for the interval
            End With
            Exit Property
        End If
    End If
    gErr vbbItemDetached, msTimerObjectName
End Property

Public Property Let Timer_Interval( _
            ByVal iWho As Long, _
            ByVal iId As Long, _
            ByVal iInt As Long)
    
    If mTimer.Timer_Active(iWho, iId) Then
        mTimer.Timer_Start iWho, iId, iInt, True
    Else
        
        Dim liClientIndex As Long
        Dim liTimerIndex As Long
        
        liClientIndex = FindClient(iWho)
        
        If liClientIndex <> Undefined Then
            liTimerIndex = FindTimer(liClientIndex, iId)
            If liTimerIndex <> Undefined Then
                mtClients(liClientIndex).oTimers(liTimerIndex).Interval = iInt
                Exit Property
            End If
        End If
        gErr vbbItemDetached, msTimerObjectName
        
    End If
    
End Property

'</cTimer Interface>

'<cTimers Interface>
Public Function Timers_Add( _
            ByVal iWho As Long, _
            ByVal iId As Long, _
            ByVal iInterval As Long) _
                As cTimer
        
    Dim liClientIndex   As Long
    Dim liTimerIndex    As Long
    Dim liFirst         As Long
        
    liClientIndex = FindClient(iWho, liFirst)
        
    If liClientIndex = Undefined Then
        If liFirst = Undefined Then
            liClientIndex = miClientCount
            miClientCount = miClientCount + 1&
            ArrRedimT mtClients, miClientCount, True
        Else
            liClientIndex = liFirst
        End If
        mtClients(liClientIndex).iPtr = iWho
    End If
    
    liTimerIndex = FindTimer(liClientIndex, iId, liFirst)
    
    If liTimerIndex = Undefined Then
        With mtClients(liClientIndex)
            If liFirst = Undefined Then
                liTimerIndex = .iTimerCount
                .iTimerCount = .iTimerCount + 1&
                ArrRedimO .oTimers, .iTimerCount, True
            Else
                liTimerIndex = liFirst
            End If
            
            Dim loTimer As pcTimer
            Set loTimer = New pcTimer
            
            
            With loTimer
                .Owner = iWho
                .ID = iId
                .Interval = iInterval
            End With
            
            Set .oTimers(liTimerIndex) = loTimer
            
            Incr .iControl
            
            Set Timers_Add = New cTimer
            Timers_Add.fInit iWho, iId
        End With
    Else
        'adding a timer id that is already there
        gErr vbbKeyAlreadyExists, "cTimers.Add"
    End If

End Function

Public Sub Timers_Remove( _
            ByVal iWho As Long, _
            ByVal iId As Long)
    
    Dim liClientIndex   As Long
    Dim liTimerIndex    As Long
    
    liClientIndex = FindClient(iWho)
        
    If liClientIndex <> Undefined Then
    
        liTimerIndex = FindTimer(liClientIndex, iId)
    
        If liTimerIndex <> Undefined Then
            With mtClients(liClientIndex)
                Set .oTimers(liTimerIndex) = Nothing
                If liTimerIndex = .iTimerCount - 1& Then
                    For .iTimerCount = liTimerIndex - 1& To 0 Step -1&
                        If Not (.oTimers(.iTimerCount) Is Nothing) Then Exit For
                    Next
                    .iTimerCount = .iTimerCount + 1&
                    If .iTimerCount = 0 Then .iPtr = Undefined
                    
                End If
                Incr .iControl
            End With
            Exit Sub
        End If
    End If
gErr vbbKeyNotFound, "cTimers.Remove"
End Sub

Public Function Timers_Exists( _
            ByVal iWho As Long, _
            ByVal iId As Long) _
                    As Boolean
    Dim liClientIndex   As Long
    
    liClientIndex = FindClient(iWho)
        
    If liClientIndex <> Undefined Then

        Timers_Exists = (FindTimer(liClientIndex, iId) <> Undefined)

    End If
End Function

Public Function Timers_Item( _
            ByVal iWho As Long, _
            ByVal iId As Long) _
                    As cTimer
    Dim liClientIndex   As Long
    
    liClientIndex = FindClient(iWho)
        
    If liClientIndex <> Undefined Then
        
        If (FindTimer(liClientIndex, iId) <> Undefined) Then
            Set Timers_Item = New cTimer
            Timers_Item.fInit iWho, iId
            Exit Function
        End If

    End If
    
    gErr vbbKeyNotFound, "cTimers.Item"
End Function

Public Function Timers_Count( _
            ByVal iWho As Long) _
                As Long
    Dim i As Long
    Dim liClientIndex   As Long
    
    liClientIndex = FindClient(iWho)
        
    If liClientIndex <> Undefined Then
        With mtClients(liClientIndex)
            For i = 0 To .iTimerCount - 1&
                
                If Not .oTimers(i) Is Nothing _
                    Then Timers_Count = Timers_Count + 1&
            
            Next
        End With
    Else
        'client not found!
        'Debug.Assert False
    End If
End Function

Public Function Timers_Clear( _
            ByVal iWho As Long) _
                As Long
    Dim i As Long
    Dim liClientIndex   As Long
    
    liClientIndex = FindClient(iWho)
        
    If liClientIndex <> Undefined Then
        With mtClients(liClientIndex)
            For i = 0 To .iTimerCount - 1&
                If Not .oTimers(i) Is Nothing Then
                    Timers_Clear = Timers_Clear + 1&
                    Set .oTimers(i) = Nothing
                End If
            Next
            .iTimerCount = 0&
            .iPtr = Undefined
            If Timers_Clear Then Incr .iControl
        End With
    Else
        'client not found!
        'Debug.Assert False
    End If
    
End Function

Public Sub Timers_NextItem( _
            ByVal iWho As Long, _
            ByRef tEnum As tEnum, _
            ByRef vNextItem As Variant, _
            ByRef bNoMore As Boolean)
            
    Dim liClientIndex As Long
    
    liClientIndex = FindClient(iWho)
    
    If liClientIndex <> Undefined Then
        With mtClients(liClientIndex)
            If .iControl = tEnum.iControl Then
                
                Dim loTimer As cTimer
                Dim i As Long
                
                For tEnum.iIndex = tEnum.iIndex + 1& To .iTimerCount - 1&
                    If Not .oTimers(tEnum.iIndex) Is Nothing Then
                        Set loTimer = New cTimer
                        loTimer.fInit iWho, .oTimers(tEnum.iIndex).ID
                        Set vNextItem = loTimer
                        Exit For
                    End If
                Next
                
                If tEnum.iIndex = .iTimerCount Then bNoMore = True
            
                Exit Sub
            End If
        End With
        
        gErr vbbCollChangedDuringEnum, "cTimers.NewEnum"
    Else
        bNoMore = True
    End If
End Sub

Public Function Timers_Control(ByVal iWho As Long)
    Timers_Control = FindClient(iWho)
    If Timers_Control <> Undefined Then
        Timers_Control = mtClients(Timers_Control).iControl
    Else
        'Debug.Assert False
    End If
End Function

'</cTimers Interface>
'</Public Interface>
    
