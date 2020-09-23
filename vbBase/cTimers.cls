VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cTimer.cls                             7/5/04
'
'           GENERAL PURPOSE:
'               Adds and removes api timers through a familiar collection interface.
'
'           MODULES CALLED FROM THIS MODULE:
'               mTimer.bas
'
'==================================================================================================

'1.  iEnumerable Interface    - For ... Each enumeration
'2.  Friendly Interface       - Used internally to initialize the modular variable.
'3.  Collection Interface     - Methods which manage the collection of cSubclass objects.

Implements iEnumerable

Private miPtr       As Long 'Used to store the objptr of the iTimer object on which to callback

'<Private Interface>
'</Private Interface>

'<iEnumerable Interface>
Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    mTimer.Timers_NextItem miPtr, tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = -1&
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
End Sub
'</iEnumerable Interface>

'<Friendly Interface>
Friend Sub fInit( _
            ByVal iPtr As Long)
    miPtr = iPtr
End Sub
'</Friendly Interface>

'<Public Interface>
Public Function Add( _
            ByVal ID As Long, _
   Optional ByVal iInterval As Long) _
                As cTimer
Attribute Add.VB_Description = "Adds a timer to the collection, returning True if successful."
    Set Add = mTimer.Timers_Add(miPtr, ID, iInterval)
End Function

Public Function Remove( _
            ByVal ID As Long) _
                As Boolean
Attribute Remove.VB_Description = "Removes a timer from the collection."
    mTimer.Timers_Remove miPtr, ID
End Function

Public Function Item( _
            ByVal ID As Long) _
                As cTimer
Attribute Item.VB_Description = "Returns a cTimer object representing the timer with the specified ID, or Nothing if the ID was not found."
    Set Item = mTimer.Timers_Item(miPtr, ID)
End Function

Public Function Exists( _
            ByVal ID As Long) _
                As Boolean
Attribute Exists.VB_Description = "Returns a value indicating whether the specified timer exists in the collection."
    Exists = mTimer.Timers_Exists(miPtr, ID)
End Function

Public Function Count() _
                As Long
Attribute Count.VB_Description = "Returns the number of timers that are in the collection."
    Count = mTimer.Timers_Count(miPtr)
End Function

Public Function Clear() _
                As Long
Attribute Clear.VB_Description = "Clears all timers from the collection."
    Clear = mTimer.Timers_Clear(miPtr)
End Function
    
Public Property Get NewEnum() _
                As IUnknown
Attribute NewEnum.VB_Description = "Returns a new enumeration to be used with a for ... each loop."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim loEnum As cEnumeration
    Set loEnum = New cEnumeration
    Set NewEnum = loEnum.GetEnum(Me, Timers_Control(miPtr))
End Property
'</Collection Interface>
'</Public Interface>
