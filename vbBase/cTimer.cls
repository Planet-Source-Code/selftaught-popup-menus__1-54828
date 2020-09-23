VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cTimer.cls                             7/5/04
'
'           GENERAL PURPOSE:
'               Wraps an api timer.
'
'           MODULES CALLED FROM THIS MODULE:
'               mTimer
'
'==================================================================================================

'1.  Friendly Interface  - Used internally to initialize the modular variable.
'2.  Timer Interface     - Methods to start/stop and get information about the timer

Private miPtr       As Long 'Used to store the objptr of the iTimer object on which to callback
Private miID        As Long 'used to store the id associated with this timer

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit( _
        ByVal iPtr As Long, _
        ByVal ID As Long)
    miPtr = iPtr
    miID = ID
End Sub
'</Friendly Interface>

'<Public Interface>
'<Timer Interface>
Public Property Get Active() As Boolean
Attribute Active.VB_Description = "Returns a value indicating whether the timer is active."
    Active = mTimer.Timer_Active(miPtr, miID)
End Property

Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns the interval of the timer."
    Interval = mTimer.Timer_Interval(miPtr, miID)
End Property

Public Property Let Interval(ByVal iNew As Long)
    mTimer.Timer_Interval(miPtr, miID) = iNew
End Property

Public Function Start( _
            Optional ByVal iInterval As Long = -1&, _
            Optional ByVal bRestartOK As Boolean = True) _
                As Boolean
Attribute Start.VB_Description = "Starts the timer, optionally using the last specified interval or a new one, and optionally allowing a restart to be done if necessary."
    Start = mTimer.Timer_Start(miPtr, iInterval, miID, bRestartOK)
End Function

Public Function StopTimer() As Boolean
Attribute StopTimer.VB_Description = "Stops the timer."
    StopTimer = mTimer.Timer_Stop(miPtr, miID)
End Function

Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns the ID that uniquely identifies the timer within the collection."
    ID = miID
End Property
'</Timer Interface>
'</Public Interface>
