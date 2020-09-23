VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiWindowClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiWindowClass.cls                  7/4/04
'
'           GENERAL PURPOSE:
'               Exposes functionality to register api window classes through a familiar collection
'               interface.  Classes are created and destroyed by adding and removing from the
'               collection, and accessed and enumerated as items.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'==================================================================================================

'1.  Friendly Interface     -  Used internally to initialize the modular variable.
'2.  Message Interface      -  Methods which manage the messages that are delived to the owner of the window.
'3.  Window Interface       -  Function that returns a collection of api class window objects

Private msClass As String

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit(ByRef sClass As String)
    msClass = sClass
End Sub
'</Friendly Interface>

'<Public Interface>
Public Function Name() As String
Attribute Name.VB_Description = "Returns the class name."
    Name = msClass
End Function

Public Function TotalWindowCount() As Long
Attribute TotalWindowCount.VB_Description = "Returns the number of windows that have been created of this class through this component by any object."
    TotalWindowCount = mWindow.ApiWindowClass_WindowCount(msClass)
End Function

'<Message Interface>
Public Function AddDefMsg( _
            ByVal iMsg As eMsg) _
                As Boolean
Attribute AddDefMsg.VB_Description = "Adds a message to the default message table.  New windows created from the class defined by this object will use these default messages."
    AddDefMsg = mWindow.ApiWindowClass_AddDefMsg(iMsg, msClass)
End Function

Public Function DelDefMsg( _
            ByVal iMsg As eMsg) _
        As Boolean
Attribute DelDefMsg.VB_Description = "Removes a message from the default message table."
    DelDefMsg = mWindow.ApiWindowClass_DelDefMsg(iMsg, msClass)
End Function

Public Function DefMsgExists( _
            ByVal iMsg As eMsg) _
        As Boolean
Attribute DefMsgExists.VB_Description = "Returns a value indicating if the specified message exists in the default table."
    DefMsgExists = mWindow.ApiWindowClass_DefMsgExists(msClass, iMsg)
End Function

Public Function DefMsgCount() As Long
Attribute DefMsgCount.VB_Description = "Returns the number of messages in the default table."
    DefMsgCount = mWindow.ApiWindowClass_DefMsgCount(msClass)
End Function

Public Function GetDefMessages( _
            ByRef iOutArray() As Long) _
                As Long
Attribute GetDefMessages.VB_Description = "Copies the default message table into the array parameter, and returns the count."
    GetDefMessages = mWindow.ApiWindowClass_GetDefMessages(iOutArray, msClass)
End Function
'</Message Interface>

'</Window Interface>
Public Function OwnedWindows( _
            ByVal oWho As iWindow) _
                As cApiClassWindows
Attribute OwnedWindows.VB_Description = "Returns a collection of the windows created from this class."
    Set OwnedWindows = mWindow.ApiWindowClass_OwnedWindows(msClass, oWho)
End Function
'</Window Interface>
'</Public Interface>
