VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cPopupMenuItem - Exposes procedures to manipulate the appearance and behavior of a menu item.
'
'Copyright free, use and abuse as you see fit.
'==================================================================================================

'1.  Friendly Interface  - Initialization
'2.  Public Interface

Private mtPointer As tMenuItemPointer
Private moOwner As cPopupMenus

'<Friendly Interface>
Friend Sub fInit(ByVal oOwner As cPopupMenus, ByRef tPointer As tMenuItemPointer)
    Set moOwner = oOwner
    LSet mtPointer = tPointer
End Sub
'</Friendly Interface>

'<Public Interface>
Public Property Get IconIndex() As Long
    IconIndex = moOwner.fPopupMenuItem_IconIndex(mtPointer)
End Property
Public Property Let IconIndex(ByVal iNew As Long)
    moOwner.fPopupMenuItem_IconIndex(mtPointer) = iNew
End Property

Public Property Get ShortCutShiftMask() As Integer
    ShortCutShiftMask = moOwner.fPopupMenuItem_ShortCutShiftMask(mtPointer)
End Property
Public Property Let ShortCutShiftMask(ByVal iNew As Integer)
    moOwner.fPopupMenuItem_ShortCutShiftMask(mtPointer) = iNew
End Property

Public Property Get ShortCutShiftKey() As Integer
    ShortCutShiftKey = moOwner.fPopupMenuItem_ShortCutShiftKey(mtPointer)
End Property
Public Property Let ShortCutShiftKey(ByVal iNew As Integer)
    moOwner.fPopupMenuItem_ShortCutShiftKey(mtPointer) = iNew
End Property

Public Property Get Caption() As String
    Caption = moOwner.fPopupMenuItem_Caption(mtPointer)
End Property
Public Property Let Caption(ByRef sNew As String)
    moOwner.fPopupMenuItem_Caption(mtPointer) = sNew
End Property

Public Property Get Index() As Long
    Index = moOwner.fPopupMenuItem_Index(mtPointer)
End Property

Public Property Get Key() As String
    Key = moOwner.fPopupMenuItem_Key(mtPointer)
End Property

Public Property Get ShortCutDisplay() As String
    ShortCutDisplay = moOwner.fPopupMenuItem_ShortCutDisplay(mtPointer)
End Property
Public Property Let ShortCutDisplay(ByRef sNew As String)
    moOwner.fPopupMenuItem_ShortCutDisplay(mtPointer) = sNew
End Property

Public Sub SetStyle(Optional ByVal iStateOr As ePopupMenuItemStyle, Optional ByVal iStateAndNot As ePopupMenuItemStyle)
    Style = ((Style() Or iStateOr) And Not iStateAndNot)
End Sub

Public Property Get Style() As ePopupMenuItemStyle
    Style = moOwner.fPopupMenuItem_Style(mtPointer)
End Property

Public Property Let Style(ByVal iNew As ePopupMenuItemStyle)
    moOwner.fPopupMenuItem_Style(mtPointer) = iNew
End Property

Public Property Get Enabled() As Boolean
    Enabled = Not CBool(Style() And mnuDisabled)
End Property

Public Property Let Enabled(ByVal bNew As Boolean)
    If bNew Then
        SetStyle , mnuDisabled
    Else
        SetStyle mnuDisabled
    End If
End Property

Public Property Get Separator() As Boolean
    Separator = CBool(Style() And mnuSeparator)
End Property

Public Property Let Separator(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuSeparator
    Else
        SetStyle , mnuSeparator
    End If
End Property

Public Property Get Default() As Boolean
    Default = CBool(Style() And mnuDefault)
End Property

Public Property Let Default(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuDefault
    Else
        SetStyle , mnuDefault
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = Not CBool(Style() And mnuInvisible)
End Property

Public Property Let Visible(ByVal bNew As Boolean)
    If bNew Then
        SetStyle , mnuInvisible
    Else
        SetStyle mnuInvisible
    End If
End Property

Public Property Get Infrequent() As Boolean
    Infrequent = CBool(Style() And mnuInfrequent)
End Property

Public Property Let Infrequent(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuInfrequent
    Else
        SetStyle , mnuInfrequent
    End If
End Property

Public Property Get OnNewVerticalLine() As Boolean
    OnNewVerticalLine = CBool(Style() And mnuNewVerticalLine)
End Property

Public Property Let OnNewVerticalLine(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuNewVerticalLine
    Else
        SetStyle , mnuNewVerticalLine
    End If
End Property

Public Property Get Checked() As Boolean
    Checked = CBool(Style() And mnuChecked)
End Property

Public Property Let Checked(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuChecked
    Else
        SetStyle , mnuChecked
    End If
End Property

Public Property Get RadioChecked() As Boolean
    RadioChecked = CBool(Style() And mnuRadioChecked)
End Property

Public Property Let RadioChecked(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuRadioChecked
    Else
        SetStyle , mnuRadioChecked
    End If
End Property

Public Property Get RedisplayOnClick() As Boolean
    RedisplayOnClick = CBool(Style() And mnuRedisplayOnClick)
End Property

Public Property Let RedisplayOnClick(ByVal bNew As Boolean)
    If bNew Then
        SetStyle mnuRedisplayOnClick
    Else
        SetStyle , mnuRedisplayOnClick
    End If
End Property

Public Property Get SubItems() As cPopupMenuItems
    Set SubItems = moOwner.fPopupMenuItem_SubItems(mtPointer)
End Property
'</Public Interface>
