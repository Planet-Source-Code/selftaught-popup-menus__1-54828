VERSION 5.00
Begin VB.UserControl ucWindowInfo 
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   ScaleHeight     =   3975
   ScaleWidth      =   5325
   Begin VB.VScrollBar vsbPos 
      Height          =   240
      Index           =   1
      Left            =   3360
      Min             =   -32767
      SmallChange     =   5
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2070
      Value           =   32767
      Width           =   240
   End
   Begin Project1.ucFindWindow FindWindow 
      Height          =   495
      Left            =   3360
      TabIndex        =   23
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VB.VScrollBar vsbPos 
      Height          =   240
      Index           =   3
      Left            =   5040
      Min             =   -32767
      SmallChange     =   5
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2070
      Value           =   32767
      Width           =   240
   End
   Begin VB.VScrollBar vsbPos 
      Height          =   240
      Index           =   2
      Left            =   4200
      Min             =   -32767
      SmallChange     =   5
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2070
      Value           =   32767
      Width           =   240
   End
   Begin VB.VScrollBar vsbPos 
      Height          =   240
      Index           =   0
      LargeChange     =   50
      Left            =   2520
      Min             =   -32767
      SmallChange     =   5
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2070
      Value           =   32767
      Width           =   240
   End
   Begin VB.CheckBox chk 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled:"
      Height          =   315
      Left            =   540
      TabIndex        =   10
      Top             =   2055
      Width           =   975
   End
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   0
      ItemData        =   "uWindowInfo.ctx":0000
      Left            =   0
      List            =   "uWindowInfo.ctx":00E8
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2520
      Width           =   2625
   End
   Begin VB.ListBox lst 
      Height          =   1410
      Index           =   1
      ItemData        =   "uWindowInfo.ctx":0237
      Left            =   2670
      List            =   "uWindowInfo.ctx":02AA
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   2520
      Width           =   2625
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtPos 
      Height          =   285
      Index           =   3
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtPos 
      Height          =   285
      Index           =   2
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtPos 
      Height          =   285
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtPos 
      Height          =   285
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Drag the finder tool to select a new parent."
      Height          =   495
      Index           =   3
      Left            =   3345
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Left            Top             Width          Height"
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   21
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   12
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Window Handle:"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Index           =   7
      Left            =   1320
      MouseIcon       =   "uWindowInfo.ctx":0443
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner:"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Window Text:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lbl 
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Class Name:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Index           =   5
      Left            =   1320
      MouseIcon       =   "uWindowInfo.ctx":074D
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Parent:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "ucWindowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eTxt
    txtLeft
    txtTop
    txtWidth
    txtHeight
End Enum

Private Enum eLbl
    lblClass = 1
    lblParent = 5
    lblOwner = 7
    lblWindow = 9
End Enum

Private Enum eLst
    lstStyle
    lstStyleEx
End Enum

Private moApiWindow As cApiWindow
Private mbFreeze As Boolean

Event RequestWindow(ByVal iWindow As Long, ByRef bCancel As Boolean)

Private Sub FindWindow_WindowSelected(ByVal hWndNew As Long)
    If Not moApiWindow Is Nothing Then
        If MsgBox("Do you want to try to change the parent of this window to 0x" & FmtHex(hWndNew) & "?", vbYesNo + vbDefaultButton2, "Change Parent") = vbYes Then
            moApiWindow.Parent = hWndNew
            hWndDisplay = moApiWindow.hwnd
        End If
    End If
End Sub

Private Sub txtPos_Change(Index As Integer)
    On Error Resume Next
    If Not mbFreeze Then
        If Not moApiWindow Is Nothing Then moApiWindow.Move txtPos(txtLeft).Text, txtPos(txtTop).Text, txtPos(txtWidth).Text, txtPos(txtHeight).Text
    End If
End Sub

Private Sub chk_Click()
    On Error Resume Next
    If Not mbFreeze Then
        If Not moApiWindow Is Nothing Then moApiWindow.Enabled = (chk.Value = vbChecked)
    End If
End Sub

Private Sub lbl_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case lblParent, lblOwner
        Dim lhWnd As Long, bCancel As Boolean
        lhWnd = HexVal(lbl(Index).Caption)
        RaiseEvent RequestWindow(lhWnd, bCancel)
        If Not bCancel Then hWndDisplay = lhWnd
    End Select
End Sub

Private Sub lst_Click(Index As Integer)
    On Error Resume Next
    If Not mbFreeze Then
        If Not moApiWindow Is Nothing Then
            Dim liStyle As Long
            liStyle = lstOrAllItemData(lst(Index))
            Select Case Index
            Case lstStyle
                moApiWindow.SetStyle liStyle, -1&
            Case lstStyleEx
                moApiWindow.SetStyleEx liStyle, -1&
            End Select
        End If
    End If
End Sub

Private Sub txt_Change()
    On Error Resume Next
    If Not mbFreeze Then
        If Not moApiWindow Is Nothing Then moApiWindow.Text = txt.Text
    End If
End Sub

Private Sub vsbPos_Change(Index As Integer)
    On Error Resume Next
    txtPos(Index).Text = -vsbPos(Index).Value
End Sub

Private Sub vsbPos_GotFocus(Index As Integer)
    On Error Resume Next
    txtPos(Index).SetFocus
End Sub

Public Property Let hWndDisplay(ByVal ihWnd As Long)
    On Error Resume Next
    Dim bVal As Boolean
    
    Set moApiWindow = xApiWindow(ihWnd)
    mbFreeze = True
    If Not moApiWindow Is Nothing Then
        With moApiWindow
            Dim liLeft As Long, liTop As Long, liWidth As Long, liHeight As Long
            .GetPos liLeft, liTop, liWidth, liHeight
            With vsbPos
                .Item(txtLeft).Value = -liLeft
                .Item(txtTop).Value = -liTop
                .Item(txtWidth).Value = -liWidth
                .Item(txtHeight).Value = -liHeight
            End With
            chk.Value = IIf(.Enabled, vbChecked, vbUnchecked)
            txt.Text = .Text
            lbl(lblClass).Caption = .ClassName
            lbl(lblOwner).Caption = "0x" & FmtHex(.Owner)
            lbl(lblParent).Caption = "0x" & FmtHex(.Parent)
            lbl(lblWindow).Caption = "0x" & FmtHex(.hwnd)
            lstSetChecksFromItemData lst(lstStyle), .GetStyle
            lstSetChecksFromItemData lst(lstStyleEx), .GetStyleEx
        End With
        bVal = True
    Else
        With vsbPos
            .Item(txtLeft).Value = 0
            .Item(txtTop).Value = 0
            .Item(txtWidth).Value = 0
            .Item(txtHeight).Value = 0
        End With
        chk.Value = vbUnchecked
        txt.Text = vbNullString
        lbl(lblClass).Caption = vbNullString
        lbl(lblOwner).Caption = vbNullString
        lbl(lblParent).Caption = vbNullString
        lbl(lblWindow).Caption = "(Invalid)"
        
        lstClear lst(lstStyle)
        lstClear lst(lstStyleEx)
    End If

    txt.Enabled = bVal
    chk.Enabled = bVal
    lbl(lblOwner).MousePointer = IIf(bVal, vbCustom, vbDefault)
    lbl(lblParent).MousePointer = IIf(bVal, vbCustom, vbDefault)
    lst(lstStyle).Enabled = bVal
    lst(lstStyleEx).Enabled = bVal
    With vsbPos
        .Item(txtLeft).Enabled = bVal
        .Item(txtTop).Enabled = bVal
        .Item(txtWidth).Enabled = bVal
        .Item(txtHeight).Enabled = bVal
    End With
    
    With txtPos
        .Item(txtLeft).Enabled = bVal
        .Item(txtTop).Enabled = bVal
        .Item(txtWidth).Enabled = bVal
        .Item(txtHeight).Enabled = bVal
    End With
            
    FindWindow.Enabled = bVal
        
    mbFreeze = False
End Property
