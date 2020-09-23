VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenuTest 
   Caption         =   "Popup Menu Test"
   ClientHeight    =   5550
   ClientLeft      =   2505
   ClientTop       =   2910
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6360
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5670
      Index           =   2
      Left            =   5520
      Picture         =   "fTest.frx":030A
      ScaleHeight     =   5670
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Index           =   1
      Left            =   5640
      Picture         =   "fTest.frx":61CC
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "&Show Menu"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Click to show a demonstration menu with sub levels."
      Top             =   60
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ilsIcons16 
      Left            =   120
      Top             =   4500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1220E
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":12528
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":12842
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":12B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":12E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":13190
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":134AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":137C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":13ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":13DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":14112
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1442C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":14746
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":14A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":14D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":15094
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":153AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":156C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":159E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":15CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":16016
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":16330
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1664A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":16964
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":16C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":16F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":172B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":175CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":178E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":17C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":17F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":18234
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1854E
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":18868
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":18B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":18E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":191B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":194D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":197EA
            Key             =   "Web"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":19B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":19E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1A138
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":1A452
            Key             =   "vbAccelerator"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H80000005&
      Height          =   2775
      Index           =   0
      Left            =   60
      ScaleHeight     =   2715
      ScaleWidth      =   6375
      TabIndex        =   2
      Top             =   480
      Width           =   6435
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Image &Process Bitmap for Highlights"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bac&kground Bitmap"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   2295
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Title Style Separators"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show &Infrequently Used"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Office XP Style"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1095
         Width           =   2295
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Gradient Highlight"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1335
         Width           =   2295
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Button Highlight"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   1575
         Width           =   2295
      End
      Begin VB.CommandButton cmdVisual 
         Caption         =   "Font"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   10
         ToolTipText     =   "Click to show a demonstration menu with sub levels."
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdVisual 
         Caption         =   "Active Fore"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Click to show a demonstration menu with sub levels."
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdVisual 
         Caption         =   "Active Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Click to show a demonstration menu with sub levels."
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdVisual 
         Caption         =   "Inactive Fore"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click to show a demonstration menu with sub levels."
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdVisual 
         Caption         =   "Inactive Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Click to show a demonstration menu with sub levels."
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Custom Font"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   2520
         TabIndex        =   5
         Top             =   15
         Width           =   2295
      End
      Begin VB.CheckBox chkVisual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Custom Colors"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   2520
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         Height          =   1515
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Right click to get an Edit popup menu"
         Top             =   1920
         Width           =   5355
      End
   End
End
Attribute VB_Name = "frmMenuTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miLastKey As Integer

' ======================================================================================
'
' Name:     vbAccelerator VB6 PopupMenu Component Demonstrator
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     18 February 2001
'
' Requires: cNewMenu6.DLL
'           SSUBTMR6.DLL
'
' Copyright © 1998-2001 Steve McMahon for vbAccelerator
'
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------

Private moCommon As cCommonDialog
Private WithEvents moPopupMenus As cPopupMenus
Attribute moPopupMenus.VB_VarHelpID = -1
Private Const mcWEBSITE = -&H8000&

Private Sub Status(ByVal sMsg As String)
   lstStatus.AddItem sMsg
   lstStatus.ListIndex = lstStatus.NewIndex
End Sub


Private Sub chkVisual_Click(Index As Integer)
    Dim lbVal As Boolean
    lbVal = chkVisual(Index).Value = vbChecked
    Select Case Index
    Case 0 'background
        If lbVal Then
            chkVisual(1).Enabled = True
            moPopupMenus.ImageProcessBitmap = (chkVisual(1).Value = vbChecked)
            Set moPopupMenus.BackgroundPicture = pic(1).Picture
        Else
            chkVisual(1).Enabled = False
            Set moPopupMenus.BackgroundPicture = Nothing
        End If
    Case 1
        moPopupMenus.ImageProcessBitmap = lbVal
    Case 2
        moPopupMenus.DrawSeparatorsAsHeaders = lbVal
    Case 3
        moPopupMenus.ShowInfrequent = lbVal
    Case 4
        moPopupMenus.OfficeXPStyle = lbVal
    Case 5
        moPopupMenus.GradientHighlight = lbVal
    Case 6
        moPopupMenus.ButtonHighlight = lbVal
    Case 7
        With cmdVisual(0)
            .Enabled = lbVal
            Set moPopupMenus.Font = IIf(lbVal, .Font, Nothing)
            Set .Font = moPopupMenus.Font
            .Caption = moPopupMenus.Font.Name
        End With
    Case 8
        With cmdVisual
            .Item(1).Enabled = lbVal
            .Item(2).Enabled = lbVal
            .Item(3).Enabled = lbVal
            .Item(4).Enabled = lbVal
            If lbVal Then
                moPopupMenus.ActiveForeColor = .Item(1).BackColor
                moPopupMenus.ActiveBackColor = .Item(2).BackColor
                moPopupMenus.InActiveForeColor = .Item(3).BackColor
                moPopupMenus.InActiveBackColor = .Item(4).BackColor
            Else
                moPopupMenus.ActiveForeColor = -1
                moPopupMenus.ActiveBackColor = -1
                moPopupMenus.InActiveForeColor = -1
                moPopupMenus.InActiveBackColor = -1
                .Item(1).BackColor = moPopupMenus.ActiveForeColor
                .Item(2).BackColor = moPopupMenus.ActiveBackColor
                .Item(3).BackColor = moPopupMenus.InActiveForeColor
                .Item(4).BackColor = moPopupMenus.InActiveBackColor
            End If
        End With
    End Select
End Sub

Private Sub cmdMenu_Click(Index As Integer)
    Dim lhWnd As Long
    lhWnd = cmdMenu(Index).hWnd
    Select Case Index
        Case 0
            moPopupMenus.Item("Demo").ShowAtWindow lhWnd, mnuPreserveVertAlign, lhWnd
    End Select
End Sub


'Private Sub cmdTest_Click()
'Dim i As Long
'Dim j As Long
'Dim k As Long
'Dim n As Long
'   With moPopupMenus
'      .Clear
'      For i = 1 To 10
'         i = .AddItem("Test" & i, , , , , , , "Test" & i)
'      Next i
'      For i = 1 To 10
'         k = .InsertItem("InsTest" & i, "Test3", , , , , , "Test" & j)
'         If i = 3 Then
'            For j = 1 To 10
'               .AddItem "SubTest" & j, , , k, , , , "SubTest" & j
'            Next j
'         End If
'      Next i
'      k = .InsertItem("InsTOP", "Test1", , , , , , "InsTOP")
'      k = .AddItem("InsTopSub 1", , , k, , , , "InsTopSub1")
'      For j = 1 To 4
'         .InsertItem "InsTopSub " & j + 1, "InsTopSub1", , , , , , "InsTopSub" & j + 1
'      Next j
'      k = .InsertItem("InsBOTTOM", "Test10", , , , , , "InsBOTTOM")
'      For j = 1 To 5
'         .AddItem "InsBottom" & j, , , k, , , , "InsBottom" & j
'      Next j
'
'      .ShowPopupMenu 0, 0
'
'      .ClearSubMenusOfItem "InsTOP"
'      k = .IndexForKey("InsTOP")
'      For j = 1 To 24
'         .AddItem "InsTopSub " & j, , , k, , , , "InsTopSub" & j
'      Next j
'      .ClearSubMenusOfItem "InsBOTTOM"
'      k = .IndexForKey("InsTopSub20")
'      For j = 1 To 24
'         i = .AddItem("InsTopSubSub " & j, , , k, , , , "InsTopSubSub" & j)
'         If j Mod 5 = 0 Then
'            For n = 1 To Rnd * 8 + 4
'               .AddItem "Testing" & n, , , i
'            Next n
'         End If
'      Next j
'
'      .ShowPopupMenu 0, 0
'
'
'
'
'   End With
'End Sub
'
'Private Sub cmdTest2_Click()
'   With moPopupMenus
'      .RestoreFromFile , "C:\Stevemac\VB\Controls\vbalTbar\Menu.dat"
'      .Restore "Main"
'      .ShowPopupMenu 0, 0
'   End With
'End Sub


Private Sub cmdVisual_Click(Index As Integer)
    Select Case Index
    Case 0
        Dim loFont As StdFont
        Set loFont = cmdVisual(Index).Font
        If moCommon.VBChooseFont(loFont, , hWnd) Then
            Set moPopupMenus.Font = loFont
            Set cmdVisual(Index).Font = loFont
            cmdVisual(Index).Caption = loFont.Name
        End If
    Case Else
        Dim liNew As OLE_COLOR
        liNew = cmdVisual(Index).BackColor
        If moCommon.VBChooseColor(liNew, True, True, False, hWnd) Then
            cmdVisual(Index).BackColor = liNew
            Select Case Index
            Case 1
                moPopupMenus.ActiveForeColor = liNew
            Case 2
                moPopupMenus.ActiveBackColor = liNew
            Case 3
                moPopupMenus.InActiveForeColor = liNew
            Case 4
                moPopupMenus.InActiveBackColor = liNew
            End Select
        End If
    End Select
End Sub

Private Sub moPopupMenus_Click(ByVal Item As cPopupMenuItem)
   Status "Clicked Item=" & Item.Index & ";Caption=" & Item.Caption
   'If Item.Key = "CHECK" Then
   Item.RadioChecked = Not Item.RadioChecked
End Sub


Private Sub moPopupMenus_InitPopupMenu(ByVal Items As cPopupMenuItems, ByVal ChevronAdded As cPopupMenuItem)
    Dim loItem As cPopupMenuItem
    Set loItem = Items.Parent

    If loItem Is Nothing Then
        Status "InitPopupMenu root popup with " & Items.Root.Key
    Else
        Status "InitPopupMenu with parent " & loItem.Key
    End If
End Sub

Private Sub moPopupMenus_ItemHighlight(ByVal Item As cPopupMenuItem)
    If Not Item Is Nothing Then Status "Highlighted  Item=" & Item.Index & ",Caption=" & Item.Caption & ", Enabled=" & Item.Enabled & ", Separator = " & Item.Separator Else Status "Highlighted Nothing"
End Sub

Private Sub moPopupMenus_RightClick(ByVal Item As cPopupMenuItem)
    Status "Right Click " & Item.Caption
End Sub

Private Sub moPopupMenus_UnInitPopupMenu(ByVal Items As cPopupMenuItems)
    Dim loItem As cPopupMenuItem
    Set loItem = Items.Parent
    
    If loItem Is Nothing Then
        Status "UnInitPopupMenu root popup with " & Items.Root.Key
    Else
        Status "UnInitPopupMenu with parent " & loItem.Key
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then Exit Sub
    If KeyCode = vbKeyF And Shift = vbCtrlMask Then
        With moPopupMenus.Item("Demo")
            If .Sidebar Is Nothing Then Set .Sidebar = pic(2).Picture Else Set .Sidebar = Nothing
        End With
    End If
    If miLastKey <> KeyCode Then moPopupMenus.AcceleratorPress KeyCode, Shift
    miLastKey = KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    miLastKey = 0
End Sub

Private Sub Form_Load()
   Set moCommon = New cCommonDialog
   Set moPopupMenus = New cPopupMenus
   moPopupMenus.hWndOwner = hWnd
   ' Make sure the ImageList has icons before setting
   ' this if it is a MS ImageList:
   moPopupMenus.ImageList = ilsIcons16
   
   ' Create some menus and store them:
   createMenus
   
   chkVisual_Click 7
   chkVisual_Click 8
   
End Sub
Private Sub createMenus()
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim iIndex As Long
Dim lIcon As Long
Dim sKey As String
Dim sCap As String
   
Dim loItem As cPopupMenuItem
   ' Create the demo menu:
   With moPopupMenus
      .Clear
      With .Add("Demo").Items
        For i = 1 To 10
           If (i = 6) Or (i = 7) Then sKey = "CHECK" Else sKey = ""
           Set loItem = .Item(.Add("Test long long long item test" & i, , sKey, IIf(sKey = "CHECK", -1&, i + 3)))
           
           If (i = 4) Or (i = 5) Then
              ' separators:
              .Add IIf(i = 4, "Title Item 1", "Title Item 2"), , , , mnuSeparator
           ElseIf (i = 8) Or (i = 9) Then
              ' Make items invisible:
             ' loItem.Visible = False
              '.ItemKey(iIndex) = "INVISIBLE" & i - 7
           ElseIf (i = 10) Then
                ' Add some submenus:
                For j = 1 To 30
                   sCap = "SubMenu Test" & j
                    If ((j - 1) Mod 10) = 0 And j > 1 Then
                       Set loItem = loItem.SubItems.Item(1)
                    End If
                    loItem.SubItems.Add sCap, , , j + 10
                    loItem.SubItems.Item(loItem.SubItems.Count).Infrequent = ((j Mod 10) Mod 3 = 0)
                Next j
           End If
           
        Next i

        .Item(6).Infrequent = True
        '.Item(7).Infrequent = True
        '.Item(5).Infrequent = True
        .Item(2).Infrequent = True
        .Item(3).Infrequent = True
        .Item(.Count).Caption = "&This is now really really really really really long"
        
        .Item(3).ShortCutShiftKey = vbKeyA
        .Item(3).ShortCutShiftMask = vbShiftMask Or vbAltMask
        End With
        '.Item("Demo").Items.ShowCheckAndIcon = True
        Set .Item("Demo").Sidebar = pic(2).Picture
        
        Dim loMenu As cPopupMenu
        For Each loMenu In moPopupMenus
            pDisplayItems loMenu.Items
        Next
      End With
      
      
      ' create a customise menu
'      .Clear
'      For i = 1 To 5
'         k = .AddItem("Test Item " & i, , , , i - 1)
'      Next i
'      .AddItem "-"
'      j = .AddItem("&Add or Remove Buttons")
'      For i = 1 To 20
'         k = .AddItem("Test Item " & i, , , j, i - 1, (i <= 5), , "CHECK")
'         .ShowCheckAndIcon(k) = True
'         .RedisplayMenuOnClick(k) = True
'      Next i
'      k = .AddItem("-", , , j)
'      k = .AddItem("&Reset Toolbar...", , , j)
'      k = .AddItem("&Customise...", , , j)
'      .Store "Customise"
'
'
'      ' Create the edit menu:
'      .Clear
'      .AddItem "Cu&t" & vbTab & "Ctrl+X", , , , ilsIcons16.ListImages("CUT").Index - 1, , , "Cut"
'      .AddItem "&Copy" & vbTab & "Ctrl+C", , , , ilsIcons16.ListImages("COPY").Index - 1, , , "Copy"
'      .AddItem "&Paste" & vbTab & "Ctrl+V", , , , ilsIcons16.ListImages("PASTE").Index - 1, , False, "Paste"
'      .Store "Edit"
'
'      ' Create the vbAccelerator menu:
'      .Clear
'      .AddItem "-vbAccelerator"
'      lIcon = ilsIcons16.ListImages("vbAccelerator").Index - 1
'      .AddItem "&vbAccelerator on the Web..." & vbTab & "F1", , , , lIcon, , , "Web"
'      .Default(2) = True
'      lIcon = ilsIcons16.ListImages("Web").Index - 1
'      .AddItem "Add vbAccelerator Active &Channel...", , mcWEBSITE, , lIcon, , , "Channel"
'      .AddItem "-Other sites"
'      i = .AddItem("VB Sites", , , , lIcon)
'      .AddItem "-VB Sites", , , i
'      .AddItem "VBWire", , mcWEBSITE, i, lIcon, , , "http://vbwire.com/"
'      .AddItem "VBNet", , mcWEBSITE, i, lIcon, , , "http://www.mvps.org/mvps"
'      .AddItem "CCRP", , mcWEBSITE, i, lIcon, , , "http://www.mvps.org/ccrp"
'      .AddItem "DevX", , mcWEBSITE, i, lIcon, , , "http://www.devx.com/"
'      i = .AddItem("Technology", , , , lIcon)
'      .AddItem "-Games", , , i
'      .AddItem "Dave's Classics", , mcWEBSITE, i, lIcon, , , "http://www.davesclassics.com/"
'      .AddItem "Future Gamer", , mcWEBSITE, i, lIcon, , , "http://www.futuregamer.com/"
'      .AddItem "-Web Site Building", , , i
'      .AddItem "Builder.com", , mcWEBSITE, i, lIcon, , , "http://www.builder.com/"
'      .AddItem "The Web Design Resource", , mcWEBSITE, i, lIcon, , , "http://www.pageresource.com/"
'      .AddItem "Web Review", , mcWEBSITE, i, lIcon, , , "http://www.webreview.com/"
'      .AddItem "-Downloads", , , i
'      .AddItem "CNet", , mcWEBSITE, i, lIcon, , , "http://www.cnet.com/"
'      .AddItem "WinFiles.com", , mcWEBSITE, i, lIcon, , , "http://www.winfiles.com/"
'      i = .AddItem("Searching and Other", , , , lIcon)
'      j = .AddItem("Pick'n'Mix", , , i)
'      .Header(j) = True
'      .AddItem "The SCHWA Corporation", , mcWEBSITE, i, lIcon, , , "http://www.theschwacorporation.com/"
'      .AddItem "Art Cars", , mcWEBSITE, i, lIcon, , , "http://www.artcars.com/"
'      .AddItem "The Onion", , mcWEBSITE, i, lIcon, , , "http://www.theonion.com/"
'      .AddItem "Virtues of a Programmer", i, mcWEBSITE, i, lIcon, , , "http://www.hhhh.org/wiml/virtues.html"
'      .AddItem "-Search", , , i
'      .AddItem "Google", , mcWEBSITE, i, lIcon, , , "http://www.google.com/"
'      .AddItem "DogPile", , mcWEBSITE, i, lIcon, , , "http://www.dogpile.com/"
'      .Store "vbAccelerator"
'
'      .Clear
'      .AddItem "First Check", , , , , True, , "Check1"
'      .AddItem "Second Check", , , , , , , "Check2"
'      .AddItem "Third Check", , , , , , , "Check3"
'      .AddItem "-"
'      i = .AddItem("First Option", , , , , , , "Option1")
'      .RadioCheck(i) = True
'      'Debug.Print .RadioCheck(i)
'      .AddItem "Second Option", , , , , , , "Option2"
'      .AddItem "Third Option", , , , , , , "Option3"
'      .AddItem "Fourth Option", , , , , , , "Option4"
'      .AddItem "-"
'      .AddItem "&vbAccelerator on the Web...", , , , lIcon, , , "Web"
'      .Store "CheckTest"
'
'      .Clear
'      .AddItem "&Back" & vbTab & "Alt+Left Arrow", , , , , , , "mnuAccel(0)"
'      .AddItem "&Next" & vbTab & "Alt+Right Arrow", , , , , , , "mnuAccel(1)"
'      .AddItem "-"
'      j = .AddItem("&Home Page" & vbTab & "Alt+Home", , , , , , , "mnuAccel(3)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("&Search the Web", , , , , , , "mnuAccel(4)")
'      .ItemInfrequentlyUsed(j) = True
'      .AddItem "-"
'      j = .AddItem("&Mail", , , , , , , "mnuAccel(6)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("&News", , , , , , , "mnuAccel(7)")
'      .ItemInfrequentlyUsed(j) = True
'      .AddItem "My &Computer", , , , , , , "mnuAccel(8)"
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("A&ddress Book", , , , , , , "mnuAccel(9)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("Ca&lendar", , , , , , , "mnuAccel(10)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("&Internet Call", , , , , , , "mnuAccel(11)")
'      .ItemInfrequentlyUsed(j) = True
'      i = .AddItem("Other &Links", , , , , , , "mnuAccel(12)")
'
'      lIcon = ilsIcons16.ListImages("Web").Index - 1
'      j = .AddItem("Planet-Mu Records", "http://www.planet-mu.com/", , i, lIcon, , , "mnuLink(0)")
'      j = .AddItem("Speedranch/Jansky Noise", "http://www.forcefield.org/", , i, lIcon, , , "mnuLink(1)")
'      j = .AddItem("LFO Discography", "http://www.sci.fi/~phinnweb/links/artists/lfo/", , i, lIcon, , , "mnuLink(2)")
'      j = .AddItem("All Tommorrow's Parties", "http://www.alltomorrowsparties.co.uk/", , i, lIcon, , , "mnuLink(3)")
'      j = .AddItem("XLR8R Magazine", "http://www.xlr8r.com/", , i, lIcon, , , "mnuLink(4)")
'      j = .AddItem("Superbad", "http://www.superbad.com/", , i, lIcon, , , "mnuLink(5)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("Stereolab", "http://www.stereolab.co.uk/", , i, lIcon, , , "mnuLink(6)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("Pixies Discography", "http://www.evo.org/html/group/pixies.html", , i, lIcon, , , "mnuLink(7)")
'      .ItemInfrequentlyUsed(j) = True
'      j = .AddItem("IconMenu Links", , , i, lIcon, , , "mnuLink(8)")
'
'      For l = 1 To 10
'         k = .AddItem("Test Menu " & l, , , j, lIcon)
'         .ItemInfrequentlyUsed(k) = (l <> 2)
'      Next l
'
'      .Store "AccelTest"
'   End With
   
End Sub

Private Sub pDisplayItems(ByVal oItems As cPopupMenuItems, Optional ByRef psPrefix As String = vbNullString)
    Exit Sub
    Dim loItem As cPopupMenuItem
    For Each loItem In oItems
        Debug.Print psPrefix & loItem.Caption
        pDisplayItems loItem.SubItems, psPrefix & vbTab
    Next
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    With pic(0)
        .Move .Left, .Top, ScaleWidth - .Left - .Left, ScaleHeight - .Top - .Left
    End With
End Sub

Private Sub lstStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   If (Button And vbRightButton) = vbRightButton Then
'      Dim iIndex As Long
'         moPopupMenus.Restore "Edit"
'         moPopupMenus.Enabled(moPopupMenus.IndexForKey("Paste")) = Clipboard.GetFormat(vbCFText)
'         iIndex = moPopupMenus.ShowPopupMenu( _
'            X + lstStatus.Left, Y + lstStatus.Top + picFrame.Top)
'         If (iIndex > 0) Then
'            Status "Clicked " & iIndex
'         End If
'   End If
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        With lstStatus
            .Move .Left, .Top, pic(0).ScaleWidth, pic(0).ScaleHeight - .Top
        End With
    End If
End Sub

