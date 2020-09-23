VERSION 5.00
Begin VB.Form fNew 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3525
   ClientLeft      =   2085
   ClientTop       =   2910
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   0
      Left            =   2160
      ScaleHeight     =   3135
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      Begin VB.ListBox lst 
         Height          =   2985
         ItemData        =   "fNew.frx":0000
         Left            =   0
         List            =   "fNew.frx":003F
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   3015
      Index           =   1
      Left            =   2160
      ScaleHeight     =   3015
      ScaleWidth      =   2295
      TabIndex        =   4
      Top             =   240
      Width           =   2295
      Begin VB.ComboBox cmb 
         Height          =   315
         ItemData        =   "fNew.frx":00F9
         Left            =   0
         List            =   "fNew.frx":0118
         TabIndex        =   5
         Text            =   "cmb"
         Top             =   0
         Width           =   1575
      End
   End
End
Attribute VB_Name = "fNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancel As Boolean

Private Sub cmd_Click(Index As Integer)
    mbCancel = (Index = 1)
    Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mbCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbCancel = True
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    cmd(0).Enabled = (lst.SelCount > 0)
End Sub

Public Function GetNewClassStyle(ByRef iStyle As eClassStyle) As Boolean
    lstClear lst
    cmd(0).Enabled = False
    picShowOne pic, 0
    Caption = "Add Class"
    Show vbModal
    If Not mbCancel Then
        GetNewClassStyle = True
        iStyle = lstOrAllItemData(lst)
    End If
End Function

Public Function GetNewWindowClass(ByRef sClass As String) As Boolean
    cmd(0).Enabled = True
    picShowOne pic, 1
    Caption = "Add Window"
    Show vbModal
    If Not mbCancel Then
        GetNewWindowClass = True
        sClass = cmb.Text
    End If
End Function

