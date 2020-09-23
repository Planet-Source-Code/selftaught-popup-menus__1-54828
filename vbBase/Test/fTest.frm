VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Testing..."
   ClientHeight    =   7995
   ClientLeft      =   1275
   ClientTop       =   1440
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox pic 
      Align           =   1  'Align Top
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   3480
      Index           =   0
      Left            =   0
      ScaleHeight     =   228
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.OptionButton optDemo 
         Caption         =   "Other Wnds"
         Height          =   375
         Index           =   5
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton optDemo 
         Caption         =   "Class Wnds"
         Height          =   375
         Index           =   4
         Left            =   6216
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton optDemo 
         Caption         =   "Wnd Classes"
         Height          =   375
         Index           =   3
         Left            =   4752
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton optDemo 
         Caption         =   "Timers"
         Height          =   375
         Index           =   2
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optDemo 
         Caption         =   "Hooks"
         Height          =   375
         Index           =   1
         Left            =   1584
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optDemo 
         Caption         =   "Subclasses"
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   2760
         Index           =   1
         Left            =   0
         ScaleHeight     =   2760
         ScaleWidth      =   9195
         TabIndex        =   3
         Top             =   600
         Width           =   9195
         Begin VB.CommandButton cmd 
            Caption         =   "Enumerate"
            Height          =   375
            Index           =   10
            Left            =   1920
            TabIndex        =   62
            Top             =   0
            Width           =   1575
         End
         Begin VB.PictureBox pic 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   465
            Index           =   9
            Left            =   2040
            ScaleHeight     =   465
            ScaleWidth      =   1650
            TabIndex        =   12
            Top             =   1245
            Width           =   1650
            Begin VB.OptionButton opt 
               Caption         =   "All messages"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   1335
            End
            Begin VB.OptionButton opt 
               Caption         =   "Selected messages"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   0
               TabIndex        =   13
               Top             =   270
               Width           =   1680
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "Before original WndProc"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1770
            TabIndex        =   11
            Top             =   990
            Width           =   2025
         End
         Begin VB.PictureBox pic 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   465
            Index           =   8
            Left            =   2040
            ScaleHeight     =   465
            ScaleWidth      =   1665
            TabIndex        =   8
            Top             =   2205
            Width           =   1665
            Begin VB.OptionButton opt 
               Caption         =   "Selected messages"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   0
               TabIndex        =   10
               Top             =   255
               Width           =   1695
            End
            Begin VB.OptionButton opt 
               Caption         =   "All messages"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   1455
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "After original WndProc"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   1815
            TabIndex        =   7
            Top             =   1920
            Width           =   1950
         End
         Begin VB.ListBox lst 
            Height          =   1410
            Index           =   0
            ItemData        =   "fTest.frx":0000
            Left            =   120
            List            =   "fTest.frx":0002
            TabIndex        =   6
            Top             =   990
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Add"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Remove"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1575
         End
         Begin MSComctlLib.ListView lvwMessages 
            Height          =   2400
            Index           =   0
            Left            =   3825
            TabIndex        =   15
            Top             =   270
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4233
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4075
            EndProperty
         End
         Begin MSComctlLib.ListView lvwMessages 
            Height          =   2400
            Index           =   1
            Left            =   6465
            TabIndex        =   16
            Top             =   255
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4233
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4075
            EndProperty
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Before                   After"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   18
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   4380
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Dbl-Click for info"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   225
            TabIndex        =   17
            Top             =   2490
            UseMnemonic     =   0   'False
            Width           =   1380
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   2730
         Index           =   5
         Left            =   0
         ScaleHeight     =   2730
         ScaleWidth      =   9195
         TabIndex        =   43
         Top             =   600
         Width           =   9195
         Begin VB.CommandButton cmd 
            Caption         =   "Enumerate"
            Height          =   495
            Index           =   14
            Left            =   1920
            TabIndex        =   66
            Top             =   0
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "Default Messages"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   4200
            TabIndex        =   60
            Top             =   0
            Width           =   2025
         End
         Begin VB.CheckBox chk 
            Caption         =   "All Messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   6480
            TabIndex        =   57
            Top             =   0
            Width           =   2025
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Add Window"
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   46
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Del Window"
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   45
            Top             =   600
            Width           =   1575
         End
         Begin VB.ListBox lst 
            Height          =   1185
            Index           =   5
            ItemData        =   "fTest.frx":0004
            Left            =   240
            List            =   "fTest.frx":0006
            TabIndex        =   44
            Top             =   1440
            Width           =   1575
         End
         Begin MSComctlLib.ListView lvwMessages 
            Height          =   2400
            Index           =   3
            Left            =   6480
            TabIndex        =   47
            Top             =   240
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4233
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4075
            EndProperty
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   $"fTest.frx":0008
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2055
            Index           =   15
            Left            =   4200
            TabIndex        =   61
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   1980
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Windows:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   49
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1740
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Double-click on a window in the list to manipulate it."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   9
            Left            =   2040
            TabIndex        =   48
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   1620
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   2730
         Index           =   4
         Left            =   0
         ScaleHeight     =   2730
         ScaleWidth      =   9195
         TabIndex        =   26
         Top             =   600
         Width           =   9195
         Begin VB.CommandButton cmd 
            Caption         =   "Enumerate"
            Height          =   495
            Index           =   13
            Left            =   1920
            TabIndex        =   65
            Top             =   0
            Width           =   1575
         End
         Begin VB.CheckBox chk 
            Caption         =   "All Messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   6360
            TabIndex        =   56
            Top             =   0
            Width           =   2025
         End
         Begin VB.ListBox lst 
            Height          =   1185
            Index           =   4
            ItemData        =   "fTest.frx":009D
            Left            =   240
            List            =   "fTest.frx":009F
            TabIndex        =   37
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Del Class"
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Add Class"
            Height          =   495
            Index           =   4
            Left            =   240
            TabIndex        =   35
            Top             =   0
            Width           =   1575
         End
         Begin MSComctlLib.ListView lvwMessages 
            Height          =   2400
            Index           =   2
            Left            =   6360
            TabIndex        =   42
            Top             =   240
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   4233
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   4075
            EndProperty
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Classes:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            UseMnemonic     =   0   'False
            Width           =   1740
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Current Class:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2655
            Index           =   7
            Left            =   3840
            TabIndex        =   39
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   2340
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   2730
         Index           =   3
         Left            =   0
         ScaleHeight     =   2730
         ScaleWidth      =   9195
         TabIndex        =   27
         Top             =   600
         Width           =   9195
         Begin VB.CommandButton cmd 
            Caption         =   "Enumerate"
            Height          =   375
            Index           =   11
            Left            =   3000
            TabIndex        =   63
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Remove"
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Add"
            Height          =   375
            Index           =   2
            Left            =   360
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.ListBox lst 
            Height          =   1590
            Index           =   3
            ItemData        =   "fTest.frx":00A1
            Left            =   360
            List            =   "fTest.frx":00A3
            Style           =   1  'Checkbox
            TabIndex        =   29
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   6
            Left            =   2880
            TabIndex        =   28
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   5340
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   2730
         Index           =   2
         Left            =   0
         ScaleHeight     =   2730
         ScaleWidth      =   9195
         TabIndex        =   23
         Top             =   600
         Width           =   9195
         Begin VB.ListBox lst 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Index           =   2
            ItemData        =   "fTest.frx":00A5
            Left            =   3360
            List            =   "fTest.frx":00CC
            Style           =   1  'Checkbox
            TabIndex        =   32
            Top             =   480
            Width           =   3015
         End
         Begin VB.ListBox lst 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Index           =   1
            ItemData        =   "fTest.frx":0183
            Left            =   120
            List            =   "fTest.frx":01A9
            Style           =   1  'Checkbox
            TabIndex        =   24
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   $"fTest.frx":0235
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2655
            Index           =   5
            Left            =   6480
            TabIndex        =   34
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   2580
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Hooks for all threads:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   33
            Top             =   120
            UseMnemonic     =   0   'False
            Width           =   1860
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Hooks for this thread:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   120
            UseMnemonic     =   0   'False
            Width           =   1860
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   2730
         Index           =   6
         Left            =   0
         ScaleHeight     =   2730
         ScaleWidth      =   9195
         TabIndex        =   50
         Top             =   600
         Width           =   9195
         Begin VB.CommandButton cmd 
            Caption         =   "Enumerate"
            Height          =   375
            Index           =   12
            Left            =   2160
            TabIndex        =   64
            Top             =   0
            Width           =   1575
         End
         Begin VB.ListBox lst 
            Height          =   1185
            Index           =   6
            ItemData        =   "fTest.frx":0356
            Left            =   240
            List            =   "fTest.frx":0358
            TabIndex        =   53
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Del Window"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   52
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Add Window"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   51
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   $"fTest.frx":035A
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2655
            Index           =   14
            Left            =   4320
            TabIndex        =   59
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   4305
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Windows:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   54
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1740
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            Caption         =   "Double-click on a window."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   240
            TabIndex        =   55
            Top             =   2400
            UseMnemonic     =   0   'False
            Width           =   2340
         End
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   10
      Left            =   0
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   58
      Top             =   3720
      Width           =   9135
   End
   Begin VB.PictureBox pic 
      Align           =   1  'Align Top
      HasDC           =   0   'False
      Height          =   300
      Index           =   7
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   9135
      TabIndex        =   1
      Top             =   3480
      Width           =   9195
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   13
         Left            =   45
         TabIndex        =   2
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   8940
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuItm 
         Caption         =   "New Instance"
         Index           =   0
      End
      Begin VB.Menu mnuItm 
         Caption         =   "E&xit"
         Index           =   1
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements vbBase.iSubclass   'Tell VB that we need to implement the methods and properties defined
Implements vbBase.iHook       'in these four classes.  This is done so that the vbBase component
Implements vbBase.iWindow     'can make early bound calls to these interfaces without having to
Implements vbBase.iTimer      'know or care what is the defining class; whether it's your form as
                               'in this case, a regular class object, or a usercontrol.

Private Enum eTest
    tstSubclass                'We'll need to tell each demo mode apart
    tstHooks
    tstTimers
    tstClasses
    tstClassWindows
    tstWindows
End Enum

Private Enum eChk
    chkSubclassBefore
    chkSubclassAfter
    chkAllClass
    chkAllClassWindow
    chkClassWindowDefault
End Enum

Private Enum eCmd
    cmdAddSubclass
    cmdDelSubclass
    cmdAddTimer
    cmdDelTimer
    cmdAddClass
    cmdDelClass
    cmdAddClassWindow
    cmdDelClassWindow
    cmdAddWindow
    cmdDelWindow
    cmdEnumSubclasses
    cmdEnumTimers
    cmdEnumWindows
    cmdEnumClasses
    cmdEnumClassWindows
End Enum

Private Enum eLbl
    lblTimer = 6
    lblCurrentClass = 7
    lblHeader = 13
End Enum

Private Enum eLst
    lstSubclass
    lstHooksThread
    lstHooksGlobal
    lstTimers
    lstClasses
    lstClassWindows
    lstWindows
End Enum

Private Enum eLvw
    lvwSubclassBefore
    lvwSubclassAfter
    lvwClass
    lvwClassWindow
End Enum

Private Enum eOpt
    optAllBefore
    optSelBefore
    optAllAfter
    optSelAfter
End Enum

Private Enum ePic
    picHeader1 = 0
    picClassWindows = 5
    picWindows = 6
    picHeader2 = 7
    picDisplay = 10
End Enum

Private miCurrentTest       As eTest
Private miTextHeight        As Long                    'Height of a text line
Private mbLoaded            As Boolean

Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, ByVal lprcScroll As Long, ByVal lprcClip As Long, ByVal hrgnUpdate As Long, ByVal lprcUpdate As Long) As Long

'This flag is set when changing item checked states in the subclass listviews or checkboxes.
'It prevents adding or removing messages when the ItemCheck events are fired.
Private mbFreezeMessages    As Boolean

'############################
'##    Event Procedures    ##
'############################

Private Sub chk_Click(Index As Integer)
On Error GoTo handler
    Dim lbVal As Boolean
    Dim lhWnd As Long
    Dim lsClass As String
    
    lbVal = (chk(Index) = vbChecked)
    Select Case CLng(Index)
    Case chkSubclassBefore
        pDeselect lvwMessages(lvwSubclassBefore)
        lvwMessages(lvwSubclassBefore).Enabled = False
        
        opt(optAllBefore).Enabled = lbVal
        opt(optSelBefore).Enabled = lbVal
        opt(optAllBefore).Value = False
        opt(optSelBefore).Value = False
        
        If lst(lstSubclass).ListIndex > -1& Then
            Subclasses(Me).Item(HexVal(lst(lstSubclass).Text)).DelMsg ALL_MESSAGES, MSG_BEFORE
        End If
        
    Case chkSubclassAfter
    
        pDeselect lvwMessages(lvwSubclassAfter)
        lvwMessages(lvwSubclassAfter).Enabled = False
        
        opt(optAllAfter).Enabled = lbVal
        opt(optSelAfter).Enabled = lbVal
        opt(optAllAfter).Value = False
        opt(optSelAfter).Value = False
    
        If lst(lstSubclass).ListIndex > -1& Then
            Subclasses(Me).Item(HexVal(lst(lstSubclass).Text)).DelMsg ALL_MESSAGES, MSG_AFTER
        End If
    
    Case chkAllClass
        lsClass = lst(lstClasses).Text
        If Not mbFreezeMessages Then
            If Len(lsClass) Then
                With ApiWindowClasses(lsClass)
                    If lbVal Then .AddDefMsg ALL_MESSAGES Else .DelDefMsg ALL_MESSAGES
                End With
                mbFreezeMessages = True
                pDeselect lvwMessages(lvwClass)
                mbFreezeMessages = False
                lvwMessages(lvwClass).Enabled = Not lbVal
            End If
        End If
    Case chkAllClassWindow
        If lst(lstClassWindows).ListIndex > -1& Then
            lhWnd = HexVal(lst(lstClassWindows).Text)
            lsClass = xApiWindow(lhWnd).ClassName
            With ApiWindowClasses(lsClass).OwnedWindows(Me).Item(lhWnd)
                If lbVal _
                    Then .AddMsg ALL_MESSAGES _
                    Else .DelMsg ALL_MESSAGES
            End With
            pShowMessages tstClassWindows
        End If
    Case chkClassWindowDefault
        If lst(lstClassWindows).ListIndex > -1& Then
            lhWnd = HexVal(lst(lstClassWindows).Text)
            lsClass = xApiWindow(lhWnd).ClassName
            ApiWindowClasses(lsClass).OwnedWindows(Me).Item(lhWnd).DefaultMessages = lbVal
            pShowMessages tstClassWindows
        End If
    End Select

    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error GoTo handler
    Dim lhWnd As Long
    Dim liId As Long
    Dim liInterval As Long
    Dim lsMsg As String
    Dim lsClass As String
    Dim loCls As cApiWindowClass
    
    
    Select Case CLng(Index)
    Case cmdAddSubclass
        fFindWindow.GetWindow Me
    Case cmdDelSubclass
        With lst(lstSubclass)
            lhWnd = HexVal(.Text)
            If lhWnd Then
                If MsgBox("Delete this subclass?   " & FmtHex(lhWnd), vbYesNo + vbQuestion) = vbYes Then
                    Subclasses(Me).Remove lhWnd
                    .RemoveItem .ListIndex
                    lst_Click lstSubclass
                End If
            End If
        End With
    Case cmdEnumSubclasses
        Dim loSub As cSubclass
        For Each loSub In Subclasses(Me)
            lsMsg = lsMsg & FmtHex(loSub.hwnd) & vbTab & xApiWindow(loSub.hwnd).ClassName & vbNewLine
        Next
        MsgBox lsMsg
    Case cmdAddTimer
        If pGetNumericInput("Enter the new ID:", "0", liId) Then
            If pGetNumericInput("Enter the interval in milliseconds:", "1000", liInterval) Then
                Timers(Me).Add liId, liInterval
                With lst(lstTimers)
                    .AddItem FmtHex(liId) & vbTab & liInterval
                    .ListIndex = .NewIndex
                End With
            End If
        End If
    Case cmdDelTimer
        With lst(lstTimers)
            If .ListIndex > -1& Then
                liId = HexVal(Mid$(.Text, 1, InStr(1, .Text, vbTab) - 1))
                .RemoveItem .ListIndex
                Timers(Me).Remove liId
            End If
        End With
    Case cmdEnumTimers
        Dim loTimer As cTimer
        For Each loTimer In Timers(Me)
            lsMsg = lsMsg & FmtHex(loTimer.ID) & vbTab & loTimer.Interval & vbNewLine
        Next
        MsgBox lsMsg
    Case cmdAddClass
        Dim liStyle As eClassStyle
        
        lsClass = InputBox("Enter a name for the window class:")
        
        If Len(lsClass) Then
            If fNew.GetNewClassStyle(liStyle) Then
                Randomize Timer
                ApiWindowClasses.Add lsClass, QBColor(CInt(Rnd * 15)) + 1, liStyle
                BroadcastClass True, lsClass, True
            End If
        End If
        
    Case cmdDelClass
        If lst(lstClasses).ListIndex > -1& Then
            lsClass = lst(lstClasses).Text
            ApiWindowClasses.Remove lsClass
            BroadcastClass False, lsClass, True
        End If
    Case cmdEnumClasses
        For Each loCls In ApiWindowClasses
            lsMsg = lsMsg & loCls.Name & vbNewLine
        Next
        MsgBox lsMsg
    Case cmdAddClassWindow
        If ApiWindowClasses.Count Then
            lsClass = InputBox("Enter the name of the class to create the window from:")
            If Len(lsClass) Then
                With ApiWindowClasses(lsClass).OwnedWindows(Me).Add(WS_OVERLAPPEDWINDOW Or WS_CAPTION Or WS_VISIBLE, WS_EX_APPWINDOW, 50, 50, 300, 300, "New Window From " & lsClass)
                    lst(lstClassWindows).AddItem "0x" & FmtHex(.hwnd)
                    lst(lstClassWindows).ListIndex = lst(lstClassWindows).NewIndex
                End With
            End If
        End If
    Case cmdDelClassWindow
        If lst(lstClassWindows).ListIndex > -1& Then
            lhWnd = HexVal(lst(lstClassWindows).Text)
            lsClass = xApiWindow(lhWnd).ClassName
            ApiWindowClasses(lsClass).OwnedWindows(Me).Remove lhWnd
            lst(lstClassWindows).RemoveItem lst(lstClassWindows).ListIndex
        End If
    Case cmdEnumClassWindows
        Dim loClsWin As cApiClassWindow
        For Each loCls In ApiWindowClasses
            For Each loClsWin In loCls.OwnedWindows(Me)
                lsMsg = lsMsg & FmtHex(loClsWin.hwnd) & vbTab & loCls.Name & vbNewLine
            Next
        Next
        MsgBox lsMsg
    Case cmdAddWindow
        If fNew.GetNewWindowClass(lsClass) Then
            lst(lstWindows).AddItem FmtHex(ApiWindows(Me).Add(lsClass, AS_WINDOWCLASS, WS_VISIBLE Or WS_CHILD, WS_EX_NOPARENTNOTIFY, 160, 60, 80, 80, , pic(picWindows).hwnd).hwnd)
        End If
    Case cmdDelWindow
        With lst(lstWindows)
            If .ListIndex > -1& Then
                lhWnd = HexVal(.Text)
                .RemoveItem .ListIndex
                ApiWindows(Me).Remove lhWnd
            End If
        End With
    Case cmdEnumWindows
        Dim loWin As cApiWindow
        For Each loWin In ApiWindows(Me)
            lsMsg = lsMsg & FmtHex(loWin.hwnd) & vbNewLine
        Next
        MsgBox lsMsg
    End Select
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub Form_Load()
  giTestForms = giTestForms + 1&
    mbLoaded = True
  Dim i As eMsg
  Dim s As String


    Dim loLV As ListView

    miTextHeight = pic(picDisplay).TextHeight("M")

  'Adjust the height of the window... like the IntegralHeight property in a listbox
  'Height = Height - (((Me.ScaleHeight Mod nTxtHeight) - 2) * Screen.TwipsPerPixelY)

  For i = 0 To &H400
    s = GetMsgName(i)
    If Asc(s) <> vbKey0 Then
      For Each loLV In lvwMessages
        loLV.ListItems.Add , "k" & i, s
      Next
    End If
  Next i

  For Each loLV In lvwMessages
    loLV.Sorted = True
  Next

  Dim loClass As cApiWindowClass
  For Each loClass In ApiWindowClasses
    lst(lstClasses).AddItem loClass.Name
  Next

  optDemo(0).Value = True
  'FadeIn hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mbLoaded = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    Dim liOffset As Long
    liOffset = pic(picHeader1).Height + pic(picHeader2).Height
    pic(picDisplay).Move 0, liOffset, ScaleWidth, ScaleHeight - liOffset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mbLoaded = False
    Subclasses(Me).Clear
    Hooks(Me).Clear
    ApiWindows(Me).Clear
    Timers(Me).Clear
    
    giTestForms = giTestForms - 1&
    
    Dim i As Long
    With ApiWindowClasses
        For i = 0 To lst(lstClasses).ListCount - 1&
            .Item(lst(lstClasses).List(i)).OwnedWindows(Me).Clear
            If giTestForms = 0& Then .Remove lst(lstClasses).List(i)
        Next
    End With
    
    If giTestForms = 0& Then
        Unload fFindWindow
        Unload fNew
        Unload fWindow
    End If

    'FadeOut hwnd
End Sub


Private Sub lst_Click(Index As Integer)
On Error GoTo handler
    If Index = lstSubclass Then
        pShowMessages tstSubclass
    ElseIf Index = lstClasses Then
        pShowMessages tstClasses
    ElseIf Index = lstClassWindows Then
        pShowMessages tstClassWindows
    ElseIf Index = lstTimers Then
        
        If lst(Index).ListIndex > -1& Then
            With Timers(Me).Item(HexVal(Mid$(lst(Index).Text, 1, InStr(1, lst(Index).Text, vbTab) - 1)))
                lbl(lblTimer).Caption = "ID: " & .ID & vbNewLine & "Interval: " & .Interval & vbNewLine & "Active: " & .Active
            End With
        Else
            lbl(lblTimer).Caption = vbNullString
        End If
        
        lbl(lblTimer).Caption = lbl(lblTimer).Caption & vbNewLine & vbNewLine & "Right-click on the list to change the interval of the timer."
        
    End If
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub lst_DblClick(Index As Integer)
On Error GoTo handler
    
    If Index <> lstTimers Then
        Dim lhWnd As Long
        lhWnd = HexVal(lst(Index).Text)
        If lhWnd Then
            fWindow.WindowInfo.hWndDisplay = lhWnd
            fWindow.Show
        End If
    End If
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
On Error GoTo handler
    
    If Index = lstTimers Then
        If lst(Index).ListIndex > -1& Then
            With Timers(Me).Item(HexVal(Mid$(lst(Index).Text, 1, InStr(1, lst(Index).Text, vbTab) - 1)))
                If lst(Index).Selected(Item) Then .Start Else .StopTimer
            End With
        End If
    ElseIf Index = lstHooksGlobal Or Index = lstHooksThread Then
        
        Dim liCode      As eHookCode: liCode = lst(Index).ItemData(Item)
        Dim lbThread    As Boolean: lbThread = (Index = lstHooksThread)
        
        If lst(Index).Selected(Item) _
            Then Hooks(Me).Add liCode, lbThread _
            Else Hooks(Me).Remove liCode, lbThread
            
    End If
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub lst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo handler
    
    If Index = lstTimers And Button = 2 Then
        If lst(Index).ListIndex > -1& Then
            Dim liInt As Long
            With Timers(Me).Item(Mid$(lst(Index).Text, 1, InStr(1, lst(Index).Text, vbTab) - 1))
                If pGetNumericInput("Enter the new interval for this timer: ", .Interval, liInt) Then .Interval = liInt
                lst(Index).List(lst(Index).ListIndex) = FmtHex(.ID) & vbTab & .Interval
            End With
        End If
    End If
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub lvwMessages_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
On Error GoTo handler
    
    If Not mbFreezeMessages Then
        Dim lhWnd As Long
        Dim liMsg As eMsg
        Dim liWhen As eMsgWhen
        Dim lbVal As Boolean
        Dim lsClass As String
        
        lbVal = Item.Checked
        liMsg = Val(Right$(Item.Key, Len(Item.Key) - 1))
        
        Select Case CLng(Index)
        Case lvwSubclassBefore, lvwSubclassAfter
            liWhen = IIf(CLng(Index) = lvwSubclassBefore, MSG_BEFORE, MSG_AFTER)
            With Subclasses(Me).Item(HexVal(lst(lstSubclass).Text))
                If lbVal _
                    Then .AddMsg liMsg, liWhen _
                    Else .DelMsg liMsg, liWhen
            End With
        Case lvwClass
            lsClass = lst(lstClasses).Text
            If Len(lsClass) Then
                With ApiWindowClasses(lsClass)
                    If lbVal _
                        Then .AddDefMsg liMsg _
                        Else .DelDefMsg liMsg
                End With
            End If
        Case lvwClassWindow
        
        End Select
    End If
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub mnuItm_Click(Index As Integer)
    If Index Then
        Unload Me
    Else
        Dim temp As New fTest
        temp.Show
    End If
End Sub

Private Sub opt_Click(Index As Integer)
On Error GoTo handler
    
    pMessageOption Choose(Index + 1, _
                      MSG_BEFORE, MSG_BEFORE, _
                      MSG_AFTER, MSG_AFTER), _
               Index = optAllBefore Or Index = optAllAfter
    
    Exit Sub
handler:
pDisplayErr
End Sub

Private Sub optDemo_Click(Index As Integer)
    miCurrentTest = Index
    lbl(lblHeader).Caption = Choose(miCurrentTest + 1&, _
        "######## When.. lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... ", _
        "######## When.. lReturn. nCode... wParam.. lParam.. Code Name/Hook Type......... ", _
        "######## ID...... Interval Elasped", _
        "######## lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... ", _
        "######## lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... ", _
        "")

    Dim loPic As PictureBox
    For Each loPic In pic
        If loPic.Index > 0& And loPic.Index < 7 _
            Then loPic.Visible = loPic.Index = Index + 1&
    Next
    pic(picDisplay).Cls
End Sub

'############################
'##  Private Procedures    ##
'############################

Private Sub pDisplay(ByRef sString As String)

  If pTextIsBelow Then
    With pic(picDisplay)
        ScrollDC .hdc, 0, -miTextHeight, 0&, 0&, 0&, 0&
        Do Until Not pTextIsBelow
            .CurrentY = .CurrentY - miTextHeight
        Loop
    End With
   End If
   
   pic(picDisplay).Print sString

End Sub

Private Function pTextIsBelow() As Boolean
    pTextIsBelow = (pic(picDisplay).CurrentY + miTextHeight + miTextHeight) >= pic(picDisplay).ScaleHeight
End Function

Private Function pGetNumericInput(sText As String, ByVal sDefault As String, iVal As Long) As Boolean
    Dim lsAnswer As String
    lsAnswer = InputBox(sText, App.Title, sDefault)
    pGetNumericInput = StrPtr(lsAnswer) <> 0&
    iVal = Val(lsAnswer)
End Function

Private Sub pShowMessagesSub(ByRef iArray() As Long, ByVal iCount As Long, ByVal oChk As CheckBox, ByVal oLvw As ListView, Optional ByVal oOptAll As OptionButton, Optional ByVal oOptSel As OptionButton)
    If oOptAll Is Nothing Or oOptSel Is Nothing Then
        If iCount = -1& Then oChk.Value = vbChecked Else oChk.Value = vbUnchecked
    Else
        If iCount > 0& Or iCount = -1& Then oChk.Value = vbChecked Else oChk.Value = vbUnchecked
        oOptAll.Value = (iCount = -1&)
        oOptSel.Value = (iCount > 0&)
    End If
    pSelectItems oLvw, iArray, iCount
End Sub

Private Sub pEnableMessages(ByVal iTest As eTest, ByVal bVal As Boolean)
    Select Case iTest
    Case tstSubclass
        pEnableMessagesSub bVal, chk(chkSubclassAfter), lvwMessages(lvwSubclassAfter), opt(optAllAfter), opt(optSelAfter)
        pEnableMessagesSub bVal, chk(chkSubclassBefore), lvwMessages(lvwSubclassBefore), opt(optAllBefore), opt(optSelBefore)
    Case tstClasses
        pEnableMessagesSub bVal, chk(chkAllClass), lvwMessages(lvwClass)
    Case tstClassWindows
        pEnableMessagesSub bVal, chk(chkAllClassWindow), lvwMessages(lvwClassWindow)
    End Select
End Sub

Private Sub pEnableMessagesSub(ByVal bVal As Boolean, ByVal oChk As CheckBox, ByVal oLvw As ListView, Optional ByVal oOptAll As OptionButton, Optional ByVal oOptSel As OptionButton)
    Dim lbVal As Boolean
    
    oChk.Enabled = bVal
    If Not bVal Then oChk.Value = vbUnchecked
    If oOptAll Is Nothing Or oOptSel Is Nothing Then
        lbVal = (oChk.Value = vbUnchecked)
    Else
        lbVal = (oChk.Value = vbChecked)
        'If lbVal Then oOptAll.Value = False: oOptSel.Value = False
        oOptAll.Enabled = lbVal: oOptSel.Enabled = lbVal
        lbVal = lbVal And oOptSel.Value
    End If
    oLvw.Enabled = bVal And lbVal
End Sub

Private Sub pRemoveAllMessages(ByVal iWhen As eMsgWhen)
    Dim lv As ListView
    Set lv = lvwMessages(iWhen And Not MSG_BEFORE)
    lv.Enabled = False
    pDeselect lv
    If Not mbFreezeMessages Then Subclasses(Me).Item(HexVal(lst(lstSubclass).Text)).DelMsg ALL_MESSAGES, iWhen
End Sub

Private Sub pDeselect(ByVal lv As ListView)
    Dim itm As MSComctlLib.ListItem
    Dim bWasFroze As Boolean
    
    bWasFroze = mbFreezeMessages
    mbFreezeMessages = True
    
    For Each itm In lv.ListItems
        itm.Checked = False
    Next
    
    mbFreezeMessages = bWasFroze
End Sub

Private Sub pDisplayErr()
    MsgBox "Error #: " & Err.Number & vbNewLine & "Source: " & Err.Source & vbNewLine & vbNewLine & Err.Description
End Sub

Private Sub pSelectItems(ByVal lv As ListView, ByRef iArray() As Long, ByVal iCount As Long)
    On Error Resume Next
    pDeselect lv
    With lv.ListItems
        For iCount = 0 To iCount - 1&
            .Item("k" & iArray(iCount)).Checked = True
        Next
    End With
End Sub

Private Sub pMessageOption(ByVal iWhen As eMsgWhen, ByVal bAll As Boolean)
    Dim lv As ListView
    Set lv = lvwMessages(iWhen And Not MSG_BEFORE)
    pDeselect lv
    lv.Enabled = Not bAll
    
    If Not mbFreezeMessages Then
        With Subclasses(Me).Item(HexVal(lst(lstSubclass).Text))
            If bAll Then
                .AddMsg ALL_MESSAGES, iWhen
            Else
                .DelMsg ALL_MESSAGES, iWhen
            End If
        End With
    End If
End Sub

Private Sub pShowMessages(ByVal iTest As eTest)
    Dim lhWnd As Long
    Dim iCount As Long
    Dim iMessages() As Long
    Dim bEnabled As Boolean
    Dim lsClass As String
    
    On Error Resume Next
    
    mbFreezeMessages = True
    Select Case iTest
    Case tstSubclass
        lhWnd = HexVal(lst(lstSubclass).Text)
        If lhWnd Then
            With Subclasses(Me).Item(lhWnd)
                iCount = .GetMessages(iMessages, MSG_BEFORE)
                pShowMessagesSub iMessages, iCount, chk(chkSubclassBefore), lvwMessages(lvwSubclassBefore), opt(optAllBefore), opt(optSelBefore)
    
                iCount = .GetMessages(iMessages, MSG_AFTER)
                pShowMessagesSub iMessages, iCount, chk(chkSubclassAfter), lvwMessages(lvwSubclassAfter), opt(optAllAfter), opt(optSelAfter)
            End With
            bEnabled = True
        Else
            bEnabled = False
        End If
    Case tstClasses
        lsClass = lst(lstClasses).Text
        If Len(lsClass) Then
            With ApiWindowClasses.Item(lsClass)
                iCount = .GetDefMessages(iMessages)
                pShowMessagesSub iMessages, iCount, chk(chkAllClass), lvwMessages(lvwClass)
                lbl(lblCurrentClass).Caption = "Current Class: " & vbNewLine & vbNewLine & "Name: " & .Name & vbNewLine & "Total Windows: " & .TotalWindowCount
            End With
            bEnabled = True
        Else
            bEnabled = False
            lbl(lblCurrentClass).Caption = vbNullString
        End If
    Case tstClassWindows
        If lst(lstClassWindows).ListIndex > -1& Then
            lhWnd = HexVal(lst(lstClassWindows).Text)
            lsClass = xApiWindow(lhWnd).ClassName
            With ApiWindowClasses.Item(lsClass).OwnedWindows(Me).Item(lhWnd)
                iCount = .GetMessages(iMessages)
                bEnabled = Not .DefaultMessages
                chk(chkClassWindowDefault).Value = Abs(Not bEnabled)
                pShowMessagesSub iMessages, iCount, chk(chkAllClassWindow), lvwMessages(lvwClassWindow)
            End With
        End If
    End Select
    pEnableMessages iTest, bEnabled
    mbFreezeMessages = False
End Sub

'############################
'## Implemented Interfaces ##
'############################

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef hwnd As Long, ByRef uMsg As eMsg, ByRef wParam As Long, ByRef lParam As Long)
Static nMsgNo As Long
    If miCurrentTest = tstSubclass And mbLoaded Then
        'If we try to Display the paint message we'll just cause another paint message... vicious circle.
        If Not ((uMsg = WM_PAINT Or uMsg = WM_ERASEBKGND) And (hwnd = Me.hwnd Or hwnd = pic(picDisplay).hwnd)) Then
            nMsgNo = nMsgNo + 1
            pDisplay FmtHex(nMsgNo) & _
                    IIf(bBefore, "Before ", "After  ") & _
                    FmtHex(lReturn) & _
                    FmtHex(hwnd) & _
                    FmtHex(uMsg) & _
                    FmtHex(wParam) & _
                    FmtHex(lParam) & _
                    GetMsgName(uMsg)

        End If
    End If
End Sub

Private Sub iHook_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, nCode As eHookCode, wParam As Long, lParam As Long)
Static nHookNo As Long
Static bInHere As Boolean
    If miCurrentTest = tstHooks And mbLoaded Then
        If bInHere Then Exit Sub
        bInHere = True
        nHookNo = nHookNo + 1&
        pDisplay FmtHex(nHookNo) & _
                IIf(bBefore, "Before ", "After  ") & _
                FmtHex(lReturn) & _
                FmtHex(nCode) & _
                FmtHex(wParam) & _
                FmtHex(lParam) & _
                GetHCName(nCode) & "/" & HookName(Hooks(Me).CurrentHook)
        bInHere = False
    End If

End Sub

Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
    Static iTimerNo As Long
    iTimerNo = iTimerNo + 1&
    If miCurrentTest = tstTimers And mbLoaded Then pDisplay FmtHex(iTimerNo) & _
                                              FmtHex(lTimerID) & _
                                              FmtHex(Timers(Me).Item(lTimerID).Interval) & _
                                              FmtHex(lElapsedMS)
End Sub

Private Sub iWindow_Proc(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As eMsg, wParam As Long, lParam As Long)
Static nMsgNo As Long
    If (miCurrentTest = tstClassWindows Or miCurrentTest = tstClassWindows) And mbLoaded Then
        nMsgNo = nMsgNo + 1
        pDisplay FmtHex(nMsgNo) & _
                FmtHex(lReturn) & _
                FmtHex(hwnd) & _
                FmtHex(uMsg) & _
                FmtHex(wParam) & _
                FmtHex(lParam) & _
                GetMsgName(uMsg)
    End If
End Sub

'############################
'##   Public Procedures    ##
'############################

Public Sub BroadcastClass(ByVal bAdd As Boolean, ByRef sClass As String, Optional ByVal bBroadcast As Boolean)
    If Not bBroadcast Then
        With lst(lstClasses)
            If bAdd Then
                .AddItem sClass
            Else
                Dim i As Long
                For i = 0 To .ListCount - 1&
                    If .List(i) = sClass Then .RemoveItem i: Exit For
                Next
            End If
        End With
    Else
        Dim loF As fTest
        Dim loTemp As Object
        For Each loTemp In Forms
            If TypeOf loTemp Is fTest Then
                Set loF = loTemp
                loF.BroadcastClass bAdd, sClass
            End If
        Next
    End If
End Sub

Public Sub AddSubclass(ByVal ihWnd As Long)
    On Error GoTo handler
    If ihWnd Then
        With Subclasses(Me)
            If Not .Exists(ihWnd) Then
                .Add ihWnd
                With lst(lstSubclass)
                    .AddItem "0x" & FmtHex(ihWnd)
                    .ListIndex = .NewIndex
                End With
            Else
                If MsgBox("This window is already being subclassed!" & vbCrLf & vbCrLf & _
                          "Do you want to try another one?", _
                          vbYesNo + vbDefaultButton1 + vbQuestion, _
                          "Window already Subclassed") _
                    = vbYes Then fFindWindow.GetWindow Me
            End If
        End With
    End If
    Exit Sub
handler:
    pDisplayErr
End Sub
