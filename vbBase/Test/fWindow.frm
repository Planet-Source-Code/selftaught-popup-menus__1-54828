VERSION 5.00
Begin VB.Form fWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window"
   ClientHeight    =   4155
   ClientLeft      =   3675
   ClientTop       =   2100
   ClientWidth     =   5505
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucWindowInfo WindowInfo 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7011
   End
End
Attribute VB_Name = "fWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    xApiWindow(Me.hwnd).SetPos 0, 0, 0, 0, HWND_TOPMOST, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub
