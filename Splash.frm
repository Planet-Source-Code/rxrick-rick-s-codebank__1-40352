VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2985
   ClientLeft      =   2310
   ClientTop       =   1815
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   10
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   0
      Picture         =   "Splash.frx":0442
      ScaleHeight     =   3000
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Center on the screen
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = (Screen.Height - Me.Height) / 2

End Sub


Private Sub Timer1_Timer()
  Load frmMain
  Unload Me
  frmMain.Show

End Sub


Sub Form_UnLoad(Cancel As Integer)
'*** Code added by HelpWriter ***
'*** Subroutine added by HelpWriter ***
    QuitHelp
'***********************************
End Sub
