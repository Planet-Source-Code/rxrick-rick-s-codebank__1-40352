VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepair 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repair database"
   ClientHeight    =   2220
   ClientLeft      =   2910
   ClientTop       =   2565
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   20
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2220
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   1500
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRepair 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Delay(NumberofSeconds As Single)
'Compute time until end of delay
  Dim EndTime As Single
  EndTime = Timer + NumberofSeconds
  
'Delay
  Do
    DoEvents  'Let other things happen
  Loop Until Timer >= EndTime
  
End Sub

Private Sub cmdOkay_Click()
'Unload repair dialog box
  Unload frmRepair
  Set frmRepair = Nothing

'Reload Code bank main form then exit
  frmMain.Show

End Sub

  
Private Sub Form_GotFocus()
'Set caption
  frmRepair.lblRepair.Caption = "Opening codebank.mdb for exclusive use"
  ProgressBar1.Value = 10
  Call Delay(1)

'Error trap for already open
  On Error GoTo DBErrorHandler

'Set database name
  Dim cDBName As String
  cDBName = App.Path + "\codebank.mdb"
  
'Open db for exclusive use to see if anyone using it
  Dim dbName As Database  ' database object
  Set dbName = OpenDatabase(cDBName, True)
    
'Close db
  dbName.Close
        
'Setup other Error Handler
  On Error GoTo OtherErrorHandler

'If "kinetics.bak" exists, erase it, then Copy current db to *.bak
  frmRepair.lblRepair.Caption = "Making backup of current codebank.mdb"
  ProgressBar1.Value = 25
  Call Delay(1)
  
  Dim bakfile$
  bakfile$ = App.Path + "\codebank.bak"
  If Dir$(bakfile$) = "codebank.bak" Then  ' Returns "filename" if exists.
     Kill bakfile$
  End If
  FileCopy cDBName, bakfile$
  
'Label RepairDatabase
  frmRepair.lblRepair.Caption = "Repairing current codebank.mdb"
  ProgressBar1.Value = 50
  Call Delay(1)
  
'RepairDatabase
  RepairDatabase cDBName

'If "tpnasist.tmp" exists, Kill it
  Dim tempfile$
  tempfile$ = App.Path + "\codebank.tmp"
  
  If Dir$(tempfile$) = "codebank.tmp" Then  ' Returns "filename" if exists.
     Kill tempfile$
  End If
     
'Label Compact Database
  frmRepair.lblRepair.Caption = "Compacting current codebank.mdb"
  ProgressBar1.Value = 75
  Call Delay(1)
  
'Compact Database
  CompactDatabase cDBName, tempfile$
  
'Copy tempfile$ to "tpnasist.mdb"
  FileCopy tempfile$, cDBName

'Kill tempfile$
  Kill tempfile$

'Unhide Okay button
  frmRepair.lblRepair.Caption = "Codebank.mdb repaired and compacted"
  ProgressBar1.Visible = False
  cmdOkay.Visible = True
    
'Disable error trapping
  On Error GoTo 0
  Exit Sub
  
DBErrorHandler:
  MsgBox "Unable to open database for exclusive use." + vbCrLf + "All other users must be logged off!", vbCritical + vbOKOnly, "Unable to repair database"
  GoTo SubExit

OtherErrorHandler:
  Dim strErr As String
  strErr = "Error " & Err.Number & " " & Err.Description + vbCrLf
  MsgBox strErr + "Unable to compact database.", vbCritical + vbOKOnly, "Compacting database failed"

SubExit:
  Unload Me
  frmMain.Show

End Sub
Private Sub Form_Load()
'Center on the screen
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = (Screen.Height - Me.Height) / 2

End Sub




