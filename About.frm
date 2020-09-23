VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2370
   ClientLeft      =   3465
   ClientTop       =   2205
   ClientWidth     =   3705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   30
   LinkTopic       =   "frmAbout"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00808000&
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "email: rtharp@ccp.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "VB Code Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "by Rick Tharp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  ' Unloads the form.
   Unload Me
   Set frmAbout = Nothing

End Sub

Private Sub Form_Load()
  Label1(2).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   
'Centers over main form
  Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
  Me.Top = frmMain.Top + 500

'Get active bar color
  Dim lngColor As Long
  lngColor = GetSysColor(COLOR_ACTIVECAPTION)
  
'Set background color to system active color
  frmAbout.BackColor = lngColor
  Label1(1).BackColor = lngColor
  Label1(2).BackColor = lngColor
  Label1(3).BackColor = lngColor
  lblEMail.BackColor = lngColor
  
End Sub


Private Sub imgImage_Click()

End Sub





Private Sub lblEMail_Click()
   Dim iRet As Long
    
   Dim Response As Integer
'   Response = MsgBox("You have chosen 'E-mail', which will" & vbCrLf & "launch your default e-mail program." & vbCrLf & vbCrLf & "Do you wish to continue?", vbInformation + vbYesNo, "E-mail")
   Response = MsgBox(LoadResString(100), vbInformation + vbYesNo, "E-mail")
   Select Case Response
     Case vbYes
         iRet = Shell("start.exe mailto:rtharp@ccp.com", vbNormal)
     Case vbNo
       Exit Sub
   End Select

End Sub


