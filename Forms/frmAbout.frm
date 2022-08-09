VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Us"
   ClientHeight    =   4350
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3002.447
   ScaleMode       =   0  'User
   ScaleWidth      =   5084.966
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   0
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":0ECA
      TabIndex        =   13
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   1905
      TabIndex        =   12
      Top             =   0
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   3615
         Left            =   0
         Picture         =   "frmAbout.frx":0F37
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   6240
      TabIndex        =   11
      Top             =   4440
      Width           =   4935
      Begin VB.CommandButton CmdSEnd 
         Caption         =   "Send Mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   6240
      TabIndex        =   6
      Top             =   0
      Width           =   4935
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":543F
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Krayknot@Yahoo.com"
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox TxtSubject 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox TxtBody 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":549C
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":5505
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":556C
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   3600
      Width           =   5415
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Index           =   11
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":55CD
         TabIndex        =   24
         Top             =   240
         Width           =   3255
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   1
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":5624
      TabIndex        =   14
      Top             =   360
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   2
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":568D
      TabIndex        =   15
      Top             =   960
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   3
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":5702
      TabIndex        =   16
      Top             =   1200
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   4
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":5771
      TabIndex        =   17
      Top             =   1440
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   5
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":57FC
      TabIndex        =   18
      Top             =   1680
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   6
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":5867
      TabIndex        =   19
      Top             =   1920
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   495
      Index           =   7
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":58FE
      TabIndex        =   20
      Top             =   2280
      Width           =   3375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   8
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":59F9
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   9
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":5A7C
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Index           =   10
      Left            =   2040
      OleObjectBlob   =   "frmAbout.frx":5AF9
      TabIndex        =   23
      Top             =   3000
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
SkinLabel5(11).Caption = USERNAME
End Sub

Private Sub Label12_Click()
Me.Width = 11350
End Sub
