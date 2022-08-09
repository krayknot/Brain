VERSION 5.00
Object = "{C3CBD80D-C8D1-11D2-9F8E-0080C7CE5CDC}#4.1#0"; "ActCndy2.ocx"
Object = "{D8A9DA2D-AB82-4962-B789-727EBE641D59}#1.0#0"; "cpvButton.ocx"
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      FillColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   5640
      ScaleHeight     =   330
      ScaleWidth      =   2085
      TabIndex        =   11
      Top             =   4220
      Width           =   2115
      Begin VB.Label lblmarks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Balls on the rampage"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   50
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   6
      Left            =   360
      Top             =   4800
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      FillColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      ScaleHeight     =   345
      ScaleWidth      =   4725
      TabIndex        =   9
      Top             =   4220
      Width           =   4750
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Balls on the rampage"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   50
         Width           =   4695
      End
   End
   Begin VB.ComboBox cmblistpaper 
      Height          =   315
      ItemData        =   "FrmMain.frx":0000
      Left            =   960
      List            =   "FrmMain.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   7
      Text            =   "cmblistpaper"
      Top             =   4845
      Width           =   3015
   End
   Begin VB.ListBox LstOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00000000&
      Height          =   2760
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin ActiveCandy.CandyHCommand cmdnext 
      Height          =   480
      Left            =   4080
      TabIndex        =   0
      Top             =   4800
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   16776960
      CandyStyle      =   1
      Caption         =   "Next>>"
      PictureMode     =   2
      UseAnimation    =   -1  'True
      ShineMode       =   2
      UseOnTop        =   0   'False
   End
   Begin ActiveCandy.CandyHCommand cmdback 
      Height          =   480
      Left            =   5160
      TabIndex        =   1
      Top             =   4800
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   16776960
      CandyStyle      =   1
      Caption         =   "<<Back"
      PictureMode     =   2
      UseAnimation    =   -1  'True
      ShineMode       =   2
      UseOnTop        =   0   'False
   End
   Begin ActiveCandy.CandyHCommand cmdcheck 
      Height          =   480
      Left            =   6240
      TabIndex        =   2
      Top             =   4800
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   847
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   16776960
      CandyStyle      =   1
      Caption         =   "#Check#"
      PictureMode     =   2
      UseAnimation    =   -1  'True
      ShineMode       =   2
      UseOnTop        =   0   'False
   End
   Begin ActiveCandy.CandyHCommand CandyHCommand4 
      Height          =   480
      Left            =   7320
      TabIndex        =   3
      Top             =   4800
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   16776960
      CandyStyle      =   1
      Caption         =   "Calculate"
      PictureMode     =   2
      UseAnimation    =   -1  'True
      ShineMode       =   2
      UseOnTop        =   0   'False
   End
   Begin Button2.cpvButton cpvButton1 
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   192
      Caption         =   "Quit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin Button2.cpvButton cpvButton2 
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   192
      Caption         =   "Help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin Button2.cpvButton cpvButton3 
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   192
      Caption         =   "Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin Button2.cpvButton CmdChoices 
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   192
      Caption         =   "Choices"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin Button2.cpvButton cpvButton4 
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   192
      Caption         =   "Voice"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   65535
   End
   Begin VB.Label lblnum 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblquestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Height          =   975
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   0
      Picture         =   "FrmMain.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
