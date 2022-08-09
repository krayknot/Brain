VERSION 5.00
Begin VB.Form FrmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN2: Report"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   80
      Top             =   7680
      Width           =   9015
      Begin VB.CommandButton CmdCLose 
         Caption         =   "Close"
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
         Left            =   8040
         MousePointer    =   1  'Arrow
         TabIndex        =   90
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdHelp 
         Caption         =   "Help"
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
         Left            =   7200
         MousePointer    =   1  'Arrow
         TabIndex        =   89
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdReply 
         Caption         =   "Replay"
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
         Left            =   6000
         MousePointer    =   1  'Arrow
         TabIndex        =   88
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdOptions 
         Caption         =   "Options"
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
         Left            =   5160
         MousePointer    =   1  'Arrow
         TabIndex        =   87
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
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
         Left            =   4320
         MousePointer    =   1  'Arrow
         TabIndex        =   86
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Next"
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
         MousePointer    =   1  'Arrow
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Next"
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
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   84
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
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
         Left            =   1800
         MousePointer    =   1  'Arrow
         TabIndex        =   83
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
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
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   82
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "Next"
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
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2FFF5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   74
         Left            =   480
         TabIndex        =   105
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   73
         Left            =   960
         TabIndex        =   104
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   72
         Left            =   1440
         TabIndex        =   103
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   71
         Left            =   1920
         TabIndex        =   102
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   70
         Left            =   2400
         TabIndex        =   101
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   69
         Left            =   2880
         TabIndex        =   100
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   68
         Left            =   3360
         TabIndex        =   99
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   67
         Left            =   3840
         TabIndex        =   98
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   66
         Left            =   4320
         TabIndex        =   97
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   65
         Left            =   4800
         TabIndex        =   96
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   5280
         TabIndex        =   95
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   5760
         TabIndex        =   94
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   6240
         TabIndex        =   93
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   61
         Left            =   6720
         TabIndex        =   92
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   7200
         TabIndex        =   91
         Top             =   1480
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5760
         TabIndex        =   79
         Top             =   7320
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   78
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5760
         TabIndex        =   77
         Top             =   6840
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Appearance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   76
         Top             =   7320
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Appearance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   75
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Candidate Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   74
         Top             =   6840
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF80&
         Index           =   4
         X1              =   360
         X2              =   8880
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   4
         Index           =   4
         X1              =   360
         X2              =   8880
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   7200
         TabIndex        =   73
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   6720
         TabIndex        =   72
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   6240
         TabIndex        =   71
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   5760
         TabIndex        =   70
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   5280
         TabIndex        =   69
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   4800
         TabIndex        =   68
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   4320
         TabIndex        =   67
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   3840
         TabIndex        =   66
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   3360
         TabIndex        =   65
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   2880
         TabIndex        =   64
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   2400
         TabIndex        =   63
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   1920
         TabIndex        =   62
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   1440
         TabIndex        =   61
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   960
         TabIndex        =   60
         Top             =   5685
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   480
         TabIndex        =   59
         Top             =   5685
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   15
         X1              =   360
         X2              =   7560
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   14
         X1              =   360
         X2              =   7560
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   63
         X1              =   7560
         X2              =   7560
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   62
         X1              =   7080
         X2              =   7080
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   61
         X1              =   6600
         X2              =   6600
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   60
         X1              =   6120
         X2              =   6120
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   59
         X1              =   5640
         X2              =   5640
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   58
         X1              =   5160
         X2              =   5160
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   57
         X1              =   4680
         X2              =   4680
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   56
         X1              =   4200
         X2              =   4200
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   55
         X1              =   3720
         X2              =   3720
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   54
         X1              =   3240
         X2              =   3240
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   53
         X1              =   2760
         X2              =   2760
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   52
         X1              =   2280
         X2              =   2280
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   51
         X1              =   1800
         X2              =   1800
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   50
         X1              =   1320
         X2              =   1320
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   49
         X1              =   840
         X2              =   840
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   48
         X1              =   360
         X2              =   360
         Y1              =   5640
         Y2              =   6480
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   13
         X1              =   360
         X2              =   7560
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   12
         X1              =   360
         X2              =   7560
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7680
         TabIndex        =   58
         Top             =   6240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7680
         TabIndex        =   57
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF80&
         Index           =   3
         X1              =   360
         X2              =   8880
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   4
         Index           =   3
         X1              =   360
         X2              =   8880
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiple Choice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   56
         Top             =   5280
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   7200
         TabIndex        =   55
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   6720
         TabIndex        =   54
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   6240
         TabIndex        =   53
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   5760
         TabIndex        =   52
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   5280
         TabIndex        =   51
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   4800
         TabIndex        =   50
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   4320
         TabIndex        =   49
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   3840
         TabIndex        =   48
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   3360
         TabIndex        =   47
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   2880
         TabIndex        =   46
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   2400
         TabIndex        =   45
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   1920
         TabIndex        =   44
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   1440
         TabIndex        =   43
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   960
         TabIndex        =   42
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   480
         TabIndex        =   41
         Top             =   4125
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   11
         X1              =   360
         X2              =   7560
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   10
         X1              =   360
         X2              =   7560
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   47
         X1              =   7560
         X2              =   7560
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   46
         X1              =   7080
         X2              =   7080
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   45
         X1              =   6600
         X2              =   6600
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   44
         X1              =   6120
         X2              =   6120
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   43
         X1              =   5640
         X2              =   5640
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   42
         X1              =   5160
         X2              =   5160
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   41
         X1              =   4680
         X2              =   4680
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   40
         X1              =   4200
         X2              =   4200
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   39
         X1              =   3720
         X2              =   3720
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   38
         X1              =   3240
         X2              =   3240
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   37
         X1              =   2760
         X2              =   2760
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   36
         X1              =   2280
         X2              =   2280
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   35
         X1              =   1800
         X2              =   1800
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   34
         X1              =   1320
         X2              =   1320
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   33
         X1              =   840
         X2              =   840
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   32
         X1              =   360
         X2              =   360
         Y1              =   4080
         Y2              =   4920
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   9
         X1              =   360
         X2              =   7560
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   8
         X1              =   360
         X2              =   7560
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7680
         TabIndex        =   40
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7680
         TabIndex        =   39
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF80&
         Index           =   2
         X1              =   360
         X2              =   8880
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   4
         Index           =   2
         X1              =   360
         X2              =   8880
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Matching Columns"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   38
         Top             =   3720
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   7200
         TabIndex        =   37
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   6720
         TabIndex        =   36
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   6240
         TabIndex        =   35
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   5760
         TabIndex        =   34
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   5280
         TabIndex        =   33
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   4800
         TabIndex        =   32
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   4320
         TabIndex        =   31
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   3840
         TabIndex        =   30
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   3360
         TabIndex        =   29
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   2880
         TabIndex        =   28
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   2400
         TabIndex        =   27
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   1920
         TabIndex        =   26
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   1440
         TabIndex        =   25
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   24
         Top             =   2565
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   23
         Top             =   2565
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   7
         X1              =   360
         X2              =   7560
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   6
         X1              =   360
         X2              =   7560
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   31
         X1              =   7560
         X2              =   7560
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   30
         X1              =   7080
         X2              =   7080
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   29
         X1              =   6600
         X2              =   6600
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   28
         X1              =   6120
         X2              =   6120
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   27
         X1              =   5640
         X2              =   5640
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   26
         X1              =   5160
         X2              =   5160
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   25
         X1              =   4680
         X2              =   4680
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   24
         X1              =   4200
         X2              =   4200
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   23
         X1              =   3720
         X2              =   3720
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   22
         X1              =   3240
         X2              =   3240
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   21
         X1              =   2760
         X2              =   2760
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   20
         X1              =   2280
         X2              =   2280
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   19
         X1              =   1800
         X2              =   1800
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   18
         X1              =   1320
         X2              =   1320
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   17
         X1              =   840
         X2              =   840
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   16
         X1              =   360
         X2              =   360
         Y1              =   2520
         Y2              =   3360
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   5
         X1              =   360
         X2              =   7560
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   4
         X1              =   360
         X2              =   7560
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   22
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   21
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF80&
         Index           =   1
         X1              =   360
         X2              =   8880
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   4
         Index           =   1
         X1              =   360
         X2              =   8880
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "True or False"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   2160
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   7200
         TabIndex        =   19
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   6720
         TabIndex        =   18
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   6240
         TabIndex        =   17
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5760
         TabIndex        =   16
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5280
         TabIndex        =   15
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   4800
         TabIndex        =   14
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   13
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   12
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   11
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   10
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   8
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   1000
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   1000
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   3
         X1              =   360
         X2              =   7560
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   2
         X1              =   360
         X2              =   7560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   15
         X1              =   7560
         X2              =   7560
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   14
         X1              =   7080
         X2              =   7080
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   13
         X1              =   6600
         X2              =   6600
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   12
         X1              =   6120
         X2              =   6120
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   11
         X1              =   5640
         X2              =   5640
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   10
         X1              =   5160
         X2              =   5160
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   9
         X1              =   4680
         X2              =   4680
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   8
         X1              =   4200
         X2              =   4200
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   7
         X1              =   3720
         X2              =   3720
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   6
         X1              =   3240
         X2              =   3240
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   5
         X1              =   2760
         X2              =   2760
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   4
         X1              =   2280
         X2              =   2280
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   3
         X1              =   1800
         X2              =   1800
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   2
         X1              =   1320
         X2              =   1320
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   1
         X1              =   840
         X2              =   840
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   960
         Y2              =   1800
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   1
         X1              =   360
         X2              =   7560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   7560
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Marks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   4
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF80&
         Index           =   0
         X1              =   360
         X2              =   8880
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   4
         Index           =   0
         X1              =   360
         X2              =   8880
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill in the Blanks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name of the Paper"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   120
         Width           =   5775
      End
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As New ADODB.Connection
Dim Rst As New ADODB.Recordset

Private Sub CmdCLose_Click()
Unload Me
End Sub

Private Sub CmdReply_Click()
frmreplay.Show vbModal
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd

Dim IntCheck As Integer, IntCnt As Integer, IntLblValue As Integer, IntCount As Integer
Dim strpn As String
Dim StrPT(4) As String

StrPT(1) = "Fill in the Blanks"
StrPT(2) = "True or False"
StrPT(3) = "Matching Columns"
StrPT(4) = "Multiple Choice"

 IntLblValue = 74
'Opening database
 DB.ConnectionString = "Provider='Microsoft.Jet.OLEDB.4.0';Data Source='" & App.Path & "\database\db.mdb';"
 DB.Open
 
 For IntCount = 1 To 4
     For IntCnt = 1 To 15
         Rst.Open "Select * From test where paperno = " & IntCnt & "and papertype = '" & StrPT(IntCount) & "' and questionright <> '' ", DB, adOpenDynamic, adLockOptimistic
        'Search for the same paperno and same papertype
         While Not Rst.EOF
               IntCheck = IntCheck + 1
               Rst.MoveNext
         Wend
         Label5(IntLblValue).Caption = IntCheck
         Label5(IntLblValue).Refresh
         IntLblValue = IntLblValue - 1
         IntCheck = 0
         Rst.Close
     Next IntCnt
 Next IntCount
 DB.Close
       
End Sub
