VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmcreatewizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Create Paper Wizard"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   17190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   9720
      TabIndex        =   84
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame7 
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
      Height          =   3735
      Left            =   0
      TabIndex        =   61
      Top             =   2400
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   120
         TabIndex        =   90
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   120
         TabIndex        =   89
         Top             =   2230
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   120
         TabIndex        =   88
         Top             =   1970
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   645
         Index           =   11
         Left            =   120
         TabIndex        =   87
         Top             =   3000
         Width           =   3615
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":0000
         TabIndex        =   66
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   65
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":1AC2
         TabIndex        =   64
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   645
         Index           =   7
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":3584
         TabIndex        =   62
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5046
         TabIndex        =   67
         Top             =   240
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   12
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":50E3
         TabIndex        =   68
         Top             =   480
         Width           =   3495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   13
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5160
         TabIndex        =   69
         Top             =   1440
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   14
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":51D5
         TabIndex        =   70
         Top             =   2760
         Width           =   3255
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   0
      TabIndex        =   39
      Top             =   6120
      Width           =   9495
      Begin VB.ListBox List5 
         Height          =   1035
         Left            =   120
         TabIndex        =   85
         Top             =   480
         Width           =   4455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Edit"
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
         Left            =   8520
         TabIndex        =   49
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Open"
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
         Left            =   7680
         TabIndex        =   48
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
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
         Left            =   6840
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdmake 
         Caption         =   "Make"
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
         Left            =   8520
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdcheck 
         Caption         =   "Check"
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
         Left            =   7680
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
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
         Left            =   6840
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "About"
         Height          =   375
         Left            =   6720
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         Height          =   375
         Left            =   6000
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdclose 
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
         Left            =   8520
         TabIndex        =   41
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton CmdAbout 
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
         Left            =   7680
         TabIndex        =   40
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   17
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5246
         TabIndex        =   86
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   4815
      Left            =   4800
      TabIndex        =   12
      Top             =   1320
      Width           =   4695
      Begin VB.CommandButton cmdremoveanswers 
         Caption         =   "Remove"
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
         Left            =   3720
         TabIndex        =   32
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton cmdremoveoptions 
         Caption         =   "Remove"
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
         Left            =   3720
         TabIndex        =   31
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton cmdremovequestion 
         Caption         =   "Remove"
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
         Left            =   3720
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox List3 
         Height          =   645
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   4455
      End
      Begin VB.ListBox List2 
         Height          =   645
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   4455
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":52A9
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5322
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   3
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5397
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   6
         Left            =   3000
         OleObjectBlob   =   "frmcreatewizard.frx":540C
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   7
         Left            =   3000
         OleObjectBlob   =   "frmcreatewizard.frx":5463
         TabIndex        =   34
         Top             =   1800
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   8
         Left            =   3000
         OleObjectBlob   =   "frmcreatewizard.frx":54BA
         TabIndex        =   35
         Top             =   3360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   945
         ScaleWidth      =   4425
         TabIndex        =   36
         Top             =   240
         Width           =   4455
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PaperName"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   4095
         End
         Begin VB.Label LblType 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paper Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   4095
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.OptionButton Option5 
         Caption         =   "True or False"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Descriptive Questions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Multiple Choice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Matching Columns"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fill in the Blanks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5511
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":559E
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5613
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":5682
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   3735
      Left            =   5760
      TabIndex        =   76
      Top             =   7920
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   12
         Left            =   120
         TabIndex        =   80
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":570F
         TabIndex        =   79
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   1725
         Index           =   10
         Left            =   120
         TabIndex        =   78
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":71D1
         TabIndex        =   77
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":8C93
         TabIndex        =   81
         Top             =   240
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   16
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":8D30
         TabIndex        =   82
         Top             =   480
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   18
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":8DC5
         TabIndex        =   83
         Top             =   2520
         Width           =   4335
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Descriptive Questions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   5640
      TabIndex        =   71
      Top             =   7800
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   2805
         Index           =   9
         Left            =   120
         TabIndex        =   73
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":8E3C
         TabIndex        =   72
         Top             =   2760
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":A8FE
         TabIndex        =   74
         Top             =   240
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   15
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":A99B
         TabIndex        =   75
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mathing Columns"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   50
      Top             =   2400
      Width           =   4695
      Begin VB.CommandButton Command6 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":AA30
         TabIndex        =   56
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":C4F2
         TabIndex        =   54
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   5
         Left            =   120
         TabIndex        =   53
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":DFB4
         TabIndex        =   52
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   4
         Left            =   120
         TabIndex        =   51
         Top             =   2880
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":FA76
         TabIndex        =   57
         Top             =   240
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   9
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":FB13
         TabIndex        =   58
         Top             =   480
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   10
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":FBA8
         TabIndex        =   59
         Top             =   1560
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   11
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":FC39
         TabIndex        =   60
         Top             =   2640
         Width           =   4335
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   3735
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   3615
      End
      Begin VB.CommandButton cmdaddfillanswers 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":FCB0
         TabIndex        =   27
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CommandButton cmdaddfilloptions 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":11772
         TabIndex        =   24
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   765
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdaddfill 
         Caption         =   "Add 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmcreatewizard.frx":13234
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":14CF6
         TabIndex        =   22
         Top             =   240
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":14D93
         TabIndex        =   23
         Top             =   480
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   4
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":14E52
         TabIndex        =   26
         Top             =   1560
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Index           =   5
         Left            =   120
         OleObjectBlob   =   "frmcreatewizard.frx":14F0D
         TabIndex        =   29
         Top             =   2640
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmcreatewizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IntFill As Integer
Dim IntFilloptions As Integer
Dim IntFillAnswers As Integer
Dim FSO As New FileSystemObject

Private Sub cmdaddfill_Click()
If Text1(1).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(1).SetFocus
   Exit Sub
End If

If List1.ListCount < 10 Then
    List1.AddItem Text1(1).Text
    Text1(1).Text = ""
    Text1(1).SetFocus
    cmdaddfill.Caption = "Add " & List1.ListCount
    SkinLabel6(6).Caption = List1.ListCount
End If
If SkinLabel6(6).Caption = "10" Then
   cmdaddfill.Enabled = False
End If
Text1(1).Text = ""
   
End Sub

Private Sub cmdaddfillanswers_Click()
If Text1(3).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(3).SetFocus
   Exit Sub
End If

If List3.ListCount < 10 Then
    List3.AddItem Text1(3).Text
    Text1(3).Text = ""
    Text1(3).SetFocus
    cmdaddfillanswers.Caption = "Add " & List3.ListCount
    SkinLabel6(8).Caption = List3.ListCount
End If
If SkinLabel6(8).Caption = "10" Then
   cmdaddfillanswers.Enabled = False
End If
Text1(3).Text = ""
End Sub

Private Sub cmdaddfilloptions_Click()
If Text1(2).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(2).SetFocus
   Exit Sub
End If

If List2.ListCount < 10 Then
    List2.AddItem Text1(2).Text
    Text1(2).Text = ""
    Text1(2).SetFocus
    cmdaddfilloptions.Caption = "Add " & List2.ListCount
    SkinLabel6(7).Caption = List2.ListCount
End If
If SkinLabel6(7).Caption = "10" Then
   cmdaddfilloptions.Enabled = False
End If
Text1(2).Text = ""
End Sub

Private Sub CmdCheck_Click()
'Checks the paper name
    If Label2.Caption = "Paper Name" Or Label2.Caption = "" Then
       MESSAGE "You cannot have a paper name 'Paper Name'. Please change it.", OkOnly, "Error"
       Text1(0).SetFocus
       Exit Sub
    End If

'Checks the paper type
    If Option1.Value = False And Option2.Value = False And Option3.Value = False And _
       Option4.Value = False And Option5.Value = False Then
       MESSAGE "Select the Paper Type", OkOnly, "Error"
       Option1.SetFocus
       Exit Sub
    End If

'Checks the paper format according to the paper type selected
    If Option1.Value = True Then  'if  fill in  the blanks
       If List1.ListCount <> 10 Or List2.ListCount <> 10 Or List3.ListCount <> 10 Then
          MESSAGE "Number of Questions, Answers and Options must be 10. Please check it", OkOnly, "Error"
          Exit Sub
       End If
    End If
    

   
End Sub

Private Sub cmdclear_Click()
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = ""
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdmake_Click()
'Checks the correct format of questions, answers an options in the list boxes
 List5.Clear

        'Checks the paper name
        If Label2.Caption = "PaperName" Or Label2.Caption = "" Then
           List5.AddItem "You cannot have a paper name 'PaperName'. Please change it."
           Text1(0).SetFocus
           Exit Sub
        End If
    
    'Checks the paper type
        If Option1.Value = False And Option2.Value = False And Option3.Value = False And _
           Option4.Value = False And Option5.Value = False Then
           List5.AddItem "Select the Paper Type"
           Option1.SetFocus
           Exit Sub
        End If
    
    'Checks the paper format according to the paper type selected
        If Option1.Value = True Then  'if  fill in  the blanks
           If List1.ListCount <> 10 Or List2.ListCount <> 10 Or List3.ListCount <> 10 Then
              List5.AddItem "Number of Questions, Answers and Options must be 10."
              Exit Sub
           End If
        End If

 

'Give a perfect name to the file of the paper according the name and type of the paper
 List5.AddItem "Checking Files and Folder"
 If Option1.Value = True Then 'if fill in  the blanks
    Dim inttemp As Integer, IntPaperNo As Integer, IntTemp1 As Integer
    Dim StrTemp As String
    Dim Flag As Boolean
    IntPaperNo = 1
    
       If Not FSO.FolderExists(App.Path & "\CreatePaper") Then
       FSO.CreateFolder (App.Path & "\createpaper")
    End If

    If Not FSO.FolderExists(App.Path & "\CreatePaper\" & Trim(Label2.Caption)) Then
       FSO.CreateFolder (App.Path & "\createpaper\" & Trim(Label2.Caption))
    End If
        
    If Not FSO.FolderExists(App.Path & "\CreatePaper\" & Trim(Label2.Caption) & "\Fillintheblanks") Then
       FSO.CreateFolder (App.Path & "\createpaper\" & Trim(Label2.Caption) & "\Fillintheblanks")
    End If
    List5.AddItem "Files and Folder checked"
    
    File1.Path = App.Path & "\createpaper\" & Trim(Label2.Caption) & "\Fillintheblanks\"
    File1.Refresh
    
    For inttemp = 1 To 100
        StrTemp = "Paper" & inttemp & ".BRN"
        If Not FSO.FileExists(App.Path & "\createpaper\" & Trim(Label2.Caption) & "\Fillintheblanks\" & StrTemp) Then
           MsgBox StrTemp
           Exit For
        End If
    Next
   
   'make the paper
    List5.AddItem "Creating Paper"
    fname = App.Path & "\createpaper\" & Trim(Label2.Caption) & "\fillintheblanks\" & Trim(StrTemp)
    Open fname For Output As #2
    Print #2, "[Paper Name]"
    Print #2, Trim(Label2.Caption)
    Print #2, "[Paper Type]"
    Print #2, Trim(LblType.Caption)

    Print #2, "[Questions]"
    For inttemp = 0 To 9
        Print #2, List1.List(inttemp)
    Next inttemp

    Print #2, "[Options]"
    For inttemp = 0 To 9
        Print #2, List2.List(inttemp)
    Next inttemp

    Print #2, "[Answers]"
    For inttemp = 0 To 9
        Print #2, List3.List(inttemp)
    Next inttemp
    Close #2
    List5.AddItem "Paper Creation completed"

ElseIf Option2.Value = True Then
'    Dim IntTemp As Integer, IntPaperNo As Integer, IntTemp1 As Integer
'    Dim StrTemp As String
'    Dim Flag As Boolean
    IntPaperNo = 1
    
       If Not FSO.FolderExists(App.Path & "\CreatePaper") Then
       FSO.CreateFolder (App.Path & "\createpaper")
    End If

    If Not FSO.FolderExists(App.Path & "\CreatePaper\" & Trim(Label2.Caption)) Then
       FSO.CreateFolder (App.Path & "\createpaper\" & Trim(Label2.Caption))
    End If
        
    If Not FSO.FolderExists(App.Path & "\CreatePaper\" & Trim(Label2.Caption) & "\MAtchingColumns") Then
       FSO.CreateFolder (App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MatchingColumns")
    End If
    List5.AddItem "Files and Folder checked"
    
    File1.Path = App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MatchingColumns\"
    File1.Refresh
    
    For inttemp = 1 To 100
        StrTemp = "Paper" & inttemp & ".BRN"
        If Not FSO.FileExists(App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MatchingColumns\" & StrTemp) Then
           MsgBox StrTemp
           Exit For
        End If
    Next
   
   'make the paper
    List5.AddItem "Creating Paper"
    fname = App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MatchingColumns\" & Trim(StrTemp)
    Open fname For Output As #2
    Print #2, "[Paper Name]"
    Print #2, Trim(Label2.Caption)
    Print #2, "[Paper Type]"
    Print #2, Trim(LblType.Caption)

    Print #2, "[Questions]"
    For inttemp = 0 To 9
        Print #2, List1.List(inttemp)
    Next inttemp

    Print #2, "[Options]"
    For inttemp = 0 To 9
        Print #2, List2.List(inttemp)
    Next inttemp

    Print #2, "[Answers]"
    For inttemp = 0 To 9
        Print #2, List3.List(inttemp)
    Next inttemp
    Close #2
    List5.AddItem "Paper Creation completed"
       
ElseIf Option3.Value = True Then
'    Dim IntTemp As Integer, IntPaperNo As Integer, IntTemp1 As Integer
'    Dim StrTemp As String
'    Dim Flag As Boolean
    IntPaperNo = 1
    
    If Not FSO.FolderExists(App.Path & "\CreatePaper") Then
       FSO.CreateFolder (App.Path & "\createpaper")
    End If

    If Not FSO.FolderExists(App.Path & "\CreatePaper\" & Trim(Label2.Caption)) Then
       FSO.CreateFolder (App.Path & "\createpaper\" & Trim(Label2.Caption))
    End If
        
    If Not FSO.FolderExists(App.Path & "\CreatePaper\" & Trim(Label2.Caption) & "\MultipleChoice") Then
       FSO.CreateFolder (App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MultipleChoice")
    End If
    List5.AddItem "Files and Folder checked"
    
    File1.Path = App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MultipleChoice\"
    File1.Refresh
    
    For inttemp = 1 To 100
        StrTemp = "Paper" & inttemp & ".BRN"
        If Not FSO.FileExists(App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MultipleChoice\" & StrTemp) Then
           MsgBox StrTemp
           Exit For
        End If
    Next
   
   'make the paper
    List5.AddItem "Creating Paper"
    fname = App.Path & "\createpaper\" & Trim(Label2.Caption) & "\MultipleChoice\" & Trim(StrTemp)
    Open fname For Output As #2
    Print #2, "[Paper Name]"
    Print #2, Trim(Label2.Caption)
    Print #2, "[Paper Type]"
    Print #2, Trim(LblType.Caption)

    Print #2, "[Questions]"
    For inttemp = 0 To 9
        Print #2, List1.List(inttemp)
    Next inttemp

    Print #2, "[Options]"
    For inttemp = 0 To 9
        Print #2, List2.List(inttemp)
    Next inttemp

    Print #2, "[Answers]"
    For inttemp = 0 To 9
        Print #2, List3.List(inttemp)
    Next inttemp
    Close #2
    List5.AddItem "Paper Creation completed"


ElseIf Option4.Value = True Then

ElseIf Option5.Value = True Then

End If

    
End Sub

Private Sub cmdremoveanswers_Click()
If List3.Text = "" Then
   MESSAGE "Please select the Question first from the list", OkOnly, "Error"
   List3.SetFocus
   Exit Sub
End If

List3.RemoveItem List3.ListIndex
End Sub

Private Sub cmdremoveoptions_Click()
If List2.Text = "" Then
   MESSAGE "Please select the Question first from the list", OkOnly, "Error"
   List2.SetFocus
   Exit Sub
End If

List2.RemoveItem List2.ListIndex
End Sub

Private Sub cmdremovequestion_Click()
If List1.Text = "" Then
   MESSAGE "Please select the Question first from the list", OkOnly, "Error"
   List1.SetFocus
   Exit Sub
End If

List1.RemoveItem List1.ListIndex

End Sub

Private Sub cmdsave_Click()
''Checks the paper name
'    If Label2.Caption = "Paper Name" Or Label2.Caption = "" Then
'       MESSAGE "You cannot have a paper name 'Paper Name'. Please change it.", OkOnly, "Error"
'       Text1(0).SetFocus
'       Exit Sub
'    End If
'
''Checks the paper type
'    If Option1.Value = False And Option2.Value = False And Option3.Value = False And _
'       Option4.Value = False And Option5.Value = False Then
'       MESSAGE "Select the Paper Type", OkOnly, "Error"
'       Option1.SetFocus
'       Exit Sub
'    End If
'
''Checks the paper format according to the paper type selected
'    If Option1.Value = True Then  'if  fill in  the blanks
'       If List1.ListCount <> 10 Or List2.ListCount <> 10 Or List3.ListCount <> 10 Then
'          MESSAGE "Number of Questions, Answers and Options must be 10. Please check it", OkOnly, "Error"
'          Exit Sub
'       End If
'    End If
    
    Dim inttemp As Integer
    fname = App.Path & "\createpaper\" & Trim(Label2.Caption) & ".BRN"
    Open fname For Output As #2
    Print #2, "                                      This file cannot be read in MS-DOS mode"
    Print #2, "[Paper Name]"
    Print #2, Trim(Label2.Caption)
    Print #2, "[Paper Type]"
    Print #2, Trim(LblType.Caption)
    
    If Option1.Value = True Then ' if fill in the blanks
       Print #2, "[Questions]"
       For inttemp = 0 To 9
           Print #2, List1.List(inttemp)
       Next inttemp
       Print #2, "[Options]"
       Print #2, "[Answers]"
       Close #2
    End If
    
     
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command1_Click()
If Text1(4).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(4).SetFocus
   Exit Sub
End If

If List3.ListCount < 10 Then
    List3.AddItem Text1(4).Text
    Text1(4).Text = ""
    Text1(4).SetFocus
    Command1.Caption = "Add " & List3.ListCount
    SkinLabel6(8).Caption = List3.ListCount
End If
If SkinLabel6(8).Caption = "10" Then
   Command1.Enabled = False
End If
Text1(4).Text = ""

End Sub

Private Sub Command10_Click()
If Text1(8).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(8).SetFocus
   Exit Sub
End If

If List2.ListCount < 10 Then
    List2.AddItem Text1(8).Text
    Text1(8).Text = ""
    Text1(8).SetFocus
    Command10.Caption = "Add " & List2.ListCount
    SkinLabel6(7).Caption = List2.ListCount
End If
If SkinLabel6(7).Caption = "10" Then
   Command10.Enabled = False
End If
Text1(8).Text = ""
End Sub

Private Sub Command11_Click()
If Text1(11).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(11).SetFocus
   Exit Sub
End If

If List3.ListCount < 10 Then
    List3.AddItem Text1(11).Text
    Text1(11).Text = ""
    Text1(11).SetFocus
    Command11.Caption = "Add " & List3.ListCount
    SkinLabel6(8).Caption = List3.ListCount
End If
If SkinLabel6(8).Caption = "10" Then
   Command11.Enabled = False
End If
Text1(11).Text = ""
End Sub

Private Sub Command5_Click()
If Text1(5).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(5).SetFocus
   Exit Sub
End If

If List2.ListCount < 10 Then
    List2.AddItem Text1(5).Text
    Text1(5).Text = ""
    Text1(5).SetFocus
    Command5.Caption = "Add " & List2.ListCount
    SkinLabel6(7).Caption = List2.ListCount
End If
If SkinLabel6(7).Caption = "10" Then
   Command5.Enabled = False
End If
Text1(5).Text = ""

End Sub

Private Sub Command6_Click()
If Text1(6).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(6).SetFocus
   Exit Sub
End If

If List1.ListCount < 10 Then
    List1.AddItem Text1(6).Text
    Text1(6).Text = ""
    Text1(6).SetFocus
    Command6.Caption = "Add " & List1.ListCount
    SkinLabel6(6).Caption = List1.ListCount
End If
If SkinLabel6(6).Caption = "10" Then
   Command6.Enabled = False
End If
Text1(6).Text = ""

End Sub

Private Sub Command7_Click()
If Text1(7).Text = "" Then
   MESSAGE "Blank entry is not allowed", OkOnly, "Error"
   Text1(6).SetFocus
   Exit Sub
End If

If List1.ListCount < 10 Then
    List1.AddItem Text1(7).Text
    Text1(7).Text = ""
    Text1(7).SetFocus
    Command7.Caption = "Add " & List1.ListCount
    SkinLabel6(6).Caption = List1.ListCount
End If
If SkinLabel6(6).Caption = "10" Then
   Command7.Enabled = False
End If
Text1(7).Text = ""
End Sub

Private Sub Form_Load()
'FrmSelection.Skin1.ApplySkin Me.hwnd

IntFill = 1
IntFilloptions = 1
IntFillAnswers = 1

End Sub

Private Sub Option1_Click()
LblType.Caption = proper(Option1.Caption)
Frame4.ZOrder 0
End Sub

Private Sub Option2_Click()
LblType.Caption = proper(Option2.Caption)
Frame6.ZOrder 0
End Sub

Private Sub Option3_Click()
LblType.Caption = proper(Option3.Caption)
Frame7.ZOrder 0
End Sub

Private Sub Option4_Click()
LblType.Caption = proper(Option4.Caption)
End Sub

Private Sub Option5_Click()
LblType.Caption = proper(Option5.Caption)
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text1_Change(Index As Integer)
Label2.Caption = proper(Text1(0).Text)
End Sub
