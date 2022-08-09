VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form FrmBrain 
   Caption         =   "BRAIN: Now Be prepare for technicals within some days"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   ControlBox      =   0   'False
   Icon            =   "FrmBrain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   440
      Left            =   6840
      ScaleHeight     =   405
      ScaleWidth      =   1170
      TabIndex        =   43
      Top             =   10200
      Width           =   1200
      Begin VB.Image ImgTotal 
         Height          =   405
         Index           =   1
         Left            =   0
         Picture         =   "FrmBrain.frx":0ECA
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image ImgTotal 
         Height          =   405
         Index           =   2
         Left            =   405
         Picture         =   "FrmBrain.frx":1BBC
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image ImgTotal 
         Height          =   405
         Index           =   3
         Left            =   765
         Picture         =   "FrmBrain.frx":28AE
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   440
      Left            =   5280
      ScaleHeight     =   405
      ScaleWidth      =   1170
      TabIndex        =   42
      Top             =   10080
      Width           =   1200
      Begin VB.Image ImgMarks 
         Height          =   405
         Index           =   3
         Left            =   765
         Picture         =   "FrmBrain.frx":35A0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image ImgMarks 
         Height          =   405
         Index           =   2
         Left            =   405
         Picture         =   "FrmBrain.frx":4292
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image ImgMarks 
         Height          =   405
         Index           =   1
         Left            =   0
         Picture         =   "FrmBrain.frx":4F84
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.PictureBox picdraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   8775
      TabIndex        =   0
      Top             =   -120
      Width           =   8775
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   0
         TabIndex        =   48
         Top             =   5800
         Width           =   4695
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1065
            ScaleWidth      =   4425
            TabIndex        =   49
            Top             =   240
            Width           =   4455
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   900
               Index           =   0
               Left            =   0
               ScaleHeight     =   870
               ScaleWidth      =   945
               TabIndex        =   59
               Top             =   840
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   975
               Index           =   1
               Left            =   120
               ScaleHeight     =   945
               ScaleWidth      =   945
               TabIndex        =   58
               Top             =   1320
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   900
               Index           =   2
               Left            =   120
               ScaleHeight     =   900
               ScaleWidth      =   975
               TabIndex        =   57
               Top             =   120
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   975
               Index           =   3
               Left            =   2520
               ScaleHeight     =   945
               ScaleWidth      =   945
               TabIndex        =   56
               Top             =   1320
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   900
               Index           =   4
               Left            =   1200
               ScaleHeight     =   900
               ScaleWidth      =   975
               TabIndex        =   55
               Top             =   120
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   900
               Index           =   5
               Left            =   2280
               ScaleHeight     =   900
               ScaleWidth      =   975
               TabIndex        =   54
               Top             =   120
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   975
               Index           =   6
               Left            =   1440
               ScaleHeight     =   945
               ScaleWidth      =   945
               TabIndex        =   53
               Top             =   1320
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   900
               Index           =   7
               Left            =   3360
               ScaleHeight     =   900
               ScaleWidth      =   975
               TabIndex        =   52
               Top             =   120
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   975
               Index           =   8
               Left            =   120
               ScaleHeight     =   975
               ScaleWidth      =   975
               TabIndex        =   51
               Top             =   4440
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   975
               Index           =   9
               Left            =   1200
               ScaleHeight     =   975
               ScaleWidth      =   975
               TabIndex        =   50
               Top             =   4440
               Visible         =   0   'False
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   0
         TabIndex        =   30
         Top             =   7150
         Width           =   8775
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   3375
            Top             =   240
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   8535
            TabIndex        =   31
            Top             =   350
            Width           =   8535
            Begin VB.Label LblCheck 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FF80&
               Height          =   375
               Left            =   45
               TabIndex        =   32
               Top             =   120
               Width           =   8295
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   4680
         TabIndex        =   33
         Top             =   5800
         Width           =   4095
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   720
            OleObjectBlob   =   "FrmBrain.frx":5C76
            TabIndex        =   60
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton CmdAbout 
            Caption         =   "About"
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
            Left            =   360
            TabIndex        =   41
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton CmdAnswer 
            Caption         =   "Result"
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
            Left            =   2880
            TabIndex        =   40
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton CmdCheck 
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
            Left            =   2040
            TabIndex        =   39
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton CmdBack 
            Caption         =   "Back"
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
            Left            =   1200
            TabIndex        =   38
            Top             =   600
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
            Left            =   360
            TabIndex        =   37
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton CmdQuit 
            Caption         =   "Quit"
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
            Left            =   2880
            TabIndex        =   36
            Top             =   960
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
            Left            =   2040
            TabIndex        =   35
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton CmdChoice 
            Caption         =   "Choice"
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
            Left            =   1200
            TabIndex        =   34
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame FrmFillintheblanks 
         Height          =   5895
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8775
         Begin VB.Frame Frame4 
            Caption         =   "Answer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   3
            Top             =   4800
            Width           =   8535
            Begin VB.Timer Timer3 
               Interval        =   1
               Left            =   7080
               Top             =   480
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblAnswer 
               Height          =   615
               Left            =   120
               OleObjectBlob   =   "FrmBrain.frx":5CF1
               TabIndex        =   4
               Top             =   240
               Width           =   8055
            End
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblPaperNum 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmBrain.frx":5D53
            TabIndex        =   2
            Top             =   120
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Index           =   0
            Left            =   5400
            OleObjectBlob   =   "FrmBrain.frx":5DB8
            TabIndex        =   5
            Top             =   480
            Width           =   2895
         End
         Begin ACTIVESKINLibCtl.SkinLabel Lblquesno 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmBrain.frx":5E49
            TabIndex        =   6
            Top             =   400
            Width           =   2895
         End
         Begin VB.Frame FraFillintheBlanks 
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
            Height          =   4095
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   8535
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2955
               Left            =   5160
               TabIndex        =   16
               Top             =   240
               Width           =   3015
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   5160
               TabIndex        =   15
               Text            =   "Paper 1"
               Top             =   3600
               Width           =   3015
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblQuestion 
               Height          =   2775
               Left            =   120
               OleObjectBlob   =   "FrmBrain.frx":5EB0
               TabIndex        =   17
               Top             =   480
               Width           =   4575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Index           =   3
               Left            =   5160
               OleObjectBlob   =   "FrmBrain.frx":5F23
               TabIndex        =   46
               Top             =   3360
               Width           =   2655
            End
         End
         Begin VB.Frame FraMatch 
            Caption         =   "Match the Columns"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   8415
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   5160
               TabIndex        =   10
               Text            =   "Paper 1"
               Top             =   3720
               Width           =   3015
            End
            Begin VB.ListBox List3 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2790
               Left            =   3480
               TabIndex        =   9
               Top             =   600
               Width           =   4695
            End
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2790
               Left            =   120
               TabIndex        =   8
               Top             =   600
               Width           =   3375
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Left            =   720
               OleObjectBlob   =   "FrmBrain.frx":5F9A
               TabIndex        =   11
               Top             =   360
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   5280
               OleObjectBlob   =   "FrmBrain.frx":6001
               TabIndex        =   12
               Top             =   360
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   495
               Left            =   120
               OleObjectBlob   =   "FrmBrain.frx":6068
               TabIndex        =   13
               Top             =   3480
               Width           =   4575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Index           =   4
               Left            =   5160
               OleObjectBlob   =   "FrmBrain.frx":6187
               TabIndex        =   47
               Top             =   3480
               Width           =   2895
            End
         End
         Begin VB.Frame FraTF 
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
            Height          =   4095
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   8535
            Begin VB.OptionButton OptTrue 
               Caption         =   "True"
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
               Left            =   6600
               TabIndex        =   28
               Top             =   1560
               Width           =   1575
            End
            Begin VB.OptionButton OptFalse 
               Caption         =   "False"
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
               Left            =   6600
               TabIndex        =   27
               Top             =   1800
               Width           =   1575
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   5160
               TabIndex        =   26
               Text            =   "Paper 1"
               Top             =   3600
               Width           =   3015
            End
            Begin ACTIVESKINLibCtl.SkinLabel Lbltf 
               Height          =   2175
               Left            =   240
               OleObjectBlob   =   "FrmBrain.frx":61FE
               TabIndex        =   29
               Top             =   360
               Width           =   6135
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Index           =   1
               Left            =   5160
               OleObjectBlob   =   "FrmBrain.frx":626F
               TabIndex        =   44
               Top             =   3360
               Width           =   2895
            End
         End
         Begin VB.Frame Framultiple 
            Caption         =   "Multiple Choice "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   8535
            Begin VB.Timer Timer2 
               Interval        =   1000
               Left            =   3480
               Top             =   2520
            End
            Begin VB.OptionButton Opt1 
               Caption         =   "option1"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   23
               Top             =   1200
               Width           =   7935
            End
            Begin VB.OptionButton Opt2 
               Caption         =   "Option1"
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
               TabIndex        =   22
               Top             =   1800
               Width           =   7935
            End
            Begin VB.OptionButton Opt3 
               Caption         =   "Option1"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   21
               Top             =   2280
               Width           =   7935
            End
            Begin VB.OptionButton Opt4 
               Caption         =   "Option1"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   20
               Top             =   2880
               Width           =   7935
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   5400
               TabIndex        =   19
               Text            =   "Combo4"
               Top             =   3600
               Width           =   2775
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblMCQues 
               Height          =   855
               Left            =   120
               OleObjectBlob   =   "FrmBrain.frx":62E6
               TabIndex        =   24
               Top             =   240
               Width           =   7815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Index           =   2
               Left            =   5400
               OleObjectBlob   =   "FrmBrain.frx":6359
               TabIndex        =   45
               Top             =   3360
               Width           =   2535
            End
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   8
      Left            =   4800
      Picture         =   "FrmBrain.frx":63D0
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   0
      Left            =   0
      Picture         =   "FrmBrain.frx":7B52
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   1
      Left            =   600
      Picture         =   "FrmBrain.frx":8844
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   2
      Left            =   1200
      Picture         =   "FrmBrain.frx":9536
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   3
      Left            =   1800
      Picture         =   "FrmBrain.frx":A228
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   4
      Left            =   2400
      Picture         =   "FrmBrain.frx":AF1A
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   5
      Left            =   3000
      Picture         =   "FrmBrain.frx":BC0C
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   6
      Left            =   3600
      Picture         =   "FrmBrain.frx":C8FE
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   7
      Left            =   4200
      Picture         =   "FrmBrain.frx":D5F0
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   8
      Left            =   4800
      Picture         =   "FrmBrain.frx":E2E2
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   435
      Index           =   9
      Left            =   5400
      Picture         =   "FrmBrain.frx":EFD4
      Top             =   10680
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   0
      Left            =   0
      Picture         =   "FrmBrain.frx":FCC6
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   1
      Left            =   600
      Picture         =   "FrmBrain.frx":109B8
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   2
      Left            =   1200
      Picture         =   "FrmBrain.frx":116AA
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   3
      Left            =   1800
      Picture         =   "FrmBrain.frx":1239C
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   4
      Left            =   2400
      Picture         =   "FrmBrain.frx":1308E
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   5
      Left            =   3000
      Picture         =   "FrmBrain.frx":13D80
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   6
      Left            =   3600
      Picture         =   "FrmBrain.frx":14A72
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   7
      Left            =   4200
      Picture         =   "FrmBrain.frx":15764
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   8
      Left            =   4800
      Picture         =   "FrmBrain.frx":16456
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   435
      Index           =   9
      Left            =   5400
      Picture         =   "FrmBrain.frx":17148
      Top             =   11160
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   4320
      Picture         =   "FrmBrain.frx":17E3A
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   3720
      Picture         =   "FrmBrain.frx":195BC
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   3120
      Picture         =   "FrmBrain.frx":1B2E6
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   2520
      Picture         =   "FrmBrain.frx":1CAD8
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "FrmBrain.frx":1E802
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   1320
      Picture         =   "FrmBrain.frx":1F4CC
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "FrmBrain.frx":1F90E
      Top             =   9960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "FrmBrain.frx":21638
      Top             =   9960
      Width           =   480
   End
End
Attribute VB_Name = "FrmBrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bldown As Boolean, BlUp As Boolean
Dim i As Integer

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim lpPoint As POINTAPI, mHwnd As Long, lHwnd As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Dim IntIndex As Integer

'Dim t As NOTIFYICONDATA
Dim cnt As Long, sa As String
Dim Resize1 As Boolean

Dim BlResize As Boolean
Dim IntCount As Integer
Dim IntQuesNo As Integer
Dim IntMarksGet As Integer
Dim Marks As String
Dim strans As String

Dim StrData As String
Private myTip As New tooltip ' ToolTip Class
Dim BlArticles As Boolean

Dim IntCnt As Integer
Dim Intlen As Integer
Dim IntMulCount As Integer
Dim IntMarks As Integer
Dim mLeft As Long
Dim mTop As Long

Private Sub CmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub CmdAnswer_Click()
If StrPType = "Matching Columns" Then
   LblAnswer.Caption = StrAnswer(List2.ListIndex + 1)
   LblAnswer.Visible = True
ElseIf FraFillintheBlanks.Visible = True Then
       LblAnswer.Caption = StrAnswer(IntQuesNo)
       LblAnswer.Visible = True
ElseIf TRUE_FALSE = True Then
       LblAnswer.Caption = StrAnswer(IntQuesNo)
       LblAnswer.Visible = True
ElseIf StrPType = "Multiple Choice" Then
       LblAnswer.Caption = StrAnswer(IntQuesNo)
       LblAnswer.Visible = True
End If
End Sub

Private Sub CmdBack_Click()
On Error Resume Next
CmdBack.Enabled = False
CmdCheck.Enabled = True
Dim temp As String

If IntQuesNo > 1 Then
   If StrPType = "Fill in the Blanks" Or StrPType = "Descriptive Questions" Then
        LblAnswer.Visible = False
        LblAnswer.Caption = ""
        Lblquesno.Visible = False
        LblQuestion.Visible = False
        IntQuesNo = IntQuesNo - 1
        Lblquesno = "Question " & IntQuesNo
        LblQuestion.Caption = StrQuestion(IntQuesNo)
        Lblquesno.Visible = True
        LblQuestion.Visible = True
        
        If BlAgent = True Then
            FrmSelection.Agent1.Characters("nitij").Stop
            FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
            FrmSelection.Agent1.Characters("nitij").Speak StrQuestionRead(IntQuesNo)
            FrmSelection.Agent1.Characters("nitij").Play "alert"
        End If
    ElseIf StrPType = "True Or False" Then
           Lblquesno.Visible = False
           Lbltf.Visible = False
           IntQuesNo = IntQuesNo - 1
           Lblquesno = "Question " & IntQuesNo
           Lbltf.Caption = StrQuestion(IntQuesNo)
           Lblquesno.Visible = True
           Lbltf.Visible = True
           If BlAgent = True Then
              FrmSelection.Agent1.Characters("nitij").Stop
              FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
              FrmSelection.Agent1.Characters("nitij").Speak Trim(Lbltf.Caption)
              FrmSelection.Agent1.Characters("nitij").Play "alert"
           End If
           
    ElseIf StrPType = "Multiple Choice" Then
           Dim temp1 As Integer, temp2 As Integer, temp3 As Integer, temp4 As Integer
                      
           LblMCQues.Visible = False
           Opt1.Visible = False
           Opt2.Visible = False
           Opt3.Visible = False
           Opt4.Visible = False
           LblAnswer.Visible = False
           LblAnswer.Caption = ""
           Lblquesno.Visible = False
           IntQuesNo = IntQuesNo - 1
           Lblquesno = "Question " & IntQuesNo
           LblMCQues.Caption = StrQuestion(IntQuesNo)
           
           Opt1.Caption = strmuloption((IntMulCount - 3) - 4)
           Opt2.Caption = strmuloption((IntMulCount - 2) - 4)
           Opt3.Caption = strmuloption((IntMulCount - 1) - 4)
           Opt4.Caption = strmuloption(IntMulCount - 4)
           IntMulCount = IntMulCount - 4
           Lblquesno.Visible = True
           If BlAgent = True Then
              FrmSelection.Agent1.Characters("nitij").Stop
              FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
              FrmSelection.Agent1.Characters("nitij").Speak Trim(StrQuestion(IntQuesNo))
              FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt1.Caption)
              FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt2.Caption)
              FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt3.Caption)
              FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt4.Caption)
              FrmSelection.Agent1.Characters("nitij").Play "alert"
          End If
           If BlAgent = True Then
              FrmSelection.Agent1.Characters("nitij").Stop
              FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
              FrmSelection.Agent1.Characters("nitij").Play "lookdown"
           End If
           LblMCQues.Visible = True
           Opt1.Visible = True
           Opt2.Visible = True
           Opt3.Visible = True
           Opt4.Visible = True
    End If
           
Else
   LblCheck.Caption = ""
   StrData = "No Questions"
   LblCheck.Caption = StrData
   If BlAgent = True Then
      FrmSelection.Agent1.Characters("nitij").Stop
      FrmSelection.Agent1.Characters("nitij").Play "decline"
      FrmSelection.Agent1.Characters("nitij").Speak "No Questions. CLick Next to move further "
      FrmSelection.Agent1.Characters("nitij").Play "alert"
   End If
End If
CmdBack.Enabled = True
End Sub

Private Sub CmdCheck_Click()
On Error Resume Next
Dim temp As Integer
Dim IntTempMarks As Integer
Dim strans As String

If MULTIPLE_CHOICE = True Then
   If Opt1.Value = True Then
      strans = Trim(Opt1.Caption)
   ElseIf Opt2.Value = True Then
          strans = Trim(Opt2.Caption)
   ElseIf Opt3.Value = True Then
          strans = Trim(Opt3.Caption)
   ElseIf Opt4.Value = True Then
          strans = Trim(Opt4.Caption)
   End If
   If UCase(strans) = UCase(StrAnswer(IntQuesNo)) Then
      LblCheck.Caption = ""
      StrData = "Correct"
      If BlAgent = True Then
         ReadCorrect
      End If
      LblCheck.Caption = StrData
   Else
      LblCheck.Caption = ""
      StrData = "Wrong"
      If BlAgent = True Then
         ReadWrong
      End If
         LblCheck.Caption = StrData
   End If
 
ElseIf TRUE_FALSE = True Then
       If OptTrue.Value = True Then
          strans = "True"
       ElseIf OptFalse.Value = True Then
              strans = "False"
       End If
      If UCase(strans) = UCase(Trim(StrAnswer(IntQuesNo))) Then
         LblCheck.Caption = ""
         StrData = "Correct"
         ReadCorrect
         LblCheck.Caption = StrData
      Else
         LblCheck.Caption = ""
         StrData = "Wrong"
         ReadWrong
         LblCheck.Caption = StrData
      End If
ElseIf MATCHING_COLUMNS = True Then
       IntQuesNo = List2.ListIndex + 1
       If UCase(Trim(List3.Text)) = UCase(Trim(StrAnswer(IntQuesNo))) Then
          LblCheck.Caption = ""
          StrData = "Correct"
          
          IntMarks = IntMarks + 1
          If BlAgent = True Then
             FrmSelection.Agent1.Characters("nitij").Play "surprised"
             FrmSelection.Agent1.Characters("nitij").Speak "Correct Answer. Go ahead"
          End If
          LblCheck.Caption = StrData
       Else
          LblCheck.Caption = ""
          StrData = "Wrong"
          
          If BlAgent = True Then
             FrmSelection.Agent1.Characters("nitij").Play "suggest"
             FrmSelection.Agent1.Characters("nitij").Speak "Sorry Wrong Answer."
          End If
          LblCheck.Caption = StrData
       End If
Else
   If UCase(List1.Text) = UCase(StrAnswer(IntQuesNo)) Then
      LblCheck.Caption = ""
      StrData = "Correct"
     'If agent is displaying
      If BlAgent = True Then
         ReadCorrect
      End If
      LblCheck.Caption = StrData
   Else
      LblCheck.Caption = ""
      StrData = "Wrong"
       If BlAgent = True Then
         ReadWrong
      End If
      LblCheck.Caption = StrData
   End If
End If
If Not MATCHING_COLUMNS = True Then
CmdCheck.Enabled = False
End If

End Sub

Private Sub CmdChoice_Click()

'MESSAGE "This will finish your current paper session. Are you willing to proceed?", YesNoonly, "BRAIN: Information"
'If BLMessage = True Then
   Unload Me
   If BlAgent = True Then
      FrmSelection.Agent1.Characters("nitij").StopAll
   End If
   BRAIN_SHOW = False
   FrmSelection.Show
'Else
'   FrmBrain.Show
'End If
End Sub

Private Sub CmdHelp_Click()
FrmHelp.FraBrain.Visible = True
FrmHelp.Show vbModal
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
CmdNext.Enabled = False
CmdCheck.Enabled = True
Dim temp As String, StrPicNum As String
Dim IntCnt As Integer, Intlen As Integer

LblCheck.Caption = ""
LblCheck.Caption = "BRAIN: A New Evolution"

If IntQuesNo < 10 Then
   If StrPType = "Fill in the Blanks" Or StrPType = "Descriptive Questions" Then
      common 'function that will perform common functions
      LblQuestion.Caption = StrQuestion(IntQuesNo)
      Lblquesno.Visible = True
      LblQuestion.Visible = True
      If BlAgent = True Then
         FrmSelection.Agent1.Characters("nitij").Stop
         FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
         FrmSelection.Agent1.Characters("nitij").Speak StrQuestionRead(IntQuesNo)
         FrmSelection.Agent1.Characters("nitij").Play "alert"
      End If
   ElseIf StrPType = "True Or False" Then
          Lblquesno.Visible = False
          Lbltf.Visible = False
          common 'function that will perform common functions
          Lbltf.Caption = StrQuestion(IntQuesNo)
          Lblquesno.Visible = True
          Lbltf.Visible = True
          If BlAgent = True Then
             FrmSelection.Agent1.Characters("nitij").Stop
             FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
             FrmSelection.Agent1.Characters("nitij").Speak Trim(Lbltf.Caption)
             FrmSelection.Agent1.Characters("nitij").Play "alert"
          End If
   ElseIf StrPType = "Multiple Choice" Then
          LblMCQues.Visible = False
          Opt1.Visible = False
          Opt2.Visible = False
          Opt3.Visible = False
          Opt4.Visible = False
          common 'function that will perform common functions
          IntMulCount = IntMulCount + 1
          Lblquesno = "Question " & IntQuesNo
          LblMCQues.Caption = StrQuestion(IntQuesNo)
          Opt1.Caption = strmuloption(IntMulCount)
          IntMulCount = IntMulCount + 1
          Opt2.Caption = strmuloption(IntMulCount)
          IntMulCount = IntMulCount + 1
          Opt3.Caption = strmuloption(IntMulCount)
          IntMulCount = IntMulCount + 1
          Opt4.Caption = strmuloption(IntMulCount)
          Lblquesno.Visible = True
          If BlAgent = True Then
             FrmSelection.Agent1.Characters("nitij").Stop
             FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
             FrmSelection.Agent1.Characters("nitij").Speak Trim(StrQuestion(IntQuesNo))
             FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt1.Caption)
             FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt2.Caption)
             FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt3.Caption)
             FrmSelection.Agent1.Characters("nitij").Speak Trim(Opt4.Caption)
             FrmSelection.Agent1.Characters("nitij").Play "alert"
          End If
          LblMCQues.Visible = True
          Opt1.Visible = True
          Opt2.Visible = True
          Opt3.Visible = True
          Opt4.Visible = True
   End If
           
Else
   LblCheck.Caption = ""
   StrData = "No Questions"
   LblCheck.Caption = StrData
   If BlAgent = True Then
      FrmSelection.Agent1.Characters("nitij").Stop
      FrmSelection.Agent1.Characters("nitij").Play "decline"
      FrmSelection.Agent1.Characters("nitij").Speak "No questions. click back to go previous or change the paper"
      FrmSelection.Agent1.Characters("nitij").Play "alert"
   End If
   
End If
CmdNext.Enabled = True
End Sub

Private Sub CmdQuit_Click()
Me.Hide
End
End Sub

Private Sub Combo1_Click()
    Combo1.Refresh
    LblPaperNum.Visible = False
    LblPaperNum.Caption = Combo1.Text
    LblPaperNum.Visible = True
    
    PaperNum = Val(Trim(Right(Combo1.Text, 2)))
    IntQuesNo = 1
    LblCheck.Caption = USERNAME & " " & USEREMAIL
    CmdCheck.Enabled = True
    FrmPrepare.Show
End Sub

Private Sub Combo2_Click()
LblPaperNum.Visible = False
List2.Visible = False
List3.Visible = False

LblPaperNum.Caption = Combo2.Text
List2.Visible = True
List3.Visible = True

LblPaperNum.Visible = True
PaperNum = Val(Trim(Right(Combo2.Text, 2)))
IntQuesNo = 1
LblCheck.Caption = USERNAME & " " & USEREMAIL
LblPaperNum.Visible = True
FrmPrepare.Show
End Sub
Private Sub Combo3_Click()
    
    Lbltf.Visible = False
    LblPaperNum.Visible = False
    LblPaperNum.Caption = Combo3.Text
    LblPaperNum.Visible = True
    
    PaperNum = Val(Trim(Right(Combo3.Text, 2)))
    IntQuesNo = 1
    LblCheck.Caption = USERNAME & " " & USEREMAIL
    CmdCheck.Enabled = True
    FrmPrepare.Show
    Lbltf.Visible = True

End Sub

Private Sub Combo4_Click()
    IntMulCount = 4
    LblMCQues.Visible = False
    Opt1.Visible = False
    Opt2.Visible = False
    Opt3.Visible = False
    Opt4.Visible = False
    
    Lbltf.Visible = False
    LblPaperNum.Visible = False
    LblPaperNum.Caption = Combo4.Text
    LblPaperNum.Visible = True
    
    PaperNum = Val(Trim(Right(Combo4.Text, 2)))
    IntQuesNo = 1
    LblCheck.Caption = USERNAME & " " & USEREMAIL
    CmdCheck.Enabled = True
    FrmPrepare.Show
    LblMCQues.Visible = True
    Opt1.Visible = True
    Opt2.Visible = True
    Opt3.Visible = True
    Opt4.Visible = True
    
    Opt1.Refresh
    Opt2.Refresh
    Opt3.Refresh
    Opt4.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
FrmSelection.Skin1.ApplySkin Me.hwnd
IntMulCount = 4

BlFirstTime = True
Bldown = True

LblCheck.Font.Name = "Tahoma"
LblCheck.Refresh
LblAnswer.Font.Name = "Tahoma"

If StrPType = "Descriptive Questions" Then
   CmdCheck.Enabled = False
Else
   CmdCheck.Enabled = True
End If

Dim temp As Integer
Dim IntCnt As Integer

Timer3_Timer

Dim PAPERNOIT As New ITPaperCount
Dim StroptionNumIT As New ITPaperCount

Dim PAPERNOPC As New PCPaperCount
Dim StroptionNumpc As New PCPaperCount

Dim PAPERNOIWPD As New IWPDPaperCount
Dim StroptionNumIWPD As New IWPDPaperCount

Dim PAPERNOC As New CPaperCount
Dim StroptionNumC As New CPaperCount

Dim PAPERNOBS As New BSPaperCount
Dim StroptionNumBS As New BSPaperCount

Dim PAPERNODBMS As New DBMSPaperCount
Dim StroptionNumDBMS As New DBMSPaperCount

Dim PAPERNOSAD As New SADPaperCount
Dim StroptionNumSAD As New SADPaperCount

Dim PAPERNODCN As New DCNPaperCount
Dim StroptionNumDCN As New DCNPaperCount

Dim PAPERNOCG As New CGPaperCount
Dim StroptionNumCG As New CGPaperCount

Dim PAPERNODS As New DSPaperCount
Dim StroptionNumDS As New DSPaperCount

Dim PAPERNOUnix As New UnixPaperCount
Dim StroptionNumUnix As New UnixPaperCount

Dim PAPERNOCPP As New CPPPaperCount
Dim StroptionNumCPP As New CPPPaperCount

Dim PAPERNOCO As New COPaperCount
Dim StroptionNumCO As New COPaperCount

LblAnswer.Font.Name = "Balls on the rampage"
LblAnswer.Font.Size = 10

IntCount = 1
LblAnswer.Caption = ""

LblCheck.Caption = ""
StrData = "Krayknot"
LblCheck.Caption = StrData

'****************************************************
'Paper setting if the paper is Information Technology
'****************************************************
If StrPName = "Information Technology" Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOIT.IT_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOIT.IT_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOIT.IT_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOIT.IT_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOIT.IT_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If
   
'****************************************************
'Paper setting if the paper is PC Technology
'****************************************************
 ElseIf StrPName = "Personal Computing Technology" Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings

        For IntCnt = 1 To PAPERNOPC.PC_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt

   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOPC.PC_MTC_PaperNum 'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt

  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOPC.PC_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt

   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOPC.PC_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOPC.PC_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt

   End If
  
'****************************************************
'Paper setting if the paper is Intenet & Web Design
'****************************************************
 ElseIf StrPName = "Internet and Web Design" Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOIWPD.IWPD_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt

   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOIWPD.IWPD_MTC_PaperNum 'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt

  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOIWPD.IWPD_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt

   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOIWPD.IWPD_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOIWPD.IWPD_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt

   End If

'****************************************************
'Paper setting if the paper is C Language
'****************************************************
 ElseIf StrPName = "Programming Through C Language" Then 'If Information Technology
       If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOC.C_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOC.C_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOC.C_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOC.C_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOC.C_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If
   
'****************************************************
'Paper setting if the paper is Business Systems
'****************************************************
 ElseIf StrPName = "Business Systems" Then 'If Information Technology
If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOBS.BS_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOBS.BS_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOBS.BS_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOBS.BS_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOBS.BS_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
   End If
  
'****************************************************
'Paper setting if the paper is Database Management Systems
'****************************************************
ElseIf StrPName = "Database Management Systems" Then 'If Database Management Systems
If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNODBMS.DBMS_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODBMS.DBMS_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNODBMS.DBMS_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODBMS.DBMS_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODBMS.DBMS_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If
  
'****************************************************
'Paper setting if the paper is System Analysis and Design and MIS
'****************************************************
 ElseIf StrPName = "System Analysis and Design and MIS" Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings

        For IntCnt = 1 To PAPERNOSAD.SAD_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt

   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOSAD.SAD_MTC_PaperNum 'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt

  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOSAD.SAD_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt

   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOSAD.SAD_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOSAD.SAD_MC_PaperNum     'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
   End If
  
'****************************************************
'Paper setting if the paper is Data Communications and Networking
'****************************************************
 ElseIf StrPName = "Data Communications and Networking" Then 'If Information Technology
 If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNODCN.DCN_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODCN.DCN_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNODCN.DCN_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODCN.DCN_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODCN.DCN_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If
 
 '****************************************************
'Paper setting if the paper is Computer Graphics
'****************************************************
 ElseIf StrPName = "Computer Graphics" Then 'If Information Technology
If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOCG.CG_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOCG.CG_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOCG.CG_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOCG.CG_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOCG.CG_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If
        
'****************************************************
'Paper setting if the paper is Data Structure
'****************************************************
 ElseIf StrPName = "Data Structure" Then 'If Information Technology
If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNODS.DS_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNODS.DS_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNODS.DS_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODS.DS_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNODS.DS_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If

'********************************************************
'Paper setting if the paper is Unix and Shell programming
'********************************************************
ElseIf StrPName = "UNIX and Shell Programming" Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOUnix.UNIX_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOUnix.UNIX_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOUnix.UNIX_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOUnix.UNIX_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOUnix.UNIX_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
          
   End If
        
'********************************************************
'Paper setting if the paper is Programming and C++
'********************************************************
ElseIf StrPName = "Programming and C++" Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        
        For IntCnt = 1 To PAPERNOCPP.CPP_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt
        
   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings
          
          For IntCnt = 1 To PAPERNOCPP.CPP_MTC_PaperNum  'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt
  
  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOCPP.CPP_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOCPP.CPP_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default 'default settings
          For IntCnt = 1 To PAPERNOCPP.CPP_MC_PaperNum    'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt
   End If
   
'**********************************
'Paper setting if the paper is COSS
'**********************************
 ElseIf StrPName = "C. O. S. S." Then 'If Information Technology
   If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings

        For IntCnt = 1 To PAPERNOCO.CO_FB_PaperNum 'Adding the options in the combo box
            Combo1.AddItem "Paper " & IntCnt
        Next IntCnt

   ElseIf StrPType = "Matching Columns" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOCO.CO_MTC_PaperNum 'Adding the options in the combo box
              Combo2.AddItem "Paper " & IntCnt
          Next IntCnt

  ElseIf StrPType = "True Or False" Then
          prepare_default 'default settings

          For IntCnt = 1 To PAPERNOCO.CO_TF_PaperNum     'Adding the options in the combo box
              Combo3.AddItem "Paper " & IntCnt
          Next IntCnt

   ElseIf StrPType = "Descriptive Questions" Then
          prepare_default
          For IntCnt = 1 To PAPERNOCO.CO_QUES_PaperNum   'Adding the options in the combo box
              Combo1.AddItem "Paper " & IntCnt
          Next IntCnt
          
   ElseIf StrPType = "Multiple Choice" Then
          prepare_default
          For IntCnt = 1 To PAPERNOCO.CO_MC_PaperNum     'Adding the options in the combo box
              Combo4.AddItem "Paper " & IntCnt
          Next IntCnt

   End If
End If

LblCheck.Font.Name = "Tahoma"
LblCheck.Refresh

IntQuesNo = 1
If BlAgent = True Then
   FrmSelection.Agent1.Characters("nitij").Top = (SELECTION_TOP / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
   FrmSelection.Agent1.Characters("nitij").Left = (SELECTION_LEFT / Screen.TwipsPerPixelX)
End If

Intlen = Len(IntMarksTotal)
End Sub

Private Sub List2_Click()
With myTip
    .ttTitle = "Brain"
    .ttStyle = TTBalloon
    .ttIcon = TTIconInfo
    .ttForeColor = vbBlack
    .ttCentered = False
    .ttBackColor = &HC0C0C0
    .msgTipText = List2.Text
Set .ParentControl = List2
    .Create
End With
    If BlAgent = True Then
       FrmSelection.Agent1.Characters("nitij").Stop
       FrmSelection.Agent1.Characters("nitij").Speak List2.Text
    End If
Lblquesno.Caption = "Question " & List2.ListIndex + 1
End Sub

Private Sub List3_Click()
With myTip
    .ttTitle = "Brain"
    .ttStyle = TTBalloon
    .ttIcon = TTIconInfo
    .ttForeColor = vbBlack
    .ttCentered = False
    .ttBackColor = &HC0C0C0
    .msgTipText = List3.Text
Set .ParentControl = List3
    .Create
End With
If BlAgent = True Then
   FrmSelection.Agent1.Characters("nitij").Stop
   FrmSelection.Agent1.Characters("nitij").Speak List3.Text
End If
End Sub

Private Sub List3_DblClick()
IntQuesNo = List2.ListIndex + 1
If Trim(List3.Text) = Trim(StrAnswer(IntQuesNo)) Then
   LblCheck.Caption = ""
   StrData = "Correct"
   If BlAgent = True Then
      FrmSelection.Agent1.Characters("nitij").Play "surprised"
      FrmSelection.Agent1.Characters("nitij").Speak "Correct Answer. Go ahead"
   End If
     
   LblCheck.Caption = StrData
Else
   LblCheck.Caption = ""
   StrData = "Wrong"
   If BlAgent = True Then
      FrmSelection.Agent1.Characters("nitij").Play "suggest"
      FrmSelection.Agent1.Characters("nitij").Speak "Sorry Wrong Answer."
   End If
   
   LblCheck.Caption = StrData
End If
End Sub

Private Sub OptFalse_Click()
If OptFalse.Value = True Then strans = "False"
End Sub

Private Sub OptTrue_Click()
If OptTrue.Value = True Then strans = "True"
End Sub

Private Sub Picture9_Click(Index As Integer)
If Index = 2 Then
'   FrmBooks.Show vbModal
   MESSAGE "Books are not available in this version.", OkOnly, "Information"
ElseIf Index = 4 Then
       FrmHelp.FraBrain.Visible = True
       FrmHelp.Show vbModal
ElseIf Index = 5 Then
       frmAbout.Show vbModal
ElseIf Index = 7 Then
       frmtest.Show vbModal
End If
End Sub

Private Sub Picture9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
IntIndex = Index
End Sub

Private Sub Timer1_Timer()
LblCheck.Caption = LblCheck.Caption & Mid(StrData, IntCount, 1)
If Len(LblCheck.Caption) = Len(StrData) Then
Timer1.Enabled = False
IntCount = 1
Else
 IntCount = IntCount + 1
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If mLeft <> (Me.Left / Screen.TwipsPerPixelX) Then
  FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelX, _
  (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10, 50
  mTop = (Me.Top / Screen.TwipsPerPixelY)
  mLeft = (Me.Left / Screen.TwipsPerPixelX)
End If

If mTop <> (Me.Top / Screen.TwipsPerPixelY) Then
  FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelX, (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10, 50
  mTop = (Me.Top / Screen.TwipsPerPixelY)
  mLeft = (Me.Left / Screen.TwipsPerPixelX)
End If
End Sub

Private Sub Timer3_Timer()
Dim i
GetCursorPos lpPoint
mHwnd = WindowFromPoint(lpPoint.X, lpPoint.Y)

If mHwnd = Picture9(IntIndex).hwnd Then
   Picture9(IntIndex).Cls
   Picture9(IntIndex).PaintPicture Image2(IntIndex).Picture, 250, 250, 600, 600
   Picture9(IntIndex).CurrentX = 230
   Picture9(IntIndex).CurrentY = 50
   Picture9(IntIndex).Font.Name = "Tahoma"
   Picture9(IntIndex).FontBold = True
   If IntIndex = 0 Then
          Picture9(IntIndex).Print "Articles"
   ElseIf IntIndex = 2 Then
          Picture9(IntIndex).Print "Books"
   ElseIf IntIndex = 4 Then
          Picture9(IntIndex).Print "  Help"
   ElseIf IntIndex = 5 Then
          Picture9(IntIndex).Print " About"
   ElseIf IntIndex = 7 Then
          Picture9(IntIndex).Print "   Test"
   ElseIf IntIndex = 8 Then
          Picture9(IntIndex).Print " Report"
   End If
Else
   For i = 0 To 8
     If mHwnd <> Picture9(i).hwnd Then
     Picture9(i).Cls
     Picture9(i).PaintPicture Image2(i).Picture, 250, 50, 500, 500
     Picture9(i).CurrentX = 220
     Picture9(i).CurrentY = 600
     Picture9(i).FontBold = False
     Picture9(i).Font.Name = "Tahoma"
     End If
   Next i
   
    Picture9(0).Print "Articles"
    Picture9(1).Print "Schedule"
    Picture9(2).Print "Books"
    Picture9(3).Print " Details"
    Picture9(4).Print "  Help"
    Picture9(5).Print " About"
    Picture9(6).Print "  Fonts"
    Picture9(7).Print "  Test"
    Picture9(8).Print " Report"
End If
End Sub

Function common()
LblAnswer.Visible = False
LblAnswer.Caption = ""
Lblquesno.Visible = False
LblQuestion.Visible = False
IntQuesNo = IntQuesNo + 1
Lblquesno = "Question " & IntQuesNo
End Function

Function ReadCorrect()
FrmSelection.Agent1.Characters("nitij").Stop
FrmSelection.Agent1.Characters("nitij").Play "surprised"
FrmSelection.Agent1.Characters("nitij").Speak "Correct Answer. You got it"
FrmSelection.Agent1.Characters("nitij").Play "alert"
End Function

Function ReadWrong()
FrmSelection.Agent1.Characters("nitij").Stop
FrmSelection.Agent1.Characters("nitij").Play "uncertain"
FrmSelection.Agent1.Characters("nitij").Speak "Sorry Wrong Answer"
FrmSelection.Agent1.Characters("nitij").Play "alert"
End Function

Function prepare_default()
If StrPType = "Fill in the Blanks" Then
        Combo1.Clear
        Combo1.Text = "Paper1 "
                
   ElseIf StrPType = "Matching Columns" Then
          Combo2.Clear
          Combo2.Text = "Paper1 "
          
  ElseIf StrPType = "True Or False" Then
          Combo3.Clear
          Combo3.Text = "Paper1 "
          
   ElseIf StrPType = "Descriptive Questions" Then
          Combo1.Clear
          Combo1.Text = "Paper1 "
          
   ElseIf StrPType = "Multiple Choice" Then
          Combo4.Clear
          Combo4.Text = "Paper1 "
          
End If
End Function

