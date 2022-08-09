VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: : Help"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   ControlBox      =   0   'False
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1080
      Top             =   5400
   End
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
      Left            =   3360
      TabIndex        =   16
      Top             =   5280
      Width           =   855
   End
   Begin VB.Frame FraBrain 
      Height          =   5175
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         ScaleHeight     =   4815
         ScaleWidth      =   3975
         TabIndex        =   32
         Top             =   240
         Width           =   3975
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   " Will exit the BRAIN2."
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
            Index           =   18
            Left            =   840
            TabIndex        =   50
            Top             =   3960
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Quit "
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
            Left            =   120
            TabIndex        =   49
            Top             =   3960
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Will display help."
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
            Index           =   17
            Left            =   840
            TabIndex        =   48
            Top             =   3720
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Help "
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
            Left            =   120
            TabIndex        =   47
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Will display the Paper Selection dialog box."
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
            Index           =   16
            Left            =   840
            TabIndex        =   46
            Top             =   3480
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Choices "
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
            Left            =   120
            TabIndex        =   45
            Top             =   3480
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Will display the information about the Creator."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   15
            Left            =   840
            TabIndex        =   44
            Top             =   3000
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "About "
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
            Left            =   120
            TabIndex        =   43
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Back button will display the previous question (Deactive in Matching Columns)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   840
            TabIndex        =   42
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Back "
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
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   40
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click on this button to know the answer."
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
            Index           =   10
            Left            =   840
            TabIndex        =   39
            Top             =   2760
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Check "
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
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Check will check the selected answer and display the results."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   840
            TabIndex        =   37
            Top             =   2280
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Next "
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
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Next button will display the next question. (Deactive in Matching Columns)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   840
            TabIndex        =   35
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   $"FrmHelp.frx":0ECA
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "BRAIN"
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
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1935
         End
      End
   End
   Begin VB.Frame FraAgent 
      Height          =   5175
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         ScaleHeight     =   4815
         ScaleWidth      =   3975
         TabIndex        =   18
         Top             =   240
         Width           =   3975
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent"
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
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the Agent you want to apply in BRAIN2 by selecting the name from the list. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click on Apply to Apply the settings"
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
            Index           =   11
            Left            =   840
            TabIndex        =   28
            Top             =   1920
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Apply "
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
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Check the Show Balloon button to display or hide the balloon."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   3735
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Click on the Show/Hide button to show or hide the Agent."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click on Preview to see a preview of the Agent"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   840
            TabIndex        =   24
            Top             =   2160
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Preview "
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
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click on Help to gather help."
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
            Index           =   13
            Left            =   840
            TabIndex        =   22
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Help "
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
            Left            =   120
            TabIndex        =   21
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click on Close to close the Agent dialog."
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
            Index           =   14
            Left            =   840
            TabIndex        =   20
            Top             =   2880
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Close "
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
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   615
         End
      End
   End
   Begin VB.Frame FraSelection 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         ScaleHeight     =   4815
         ScaleWidth      =   3975
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Information about the creator of this software."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   840
            TabIndex        =   15
            Top             =   3000
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "About "
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
            Left            =   120
            TabIndex        =   14
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Quit to exit from the software."
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
            Index           =   5
            Left            =   840
            TabIndex        =   13
            Top             =   2760
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Quit "
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
            Left            =   120
            TabIndex        =   12
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click the Help button to gather help."
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
            Index           =   4
            Left            =   840
            TabIndex        =   11
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Help "
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
            Left            =   120
            TabIndex        =   10
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Agent to appear dissapear a animated character around the software that will help you a lot"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   1920
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent "
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
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Skin to change the skin of the software"
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
            Index           =   0
            Left            =   840
            TabIndex        =   7
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Skin "
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
            Left            =   120
            TabIndex        =   6
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Click Next to proceed further."
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
            Left            =   840
            TabIndex        =   5
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"FrmHelp.frx":0F82
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Paper Selection"
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
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bldown As Boolean, BlUp As Boolean
Dim i As Integer
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
Bldown = True
End Sub

Private Sub Label16_Click()

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)
End Sub

Private Sub Timer1_Timer()
If Bldown = True Then
   Me.Top = Me.Top + 2 'Frequency of the movment of the form
   i = i + 1
   If i = 50 Then
      Bldown = False
      BlUp = True
   End If
End If

If BlUp = True Then
   Me.Top = Me.Top - 2
   i = i - 1
   If i = 0 Then
      Bldown = True
      BlUp = False
   End If
End If
End Sub
