VERSION 5.00
Object = "{D8A9DA2D-AB82-4962-B789-727EBE641D59}#1.0#0"; "cpvButton.ocx"
Begin VB.Form FrmSelection 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   4680
      TabIndex        =   22
      Top             =   1560
      Width           =   3855
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Project samples"
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
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   3360
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Programs without solutions"
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
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   3000
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Programs with Solutions"
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
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Questions without Answers"
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
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Extra Questions with Solutions"
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
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "True Or False"
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
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Descriptive Questions"
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
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Multiple Choice"
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
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Matching Columns"
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00000000&
         Caption         =   "Fill in the Blanks"
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
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   3735
      End
   End
   Begin Button2.cpvButton CmdQuit 
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "About"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Data Structure"
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
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Computer Graphics"
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
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Data Communications and Networking"
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
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "System Analysis and Design"
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
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Database Management Systems"
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
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Business Systems"
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
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Programming Through C Language"
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
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Internet and Web Design"
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
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Personal Computing Technology"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton OptSubject 
         BackColor       =   &H00000000&
         Caption         =   "Information Technology"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   3735
      End
   End
   Begin Button2.cpvButton cpvButton2 
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Quit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton CmdBa 
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton CmdNe 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton cpvButton1 
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Create"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton cpvButton3 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Character"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton cpvButton4 
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Skin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton CmdBack 
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "<<Back"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Button2.cpvButton CmdNext 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   5625
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonColor     =   16744576
      Caption         =   "Next>>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the category you want to solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Subject you want to solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BRAIN 2"
      BeginProperty Font 
         Name            =   "BalletEngraved"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AFE054&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   0
      Picture         =   "frmstep1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "FrmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tname As String
Dim i As Integer

Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Private Declare Sub ReleaseCapture Lib "User32" ()

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2

Private Sub CandyCommand3_Click()

End Sub

Private Sub CmdNext_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Sets the default paper values for use in the Frmprepare
 StrPaperName = OptSubject(Index).Caption
 StrPaperType = OptCategory(Index).Caption
 
End Sub

Private Sub OptCategory_Click(Index As Integer)
StrPaperType = OptCategory(Index).Caption
FrmPrepare.LblCategory = StrPaperType
End Sub

Private Sub OptSubject_Click(Index As Integer)
StrPaperName = OptSubject(Index).Caption
FrmPrepare.LblSubject = StrPaperName
End Sub
