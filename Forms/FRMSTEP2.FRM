VERSION 5.00
Object = "{D8A9DA2D-AB82-4962-B789-727EBE641D59}#1.0#0"; "cpvButton.ocx"
Begin VB.Form FrmStep2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      ScaleHeight     =   465
      ScaleWidth      =   4185
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fill in the Blanks"
         BeginProperty Font 
            Name            =   "12SaruYellowFog"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   4200
   End
   Begin Button2.cpvButton CmdQuit 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   5520
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
   Begin Button2.cpvButton cpvButton1 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   5520
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
   Begin Button2.cpvButton CmdBack 
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   5520
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
      Left            =   4800
      TabIndex        =   4
      Top             =   5520
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
   Begin Button2.cpvButton cpvButton2 
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   5520
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
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the name of the paper you want to solve"
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
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Step 2"
      BeginProperty Font 
         Name            =   "BalletEngraved"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   0
      Picture         =   "frmstep2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8910
   End
End
Attribute VB_Name = "frmstep2"
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

Private Sub CmdBack_Click()
Me.Hide
frmstep1.Show
End Sub

Private Sub cmdnext_Click()
 tname = "Processing"
 Label6.Caption = ""
 Timer1.Enabled = True
 
 If optfb.Value = True Then
  paper_type = "Fill in the Blanks"
 Else
 If optmc.Value = True Then
  paper_type = "Matching Columns"
 Else
 If optmultiple.Value = True Then
  paper_type = "Multiple Choice"
 Else
 If optdq.Value = True Then
  paper_type = "Descriptive Questions"
 Else
 If opttf.Value = True Then
  paper_type = "True or False"
 Else
 If optextra.Value = True Then
  paper_type = "Extra Questions"
 Else
 If optq.Value = True Then
  paper_type = "Questions"
 Else
 If optprograms.Value = True Then
  paper_type = "Programs"
 Else
 If optpws.Value = True Then
  paper_type = "Programs with Solutions"
 Else
 If optproject.Value = True Then
  paper_type = "Project Samples"
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 
 Me.Hide
 frmmain.Show
  
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
 frmsteps.Agent1.Characters("manu").StopAll
 frmsteps.Agent1.Characters("manu").MoveTo 550, 400, 1
 frmsteps.Agent1.Characters("manu").Play "gestureright"
 frmsteps.Agent1.Characters("manu").SoundEffectsOn = True
 frmsteps.Agent1.Characters("manu").Speak "Select the paper category from the list mentioned here."
 frmsteps.Agent1.Characters("manu").Play "blink"
 i = 1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
End Sub

Private Sub Option1_Click()

End Sub

Private Sub optdq_Click()
 tname = "Descriptive Questions"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optextra_Click()
 tname = "Extra Questions"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optfb_Click()
 tname = "Fill in the Blanks"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optmc_Click()
 tname = "Matching Columns"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optmultiple_Click()
 tname = "Multiple Choice"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optprograms_Click()
 tname = "Program with Solutions"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optproject_Click()
 tname = "Project Samples"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optpws_Click()
 tname = "Program without Solutions"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub optq_Click()
 tname = "Questions without Answers"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub opttf_Click()
 tname = "True or False"
 Label6.Caption = ""
 Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If tname <> "" Then
Label6.Caption = Label6.Caption & Mid(tname, i, 1)
If Len(Label6.Caption) = Len(tname) Then
Timer1.Enabled = False
i = 1
'cmdgenerate.Enabled = True
'cmdlapse.Enabled = True
Else
 i = i + 1
End If
End If
End Sub
