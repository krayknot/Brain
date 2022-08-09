VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{C3CBD80D-C8D1-11D2-9F8E-0080C7CE5CDC}#4.1#0"; "ActCndy2.ocx"
Begin VB.Form frmsteps 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ActiveCandy.CandyCommand cmdnext 
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      ToolTipText     =   "Click here to move to paper list."
      Top             =   6095
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackPicture     =   5
      Caption         =   "Next>>"
      ForeColor       =   65280
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
      PictureMode     =   1
      PictureGoDown   =   -1  'True
      ShortCutKey     =   "N"
      GradientGreenDividend=   20
      GradientBlueDividend=   150
      UseSound        =   -1  'True
   End
   Begin ActiveCandy.CandyCommand CandyCommand1 
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      ToolTipText     =   "Click here to move to paper list."
      Top             =   6095
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackPicture     =   5
      Caption         =   "Help"
      ForeColor       =   65280
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
      PictureMode     =   1
      PictureGoDown   =   -1  'True
      ShortCutKey     =   "N"
      GradientGreenDividend=   20
      GradientBlueDividend=   150
      UseSound        =   -1  'True
   End
   Begin ActiveCandy.CandyCommand cmdquit 
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      ToolTipText     =   "Click here to move to paper list."
      Top             =   6550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackPicture     =   5
      Caption         =   "Quit"
      ForeColor       =   65280
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
      PictureMode     =   1
      PictureGoDown   =   -1  'True
      ShortCutKey     =   "N"
      GradientGreenDividend=   20
      GradientBlueDividend=   150
      UseSound        =   -1  'True
   End
   Begin ActiveCandy.CandyCommand CandyCommand3 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      ToolTipText     =   "Click here to move to paper list."
      Top             =   6550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackPicture     =   5
      Caption         =   "Options"
      ForeColor       =   65280
      FontName        =   "Times New Roman"
      FontSize        =   9.75
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
      PictureMode     =   1
      PictureGoDown   =   -1  'True
      ShortCutKey     =   "N"
      GradientGreenDividend=   20
      GradientBlueDividend=   150
      UseSound        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmsteps.frx":0000
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
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   4695
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   8760
      Top             =   240
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   5055
      Left            =   0
      MousePointer    =   4  'Icon
      Picture         =   "frmsteps.frx":0094
      Stretch         =   -1  'True
      Top             =   960
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   7095
      Left            =   0
      Picture         =   "frmsteps.frx":8093
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmsteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CandyCommand2_Click()

End Sub

Private Sub cmdok_Click()

End Sub

Private Sub CmdNext_Click()
Me.Hide
' Agent1.Characters("manu").Top = 400
' Agent1.Characters("manu").Left = 300
frmstep1.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub Form_Load()
 Agent1.Characters.Load "Manu", App.Path & "\chars\merlin.acs"
 Agent1.Characters("manu").Balloon.Style = 2
 Agent1.Characters("manu").Top = 415
 Agent1.Characters("manu").Left = 180
 Agent1.Characters("manu").Show
 Agent1.Characters("manu").Play "greet"
 Agent1.Characters("manu").Play "acknowledge"
 Agent1.Characters("manu").SoundEffectsOn = True
 Agent1.Characters("manu").Speak "Welcome to Brain. Before starting you will have to follow some steps. Press Next to continue"
 
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.Refresh
'Agent1.Characters("manu").StopAll
''Agent1.Characters("manu").Speak "Please wait"
'Agent1.Characters("manu").MoveTo cmdnext.Top, cmdnext.Left
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox "left: " & lblpos.Left & "     " & "top : " & lblpos.Top

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
 
