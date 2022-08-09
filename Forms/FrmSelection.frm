VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmSelection 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BRAIN:  Paper Selection"
   ClientHeight    =   4470
   ClientLeft      =   5865
   ClientTop       =   2520
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSelection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmSelection.frx":0ECA
      TabIndex        =   24
      Top             =   120
      Width           =   3375
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "FrmSelection.frx":0F59
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   2040
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   3720
      Width           =   7140
      Begin VB.CommandButton CmdAbout 
         Caption         =   "Settings"
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdQuit 
         Caption         =   "Quit"
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdHelp 
         Caption         =   "Help"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdAgent 
         Caption         =   "Agent"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSkin 
         Caption         =   "Skin"
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "Next"
         Height          =   375
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   3960
      OleObjectBlob   =   "FrmSelection.frx":3DC38
      TabIndex        =   25
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   75
      TabIndex        =   13
      Top             =   360
      Width           =   3855
      Begin VB.OptionButton OptSubject 
         Caption         =   "C. O. S. S."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   28
         Top             =   3050
         Width           =   3495
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Programming and C++"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   27
         Top             =   2800
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "UNIX and Shell Programming"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   26
         Top             =   2550
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Data Structure"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   2300
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Computer Graphics"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   3375
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Data Communications and Networking"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   3615
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "System Analysis and Design and MIS"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1500
         Width           =   3615
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Database Management Systems"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Business Systems"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Programming Through C Language"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Internet and Web Design"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Personal Computing Technology"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton OptSubject 
         Caption         =   "Information Technology"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   7
      Top             =   360
      Width           =   3255
      Begin VB.OptionButton OptCategory 
         Caption         =   "True Or False"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   2415
      End
      Begin VB.OptionButton OptCategory 
         Caption         =   "Descriptive Questions"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton OptCategory 
         Caption         =   "Multiple Choice"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton OptCategory 
         Caption         =   "Matching Columns"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton OptCategory 
         Caption         =   "Fill in the Blanks"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   7320
      Top             =   3360
   End
End
Attribute VB_Name = "FrmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IntCount As Integer
Dim mLeft  As Integer, mTop As Integer
Dim BlUp As Boolean, Bldown As Boolean
Dim i As Integer

Private Sub CmdAbout_Click()
FrmPostInstall.Show vbModal
End Sub

Private Sub CmdAgent_Click()
FrmAgent.Show vbModal
End Sub


Private Sub CmdHelp_Click()
If CmdHelp.Caption = "Help" Then

    If BlAgent = False Then
       FrmHelp.FraSelection.Visible = True
       FrmHelp.Show
    Else
       If BlBalloon = True Then
          FrmSelection.Agent1.Characters("nitij").Balloon.FontName = "Tahoma"
          FrmSelection.Agent1.Characters("nitij").Balloon.FontSize = 8
          FrmSelection.Agent1.Characters("nitij").Balloon.Style = 3
       ElseIf BlBalloon = False Then
            FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
       End If
       CmdHelp.Caption = "Stop"
       BLSpeak = True
       FrmSelection.Agent1.Characters("nitij").StopAll
       FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 150, Me.Top / Screen.TwipsPerPixelY + 100
       FrmSelection.Agent1.Characters("nitij").Play "gestureright"
       FrmSelection.Agent1.Characters("nitij").Speak "Here is the list of the papers. Select the name of the paper you would like to solve"
       FrmSelection.Agent1.Characters("nitij").Play "blink"
    
       FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 450, Me.Top / Screen.TwipsPerPixelY + 100
       FrmSelection.Agent1.Characters("nitij").Play "gestureright"
       FrmSelection.Agent1.Characters("nitij").Speak "Select the type of the paper you would like to solve"
       FrmSelection.Agent1.Characters("nitij").Play "blink"
    
       FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 50, Me.Top / Screen.TwipsPerPixelY + 300
       FrmSelection.Agent1.Characters("nitij").Play "gestureright"
       FrmSelection.Agent1.Characters("nitij").Speak "Click on the next button to proceed further"
       FrmSelection.Agent1.Characters("nitij").Play "blink"
    
       FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 120, Me.Top / Screen.TwipsPerPixelY + 300
       FrmSelection.Agent1.Characters("nitij").Play "gestureright"
       FrmSelection.Agent1.Characters("nitij").Speak "Click on the skin button to change the skin of brain"
       FrmSelection.Agent1.Characters("nitij").Play "blink"
    
       FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 170, Me.Top / Screen.TwipsPerPixelY + 300
       FrmSelection.Agent1.Characters("nitij").Play "gestureright"
       FrmSelection.Agent1.Characters("nitij").Speak "Click on the Agent button to display or change the agent character."
       FrmSelection.Agent1.Characters("nitij").Play "blink"
    
    
       FrmSelection.Agent1.Characters("nitij").MoveTo (Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
       CmdHelp.Caption = "Help"
    End If
ElseIf CmdHelp.Caption = "Stop" Then
       FrmSelection.Agent1.Characters("nitij").Stop
       FrmSelection.Agent1.Characters("nitij").MoveTo (Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
End If
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
'Nullify all variables
 TRUE_FALSE = False
 MATCHING_COLUMNS = False
 MULTIPLE_CHOICE = False
 
'As we are providing only one paper in the demo version thus we have set that papaer name
'defaultly in the Strname variable  in the form load event otherwise it should be check first
 For IntCount = 0 To 12 ' Checks for the checked button and places its caption in a variable
     If OptSubject(IntCount).Value = True Then
        StrPName = OptSubject(IntCount).Caption
        Exit For
     End If
 Next IntCount

For IntCount = 0 To 10  ' Checks for the checked button and places its caption in a variable
    If OptCategory(IntCount).Value = True Then
       StrPType = OptCategory(IntCount).Caption
       Exit For
    End If
Next IntCount

If StrPType = "" Or StrPName = "" Then
   If BlAgent = True Then
      If BlBalloon = True Then
         FrmSelection.Agent1.Characters("nitij").Balloon.Style = 1
      Else
         FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
      End If
      
      FrmSelection.Agent1.Characters("nitij").StopAll
      FrmSelection.Agent1.Characters("nitij").Play "decline"
      FrmSelection.Agent1.Characters("nitij").Speak "You did not select the Paper Name or Paper Type."
      FrmSelection.Agent1.Characters("nitij").Play "blink"
   Else
      MESSAGE "You did not select the Paper Name or Paper Type.", OkOnly, "Error"
   End If
   Exit Sub
End If

CmdNext.Refresh
Me.Hide


FrmPrepare.Show

End Sub


Private Sub CmdQuit_Click()
End
End Sub

Private Sub CmdSkin_Click()
FrmSkin.Show vbModal
End Sub


Private Sub Form_GotFocus()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
On Error GoTo ErrorHandler

SELECTION_TOP = Me.Top
SELECTION_LEFT = Me.Left
    
Skin1.LoadSkin GetSetting("BRAIN", "Values", "Skin")
Skin1.ApplySkin Me.hwnd

Bldown = True
BLMessage = False

Timer1.Enabled = True

Dim temp As String
PaperNum = 1

ErrorHandler:   ' Error-handling routine
  Select Case Err.Number
         Case 13
         MESSAGE "Type Mismatch. Try to retreving  values from registry, and not getting positive response. Loading default settings", OkOnly, "Brain: Error"
         DefaultSettings
  End Select
 
End Sub

Private Sub OptCategory_Click(Index As Integer)
StrPType = OptCategory(Index).Caption
End Sub

Private Sub OptSubject_Click(Index As Integer)
Dim IntCt As Integer
For IntCt = 0 To 4
    OptCategory(IntCt).Enabled = True
Next
End Sub

Private Sub Timer1_Timer()
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

Function DefaultSettings()
'Defalut setting in case there is any error
On Error Resume Next
FrmSelection.Skin1.LoadSkin App.Path & "/skins/green.skn"
Skin1.ApplySkin Me.hwnd

Dim temp As String
        
temp = "merlin.acs"
Timer1.Enabled = True
BlAgent = True
FrmSelection.Agent1.Characters.Unload "nitij"
FrmSelection.Agent1.Characters.Load "nitij", temp
FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
FrmSelection.Agent1.Characters("nitij").Show
FrmSelection.Agent1.Characters("nitij").Height = 140
FrmSelection.Agent1.Characters("nitij").Width = 130
FrmSelection.Agent1.Characters("nitij").Top = (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
FrmSelection.Agent1.Characters("nitij").Left = (Me.Left / Screen.TwipsPerPixelX)

If Me.Visible = False Then Me.Visible = True
End Function

