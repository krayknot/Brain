VERSION 5.00
Begin VB.Form FrmAgent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Agent"
   ClientHeight    =   3870
   ClientLeft      =   7320
   ClientTop       =   4635
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "Frmagent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2355
      ScaleWidth      =   4155
      TabIndex        =   13
      Top             =   4080
      Width           =   4215
      Begin VB.Image imgGenie 
         Height          =   855
         Left            =   120
         Picture         =   "Frmagent.frx":0ECA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   735
      End
      Begin VB.Image ImgMerlin 
         Height          =   855
         Left            =   960
         Picture         =   "Frmagent.frx":B874
         Stretch         =   -1  'True
         Top             =   240
         Width           =   735
      End
      Begin VB.Image ImgPeedy 
         Height          =   855
         Left            =   1800
         Picture         =   "Frmagent.frx":15F1E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   735
      End
      Begin VB.Image ImgRobby 
         Height          =   855
         Left            =   2640
         Picture         =   "Frmagent.frx":296F8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgal 
         Height          =   855
         Left            =   3360
         Picture         =   "Frmagent.frx":3B372
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgelectra 
         Height          =   855
         Left            =   120
         Picture         =   "Frmagent.frx":3C84F
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imghanz 
         Height          =   855
         Left            =   960
         Picture         =   "Frmagent.frx":3DEF8
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgozzar 
         Height          =   855
         Left            =   1800
         Picture         =   "Frmagent.frx":3EBD5
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgsanta 
         Height          =   855
         Left            =   2640
         Picture         =   "Frmagent.frx":401D3
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgwartnose 
         Height          =   855
         Left            =   3360
         Picture         =   "Frmagent.frx":41651
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   30
      TabIndex        =   8
      Top             =   2400
      Width           =   4095
      Begin VB.CommandButton CmdShow 
         Caption         =   "Hide Agent"
         Enabled         =   0   'False
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
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkBalloon 
         Caption         =   "Show Balloon"
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
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   30
      TabIndex        =   3
      Top             =   3120
      Width           =   4095
      Begin VB.CommandButton CmdClose 
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
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1920
         Top             =   120
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
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdPreview 
         Caption         =   "Preview"
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
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
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
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2235
         Left            =   0
         ScaleHeight     =   2205
         ScaleWidth      =   1905
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         Begin VB.Image Image1 
            Height          =   2175
            Left            =   0
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No Image to display"
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
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "FrmAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim BlUp As Boolean, Bldown As Boolean
Dim i As Integer
Dim PrChar As String
Dim mLeft, mTop
Dim Char As String

Private Sub ChkBalloon_Click()
On Error Resume Next
If CmdShow.Caption = "Show Agent" Then
   MESSAGE "Agent display is off", OkOnly, "BRAIN: Error"
   Exit Sub
End If

If ChkBalloon.Value = 1 Then
   BlBalloon = True
   FrmSelection.Agent1.Characters("nitij").Balloon.Style = 1
   FrmSelection.Agent1.Characters("nitij").Speak "Now you can see Balloon"
ElseIf ChkBalloon.Value = 0 Then
       BlBalloon = False
       FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
       FrmSelection.Agent1.Characters("nitij").Speak "Balloon display is off."
End If

End Sub

Private Sub CmdApply_Click()
On Error Resume Next
 If BlAgent = True Then
    FrmSelection.Agent1.Characters.Unload "nitij"
    BlAgent = False
 End If
   
 PREVIEW = False
 BlAgent = True
 Char = WindowsDirectory & "\BSystem2\chars\" & List1.Text & ".Acs"
 PrChar = Char
 FrmSelection.Agent1.Characters.Load "nitij", Char
 FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
 FrmSelection.Agent1.Characters("nitij").Height = 140
 FrmSelection.Agent1.Characters("nitij").Width = 130
 FrmSelection.Agent1.Characters("nitij").Top = (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
 FrmSelection.Agent1.Characters("nitij").Left = (Me.Left / Screen.TwipsPerPixelX)
 FrmSelection.Agent1.Characters("nitij").Show
   
'Saves the settings in the registry for the future use
 SaveSetting "BRAIN", "Values", "Character_Name", File1.filename
 SaveSetting "BRAIN", "Values", "AGENT_ON", True
 
 CmdApply.Enabled = False
 CmdShow.Caption = "Hide Agent"
 CmdShow.Refresh
End Sub

Private Sub cmdclose_Click()
On Error Resume Next
If PREVIEW = True Then
   FrmSelection.Agent1.Characters("nitij").StopAll
   FrmSelection.Agent1.Characters.Unload "nitij"
   PREVIEW = False
End If
   Unload Me
   FrmSelection.Agent1.Characters("nitij").StopAll
End Sub

Private Sub CmdHelp_Click()
On Error Resume Next
If PREVIEW <> True Or BlAgent <> True Then
   FrmHelp.FraAgent.Visible = True
   FrmHelp.Show vbModal
Else
   If BlBalloon = True Then
      FrmSelection.Agent1.Characters("nitij").Balloon.FontName = "Tahoma"
      FrmSelection.Agent1.Characters("nitij").Balloon.FontSize = 8
      FrmSelection.Agent1.Characters("nitij").Balloon.Style = 3
   ElseIf BlBalloon = False Then
        FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
   End If
   
   FrmSelection.Agent1.Characters("nitij").StopAll
   FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 250, Me.Top / Screen.TwipsPerPixelY + 50
   FrmSelection.Agent1.Characters("nitij").Play "gestureright"
   FrmSelection.Agent1.Characters("nitij").Speak "Select the name of the Agent from this list. Selecting will display the image of the particule Agent in the Preview window."
   FrmSelection.Agent1.Characters("nitij").Play "blink"

   FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 250, Me.Top / Screen.TwipsPerPixelY + 150
   FrmSelection.Agent1.Characters("nitij").Play "gestureright"
   FrmSelection.Agent1.Characters("nitij").Speak "Click on the Hide Button to hide or show the agent"
   FrmSelection.Agent1.Characters("nitij").Play "blink"
   
   FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 100, Me.Top / Screen.TwipsPerPixelY + 200
   FrmSelection.Agent1.Characters("nitij").Play "gestureright"
   FrmSelection.Agent1.Characters("nitij").Speak "Click on the Preview button to take a preview of the selected agent."
   FrmSelection.Agent1.Characters("nitij").Play "blink"
   
   FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 50, Me.Top / Screen.TwipsPerPixelY + 200
   FrmSelection.Agent1.Characters("nitij").Play "gestureright"
   FrmSelection.Agent1.Characters("nitij").Speak "Click on the Apply button to apply the selected agent on Brain"
   FrmSelection.Agent1.Characters("nitij").Play "blink"
 
   BlBalloon = False
   
   FrmSelection.Agent1.Characters("nitij").MoveTo (Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
End If

End Sub

Private Sub CmdPreview_Click()
On Error Resume Next
If CmdShow.Caption = "Show Agent" Then
   MESSAGE "Agent display is off.", OkOnly, "BRAIN@: Error"
   Exit Sub
Else
   FrmSelection.Agent1.Characters.Unload "nitij"
   Char = WindowsDirectory & "\BSystem2\chars\" & List1.Text & ".Acs"
   PrChar = Char
   FrmSelection.Agent1.Characters.Load "nitij", Char
   BlAgent = True
   FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
   FrmSelection.Agent1.Characters("nitij").Height = 140
   FrmSelection.Agent1.Characters("nitij").Width = 130
   FrmSelection.Agent1.Characters("nitij").Top = (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
   FrmSelection.Agent1.Characters("nitij").Left = (Me.Left / Screen.TwipsPerPixelX)

   FrmSelection.Agent1.Characters("nitij").Show
   FrmSelection.Agent1.Characters("nitij").Speak "Hi, You all are Welcome in brain"
   PREVIEW = True
   CmdShow.Enabled = True
End If
End Sub

Private Sub CmdShow_Click()
On Error Resume Next
 If CmdShow.Caption = "Show Agent" Then
    Char = WindowsDirectory & "\BSystem2\chars\" & List1.Text & ".Acs"
    PrChar = Char
    FrmSelection.Agent1.Characters.Load "nitij", Char
    FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
    FrmSelection.Agent1.Characters("nitij").Height = 140
    FrmSelection.Agent1.Characters("nitij").Width = 130
    FrmSelection.Agent1.Characters("nitij").Top = (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
    FrmSelection.Agent1.Characters("nitij").Left = (Me.Left / Screen.TwipsPerPixelX)
    FrmSelection.Agent1.Characters("nitij").Show
    CmdShow.Caption = "Hide Agent"
 ElseIf CmdShow.Caption = "Hide Agent" Then
        FrmSelection.Agent1.Characters("nitij").Hide
        FrmSelection.Agent1.Characters.Unload "nitij"
        CmdApply.Enabled = False
        CmdPreview.Enabled = False
        CmdShow.Caption = "Show Agent"
 End If
 PREVIEW = False
  
 CmdApply.Enabled = True
 CmdPreview.Enabled = True

End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
On Error Resume Next
Dim IntCnt As Integer
Dim StrTemp As String
Dim StrWinPath As String
Dim BlFlag As Boolean
'Seeks the Speech Engine
 StrWinPath = WindowsDirectory & "\"
 If Not FSO.FolderExists(StrWinPath & "\" & "Speech") Then ' check the folder
       BlFlag = True
 End If

 For IntCnt = 0 To 14
     If Not FSO.FileExists(StrWinPath & "\" & "Speech\" & StrFileNames(IntCnt)) Then
        BlFlag = True
     End If
 Next

If BlBalloon = True Then ChkBalloon.Value = 1 Else ChkBalloon.Value = 0
Bldown = True
File1.Path = WindowsDirectory & "\BSystem2\Chars\"
File1.Selected(1) = True

For IntCnt = 0 To File1.ListCount - 1
    File1.Selected(IntCnt) = True
    StrTemp = Mid(File1.filename, 1, InStr(1, File1.filename, ".") - 1)
    List1.AddItem StrTemp
    StrTemp = ""
Next

'If BlAgent = True Then
   'CmdShow.Enabled = True
'Else
   CmdShow.Enabled = False
'End If

mTop = (Me.Top / Screen.TwipsPerPixelY)
mLeft = (Me.Left / Screen.TwipsPerPixelX)
List1.Selected(0) = True

End Sub

Private Sub List1_Click()
Dim StrPicture As String
If UCase(List1.Text) = "GENIE" Then
   Label1.Visible = False
   Image1.Picture = imgGenie.Picture
ElseIf UCase(List1.Text) = "MERLIN" Then
   Label1.Visible = False
   Image1.Picture = ImgMerlin.Picture
ElseIf UCase(List1.Text) = "PEEDY" Then
   Label1.Visible = False
   Image1.Picture = ImgPeedy.Picture
ElseIf UCase(List1.Text) = "ROBBY" Then
   Label1.Visible = False
   Image1.Picture = ImgRobby.Picture
ElseIf UCase(List1.Text) = "AL" Then
   Label1.Visible = False
   Image1.Picture = imgal.Picture
ElseIf UCase(List1.Text) = "ELECTRA" Then
   Label1.Visible = False
   Image1.Picture = imgelectra.Picture
ElseIf UCase(List1.Text) = "HANZ" Then
   Label1.Visible = False
   Image1.Picture = imghanz.Picture
ElseIf UCase(List1.Text) = "IMGOZZAR" Then
   Label1.Visible = False
   Image1.Picture = imgozzar.Picture
ElseIf UCase(List1.Text) = "SANTA" Then
   Label1.Visible = False
   Image1.Picture = imgsanta.Picture
ElseIf UCase(List1.Text) = "WARTNOSE" Then
   Label1.Visible = False
   Image1.Picture = imgwartnose.Picture
Else
   Image1.Picture = Nothing
   Label1.Visible = True
End If
CmdApply.Enabled = True

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
