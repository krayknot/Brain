VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frmoptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Options"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   ControlBox      =   0   'False
   Icon            =   "Frmoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   10
      TabIndex        =   12
      Top             =   1320
      Width           =   2970
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   120
         Top             =   120
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "BRAIN2 : Select Article File "
         Filter          =   "Text (*.txt)|*.txt"
         InitDir         =   "app.path "
         MaxFileSize     =   500
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
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1200
         TabIndex        =   14
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
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.FileListBox File3 
      Height          =   2430
      Left            =   3120
      TabIndex        =   39
      Top             =   -120
      Width           =   2055
   End
   Begin VB.Frame Frame7 
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
      Height          =   2535
      Left            =   2160
      TabIndex        =   36
      Top             =   3000
      Width           =   2895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   2655
         TabIndex        =   37
         Top             =   240
         Width           =   2655
         Begin VB.Image Image1 
            Height          =   1695
            Left            =   0
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Picture to display"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   960
            Width           =   1725
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6360
      Top             =   3240
   End
   Begin VB.Frame Frame6 
      Caption         =   "Articles Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Top             =   960
         Width           =   975
      End
      Begin VB.FileListBox File2 
         Height          =   480
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton CmdAttach 
         Caption         =   "Attach"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtArticleName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frmoptions.frx":000C
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frmoptions.frx":0073
         TabIndex        =   26
         Top             =   240
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Text            =   "Choose articles from the list below"
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Change Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5280
      TabIndex        =   18
      Top             =   0
      Width           =   3015
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   360
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdPicApply 
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
         Height          =   405
         Left            =   1800
         TabIndex        =   33
         Top             =   2000
         Visible         =   0   'False
         Width           =   975
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
         Height          =   1425
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frmoptions.frx":010A
         TabIndex        =   31
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Theme"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   50
      TabIndex        =   16
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton CmdThemeDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdTheme 
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
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
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
         Height          =   1230
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frmoptions.frx":0199
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings"
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
      Height          =   1335
      Left            =   5160
      TabIndex        =   7
      Top             =   3840
      Width           =   2895
      Begin VB.OptionButton OptOff 
         Caption         =   "OFF"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptOn 
         Caption         =   "ON"
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
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "Frmoptions.frx":022B
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Chkminimize 
         Caption         =   "Minimize to TaskBar on Minimizing"
         Enabled         =   0   'False
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
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agent"
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
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   2640
      Width           =   2895
      Begin VB.CheckBox ChkAgent 
         Caption         =   "Start Agent at program startup"
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
         TabIndex        =   6
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtwidth 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Text            =   "130"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txtheight 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "160"
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "Frmoptions.frx":029E
         TabIndex        =   2
         Top             =   510
         Width           =   1095
      End
      Begin VB.CheckBox Chkagentonoff 
         Caption         =   "Show Agent"
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
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   1
         Left            =   1560
         OleObjectBlob   =   "Frmoptions.frx":0303
         TabIndex        =   3
         Top             =   510
         Width           =   495
      End
   End
End
Attribute VB_Name = "Frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bldown As Boolean, BlUp As Boolean
Dim i As Integer
Dim BlTemp As Boolean
Dim BlRegistryCombo
Dim StrFileName As String

Private Sub Chkagentonoff_Click()
If BlTemp = True Then
    If Chkagentonoff.Caption = "Hide Agent" Then
       FrmSelection.Agent1.Characters.Character("nitij").Play "sad"
       FrmSelection.Agent1.Characters.Character("nitij").Hide
       Chkagentonoff.Value = 0
       Chkagentonoff.Caption = "Show Agent"
       BlAgent = False
    ElseIf Chkagentonoff.Caption = "Show Agent" Then
           FrmAgent.Show vbModal
    End If
End If
 
End Sub

Private Sub CmdApply_Click()
On Error Resume Next
'Saves the settings for the AGent
 If Chkagentonoff.Value = 1 Then
    FrmSelection.Agent1.Characters.Character("nitij").Hide
    FrmSelection.Agent1.Characters.Character("nitij").Height = Val(Txtheight.Text)
    FrmSelection.Agent1.Characters.Character("nitij").Width = Val(txtwidth.Text)
    
    FrmSelection.Agent1.Characters("nitij").Top = (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
    FrmSelection.Agent1.Characters("nitij").Left = (Me.Left / Screen.TwipsPerPixelX)
    FrmSelection.Agent1.Characters.Character("nitij").Show
 End If
  
 SaveSetting "BRAIN", "Values", "AGENT_HEIGHT", Val(Txtheight.Text)
 SaveSetting "BRAIN", "Values", "AGENT_WIDTH", Val(txtwidth.Text)
 
 If ChkAgent.Value = 1 Then
    SaveSetting "BRAIN", "Values", "AGENT_STARTUP", True
 ElseIf ChkAgent.Value = 0 Then
    SaveSetting "BRAIN", "Values", "AGENT_STARTUP", False
 End If
 
'Sets the picture logo in the frmbrain form
' FrmBrain.Image2.Picture = LoadPicture(App.Path & "\logo\" & File1.List(List2.ListIndex))
' SaveSetting "BRAIN", "Values", "LOGO ", File1.List(List2.ListIndex)

'Saves the settings for the Settings

'Saves the settings for the Thems

'Saves the settings for the Font

Unload Me
End Sub

Private Sub CmdAttach_Click()
On Error Resume Next
If CommonDialog1.filename = "" Then
   MESSAGE "Type the name of Article and then select the Article by clicking on BROWSE button", OkOnly, "BRAIN@: Error"
   Exit Sub
End If

Dim FSO As New FileSystemObject
FSO.CopyFile CommonDialog1.filename, App.Path & "\articles\", True
File3.Refresh

If BlArticleOpen = True Then FrmArticle.File1.Refresh
End Sub

Private Sub CmdBrowse_Click()

CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.filename

End Sub

Private Sub cmdclose_Click()

Unload Me
End Sub

Private Sub CmdFont_Click()
StrFont = Mid(File1.filename, 1, InStr(1, File1.filename, ".") - 1)

'Dim IntCt As Integer
'For IntCt = 0 To 9
'    FrmSelection.OptSubject(IntCt).FontName = StrFont
'    FrmSelection.OptSubject(IntCt).Refresh
'Next

SkinLabel4.Font.name = StrFont
End Sub

Private Sub CmdDelete_Click()
FrmListLabels.Show vbModal
End Sub

Private Sub CmdTheme_Click()

SaveSetting "BRAIN", "Values", "THEME", List1.Text
Frmoptions.MouseIcon = LoadPicture(App.Path & "\Theme\" & List1.Text & "\pointer.cur")
Frmoptions.MousePointer = 99
FrmSelection.MouseIcon = LoadPicture(App.Path & "\Theme\" & List1.Text & "\pointer.cur")
FrmSelection.MousePointer = 99
THEME = List1.Text

End Sub

Private Sub CmdThemeDelete_Click()
SaveSetting "BRAIN", "Values", "THEME", "Default"

Frmoptions.MousePointer = 1
FrmSelection.MousePointer = 1

THEME = "Default"
End Sub

Private Sub Combo1_Click()
TxtArticleName = Combo1.Text
TxtArticleName.SetFocus
 
End Sub

Private Sub Command2_Click()
If BlAgent = False Then
   FrmHelp.FraOptions.Visible = True
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
   FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 100, Me.Top / Screen.TwipsPerPixelY + 100
   FrmSelection.Agent1.Characters("nitij").Play "gestureright"
   FrmSelection.Agent1.Characters("nitij").Speak "Select the cursor theme of your choice to change the cursor and click apply to apply it"
   FrmSelection.Agent1.Characters("nitij").Play "blink"

   FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelY + 350, Me.Top / Screen.TwipsPerPixelY + 50
   FrmSelection.Agent1.Characters("nitij").Play "gestureright"
   FrmSelection.Agent1.Characters("nitij").Speak "Set the article of your choice by selecting the text file article location and click to Attach button. Click to Delete button to delete the selected article."
   FrmSelection.Agent1.Characters("nitij").Play "blink"

   FrmSelection.Agent1.Characters("nitij").MoveTo (Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10
End If
End Sub


Private Sub Form_Load()
On Error Resume Next
FrmSelection.Skin1.ApplySkin Me.hwnd
FlatBorder Txtheight.hwnd
FlatBorder txtwidth.hwnd
FlatBorder List1.hwnd
FlatBorder TxtArticleName.hwnd

Bldown = True
Dim IntCnt As Integer

Chkagentonoff.Visible = False

BlTemp = False

If GetSetting("BRAIN", "Values", "AGENT_STARTUP") = True Then
    If GetSetting("BRAIN", "Values", "AGENT_ON") = True Then
       Chkagentonoff.Value = 1
       Chkagentonoff.Caption = "Hide Agent"
    End If
End If

'Settings for the Logo
 File1.Path = App.Path & "\logo"
 
 For IntCnt = 0 To File1.ListCount - 1
     temp = Mid(File1.List(IntCnt), 1, InStr(1, File1.List(IntCnt), ".") - 1)
     List2.AddItem temp
 Next
      
'Settings for the Themes
 Dir1.Path = App.Path & "\theme"
 
 For IntCnt = 0 To Dir1.ListCount
     temp = Extract(Dir1.List(IntCnt))
     List1.AddItem temp
 Next IntCnt
 
'Settings for the Articles
 File2.Path = App.Path & "\articles"
 
 For IntCnt = 0 To File2.ListCount - 1
     temp = Mid(File2.List(IntCnt), 1, InStr(1, File2.List(IntCnt), ".") - 1)
     Combo1.AddItem temp
 Next
 BlTemp = True
 
 File3.Path = App.Path & "\articles\"
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
BlTemp = True
Timer1.Enabled = False
FrmSelection.Timer1.Enabled = True
End Sub

Private Sub List2_Click()
Image1.Picture = LoadPicture(App.Path & "\logo\" & File1.List(List2.ListIndex))
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If mLeft <> (Me.Left / Screen.TwipsPerPixelX) Then
  FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelX, (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10, 50
  mTop = (Me.Top / Screen.TwipsPerPixelY)
  mLeft = (Me.Left / Screen.TwipsPerPixelX)
End If

If mTop <> (Me.Top / Screen.TwipsPerPixelY) Then
  FrmSelection.Agent1.Characters("nitij").MoveTo Me.Left / Screen.TwipsPerPixelX, (Me.Top / Screen.TwipsPerPixelY) - FrmSelection.Agent1.Characters("nitij").Height + 10, 50
  mTop = (Me.Top / Screen.TwipsPerPixelY)
  mLeft = (Me.Left / Screen.TwipsPerPixelX)
End If
End Sub

Private Sub Timer2_Timer()
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

Private Sub TxtArticleName_GotFocus()
sfield TxtArticleName
End Sub
