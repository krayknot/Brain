VERSION 5.00
Begin VB.Form FrmArticle 
   Caption         =   "BRAIN: Articles"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5760
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   6975
      Left            =   4800
      TabIndex        =   8
      Top             =   0
      Width           =   5655
      Begin VB.TextBox TxtInfo 
         Height          =   6615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   10
      TabIndex        =   2
      Top             =   6240
      Width           =   4680
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
         Left            =   3720
         TabIndex        =   7
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
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdCreate 
         Caption         =   "Create"
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
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
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
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdOpen 
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click on the name of the Article you want to read"
      Height          =   6255
      Left            =   10
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   2760
         Top             =   2040
      End
      Begin VB.Timer Timer1 
         Interval        =   10000
         Left            =   1200
         Top             =   3720
      End
      Begin VB.ListBox List1 
         Height          =   5910
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "FrmArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BlUp As Boolean, Bldown As Boolean
Dim i As Integer
Dim FSO As New FileSystemObject
Dim IntCnt As Integer
Dim StrTemp As String

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub CmdCreate_Click()
Frmoptions.Show vbModal

List1.Clear
For IntCnt = 0 To File1.ListCount - 1
    StrTemp = Mid(File1.List(IntCnt), 1, InStr(1, File1.List(IntCnt), ".") - 1)
    List1.AddItem StrTemp
Next IntCnt
End Sub

Private Sub CmdDelete_Click()
If List1.Text = "" Then
   MESSAGE "Select the Article to delete", OkOnly, "BRAIN: File Selection Error"
   Exit Sub
End If

MESSAGE "Are you sure to delete the Article " & List1.Text, YesNoonly, "BRAIN: Information"

If BLMessage = True Then
   FSO.DeleteFile App.Path & "\articles\" & Trim(List1.Text) & ".txt", True
   MESSAGE "File has been deleted", OkOnly, "BRAIN: Information"
   List1.Clear
   File1.Refresh
   For IntCnt = 0 To File1.ListCount - 1
       StrTemp = Mid(File1.List(IntCnt), 1, InStr(1, File1.List(IntCnt), ".") - 1)
       List1.AddItem StrTemp
   Next IntCnt
   BLMessage = False
End If

End Sub

Private Sub CmdHelp_Click()
If BlBalloon = True Then
      FrmSelection.Agent1.Characters("nitij").Balloon.FontName = "Tahoma"
      FrmSelection.Agent1.Characters("nitij").Balloon.FontSize = 8
      FrmSelection.Agent1.Characters("nitij").Balloon.Style = 3
   ElseIf BlBalloon = False Then
        FrmSelection.Agent1.Characters("nitij").Balloon.Style = 4
   End If
End Sub

Private Sub CmdOpen_Click()
If List1.Text = "" Then
   MESSAGE "Select the Article to Open", OkOnly, "BRAIN: File Selection Error"
   Exit Sub
End If


Dim pfile As File, mtext As TextStream
Dim StrTmp As String

StrTmp = App.Path & "\articles\" & Trim(List1.Text) & ".txt"

If FSO.FileExists(StrTmp) Then
   Set pfile = FSO.GetFile(StrTmp)
   Set mtext = pfile.OpenAsTextStream(ForReading, TristateMixed)
   TxtInfo.Text = mtext.ReadAll
Else
   MESSAGE "FIle not exists", OkOnly, "Info"
End If
Frame3.Caption = List1.Text
Me.Width = 10620
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
Bldown = True

List1.Clear
File1.Path = App.Path & "\Articles"
For IntCnt = 0 To File1.ListCount - 1
    StrTemp = Mid(File1.List(IntCnt), 1, InStr(1, File1.List(IntCnt), ".") - 1)
    List1.AddItem StrTemp
Next IntCnt

End Sub

Private Sub List1_DblClick()
Dim pfile As File, mtext As TextStream
Dim StrTmp As String

StrTmp = App.Path & "\articles\" & Trim(List1.Text) & ".txt"

If FSO.FileExists(StrTmp) Then
   Set pfile = FSO.GetFile(StrTmp)
   Set mtext = pfile.OpenAsTextStream(ForReading, TristateMixed)
   TxtInfo.Text = mtext.ReadAll
Else
   MESSAGE "FIle not exists", OkOnly, "Info"
End If
Frame3.Caption = List1.Text
Me.Width = 10620

End Sub

Private Sub Timer1_Timer()
List1.Clear
File1.Path = App.Path & "\Articles"
File1.Refresh
For IntCnt = 0 To File1.ListCount - 1
    StrTemp = Mid(File1.List(IntCnt), 1, InStr(1, File1.List(IntCnt), ".") - 1)
    List1.AddItem StrTemp
Next IntCnt

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
