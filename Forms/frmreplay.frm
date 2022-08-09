VERSION 5.00
Begin VB.Form frmreplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN2: Replay"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List6 
      Height          =   2790
      Left            =   2160
      TabIndex        =   19
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   2790
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ListBox List5 
      Height          =   2790
      Left            =   6240
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   2790
      Left            =   4200
      TabIndex        =   14
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   10
      TabIndex        =   4
      Top             =   4320
      Width           =   8280
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
         Left            =   7200
         MousePointer    =   1  'Arrow
         TabIndex        =   13
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
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdTellUs 
         Caption         =   "Tell Us"
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
         Left            =   5160
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdDisplay 
         Caption         =   "Display"
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
         Left            =   4320
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
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
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command2 
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
         Left            =   1800
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
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
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
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
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "List of questions you solved right"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      Begin VB.ListBox List2 
         Height          =   3960
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of questions you solved wrong"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label Label4 
      Caption         =   "answer"
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "user input"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "answer"
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "user input"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   5400
      Width           =   975
   End
End
Attribute VB_Name = "frmreplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As New ADODB.Connection
Dim Rst As New ADODB.Recordset

Private Sub CmdCLose_Click()
Unload Me
End Sub

Private Sub CmdDisplay_Click()
If List1.Text = "" And List2.Text = "" Then
   MESSAGE "Select the question to see details.", OkOnly, "BRAIN2: Error"
   Exit Sub
End If





FrmDisplay.Show vbModal
End Sub

Private Sub CmdTellUs_Click()
frmAbout.Width = 11325
frmAbout.TxtSubject.Text = StrPType & " related query"
frmAbout.TxtBody.Text = "<Type the query related to the question>"
frmAbout.Show vbModal
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
' StrPType = "Fill in the Blanks"
'Opening database
 DB.ConnectionString = "Provider='Microsoft.Jet.OLEDB.4.0';Data Source='" & App.Path & "\database\db.mdb';"
 DB.Open
 
 Rst.Open "Select * From test where papertype = '" & Trim(StrPType) & "'", DB, adOpenDynamic, adLockOptimistic
     
    While Not Rst.EOF
          List1.AddItem Rst!questionwrong
          List3.AddItem Rst!userans
          List6.AddItem Rst!questionans
          
          List2.AddItem Rst!questionright
          List4.AddItem Rst!userans
          List5.AddItem Rst!questionans
          Rst.MoveNext
    Wend
    Rst.Close
 
 
 
 
End Sub

Private Sub List1_Click()

FrmDisplay.Label1.Caption = List1.Text
FrmDisplay.Label2.Caption = List6.List(List1.ListIndex)
FrmDisplay.Label3.Caption = List3.List(List1.ListIndex)


End Sub

Private Sub List2_Click()
FrmDisplay.Label1.Caption = List2.Text
FrmDisplay.Label2.Caption = List4.List(List2.ListIndex)
FrmDisplay.Label3.Caption = List5.List(List2.ListIndex)
End Sub
