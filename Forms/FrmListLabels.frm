VERSION 5.00
Begin VB.Form FrmListLabels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Select Heading"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2310
   ControlBox      =   0   'False
   Icon            =   "FrmListLabels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
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
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   975
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
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
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "FrmListLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
If List1.Text = "" Then
   MESSAGE "No Article is selected. Select the article to delete first", OkOnly, "Error: Delete Article"
   Exit Sub
End If

If Not FSO.FileExists(App.Path & "\Articles\" & List1.Text & ".txt") Then
   MESSAGE "Article file " & App.Path & "\Articles\" & List1.Text & ".txt not found.", OkOnly, "Error: Delete Article"
   Exit Sub
End If

FSO.DeleteFile App.Path & "\Articles\" & List1.Text & ".txt", True
Frmoptions.File3.Refresh

Form_Load
MESSAGE "Selected Article has been deleted.", OkOnly, "BRAIN:Article Deleted"


End Sub

Private Sub Form_Load()
On Error Resume Next
FrmSelection.Skin1.ApplySkin Me.hwnd
FlatBorder List1.hwnd

Dim IntCnt As Integer
Dim StrTemp As String
List1.Clear
For IntCnt = 0 To Frmoptions.File3.ListCount - 1
    StrTemp = Mid(Frmoptions.File3.List(IntCnt), 1, InStr(1, Frmoptions.File3.List(IntCnt), ".") - 1)
    List1.AddItem StrTemp
Next
End Sub

