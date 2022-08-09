VERSION 5.00
Begin VB.Form FrmSkin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Skin "
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   ControlBox      =   0   'False
   Icon            =   "FrmSkin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   960
      Top             =   1560
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   3735
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
         Left            =   2760
         TabIndex        =   5
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
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
         Height          =   3150
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "FrmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bldown As Boolean, BlUp As Boolean
Dim i As Integer
Private Sub Command1_Click()

End Sub

Private Sub CmdApply_Click()
On Error Resume Next
File1.Selected(List1.ListIndex) = True
If File1.filename = "" Then
   MsgBox "Please select the skin name you want to change"
Else
   FrmSelection.Skin1.RemoveSkin FrmSelection.hwnd
   FrmSelection.Skin1.LoadSkin File1.Path & "\" & File1.filename
  'Saves the settings in the registry for the future use
   SaveSetting "BRAIN", "Values", "Skin", File1.Path & "\" & File1.filename
   FrmSelection.Skin1.ApplySkin FrmSelection.hwnd
   Unload Me
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim IntCnt As Integer
Dim StrTemp As String
Bldown = True

FrmSelection.Skin1.ApplySkin Me.hwnd
File1.Path = WindowsDirectory & "\BSystem2\Skins\"
'File1.Selected(1) = True

For IntCnt = 0 To File1.ListCount - 1
    File1.Selected(IntCnt) = True
    StrTemp = Mid(File1.filename, 1, InStr(1, File1.filename, ".") - 1)
    List1.AddItem StrTemp
    StrTemp = ""
Next
End Sub

Private Sub List1_Click()
CmdApply.Enabled = True
End Sub

Private Sub List1_DblClick()
File1.Selected(List1.ListIndex) = True
If File1.filename = "" Then
   MsgBox "Please select the skin name you want to change"
Else
   FrmSelection.Skin1.RemoveSkin FrmSelection.hwnd
   FrmSelection.Skin1.LoadSkin File1.Path & "\" & File1.filename
  'Saves the settings in the registry for the future use
   SaveSetting "BRAIN", "Values", "Skin", File1.Path & "\" & File1.filename
   FrmSelection.Skin1.ApplySkin FrmSelection.hwnd
   Unload Me
End If
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
