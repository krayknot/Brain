VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frmmessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "Frmmessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   1080
         Top             =   120
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
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
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
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
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   240
         Width           =   495
         Begin VB.Image Imgyesno 
            Height          =   495
            Left            =   0
            Picture         =   "Frmmessage.frx":0ECA
            Stretch         =   -1  'True
            Top             =   0
            Width           =   495
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lblmessage 
         Height          =   1215
         Left            =   840
         OleObjectBlob   =   "Frmmessage.frx":2BF4
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Frmmessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bldown As Boolean, BlUp As Boolean
Dim i As Integer

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
BLMessage = True
Unload Me
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
Bldown = True
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
