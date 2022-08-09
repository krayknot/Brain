VERSION 5.00
Begin VB.Form frmStep1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Uninstall"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmStep1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
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
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
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
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   2655
      Left            =   120
      Picture         =   "FrmStep1.frx":058A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Are you sure to uninstall BRAIN"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmStep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
End
End Sub

Private Sub cmdYes_Click()
Unload Me
frmStep2.Show vbModal
End Sub

Private Sub Image2_Click()
End Sub
