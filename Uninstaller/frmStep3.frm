VERSION 5.00
Begin VB.Form frmStep3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Uninstaller"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Contact: krayknot@Yahoo.com"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Thanks for using BRAIN. Please do support us by your advice"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   120
      OLEDropMode     =   1  'Manual
      Picture         =   "frmStep3.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmStep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub
