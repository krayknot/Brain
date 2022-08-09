VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BRAIN2: Creating Paper..."
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
End Sub
