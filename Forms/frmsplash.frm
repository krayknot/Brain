VERSION 5.00
Begin VB.Form FrmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   6120
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: krayknot@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Krayknot"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Win-2000 and Win-xp preferred."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   0
      Picture         =   "frmsplash.frx":0ECA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IntCount As Integer
Dim FSO As New FileSystemObject

Private Sub Form_Load()
USERNAME = Label3(1).Caption
USEREMAIL = Label3(2).Caption

If FrmSplash.Visible = False Then FrmSplash.Visible = True
FrmSplash.Refresh
On Error Resume Next
Dim BlSetting As Boolean
Label2.Caption = ""

'Label2.Caption = "Checking Agent"
''First check the value of POSTINSTALL in the registry if it is TRUE only than proceed otherwise
''display a message like POSTINSTALLATION has not been completed.
' BlSetting = GetSetting("BRAIN", "POSTINSTALL", "POSTINSTALL")
' If BlSetting = False Then
'    Unload Me
'    FrmPostInstall.Show vbModal
' End If
'
' Label2.Caption = "Checking Fonts"
'Check for the fonts applied on different labels and if not found the font in the fonts
'folder then try to reinstall it or use a different alternative font
 If Not FSO.FileExists(WindowsDirectory & "\Fonts\" & "Ballsont.ttf") Then
    If FSO.FileExists(App.Path & "\Fonts\" & "Ballsont.ttf") Then
       FSO.CopyFile App.Path & "\fonts\ballsont.ttf", WindowsDirectory & "\Fonts\", True
    End If
 End If

 Label2.Caption = "Checking Papers"
'Sets the required files for speech
 StrFileNames(0) = "spchtel.dll"
 StrFileNames(1) = "vcauto.tlb"
 StrFileNames(2) = "VText.dll"
 StrFileNames(3) = "Xlisten.dll"
 StrFileNames(4) = "XTel.dll"
 StrFileNames(5) = "vtxtauto.tlb"
 StrFileNames(6) = "vcmd.exe"
 StrFileNames(7) = "speech.hlp"
 StrFileNames(8) = "speech.cnt"
 StrFileNames(9) = "vcmshl.dll"
 StrFileNames(10) = "WrapSAPI.dll"
 StrFileNames(11) = "Xvoice.dll"
 StrFileNames(12) = "speech.dll"
 StrFileNames(13) = "Vdict.dll"
 StrFileNames(14) = "Xcommand.dll"
 
 Dim Char As String, PrChar As String
 Dim IntCnt As Integer
 If GetSetting("BRAIN", "Values", "THEME") = "Default" Then
    FrmSelection.MousePointer = 1
 Else
    THEME = GetSetting("BRAIN", "Values", "THEME")
 End If

 Load FrmSelection
 Load FrmBrain

' Unload Me
' FrmSelection.Show
 LABEL_UP = True 'frmbrain > label5 > moveup setting
 Unload FrmBrain
End Sub

Private Sub Timer1_Timer()
If IntCount <= 5 Then
   IntCount = IntCount + 1
   If IntCount = 5 Then
      Unload Me
      FrmSelection.Show
  End If
End If

End Sub
