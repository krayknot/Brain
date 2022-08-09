VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmInstallSpeech 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Installing Speech Engine"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "FrmInstallSpeech.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   3735
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   9000
         Left            =   480
         Top             =   240
      End
      Begin VB.CommandButton CmdClose 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin Brain.UserControl1 Bar 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Color           =   6956042
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmInstallSpeech.frx":0ECA
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "FrmInstallSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
#If Win32 Then
    Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
#ElseIf Win16 Then
        Private Declare Function waveOutGetNumDevs Lib "mmsystem" () As Integer
#End If

Private Sub cmdclose_Click()
Unload Me
FrmAgent.Show vbModal
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
If Me.Visible = False Then Me.Visible = True
Me.Refresh
Dim IntCnt As Integer
Dim StrWinPath As String
 
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

Bar.Value = Bar.Value + 10
Delay

'Detect the soundcard and provide appropriate message
 #If Win32 Then
     Dim i As Long
 #ElseIf Win16 Then
         Dim i As Integer
 #End If
 i = waveOutGetNumDevs()
 If i > 0 Then         ' There is at least one device.
    List1.AddItem "Sound Card Detected."
    List1.Selected(List1.ListCount - 1) = True
    List1.Refresh
    
    Bar.Value = Bar.Value + 10
    Delay
    
   'if soundcard found then copy the appropriate files to the proper place
    StrWinPath = WindowsDirectory & "\"
    If Not FSO.FolderExists(App.Path & "\Installation") Then 'Seeking the installation folder
       List1.AddItem "Installation Folder not found"
       List1.Selected(List1.ListCount - 1) = True
       List1.AddItem "Aborting the Installation."
       List1.Selected(List1.ListCount - 1) = True
       List1.Refresh
       MESSAGE "Installation Folder not found. Aborting?", OkOnly, "BRAIN: Error"
       Unload Me
       FrmSelection.Show vbModal
    End If
    Bar.Value = Bar.Value + 10
    Delay
    
    If Not FSO.FolderExists(StrWinPath & "Speech") Then
       FSO.CreateFolder StrWinPath & "Speech" 'Creates the main speech folder
    End If
    
    Bar.Value = Bar.Value + 10
    Delay
    
    If Not FSO.FolderExists(App.Path & "\Installation\Speech") Then 'Seeking the installation folder
       List1.AddItem "Installation Folder not found"
       List1.Selected(List1.ListCount - 1) = True
       List1.AddItem "Aborting the Installation."
       List1.Selected(List1.ListCount - 1) = True
       List1.Refresh
       MESSAGE "Installation Folder not found. Aborting?", OkOnly, "BRAIN: Error"
       Unload Me
       FrmSelection.Show vbModal
    End If
    
    For IntCnt = 0 To 14 'Deletes all existing files for fresh installation
        If FSO.FileExists(StrWinPath & "Speech\" & StrFileNames(IntCnt)) Then
           FSO.DeleteFile StrWinPath & "Speech\" & StrFileNames(IntCnt), True
        End If
    Next
    
    Bar.Value = Bar.Value + 10
    Delay
    
    For IntCnt = 0 To 14
        FSO.CopyFile App.Path & "\Installation\Speech\" & StrFileNames(IntCnt), StrWinPath & "Speech\", True
        List1.AddItem StrFileNames(IntCnt)
        List1.Selected(List1.ListCount - 1) = True
        List1.Refresh
    Next
    
    Bar.Value = Bar.Value + 50
    Delay
       
    List1.AddItem "Speech Engine installed"
    List1.Selected(List1.ListCount - 1) = True
    List1.Refresh
    Delay
    MESSAGE "Speech Engine Installed. If it not works yet please restart your system", OkOnly, "BRAIN: Information"
    Unload Me
    FrmSelection.Show vbModal
 Else
    List1.AddItem "Sound Card not detected."
    List1.AddItem "You cannot use Agent Speech functions."
    Exit Sub
 End If
End Sub

Function Delay()
Dim IntCount As Long
For IntCount = 1 To 450000
Timer1.Enabled = True
Next
End Function
