VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmdinformation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BRAIN: Detailed Paper Information"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   4080
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   6735
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "Select the paper type to view details"
         Top             =   840
         Width           =   3150
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   1
         Left            =   600
         OleObjectBlob   =   "frmdinformation.frx":0000
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "frmdinformation.frx":0063
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   3
         Left            =   120
         OleObjectBlob   =   "frmdinformation.frx":00D8
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel label2 
         Height          =   255
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "frmdinformation.frx":0155
         TabIndex        =   13
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "frmdinformation.frx":01B8
         TabIndex        =   9
         Top             =   200
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmdinformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim IT1 As New ITPaperCount
Dim IWPD1 As New IWPDPaperCount
Dim PC1 As New PCPaperCount
Dim C1 As New CPaperCount
Dim BS1 As New BSPaperCount
Dim DBMS1 As New DBMSPaperCount
'Dim SAD1 As New sadpapercount
Dim DCN1 As New DCNPaperCount
Dim CG1 As New CGPaperCount
Dim dS1 As New DSPaperCount
Dim UNIX1 As New UnixPaperCount
Dim CPP1 As New CPPPaperCount
'Dim coss1 As New cosspapercount

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
Dim inttemp As Integer
Dim StrTemp As String

 If List1.Text = "Information Technology" Then
    
    If Combo1.Text = "Fill in the Blanks" Then
       Label2(2).Caption = ""
       Label2(2).Caption = IT1.IT_FB_PaperNum & " Papers with 10 Questions each"
    ElseIf Combo1.Text = "Matching Columns" Then
           Label2(2).Caption = ""
           Label2(2).Caption = IT1.IT_MTC_PaperNum & " Papers with 10 Questions each"
    ElseIf Combo1.Text = "Multiple Choice" Then
           Label2(2).Caption = ""
           Label2(2).Caption = IT1.IT_MC_PaperNum & " Papers with 10 Questions each"
    ElseIf Combo1.Text = "Descriptive Questions" Then
           Label2(2).Caption = ""
           Label2(2).Caption = IT1.IT_QUES_PaperNum & " Papers with 10 Questions each"
    ElseIf Combo1.Text = "True or False" Then
           Label2(2).Caption = ""
           Label2(2).Caption = IT1.IT_TF_PaperNum & " Papers with 10 Questions each"
    End If
    
 ElseIf List1.Text = "Internet and Web Designing" Then
        If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = IWPD1.IWPD_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = IWPD1.IWPD_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = IWPD1.IWPD_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = IWPD1.IWPD_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = IWPD1.IWPD_MC_PaperNum & " Papers with 10 Questions each"
        End If

 ElseIf List1.Text = "Personal Computing Technology" Then
        If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = PC1.PC_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = PC1.PC_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = PC1.PC_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = PC1.PC_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = PC1.PC_TF_PaperNum & " Papers with 10 Questions each"
        End If
 
ElseIf List1.Text = "C Language" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = C1.C_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = C1.C_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = C1.C_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = PC1.PC_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = PC1.PC_TF_PaperNum & " Papers with 10 Questions each"
        End If
        
ElseIf List1.Text = "Business Systems and FoxPro" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = BS1.BS_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = BS1.BS_MTC_Option_Count & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = BS1.BS_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = BS1.BS_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = BS1.BS_TF_PaperNum & " Papers with 10 Questions each"
        End If

ElseIf List1.Text = "Database Management Systems" Then   '
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = DBMS1.DBMS_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DBMS1.DBMS_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DBMS1.DBMS_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DBMS1.DBMS_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DBMS1.DBMS_TF_PaperNum & " Papers with 10 Questions each"
        End If

'ElseIf List1.Text = "Systen Analysis and Designing" Then
'If Combo1.Text = "Fill in the Blanks" Then
'           Label2(2).Caption = ""
'           Label2(2).Caption = SAD1.sad_FB_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "Matching Columns" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = SAD1.sad_MTC_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "Multiple Choice" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = SAD1.sad_MC_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "Descriptive Questions" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = SAD1.sad_QUES_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "True or False" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = SAD1.sad_TF_PaperNum & " Papers with 10 Questions each"
'        End If

ElseIf List1.Text = "Data Communications and Networking" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = DCN1.DCN_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DCN1.DCN_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DCN1.DCN_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DCN1.DCN_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = DCN1.DCN_TF_PaperNum & " Papers with 10 Questions each"
        End If

ElseIf List1.Text = "Comupter Graphics" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = CG1.CG_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CG1.CG_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CG1.CG_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CG1.CG_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CG1.CG_TF_PaperNum & " Papers with 10 Questions each"
        End If

ElseIf List1.Text = "Data Structure" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = dS1.DS_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = dS1.DS_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = dS1.DS_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = dS1.DS_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = dS1.DS_TF_PaperNum & " Papers with 10 Questions each"
        End If

ElseIf List1.Text = "Unix" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = UNIX1.UNIX_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = UNIX1.UNIX_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = UNIX1.UNIX_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = UNIX1.UNIX_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = UNIX1.UNIX_TF_PaperNum & " Papers with 10 Questions each"
        End If

ElseIf List1.Text = "C++ Language" Then
If Combo1.Text = "Fill in the Blanks" Then
           Label2(2).Caption = ""
           Label2(2).Caption = CPP1.CPP_FB_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Matching Columns" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CPP1.CPP_MTC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Multiple Choice" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CPP1.CPP_MC_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "Descriptive Questions" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CPP1.CPP_QUES_PaperNum & " Papers with 10 Questions each"
        ElseIf Combo1.Text = "True or False" Then
               Label2(2).Caption = ""
               Label2(2).Caption = CPP1.CPP_TF_PaperNum & " Papers with 10 Questions each"
        End If

'ElseIf List1.Text = "Computer Organization and System Software" Then
'If Combo1.Text = "Fill in the Blanks" Then
'           Label2(2).Caption = ""
'           Label2(2).Caption = coss1.coss_FB_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "Matching Columns" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = coss1.coss_MTC_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "Multiple Choice" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = coss1.coss_MC_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "Descriptive Questions" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = coss1.coss_QUES_PaperNum & " Papers with 10 Questions each"
'        ElseIf Combo1.Text = "True or False" Then
'               Label2(2).Caption = ""
'               Label2(2).Caption = coss1.coss_TF_PaperNum & " Papers with 10 Questions each"
'        End If

Else
    File1.Path = App.Path & "\createpaper\" & List1.Text & "\" & Combo1.Text
    File1.Refresh
    
    Label2(2).Caption = ""
    Label2(2).Caption = File1.ListCount
End If
 
     
     
     
End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd

Dim inttemp As Integer
Dim StrTemp As String

'Check the paper dlls in the directory and if notfound then run the first time installation
 If FSO.FileExists(WindowsDirectory & "\system32\it_brain.dll") Then
    List1.AddItem "Information Technology"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\pc_brain.dll") Then
        List1.AddItem "Personal Computing Technology"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\iwpd_brain.dll") Then
        List1.AddItem "Internet and Web Designing"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\c_brain.dll") Then
        List1.AddItem "C Language"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\bs_brain.dll") Then
        List1.AddItem "Business Systems and FoxPro"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\dbms_brain.dll") Then
        List1.AddItem "Database Management Systems"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\sad_brain.dll") Then
        List1.AddItem "System Analysis and Designing"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\dcn_brain.dll") Then
        List1.AddItem "Data Communications and Networking"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\cg_brain.dll") Then
        List1.AddItem "Computer Graphics"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\ds_brain.dll") Then
        List1.AddItem "Data Structure"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\unix_brain.dll") Then
        List1.AddItem "Unix"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\cpp_brain.dll") Then
        List1.AddItem "C++ Language"
 ElseIf FSO.FileExists(WindowsDirectory & "\system32\coss_brain.dll") Then
        List1.AddItem "Computer Organization and System Software"
 End If
 
'Check the user created papers
 Dir1.Path = App.Path & "\createpaper"
 Dir1.Refresh
 
 For inttemp = 0 To Dir1.ListCount - 1
     StrTemp = StrReverse(Mid(StrReverse(Dir1.List(inttemp)), 1, InStr(1, StrReverse(Dir1.List(inttemp)), "\") - 1))
     List1.AddItem StrTemp
 Next inttemp
End Sub

Private Sub List1_Click()
Dim inttemp As Integer
Dim StrTemp As String

Label2(2).Caption = ""
'Checks for the paper name selected and take the appropriate action
 If List1.Text = "Information Technology" Or List1.Text = "Internet and Web Designing" Or _
                 List1.Text = "Personal Computing Technology" Or List1.Text = "Internet and Web Designing" Or _
                 List1.Text = "C Language" Or List1.Text = "Business Systems and FoxPro" Or _
                 List1.Text = "Database Management Systems" Or List1.Text = "Systen Analysis and Designing" Or _
                 List1.Text = "Data Communications and Networking" Or List1.Text = "Comupter Graphics" Or _
                 List1.Text = "Data Structure" Or List1.Text = "Unix" Or List1.Text = "C++ Language" Or _
                 List1.Text = "Computer Organization and System Software" Then
    Combo1.Clear
    Combo1.AddItem "Fill in the Blanks"
    Combo1.AddItem "Matching Columns"
    Combo1.AddItem "Multiple Choice"
    Combo1.AddItem "Descriptive Questions"
    Combo1.AddItem "True or False"
    Combo1.Text = "Select the Type to view details"
    
 Else
    Dir1.Path = App.Path & "\createpaper\" & List1.Text
    Dir1.Refresh
    Combo1.Clear
    For inttemp = 0 To Dir1.ListCount - 1
        StrTemp = StrReverse(Mid(StrReverse(Dir1.List(inttemp)), 1, InStr(1, StrReverse(Dir1.List(inttemp)), "\") - 1))
        Combo1.AddItem StrTemp
    Next inttemp
    Combo1.Text = "Select the Type to view details"
 End If
 
End Sub


