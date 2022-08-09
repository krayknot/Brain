VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDetails 
   Caption         =   "BRAIN@: Details"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9915
      TabIndex        =   2
      Top             =   3480
      Width           =   9975
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
         Left            =   8880
         TabIndex        =   9
         Top             =   240
         Width           =   975
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
         Left            =   7920
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command1"
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
         Left            =   6960
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command1"
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
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Accessing. Please wait..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   6240
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MsTable 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6165
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim MFolder As Folder, MFile As File

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub File1_Click()
Set MFile = FSO.GetFile(File1.Path & "\" & File1.filename)

End Sub

Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
Dim IntCnt As Integer

PBar.Value = PBar.Value + 10
MsTable.Cols = 5
File1.Path = WindowsDirectory & "\system32\"
File1.Pattern = "*_BRAIN.DLL"

MsTable.ColWidth(0) = 1500
MsTable.TextMatrix(0, 0) = "Name of the paper"
MsTable.ColWidth(1) = 200 * Len("Type of paper")
MsTable.TextMatrix(0, 1) = "Type of the paper"
MsTable.ColWidth(2) = 120 * Len("Type of paper")
MsTable.ColWidth(3) = 100 * Len("Type of paper")
MsTable.Rows = File1.ListCount + 1

PBar.Value = PBar.Value + 10

''Filling the table with the file names
' For IntCnt = 0 To File1.ListCount - 1
'     MsTable.TextMatrix(IntCnt + 1, 1) = File1.List(IntCnt)
' Next IntCnt
'

PBar.Value = PBar.Value + 10
'Filling the table with the file description
 Dim StrTemp As String
 Dim IntCount As Integer
 MsTable.Cols = File1.ListCount + 1
 For IntCnt = 0 To File1.ListCount - 1
     StrTemp = File1.List(IntCnt)
      
     Select Case StrTemp
            Case "IT_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Information Technoloy")
            MsTable.TextMatrix(0, IntCnt + 1) = "Information Technology"
            Case "BS_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Business Systems and Programming through FoxPro")
            MsTable.TextMatrix(0, IntCnt + 1) = "Business Systems and Programming through FoxPro"
            Case "C_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Programming and Problem Solving through C language")
            MsTable.TextMatrix(0, IntCnt + 1) = "Programming and Problem Solving through C language"
            Case "CG_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Computer Graphics")
            MsTable.TextMatrix(0, IntCnt + 1) = "Computer Graphics"
            Case "COSS_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Computer Organization and System Software")
            MsTable.TextMatrix(0, IntCnt + 1) = "Computer Organization and System Software"
            Case "CPP_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Object Oriented programming through C++")
            MsTable.TextMatrix(0, IntCnt + 1) = "Object Oriented programming through C++"
            Case "DBMS_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Database Management System")
            MsTable.TextMatrix(0, IntCnt + 1) = "Database Management System"
            Case "DCN_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Data Communications and Networking")
            MsTable.TextMatrix(0, IntCnt + 1) = "Data Communications and Networking"
            Case "DS_BRAIN.dll"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Data Structure")
            MsTable.TextMatrix(0, IntCnt + 1) = "Data Structure"
            Case "IWPD_BRAIN.dll"
            MsTable.TextMatrix(0, IntCnt + 1) = "Business Systems and Programming through FoxPro"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Business Systems and Programming through FoxPro")
            Case "PC_BRAIN.dll"
            MsTable.TextMatrix(0, IntCnt + 1) = "Business Systems and Programming through FoxPro"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Business Systems and Programming through FoxPro")
            Case "UNIX_BRAIN.dll"
            MsTable.TextMatrix(0, IntCnt + 1) = "Business Systems and Programming through FoxPro"
            MsTable.ColWidth(IntCnt + 1) = 90 * Len("Business Systems and Programming through FoxPro")
     End Select
 Next IntCnt
 
 PBar.Value = PBar.Value + 10
'Filling the table with the file Categories
 For IntCnt = 0 To 4
     MsTable.TextMatrix(IntCnt + 1, 0) = "Component" & " " & IntCnt
 Next IntCnt
 
 PBar.Value = PBar.Value + 10
 Dim StrComponent(6) As String
 StrComponent(1) = "Fill in the Blanks"
 StrComponent(2) = "Matching Colimns"
 StrComponent(3) = "Multiple Choice"
 StrComponent(4) = "Descriptive Questions"
 StrComponent(5) = "True or False"
 StrComponent(6) = "Others"

 PBar.Value = PBar.Value + 10
 For IntCount = 1 To 5
     For IntCnt = 0 To File1.ListCount - 1
         MsTable.TextMatrix(IntCount, IntCnt) = StrComponent(IntCount)
     Next IntCnt
 Next IntCount

 PBar.Value = PBar.Value + 10
'Filling the table with the file date created
 MsTable.TextMatrix(6, 0) = "Date Created"
 For IntCnt = 0 To File1.ListCount - 1
     Set MFile = FSO.GetFile(File1.Path & "/" & File1.List(IntCnt))
     MsTable.WordWrap = True
     MsTable.TextMatrix(6, IntCnt + 1) = Str(MFile.DateCreated)
 Next IntCnt
  
 PBar.Value = PBar.Value + 10
'Filling the table with the file date created
 MsTable.TextMatrix(7, 0) = "Size of File"
 For IntCnt = 0 To File1.ListCount - 1
     Set MFile = FSO.GetFile(File1.Path & "/" & File1.List(IntCnt))
     MsTable.TextMatrix(7, IntCnt + 1) = MFile.Size / 1024 & " KB"
 Next IntCnt

 PBar.Value = PBar.Value + 10
 PBar.Value = PBar.Value + 10
 Label1.Visible = False
 PBar.Visible = False
End Sub

