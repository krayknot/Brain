VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmPrepare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BRAIN: Preparing Please wait..."
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "FrmPrepare.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1470
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin Brain.UserControl1 prgbar 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
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
         Color           =   16744576
         Scrolling       =   2
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2400
         Top             =   720
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lblpaper 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmPrepare.frx":0ECA
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblMsg 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmPrepare.frx":0F2B
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin ACTIVESKINLibCtl.SkinLabel Lblpapertype 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmPrepare.frx":0FC4
         TabIndex        =   3
         Top             =   720
         Width           =   3615
      End
   End
End
Attribute VB_Name = "FrmPrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As String
Private Sub Form_Load()
FrmSelection.Skin1.ApplySkin Me.hwnd
If Me.Visible = False Then Me.Visible = True
Call Prepare
End Sub

Function prepare_default()
   '****************************************************
   'Paper setting if the paper is Fill in the Blanks
   '****************************************************
    If StrPType = "Fill in the Blanks" Then
        FrmBrain.FraFillintheBlanks.Visible = True
        FrmBrain.FraMatch.Visible = False
        FrmBrain.Framultiple.Visible = False
        FrmBrain.FraTF.Visible = False
        LblMsg.Caption = "Reading Questions"
        Me.Refresh
    ElseIf StrPType = "Matching Columns" Then
            FrmBrain.FraFillintheBlanks.Visible = False
            FrmBrain.FraMatch.Visible = True
            FrmBrain.Framultiple.Visible = False
            FrmBrain.FraTF.Visible = False
            FrmBrain.CmdNext.Enabled = False
            FrmBrain.CmdBack.Enabled = False
            LblMsg.Caption = "Reading Questions"
            Me.Refresh
            MATCHING_COLUMNS = True
            FrmBrain.List2.Clear
            FrmBrain.List3.Clear
            FrmBrain.CmdNext.Enabled = False
            FrmBrain.CmdBack.Enabled = False
      ElseIf StrPType = "True Or False" Then
            FrmBrain.FraFillintheBlanks.Visible = False
            FrmBrain.FraMatch.Visible = False
            FrmBrain.Framultiple.Visible = False
            FrmBrain.FraTF.Visible = True
            LblMsg.Caption = "Reading Questions"
            Me.Refresh
            TRUE_FALSE = True
      ElseIf StrPType = "Descriptive Questions" Then
             FrmBrain.LblQuestion.Width = 8535
             FrmBrain.LblQuestion.Height = 4215
             FrmBrain.List1.Visible = False
             FrmBrain.FraFillintheBlanks.Visible = True
             FrmBrain.FraFillintheBlanks.Caption = "Descriptive Questions"
             FrmBrain.FraMatch.Visible = False
             FrmBrain.Framultiple.Visible = False
             FrmBrain.FraTF.Visible = False
             LblMsg.Caption = "Reading Questions"
             Me.Refresh
       ElseIf StrPType = "Multiple Choice" Then
              FrmBrain.FraFillintheBlanks.Visible = False
              FrmBrain.FraMatch.Visible = False
              FrmBrain.Framultiple.Visible = True
              FrmBrain.FraTF.Visible = False
              IntMulCnt = 0
              LblMsg.Caption = "Reading Questions"
              Me.Refresh
              MULTIPLE_CHOICE = True
       End If
End Function
Public Function Prepare()
On Error Resume Next
 Lblpaper.Caption = StrPName
 Lblpapertype.Caption = StrPType
 Me.Refresh
 
 If BlAgent = True Then
    FrmSelection.Agent1.Characters("nitij").Stop
    FrmSelection.Agent1.Characters("nitij").Play "processing"
 End If
 
 Dim IntCnt As Integer
 For IntCnt = 1 To 10
        StrQuestion(IntCnt) = ""
        StrOption(IntCnt) = ""
        StrAnswer(IntCnt) = ""
 Next
 
'****************************************************
'Paper setting if the paper is Information Technology
'****************************************************
 If StrPName = "Information Technology" Then
    Dim BRAINIT As New ItPapers

    prgbar.Value = prgbar.Value + 10

    Dim StroptionNumIT As ITPaperCount
    Set StroptionNumIT = New ITPaperCount

   '****************************************************
   'Paper setting if the paper is Fill in the Blanks
   '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default 'Funtion that will do default settings
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINIT.IT(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default 'Funtion that will do default settings
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
        
            For IntCnt = 1 To 10 'StroptionNumIT.IT_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

            FrmBrain.CmdNext.Enabled = False
            FrmBrain.CmdBack.Enabled = False

      ElseIf StrPType = "True Or False" Then
              prepare_default 'Funtion that will do default settings
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINIT.IT(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINIT.IT(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)
              
      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default 'Funtion that will do default settings
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINIT.IT(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default 'Funtion that will do default settings
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINIT.IT(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIT.IT(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIT.IT(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIT.IT(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIT.IT(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINIT.IT(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is Personal Technology
'****************************************************
ElseIf StrPName = "Personal Computing Technology" Then
    Dim BRAINPC As New PCPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumpc As PCPaperCount
  Set StroptionNumpc = New PCPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default 'Funtion that will do default settings
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINPC.PC(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default 'Funtion that will do default settings
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            For IntCnt = 1 To 10 'StroptionNumPC.PC_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt
      ElseIf StrPType = "True Or False" Then
              prepare_default 'Funtion that will do default settings
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINPC.PC(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINPC.PC(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)
              
      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default 'Funtion that will do default settings
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINPC.PC(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default 'Funtion that will do default settings
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINPC.PC(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINPC.PC(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINPC.PC(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINPC.PC(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINPC.PC(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINPC.PC(StrPType, PaperNum, MC_Answer, IntCnt)

                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

''****************************************************
''Paper setting if the paper is Internet & Web Design
''****************************************************
ElseIf StrPName = "Internet and Web Design" Then
    Dim BRAINIWPD As New IWPDpapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumIWPD As IWPDPaperCount
  Set StroptionNumIWPD = New IWPDPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINIWPD.IWPD(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default 'default settings
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumIWPD.IWPD_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

      ElseIf StrPType = "True Or False" Then
            prepare_default 'default settings
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default 'default settings
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default 'default settings
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINIWPD.IWPD(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is C language
'****************************************************
ElseIf StrPName = "Programming Through C Language" Then
    Dim BRAINC As CPapers
    Set BRAINC = New CPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumC As CPaperCount
  Set StroptionNumC = New CPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default 'default settings
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINC.C(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINC.C(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINC.C(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINC.C(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default 'default settings
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINC.C(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINC.C(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINC.C(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumC.C_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

      ElseIf StrPType = "True Or False" Then
            prepare_default 'default settings
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINC.C(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINC.C(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default 'default settings
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINC.C(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default 'default settings
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINC.C(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINC.C(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINC.C(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINC.C(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINC.C(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINC.C(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is Business Systems
'****************************************************
ElseIf StrPName = "Business Systems" Then
Dim BRAINBS As BSPapers
    Set BRAINBS = New BSPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumBS As BSPaperCount
  Set StroptionNumBS = New BSPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINBS.BS(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestion(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumBS.BS_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINBS.BS(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINBS.BS(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINBS.BS(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINBS.BS(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINBS.BS(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINBS.BS(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINBS.BS(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINBS.BS(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINBS.BS(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
        End If

'****************************************************
'Paper setting if the paper is Database Management Systems
'****************************************************
ElseIf StrPName = "Database Management Systems" Then
Dim BRAINDBMS As DBMSPapers
    Set BRAINDBMS = New DBMSPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumDBMS As DBMSPaperCount
  Set StroptionNumDBMS = New DBMSPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINDBMS.DBMS(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestion(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumDBMS.DBMS_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINDBMS.DBMS(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is system analysis and design and mis
'****************************************************
 ElseIf StrPName = "System Analysis and Design and MIS" Then
    Dim BRAINSAD As New SADPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumSAD As New SADPaperCount
  
    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default 'Funtion that will do default settings
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINSAD.SAD(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default 'Funtion that will do default settings
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            For IntCnt = 1 To 10 'StroptionNumsad.PC_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt
      ElseIf StrPType = "True Or False" Then
              prepare_default 'Funtion that will do default settings
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)
              
      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default 'Funtion that will do default settings
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default 'Funtion that will do default settings
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINSAD.SAD(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINSAD.SAD(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINSAD.SAD(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINSAD.SAD(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINSAD.SAD(StrPType, PaperNum, MC_Answer, IntCnt)

                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is Data Communications and Networking
'****************************************************
ElseIf StrPName = "Data Communications and Networking" Then
Dim BRAINDCN As DCNPapers
    Set BRAINDCN = New DCNPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumDCN As DCNPaperCount
  Set StroptionNumDCN = New DCNPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINDCN.DCN(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumDCN.DCN_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDCN.DCN(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDCN.DCN(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDCN.DCN(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDCN.DCN(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINDCN.DCN(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
        End If
'****************************************************
'Paper setting if the paper is Computer Graphics
'****************************************************
ElseIf StrPName = "Computer Graphics" Then
 Dim BRAINCG As CGPapers
    Set BRAINCG = New CGPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumCG As CGPaperCount
  Set StroptionNumCG = New CGPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINCG.CG(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumCG.CG_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt
            
      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINCG.CG(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINCG.CG(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINCG.CG(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINCG.CG(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCG.CG(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCG.CG(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCG.CG(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCG.CG(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINCG.CG(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is Data Structure
'****************************************************
ElseIf StrPName = "Data Structure" Then
    Dim BRAINDS As DSPapers
    Set BRAINDS = New DSPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumDS As DSPaperCount
  Set StroptionNumDS = New DSPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
         For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINDS.DS(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt

            For IntCnt = 1 To 10 'StroptionNumDS.DS_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINDS.DS(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINDS.DS(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINDS.DS(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINDS.DS(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDS.DS(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDS.DS(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDS.DS(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINDS.DS(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINDS.DS(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
              
       End If

'*********************************************************
'Paper setting if the paper is Unix and Shell Programming
'*********************************************************
 ElseIf StrPName = "UNIX and Shell Programming" Then
    Dim BRAINunix As UnixPapers
    Set BRAINunix = New UnixPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumUnix As UnixPaperCount
  Set StroptionNumUnix = New UnixPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINunix.Unix(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumUnix.Unix_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt
            
      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINunix.Unix(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINunix.Unix(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINunix.Unix(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINunix.Unix(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINunix.Unix(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If

'****************************************************
'Paper setting if the paper is C++
'****************************************************
 ElseIf StrPName = "Programming and C++" Then
    Dim BRAINCPP As CPPPapers
    Set BRAINCPP = New CPPPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumCPP As CPPPaperCount
  Set StroptionNumCPP = New CPPPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINCPP.CPP(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt
            
            For IntCnt = 1 To 10 'StroptionNumCPP.CPP_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt
            
      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCPP.CPP(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCPP.CPP(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCPP.CPP(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCPP.CPP(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINCPP.CPP(StrPType, PaperNum, MC_Answer, IntCnt)
                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If
'**********************************
'Paper setting if the paper is COSS
'**********************************
ElseIf StrPName = "C. O. S. S." Then
    Dim BRAINCO As COPapers
    Set BRAINCO = New COPapers

    prgbar.Value = prgbar.Value + 10
    LblMsg.Caption = "Setting Paper"
    Me.Refresh

  Dim StroptionNumCO As COPaperCount
  Set StroptionNumCO = New COPaperCount

    '****************************************************
    'Paper setting if the paper is Fill in the Blanks
    '****************************************************
    If StrPType = "Fill in the Blanks" Then
        prepare_default
        For IntCnt = 1 To 10
            StrQuestion(IntCnt) = Replace(BRAINCO.CO(StrPType, PaperNum, Fill_Questions, IntCnt), "dash", "_______")
            StrQuestionRead(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Fill_Questions, IntCnt)
            StrOption(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Fill_Options, IntCnt)
            StrAnswer(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Fill_Answers, IntCnt)
            prgbar.Value = prgbar.Value + 5
        Next

    ElseIf StrPType = "Matching Columns" Then
            prepare_default
            For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Matching_ColumnA, IntCnt)
                StrOption(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Matching_ColumnB, IntCnt)
                StrAnswer(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Matching_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
            Next IntCnt

            For IntCnt = 1 To 10 'StroptionNumco.co_MTC_PaperNum    'Adding the options in the list box
                FrmBrain.List2.AddItem Trim(StrQuestion(IntCnt))
                FrmBrain.List3.AddItem Trim(StrOption(IntCnt))
            Next IntCnt

            
      ElseIf StrPType = "True Or False" Then
            prepare_default
              For IntCnt = 1 To 10
                StrQuestion(IntCnt) = BRAINCO.CO(StrPType, PaperNum, TF_Questions, IntCnt)
                StrAnswer(IntCnt) = BRAINCO.CO(StrPType, PaperNum, TF_Answers, IntCnt)
                prgbar.Value = prgbar.Value + 5
              Next IntCnt
              TRUE_FALSE = True

           FrmBrain.Lbltf.Caption = StrQuestion(1)

      ElseIf StrPType = "Descriptive Questions" Then
             prepare_default
             For IntCnt = 1 To 10
               StrQuestion(IntCnt) = BRAINCO.CO(StrPType, PaperNum, Questions, IntCnt)
               prgbar.Value = prgbar.Value + 5
             Next IntCnt

       ElseIf StrPType = "Multiple Choice" Then
              prepare_default
              For IntCnt = 1 To 10
                  StrQuestion(IntCnt) = BRAINCO.CO(StrPType, PaperNum, MC_Questions, IntCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCO.CO(StrPType, PaperNum, MC_Options_1, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCO.CO(StrPType, PaperNum, MC_Options_2, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCO.CO(StrPType, PaperNum, MC_Options_3, IntMulCnt)
                  IntMulCnt = IntMulCnt + 1
                  strmuloption(IntMulCnt) = BRAINCO.CO(StrPType, PaperNum, MC_Options_4, IntMulCnt)
                  StrAnswer(IntCnt) = BRAINCO.CO(StrPType, PaperNum, MC_Answer, IntCnt)

                  prgbar.Value = prgbar.Value + 5
              Next IntCnt
       End If
     End If

LblMsg.Caption = "Reading Questions"
If StrPName = "Information Technology" Then 'If Information Technology
        Set StroptionNumIT = New ITPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumIT.IT_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "Personal Computing Technology" Then 'If Information Technology
        Set StroptionNumpc = New PCPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumpc.PC_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "Internet and Web Design" Then 'If Information Technology
        Set StroptionNumIWPD = New IWPDPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumIWPD.IWPD_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "Programming Through C Language" Then 'If Information Technology
        Set StroptionNumC = New CPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumC.C_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "Business Systems" Then 'If Information Technology
        Set StroptionNumBS = New BSPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumBS.BS_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt
        
ElseIf StrPName = "Database Management Systems" Then 'If Information Technology
        Set StroptionNumDBMS = New DBMSPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumDBMS.DBMS_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt
      
ElseIf StrPName = "System Analysis and Design and MIS" Then 'If Information Technology
        Set StroptionNumSAD = New SADPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumSAD.SAD_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt
     
ElseIf StrPName = "Data Communications and Networking" Then 'If Information Technology
        Set StroptionNumDCN = New DCNPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumDCN.DCN_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "Computer Graphics" Then 'If Information Technology
        Set StroptionNumCG = New CGPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumCG.CG_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "Data Structure" Then 'If data structure
        Set StroptionNumDS = New DSPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumDS.DS_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt
        
ElseIf StrPName = "UNIX and Shell Programming" Then 'If Information Technology
        Set StroptionNumUnix = New UnixPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumUnix.UNIX_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt
        
ElseIf StrPName = "Programming and C++" Then 'If Information Technology
        Set StroptionNumCPP = New CPPPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumCPP.CPP_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt

ElseIf StrPName = "C. O. S. S." Then 'If Information Technology
        Set StroptionNumCO = New COPaperCount
        FrmBrain.List1.Clear
        For IntCnt = 1 To StroptionNumCO.CO_FB_PaperNum   'Adding the options in the list box
            FrmBrain.List1.AddItem StrOption(IntCnt)
        Next IntCnt
End If
prgbar.Value = prgbar.Value + 10
Me.Refresh

LblMsg.Caption = "Setting Paper"
Me.Refresh
prgbar.Value = prgbar.Value + 10
FrmBrain.Lblquesno.Visible = False
FrmBrain.LblQuestion.Visible = False
FrmBrain.LblAnswer.Visible = False
FrmBrain.Lblquesno = "Question 1"

If StrPType = "Multiple Choice" Then
   FrmBrain.LblMCQues.Caption = StrQuestion(1)
   FrmBrain.Opt1.Caption = strmuloption(1)
   FrmBrain.Opt2.Caption = strmuloption(2)
   FrmBrain.Opt3.Caption = strmuloption(3)
   FrmBrain.Opt4.Caption = strmuloption(4)
Else
   FrmBrain.LblQuestion.Caption = StrQuestion(1)
End If

prgbar.Value = prgbar.Value + 10
FrmBrain.Lblquesno.Visible = True
FrmBrain.LblQuestion.Visible = True

LblMsg.Caption = "Paper is Ready"
prgbar.Value = prgbar.Value + 10
Me.Refresh

If BlAgent = True Then
   FrmSelection.Agent1.Characters("nitij").Stop
   FrmSelection.Agent1.Characters("nitij").Play "gesturedown"
   If StrPType = "Fill in the Blanks" Then
      FrmSelection.Agent1.Characters("nitij").Speak StrQuestionRead(1)
      FrmSelection.Agent1.Characters("nitij").Play "think"
   End If
   If StrPType = "Matching Columns" Then
      FrmSelection.Agent1.Characters("nitij").Speak _
      "Select the Question from Column A then select the appropriae answer from Column B" & _
      " and click on the Check Button below"
   End If
   If StrPType = "Multiple Choice" Then
      FrmSelection.Agent1.Characters("nitij").Speak Trim(StrQuestion(1))
      FrmSelection.Agent1.Characters("nitij").Speak Trim(strmuloption(1))
      FrmSelection.Agent1.Characters("nitij").Speak Trim(strmuloption(2))
      FrmSelection.Agent1.Characters("nitij").Speak Trim(strmuloption(3))
      FrmSelection.Agent1.Characters("nitij").Speak Trim(strmuloption(4))
      FrmSelection.Agent1.Characters("nitij").Play "think"
   End If
   If StrPType = "True Or False" Then
      FrmSelection.Agent1.Characters("nitij").Speak StrQuestion(1)
      FrmSelection.Agent1.Characters("nitij").Play "think"
   End If
End If

If BRAIN_SHOW = False Then
   FrmBrain.Show
   BRAIN_SHOW = True
End If
End Function

Private Sub Timer1_Timer()
If prgbar.Value = 100 Then Unload Me
End Sub
