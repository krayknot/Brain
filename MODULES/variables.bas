Attribute VB_Name = "variables"
Option Explicit
Enum Category
 Fill_Questions = 0
 Fill_Options = 1
 Fill_Answers = 2
 Matching_ColumnA = 3
 Matching_ColumnB = 4
 Matching_Answers = 5
 Questions = 6
 TF_Questions = 7
 TF_Answers = 8
 MC_Questions = 9
 MC_Options_1 = 10
 MC_Options_2 = 11
 MC_Options_3 = 12
 MC_Options_4 = 13
 MC_Answer = 14
 End Enum

Enum Paper
 Paper1 = 1
 Paper2 = 2
 Paper3 = 3
 Paper4 = 4
 Paper5 = 5
 Paper6 = 6
 Paper7 = 7
 Paper8 = 8
 Paper9 = 9
 Paper10 = 10
 Paper11 = 11
 Paper12 = 12
 Paper13 = 13
 Paper14 = 14
 Paper15 = 15
 Paper16 = 16
 Paper17 = 17
 Paper18 = 18
 Paper19 = 19
 Paper20 = 20
End Enum

 Global StrPaperType As String
 Global IntPaper As Integer
'Variable that will hold the name of the paper'
 Public StrPName As String
'Variable that will hold the type of the paper'
 Public StrPType As String
'Variable that will hold the number of the paper
' Public PaperNum As Integer
 Public PaperNum As Paper
  
'Variables for fill in the blanks
 Global StrQuestion(10) As String * 30000
 Global StrQuestionRead(10) As String * 30000
 Global StrOption(50) As String
 Global StrAnswer(10) As String
 Global StrRemind(10) As String
 Global strmuloption(40) As String
 Global IntMulCnt As Integer
 Global IntITFBPaperNum As Integer
 Global IntPCFBPaperNum As Integer
 Global IntIWPDFBPaperNum As Integer
 Global IntCFBPaperNum As Integer
 Global IntBSFBPaperNum As Integer
 Global IntDBMSFBPaperNum As Integer
 Global IntSADFBPaperNum As Integer
 Global IntDCNFBPaperNum As Integer
 Global IntCGFBPaperNum As Integer
 Global IntDSFBPaperNum As Integer
 Global IntUnixFBPaperNum As Integer
 Global IntCPPFBPaperNum As Integer
 Global IntCOFBPaperNum As Integer
   
'Variables for Matching Columns
 Global strcolumnA(10) As String
 Global strcolumnB(10) As String
 
'Variable that keeps the number of the current question
' Dim IntQuesNo As Integer

'VAriable for the colored tooltip color
 Dim Backcolor As Integer
 Dim Forecolor As Integer
 
'Variable that informs about the Agent Situation
 Global BlAgent As Boolean
  
'Variable for the messagebox this variable will be TRUE if the user will press OK
 Global BLMessage As Boolean
 
 Global StrFont As String ' Font Variable
 Global IntMarksTotal 'Total marks variable
 Global THEME As String 'Pointer setting
 Global LABEL_DOWN As Boolean
 Global LABEL_UP As Boolean
 
 Global TRUE_FALSE As Boolean
 Global MULTIPLE_CHOICE As Boolean
 Global MATCHING_COLUMNS As Boolean
 
 Global BRAIN_SHOW As Boolean
 Global BlBalloon As Boolean
 Global IntCode As Integer
 Global BlFirstTime As Boolean
 Global BlArticleOpen As Boolean
 
'File names of speech engine
 Global StrFileNames(14) As String
 
 Global BLSpeak As Boolean
 Global SELECTION_TOP As Long
 Global SELECTION_LEFT As Long
 
 Global PREVIEW As Boolean
 
'User Information
 Global USERNAME As String
 Global USEREMAIL As String
 

